from flask import Flask, request, jsonify, send_file, render_template
import io, os, zipfile, shutil, tempfile, re

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# Arquivos que NÃO devem ser copiados dos docs adicionais (vêm do base)
ARQUIVOS_BASE = {'word/styles.xml', 'word/settings.xml', 'word/webSettings.xml',
                 'word/fontTable.xml', 'word/theme/theme1.xml',
                 'docProps/app.xml', 'docProps/core.xml'}

def unir_documentos_zip(arquivos_bytes):
    tmp = tempfile.mkdtemp()
    try:
        pastas = []
        for i, buf in enumerate(arquivos_bytes):
            pasta = os.path.join(tmp, f'doc{i}')
            os.makedirs(pasta)
            with zipfile.ZipFile(io.BytesIO(buf), 'r') as z:
                z.extractall(pasta)
            pastas.append(pasta)

        # Base = primeiro documento
        pasta_saida = os.path.join(tmp, 'saida')
        shutil.copytree(pastas[0], pasta_saida)

        doc_xml_base = os.path.join(pasta_saida, 'word', 'document.xml')
        rels_path    = os.path.join(pasta_saida, 'word', '_rels', 'document.xml.rels')
        ct_path      = os.path.join(pasta_saida, '[Content_Types].xml')

        with open(doc_xml_base, 'r', encoding='utf-8') as f: xml_base = f.read()
        with open(rels_path,    'r', encoding='utf-8') as f: rels_base = f.read()
        with open(ct_path,      'r', encoding='utf-8') as f: ct_base = f.read()

        def extrair_body(xml):
            m = re.search(r'<w:body>(.*)</w:body>', xml, re.DOTALL)
            if not m: return '', '<w:sectPr/>'
            body = m.group(1)
            sp = re.search(r'(<w:sectPr[\s\S]*?</w:sectPr>)\s*$', body)
            if sp: return body[:sp.start()].strip(), sp.group(1)
            return body.strip(), '<w:sectPr/>'

        body_base, sectPr_base = extrair_body(xml_base)
        bodies  = [body_base]
        sectPrs = [sectPr_base]
        id_counter = [2000]

        for i, pasta in enumerate(pastas[1:], start=1):
            doc_xml_path  = os.path.join(pasta, 'word', 'document.xml')
            rels_doc_path = os.path.join(pasta, 'word', '_rels', 'document.xml.rels')

            with open(doc_xml_path, 'r', encoding='utf-8') as f: xml_doc = f.read()
            body_doc, sectPr_doc = extrair_body(xml_doc)
            sectPrs.append(sectPr_doc)

            if not os.path.exists(rels_doc_path):
                bodies.append(body_doc)
                continue

            with open(rels_doc_path, 'r', encoding='utf-8') as f: rels_doc = f.read()

            rels_matches = re.findall(
                r'<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"',
                rels_doc
            )

            rels_map = {}  # rId antigo -> novo rId

            for rid, rtype, target in rels_matches:
                # Ignorar arquivos base (styles, settings etc)
                target_full = f'word/{target}' if not target.startswith('/') else target.lstrip('/')
                if target_full in ARQUIVOS_BASE:
                    # Reutilizar o rId do base para este tipo
                    # Encontrar rId equivalente no base
                    tipo_curto = rtype.split('/')[-1]
                    m = re.search(rf'Id="([^"]+)"[^>]*Type="[^"]*/{tipo_curto}"', rels_base)
                    if m:
                        rels_map[rid] = m.group(1)
                    continue

                src = os.path.join(pasta, 'word', target)
                if not os.path.exists(src):
                    continue

                # Renomear arquivo para evitar colisão
                ext      = os.path.splitext(target)[1]
                base_nome = os.path.splitext(os.path.basename(target))[0]
                novo_nome = f'{base_nome}_d{i}_{id_counter[0]}{ext}'

                # Destino na pasta saida
                if '/' in target:
                    subdir = os.path.dirname(target)
                    dst = os.path.join(pasta_saida, 'word', subdir, novo_nome)
                else:
                    dst = os.path.join(pasta_saida, 'word', novo_nome)

                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(src, dst)

                novo_rid    = f'rId_d{i}_{id_counter[0]}'
                novo_target = f'{os.path.dirname(target)}/{novo_nome}'.lstrip('/')
                id_counter[0] += 1

                # Adicionar ao rels base
                rels_base = rels_base.replace(
                    '</Relationships>',
                    f'<Relationship Id="{novo_rid}" Type="{rtype}" Target="{novo_target}"/></Relationships>'
                )

                # Adicionar ao Content_Types se necessário
                ext_clean = ext.lstrip('.')
                mime_map  = {
                    'png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg',
                    'gif':'image/gif','bmp':'image/bmp','wmf':'image/x-wmf',
                    'emf':'image/x-emf','tiff':'image/tiff','svg':'image/svg+xml',
                    'xml':'application/xml',
                }
                if ext_clean and f'Extension="{ext_clean}"' not in ct_base:
                    mime = mime_map.get(ext_clean, 'application/octet-stream')
                    ct_base = ct_base.replace('</Types>', f'<Default Extension="{ext_clean}" ContentType="{mime}"/></Types>')

                # Adicionar Override no Content_Types para xml específico
                if ext_clean == 'xml' and novo_target:
                    part_name = f'/word/{novo_target}'
                    # Descobrir ContentType do original
                    orig_part = f'/word/{target}'
                    m_ct = re.search(rf'PartName="{re.escape(orig_part)}"\s+ContentType="([^"]+)"', ct_base)
                    if m_ct:
                        ct_novo = f'<Override PartName="{part_name}" ContentType="{m_ct.group(1)}"/>'
                        ct_base = ct_base.replace('</Types>', ct_novo + '</Types>')

                rels_map[rid] = novo_rid

            # Substituir rIds no body
            for old, new in rels_map.items():
                body_doc = body_doc.replace(f'r:embed="{old}"', f'r:embed="{new}"')
                body_doc = body_doc.replace(f'r:id="{old}"',    f'r:id="{new}"')
                body_doc = re.sub(rf'\br:link="{re.escape(old)}"', f'r:link="{new}"', body_doc)

            bodies.append(body_doc)

        # Montar body final
        body_final = ''
        for i, (body, sectPr) in enumerate(zip(bodies, sectPrs)):
            body_final += body
            if i < len(bodies) - 1:
                body_final += f'\n<w:p><w:pPr>{sectPr}</w:pPr></w:p>\n'
        body_final += '\n' + sectPrs[-1]

        xml_final = re.sub(
            r'<w:body>.*</w:body>',
            f'<w:body>{body_final}</w:body>',
            xml_base, flags=re.DOTALL
        )

        with open(doc_xml_base, 'w', encoding='utf-8') as f: f.write(xml_final)
        with open(rels_path,    'w', encoding='utf-8') as f: f.write(rels_base)
        with open(ct_path,      'w', encoding='utf-8') as f: f.write(ct_base)

        buf_saida = io.BytesIO()
        with zipfile.ZipFile(buf_saida, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(pasta_saida):
                for file in files:
                    filepath = os.path.join(root, file)
                    zout.write(filepath, os.path.relpath(filepath, pasta_saida))
        buf_saida.seek(0)
        return buf_saida

    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/unir', methods=['POST'])
def unir():
    if 'arquivos' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400
    arquivos = request.files.getlist('arquivos')
    if not arquivos:
        return jsonify({'erro': 'Nenhum arquivo válido'}), 400
    try:
        buf = unir_documentos_zip([f.read() for f in arquivos])
    except Exception as e:
        return jsonify({'erro': str(e)}), 500
    return send_file(buf, as_attachment=True,
                     download_name='CERTIFICADOS_UNIFICADOS.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
