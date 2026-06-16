from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os, zipfile, shutil, tempfile, re

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

def unir_documentos_zip(arquivos_bytes):
    tmp = tempfile.mkdtemp()
    try:
        # Extrair todos os docx
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

        with open(doc_xml_base, 'r', encoding='utf-8') as f:
            xml_base = f.read()
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_base = f.read()
        with open(ct_path, 'r', encoding='utf-8') as f:
            ct_base = f.read()

        def extrair_body(xml):
            m = re.search(r'<w:body>(.*)</w:body>', xml, re.DOTALL)
            if not m:
                return '', '<w:sectPr/>'
            body = m.group(1)
            sp = re.search(r'(<w:sectPr[\s\S]*?</w:sectPr>)\s*$', body)
            if sp:
                return body[:sp.start()].strip(), sp.group(1)
            return body.strip(), '<w:sectPr/>'

        body_base, sectPr_base = extrair_body(xml_base)
        bodies  = [body_base]
        sectPrs = [sectPr_base]
        id_counter = [1000]

        # Processar documentos adicionais
        for i, pasta in enumerate(pastas[1:], start=1):
            doc_xml_path  = os.path.join(pasta, 'word', 'document.xml')
            rels_doc_path = os.path.join(pasta, 'word', '_rels', 'document.xml.rels')

            with open(doc_xml_path, 'r', encoding='utf-8') as f:
                xml_doc = f.read()

            body_doc, sectPr_doc = extrair_body(xml_doc)
            sectPrs.append(sectPr_doc)

            rels_map = {}
            if os.path.exists(rels_doc_path):
                with open(rels_doc_path, 'r', encoding='utf-8') as f:
                    rels_doc = f.read()

                rels_matches = re.findall(
                    r'<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"',
                    rels_doc
                )

                for rid, rtype, target in rels_matches:
                    if 'image' in rtype.lower() or 'media' in target.lower():
                        src = os.path.join(pasta, 'word', target)
                        if os.path.exists(src):
                            ext      = os.path.splitext(target)[1].lower()
                            novo_nome = f'image_doc{i}_{id_counter[0]}{ext}'
                            dst = os.path.join(pasta_saida, 'word', 'media', novo_nome)
                            os.makedirs(os.path.dirname(dst), exist_ok=True)
                            shutil.copy2(src, dst)

                            novo_rid = f'rId_ext_{id_counter[0]}'
                            id_counter[0] += 1

                            novo_rel = f'<Relationship Id="{novo_rid}" Type="{rtype}" Target="media/{novo_nome}"/>'
                            rels_base = rels_base.replace('</Relationships>', novo_rel + '</Relationships>')

                            # Garantir tipo no Content_Types
                            mime_map = {
                                '.png':  'image/png',
                                '.jpg':  'image/jpeg',
                                '.jpeg': 'image/jpeg',
                                '.gif':  'image/gif',
                                '.bmp':  'image/bmp',
                                '.wmf':  'image/x-wmf',
                                '.emf':  'image/x-emf',
                                '.tiff': 'image/tiff',
                                '.svg':  'image/svg+xml',
                            }
                            mime = mime_map.get(ext, 'application/octet-stream')
                            ext_clean = ext.lstrip('.')
                            ct_decl = f'<Default Extension="{ext_clean}" ContentType="{mime}"/>'
                            if ext_clean not in ct_base:
                                ct_base = ct_base.replace('</Types>', ct_decl + '</Types>')

                            rels_map[rid] = novo_rid

            for old_rid, new_rid in rels_map.items():
                body_doc = body_doc.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')
                body_doc = body_doc.replace(f'r:id="{old_rid}"',    f'r:id="{new_rid}"')

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
            xml_base,
            flags=re.DOTALL
        )

        with open(doc_xml_base, 'w', encoding='utf-8') as f:
            f.write(xml_final)
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_base)
        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct_base)

        # Empacotar
        buf_saida = io.BytesIO()
        with zipfile.ZipFile(buf_saida, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(pasta_saida):
                for file in files:
                    filepath = os.path.join(root, file)
                    arcname  = os.path.relpath(filepath, pasta_saida)
                    zout.write(filepath, arcname)

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
        arquivos_bytes = [f.read() for f in arquivos]
        buf = unir_documentos_zip(arquivos_bytes)
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

    return send_file(buf, as_attachment=True,
                     download_name='CERTIFICADOS_UNIFICADOS.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
