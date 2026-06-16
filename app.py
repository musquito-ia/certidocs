from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os, json, zipfile, shutil, tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

def unir_documentos_zip(arquivos_bytes, nome_funcao):
    """
    Une múltiplos docx manipulando diretamente os ZIPs.
    Copia todas as mídias (imagens) de cada arquivo para o resultado final.
    """
    import uuid

    # Criar pasta temporária
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

        # Pasta base = primeiro documento
        pasta_base = pastas[0]
        pasta_saida = os.path.join(tmp, 'saida')
        shutil.copytree(pasta_base, pasta_saida)

        # Ler o document.xml base
        doc_xml_base = os.path.join(pasta_saida, 'word', 'document.xml')
        with open(doc_xml_base, 'r', encoding='utf-8') as f:
            xml_base = f.read()

        # Ler o _rels/document.xml.rels base
        rels_path = os.path.join(pasta_saida, 'word', '_rels', 'document.xml.rels')
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_base = f.read()

        # Extrair body do base (sem sectPr final)
        import re
        def extrair_body(xml):
            m = re.search(r'<w:body>(.*)</w:body>', xml, re.DOTALL)
            if not m:
                return '', ''
            body = m.group(1)
            # Separar sectPr final
            sp = re.search(r'(<w:sectPr[\s\S]*?</w:sectPr>)\s*$', body)
            if sp:
                sectPr = sp.group(1)
                body = body[:sp.start()].strip()
            else:
                sectPr = '<w:sectPr/>'
            return body, sectPr

        body_base, sectPr_base = extrair_body(xml_base)
        bodies = [body_base]
        sectPrs = [sectPr_base]

        # Contador para IDs únicos de imagens
        id_counter = [1000]

        # Processar documentos adicionais
        for i, pasta in enumerate(pastas[1:], start=1):
            doc_xml_path = os.path.join(pasta, 'word', 'document.xml')
            rels_doc_path = os.path.join(pasta, 'word', '_rels', 'document.xml.rels')

            with open(doc_xml_path, 'r', encoding='utf-8') as f:
                xml_doc = f.read()

            body_doc, sectPr_doc = extrair_body(xml_doc)
            sectPrs.append(sectPr_doc)

            # Ler rels deste documento
            rels_map = {}  # rId original -> novo rId
            if os.path.exists(rels_doc_path):
                with open(rels_doc_path, 'r', encoding='utf-8') as f:
                    rels_doc = f.read()

                # Encontrar todas as relações
                rels_matches = re.findall(
                    r'<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"',
                    rels_doc
                )

                for rid, rtype, target in rels_matches:
                    if 'image' in rtype.lower() or 'media' in target.lower():
                        # Copiar imagem para pasta saída com novo nome
                        src = os.path.join(pasta, 'word', target)
                        if os.path.exists(src):
                            ext = os.path.splitext(target)[1]
                            novo_nome = f'image_doc{i}_{id_counter[0]}{ext}'
                            dst = os.path.join(pasta_saida, 'word', 'media', novo_nome)
                            os.makedirs(os.path.dirname(dst), exist_ok=True)
                            shutil.copy2(src, dst)

                            novo_rid = f'rId_doc{i}_{id_counter[0]}'
                            id_counter[0] += 1

                            # Adicionar ao rels base
                            novo_rel = f'<Relationship Id="{novo_rid}" Type="{rtype}" Target="media/{novo_nome}"/>'
                            rels_base = rels_base.replace('</Relationships>', novo_rel + '</Relationships>')

                            rels_map[rid] = novo_rid

            # Substituir rIds no body deste documento
            for old_rid, new_rid in rels_map.items():
                body_doc = body_doc.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')
                body_doc = body_doc.replace(f'r:id="{old_rid}"', f'r:id="{new_rid}"')
                body_doc = body_doc.replace(f'"{old_rid}"', f'"{new_rid}"')

            bodies.append(body_doc)

        # Montar body final — cada doc separado por quebra de seção com sua orientação
        body_final = ''
        for i, (body, sectPr) in enumerate(zip(bodies, sectPrs)):
            body_final += body
            if i < len(bodies) - 1:
                # Quebra de seção com a orientação deste documento
                body_final += f'\n<w:p><w:pPr>{sectPr}</w:pPr></w:p>\n'

        # Último sectPr como raiz
        body_final += '\n' + sectPrs[-1]

        # Montar XML final
        xml_final = re.sub(
            r'<w:body>.*</w:body>',
            f'<w:body>{body_final}</w:body>',
            xml_base,
            flags=re.DOTALL
        )

        # Salvar document.xml
        with open(doc_xml_base, 'w', encoding='utf-8') as f:
            f.write(xml_final)

        # Salvar rels atualizado
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_base)

        # Empacotar saída como docx
        buf_saida = io.BytesIO()
        with zipfile.ZipFile(buf_saida, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(pasta_saida):
                for file in files:
                    filepath = os.path.join(root, file)
                    arcname = os.path.relpath(filepath, pasta_saida)
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

    arquivos    = request.files.getlist('arquivos')
    nome_funcao = request.form.get('nome_funcao', 'MODELO')

    if not arquivos:
        return jsonify({'erro': 'Nenhum arquivo válido'}), 400

    try:
        # Ler todos os bytes antes de processar
        arquivos_bytes = [f.read() for f in arquivos]
        buf = unir_documentos_zip(arquivos_bytes, nome_funcao)
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

    nome = f"CERTIFICADOS_{nome_funcao.upper().replace(' ', '_')}.docx"
    return send_file(buf, as_attachment=True, download_name=nome,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
