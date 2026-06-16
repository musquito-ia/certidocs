from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

def unir_documentos(docs):
    """Une múltiplos docs preservando formatação e orientação de cada um."""
    doc_base = Document()

    # Limpar body padrão
    for el in list(doc_base.element.body):
        doc_base.element.body.remove(el)

    for idx, doc in enumerate(docs):
        body = doc.element.body
        sectPr_original = body.find(qn('w:sectPr'))

        # Copiar todos os elementos exceto o sectPr raiz
        for el in body:
            if el.tag == qn('w:sectPr'):
                continue
            doc_base.element.body.append(copy.deepcopy(el))

        if idx < len(docs) - 1:
            # Inserir quebra de seção com orientação original antes do próximo doc
            p_break = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            if sectPr_original is not None:
                pPr.append(copy.deepcopy(sectPr_original))
            p_break.append(pPr)
            doc_base.element.body.append(p_break)
        else:
            # Último: adicionar sectPr como raiz
            if sectPr_original is not None:
                doc_base.element.body.append(copy.deepcopy(sectPr_original))

    return doc_base

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/unir', methods=['POST'])
def unir():
    if 'arquivos' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    arquivos     = request.files.getlist('arquivos')
    nome_funcao  = request.form.get('nome_funcao', 'MODELO')

    docs = []
    for f in arquivos:
        try:
            docs.append(Document(f.stream))
        except Exception as e:
            return jsonify({'erro': f'Erro ao abrir {f.filename}: {str(e)}'}), 500

    if not docs:
        return jsonify({'erro': 'Nenhum documento válido'}), 400

    try:
        doc_final = unir_documentos(docs)
    except Exception as e:
        return jsonify({'erro': f'Erro ao unir: {str(e)}'}), 500

    buf = io.BytesIO()
    doc_final.save(buf)
    buf.seek(0)

    nome = f"CERTIFICADOS_{nome_funcao.upper().replace(' ', '_')}.docx"
    return send_file(buf, as_attachment=True, download_name=nome,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
