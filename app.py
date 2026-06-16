from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os, json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

def extrair_texto_doc(doc):
    partes = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            partes.append(t)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for para in celula.paragraphs:
                    t = para.text.strip()
                    if t:
                        partes.append(t)
    return ' '.join(partes)[:1500]

def inserir_campo_mala_direta(paragrafo, nome_campo, texto_original):
    for run in paragrafo.runs:
        if texto_original.lower() in run.text.lower():
            rPr = run._r.find(qn('w:rPr'))
            fld = OxmlElement('w:fldSimple')
            fld.set(qn('w:instr'), f' MERGEFIELD {nome_campo} \\* MERGEFORMAT ')
            r_novo = OxmlElement('w:r')
            if rPr is not None:
                r_novo.append(copy.deepcopy(rPr))
            t_novo = OxmlElement('w:t')
            t_novo.text = f'«{nome_campo}»'
            r_novo.append(t_novo)
            fld.append(r_novo)
            run._r.getparent().replace(run._r, fld)
            return True
    return False

def inserir_campos_em_doc(doc, campos):
    for campo in campos:
        if campo.get('campo_sugerido') == 'IGNORAR':
            continue
        texto = campo['texto_original']
        nome  = campo['campo_sugerido']
        for para in doc.paragraphs:
            inserir_campo_mala_direta(para, nome, texto)
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for para in celula.paragraphs:
                        inserir_campo_mala_direta(para, nome, texto)

def unir_documentos(docs):
    doc_base = Document()
    for el in list(doc_base.element.body):
        doc_base.element.body.remove(el)

    for idx, doc in enumerate(docs):
        body = doc.element.body
        sectPr_original = body.find(qn('w:sectPr'))

        for el in body:
            if el.tag == qn('w:sectPr'):
                continue
            doc_base.element.body.append(copy.deepcopy(el))

        if idx < len(docs) - 1:
            p_break = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            if sectPr_original is not None:
                pPr.append(copy.deepcopy(sectPr_original))
            p_break.append(pPr)
            doc_base.element.body.append(p_break)
        else:
            if sectPr_original is not None:
                doc_base.element.body.append(copy.deepcopy(sectPr_original))

    return doc_base

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analisar', methods=['POST'])
def analisar():
    if 'arquivos' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400
    arquivos = request.files.getlist('arquivos')
    textos = []
    for f in arquivos:
        try:
            doc = Document(f.stream)
            txt = extrair_texto_doc(doc)
            textos.append({'nome': f.filename, 'texto': txt})
        except Exception as e:
            textos.append({'nome': f.filename, 'texto': '', 'erro': str(e)})
    return jsonify({'textos': textos})

@app.route('/gerar', methods=['POST'])
def gerar():
    if 'arquivos' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    arquivos_raw  = request.files.getlist('arquivos')
    campos_json   = request.form.get('campos', '[]')
    nome_funcao   = request.form.get('nome_funcao', 'MODELO')

    try:
        campos = json.loads(campos_json)
    except:
        campos = []

    docs = []
    for f in arquivos_raw:
        try:
            doc = Document(f.stream)
            inserir_campos_em_doc(doc, campos)
            docs.append(doc)
        except Exception as e:
            return jsonify({'erro': f'Erro ao processar {f.filename}: {str(e)}'}), 500

    if not docs:
        return jsonify({'erro': 'Nenhum documento válido'}), 400

    try:
        doc_final = unir_documentos(docs)
    except Exception as e:
        return jsonify({'erro': f'Erro ao unir documentos: {str(e)}'}), 500

    buf = io.BytesIO()
    doc_final.save(buf)
    buf.seek(0)

    nome_arquivo = f"CERTIFICADOS_{nome_funcao.upper().replace(' ', '_')}.docx"
    return send_file(buf, as_attachment=True, download_name=nome_arquivo,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
