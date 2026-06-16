from flask import Flask, request, jsonify, send_file, render_template
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import copy, io, os, json, tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def extrair_texto_doc(doc):
    """Extrai texto limpo de um Document para análise da IA."""
    partes = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            partes.append(t)
    return ' '.join(partes)[:1500]


def inserir_campo_mala_direta(paragrafo, nome_campo, texto_original):
    """
    Substitui texto em um parágrafo por campo de mala direta Word.
    Preserva formatação do run original.
    """
    for run in paragrafo.runs:
        if texto_original.lower() in run.text.lower():
            # Guardar formatação
            rPr = run._r.find(qn('w:rPr'))

            # Criar elemento fldSimple (campo de mala direta)
            fld = OxmlElement('w:fldSimple')
            fld.set(qn('w:instr'), f' MERGEFIELD {nome_campo} \\* MERGEFORMAT ')

            # Criar run com o placeholder
            r_novo = OxmlElement('w:r')
            if rPr is not None:
                r_novo.append(copy.deepcopy(rPr))
            t_novo = OxmlElement('w:t')
            t_novo.text = f'«{nome_campo}»'
            r_novo.append(t_novo)
            fld.append(r_novo)

            # Substituir run pelo campo
            run._r.getparent().replace(run._r, fld)
            return True
    return False


def inserir_campos_em_doc(doc, campos):
    """Percorre todos os parágrafos e tabelas inserindo os campos."""
    for campo in campos:
        if campo.get('campo_sugerido') == 'IGNORAR':
            continue
        texto = campo['texto_original']
        nome  = campo['campo_sugerido']

        # Parágrafos normais
        for para in doc.paragraphs:
            inserir_campo_mala_direta(para, nome, texto)

        # Parágrafos dentro de tabelas
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for para in celula.paragraphs:
                        inserir_campo_mala_direta(para, nome, texto)


def unir_documentos(docs):
    """
    Une múltiplos Documents em um único, preservando
    orientação (retrato/paisagem) de cada um.
    """
    doc_base = Document()

    # Remover parágrafo vazio padrão
    for el in list(doc_base.element.body):
        doc_base.element.body.remove(el)

    for idx, doc in enumerate(docs):
        body = doc.element.body

        # Copiar cada elemento do body
        for el in body:
            # Ignorar o sectPr raiz — vamos tratar separado
            if el.tag == qn('w:sectPr'):
                continue
            doc_base.element.body.append(copy.deepcopy(el))

        # Pegar o sectPr deste documento (define orientação/tamanho)
        sectPr_original = body.find(qn('w:sectPr'))

        if idx < len(docs) - 1:
            # Inserir quebra de seção contínua com o sectPr original
            # para preservar orientação antes de iniciar próximo doc
            p_break = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')

            if sectPr_original is not None:
                sect_copy = copy.deepcopy(sectPr_original)
                # Garantir que é quebra de página (nextPage)
                pgSz = sect_copy.find(qn('w:pgSz'))
                # Manter o pgSz original (preserva retrato/paisagem)
                pPr.append(sect_copy)

            p_break.append(pPr)
            doc_base.element.body.append(p_break)
        else:
            # Último documento: adicionar sectPr como raiz do body
            if sectPr_original is not None:
                doc_base.element.body.append(copy.deepcopy(sectPr_original))

    return doc_base


# ─────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analisar', methods=['POST'])
def analisar():
    """
    Recebe os arquivos, extrai texto e retorna para o frontend
    chamar a IA (Claude) e identificar os campos.
    """
    if 'arquivos' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    arquivos = request.files.getlist('arquivos')
    textos   = []

    for f in arquivos:
        try:
            doc  = Document(f.stream)
            txt  = extrair_texto_doc(doc)
            textos.append({'nome': f.filename, 'texto': txt})
        except Exception as e:
            textos.append({'nome': f.filename, 'texto': '', 'erro': str(e)})

    return jsonify({'textos': textos})


@app.route('/gerar', methods=['POST'])
def gerar():
    """
    Recebe os arquivos na ordem correta + campos identificados.
    Une os documentos, insere mala direta e devolve o .docx.
    """
    if 'arquivos' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    arquivos_raw = request.files.getlist('arquivos')
    campos_json  = request.form.get('campos', '[]')
    nome_funcao  = request.form.get('nome_funcao', 'MODELO')

    try:
        campos = json.loads(campos_json)
    except:
        campos = []

    docs = []
    for f in arquivos_raw:
        try:
            doc = Document(f.stream)
            # Inserir campos de mala direta neste documento
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

    # Salvar em memória e enviar
    buf = io.BytesIO()
    doc_final.save(buf)
    buf.seek(0)

    nome_arquivo = f"CERTIFICADOS_{nome_funcao.upper().replace(' ', '_')}.docx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
