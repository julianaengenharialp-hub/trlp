import io
import json
import re
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

import streamlit as st
from docx import Document
from docx.shared import Cm, Pt

st.set_page_config(page_title='TR | Secretaria de Obras', page_icon='📄', layout='wide')
st.title('Termos de Referência | Secretaria de Obras')
st.caption('Versão 0.1 • Inventário do modelo e protótipo de preenchimento. Não gera TR validado para licitação.')

CAPITULOS = [
    ('1', 'OBJETO'),
    ('2', 'JUSTIFICATIVA E OBJETIVO DA CONTRATAÇÃO'),
    ('3', 'DESCRIÇÃO DA SOLUÇÃO COMO UM TODO'),
    ('4', 'REQUISITOS DA CONTRATAÇÃO'),
    ('5', 'FORMA DE EXECUÇÃO DO OBJETO'),
    ('6', 'GESTÃO DO CONTRATO'),
    ('7', 'MEDIÇÃO E PAGAMENTO'),
    ('8', 'SELEÇÃO DO FORNECEDOR'),
    ('9', 'ESTIMATIVA DO PREÇO'),
    ('10', 'ADEQUAÇÃO ORÇAMENTÁRIA'),
    ('11', 'DISPOSIÇÕES FINAIS'),
]

@st.cache_data
def ler_odt(bytes_arquivo):
    with zipfile.ZipFile(io.BytesIO(bytes_arquivo)) as arq:
        xml = ET.fromstring(arq.read('content.xml'))
    ns = {'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'}
    corpo = xml.find('.//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body')
    linhas = []
    for el in corpo.iter():
        if el.tag in ('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p',
                      '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}h'):
            texto = ''.join(el.itertext()).strip()
            if texto:
                linhas.append(texto)
    return linhas

st.sidebar.header('Modelo de referência')
modelo = st.sidebar.file_uploader('Selecione o modelo oficial (.odt)', type=['odt'])
st.sidebar.info('Use o arquivo original da Prefeitura. Exemplos de contratações anteriores não são textos obrigatórios.')

if 'dados' not in st.session_state:
    st.session_state.dados = {}
if 'classificacoes' not in st.session_state:
    st.session_state.classificacoes = {}

aba1, aba2, aba3 = st.tabs(['1. Inventário do modelo', '2. Preenchimento experimental', '3. Conferência e exportação'])

with aba1:
    st.subheader('Classificar antes de automatizar')
    st.write('Cada parágrafo do modelo deve ser classificado. Nenhuma cláusula será considerada obrigatória apenas por constar no exemplo.')
    if modelo:
        linhas = ler_odt(modelo.getvalue())
        st.success(f'{len(linhas)} blocos de texto identificados no ODT. A classificação é manual nesta primeira versão.')
        busca = st.text_input('Filtrar texto do modelo', placeholder='Ex.: 5.1.1, subcontratação, pagamento')
        opcoes = ['Não analisado', 'Texto padronizado', 'Campo variável', 'Instrução de preenchimento', 'Exemplo específico (ignorar)', 'Alternativa condicional', 'Exige decisão técnica', 'Revisão do modelo necessária']
        filtrados = [(i, s) for i, s in enumerate(linhas) if not busca or busca.casefold() in s.casefold()]
        st.caption(f'{len(filtrados)} blocos exibidos. Para modelos longos, filtre por número do item.')
        for i, texto in filtrados[:70]:
            with st.expander(f'Bloco {i+1}: {texto[:105]}'):
                chave = str(i)
                anterior = st.session_state.classificacoes.get(chave, 'Não analisado')
                valor = st.selectbox('Classificação', opcoes, index=opcoes.index(anterior), key=f'classe_{i}')
                st.session_state.classificacoes[chave] = valor
                st.write(texto)
        if len(filtrados) > 70:
            st.warning('Exibindo os primeiros 70 resultados. Refine a busca para visualizar os demais.')
        inventario = [{'indice': i+1, 'texto': s, 'classificacao': st.session_state.classificacoes.get(str(i), 'Não analisado')} for i, s in enumerate(linhas)]
        st.download_button('Baixar inventário em JSON', json.dumps(inventario, ensure_ascii=False, indent=2).encode('utf-8'), 'inventario_modelo_tr.json', 'application/json')
    else:
        st.warning('Carregue o arquivo ODT na barra lateral para iniciar a classificação.')

with aba2:
    st.subheader('Rascunho técnico orientado')
    st.warning('Os campos abaixo são experimentais. A redação e as regras de cada item ainda serão definidas a partir do inventário completo.')
    unidade = st.text_input('Unidade solicitante', value=st.session_state.dados.get('unidade', 'Secretaria de Obras e Infraestrutura'))
    responsavel = st.text_input('Responsável pela elaboração', value=st.session_state.dados.get('responsavel', ''))
    tipo = st.selectbox('Natureza', ['Obra de engenharia', 'Serviço de engenharia', 'Outro (requer revisão do modelo)'])
    regime = st.selectbox('Regime de execução (a justificar)', ['Não definido', 'Empreitada por preço global', 'Empreitada por preço unitário', 'Tarefa'])
    st.session_state.dados.update({'unidade': unidade, 'responsavel': responsavel, 'tipo': tipo, 'regime': regime})
    for num, titulo in CAPITULOS:
        with st.expander(f'{num}. {titulo}', expanded=(num == '1')):
            st.caption('Insira somente as informações da contratação atual. Não copie exemplos do modelo.')
            campo = f'capitulo_{num}'
            st.session_state.dados[campo] = st.text_area('Conteúdo técnico preliminar', value=st.session_state.dados.get(campo, ''), height=150, key=f'entrada_{num}')

with aba3:
    st.subheader('Conferência preliminar')
    dados = st.session_state.dados
    pendencias = []
    if not dados.get('unidade', '').strip():
        pendencias.append('Unidade solicitante não informada.')
    if not dados.get('responsavel', '').strip():
        pendencias.append('Responsável pela elaboração não informado.')
    if dados.get('regime') == 'Não definido':
        pendencias.append('Regime de execução ainda não definido.')
    for num, titulo in CAPITULOS:
        if not dados.get(f'capitulo_{num}', '').strip():
            pendencias.append(f'Capítulo {num} sem conteúdo técnico.')
    if pendencias:
        st.error(f'{len(pendencias)} pendência(s) de preenchimento detectada(s).')
        for item in pendencias:
            st.write('• ' + item)
    else:
        st.success('Todos os campos preliminares estão preenchidos. Isso NÃO significa conformidade com o modelo ou com a legislação.')
    st.info('A validação de textos oficiais, subitens, escolhas condicionais, referências cruzadas e compatibilidade jurídica será incorporada após a classificação integral do modelo.')
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(3)
    sec.right_margin = Cm(2)
    normal = doc.styles['Normal']
    normal.font.name = 'Liberation Serif'
    normal.font.size = Pt(11)
    doc.add_heading('RASCUNHO PRELIMINAR DE TERMO DE REFERÊNCIA', 0)
    doc.add_paragraph('Documento de trabalho. Não utilizar como TR final sem revisão e adequação ao modelo municipal.')
    doc.add_paragraph('UNIDADE SOLICITANTE: ' + dados.get('unidade', ''))
    doc.add_paragraph('RESPONSÁVEL: ' + dados.get('responsavel', ''))
    doc.add_paragraph('NATUREZA: ' + dados.get('tipo', ''))
    doc.add_paragraph('REGIME INDICADO: ' + dados.get('regime', ''))
    for num, titulo in CAPITULOS:
        doc.add_heading(f'{num}. {titulo}', level=1)
        texto = dados.get(f'capitulo_{num}', '')
        for paragrafo in texto.split('\n'):
            if paragrafo.strip():
                doc.add_paragraph(paragrafo.strip())
        if not texto.strip():
            doc.add_paragraph('[PENDENTE DE PREENCHIMENTO]')
    saida = io.BytesIO()
    doc.save(saida)
    st.download_button('Baixar rascunho DOCX', saida.getvalue(), 'rascunho_tr_nao_validado.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    st.download_button('Salvar respostas em JSON', json.dumps(dados, ensure_ascii=False, indent=2).encode('utf-8'), 'respostas_tr.json', 'application/json')
