"""TRLP 0.2: editor orientado, sem declarar conformidade jurídica automática."""
import io
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import streamlit as st
from docx import Document
from docx.shared import Cm, Pt

st.set_page_config(page_title="TRLP | Termos de Referência", page_icon="📄", layout="wide")
CAPITULOS = [
    ("1", "OBJETO"), ("2", "JUSTIFICATIVA E OBJETIVO DA CONTRATAÇÃO"),
    ("3", "DESCRIÇÃO DA SOLUÇÃO COMO UM TODO"), ("4", "REQUISITOS DA CONTRATAÇÃO"),
    ("5", "FORMA DE EXECUÇÃO DO OBJETO"), ("6", "GESTÃO DO CONTRATO"),
    ("7", "MEDIÇÃO E PAGAMENTO"), ("8", "SELEÇÃO DO FORNECEDOR"),
    ("9", "ESTIMATIVA DO PREÇO"), ("10", "ADEQUAÇÃO ORÇAMENTÁRIA"),
    ("11", "DISPOSIÇÕES FINAIS"),
]
MODELO = Path(__file__).with_name("modelo_tr.odt")
NS_TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
NS_OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"

@st.cache_data

def extrair_paragrafos(conteudo):
    """Lê somente parágrafos do corpo, sem comentários/anotações ODT."""
    with zipfile.ZipFile(io.BytesIO(conteudo)) as pacote:
        raiz = ET.fromstring(pacote.read("content.xml"))
    corpo = raiz.find(f".//{NS_OFFICE}body/{NS_OFFICE}text")
    if corpo is None:
        return []
    resultado = []
    for el in corpo.iter():
        if el.tag not in (NS_TEXT + "p", NS_TEXT + "h"):
            continue
        if any(p.tag == NS_OFFICE + "annotation" for p in el.iter()):
            # Comentários podem estar dentro de parágrafos: extrair somente conteúdo fora deles.
            pass
        partes = []
        def coletar(no):
            if no.tag == NS_OFFICE + "annotation":
                return
            if no.text:
                partes.append(no.text)
            for filho in no:
                coletar(filho)
                if filho.tail:
                    partes.append(filho.tail)
        coletar(el)
        texto = " ".join("".join(partes).split())
        if texto:
            resultado.append(texto)
    return resultado

@st.cache_data

def capitulos_do_modelo(conteudo):
    linhas = extrair_paragrafos(conteudo)
    blocos = {num: [] for num, _ in CAPITULOS}
    atual = None
    for linha in linhas:
        titulo = next((num for num, nome in CAPITULOS if re.match(rf"^{num}\.\s+{re.escape(nome)}\s*$", linha, re.I)), None)
        if titulo:
            atual = titulo
            continue
        if atual:
            blocos[atual].append(linha)
    return blocos


def inicial():
    return {"unidade": "Secretaria de Obras e Infraestrutura", "responsavel": "",
            "natureza": "Obra de engenharia", "regime": "Não definido", "objeto": "",
            "capitulos": {n: "" for n, _ in CAPITULOS},
            "alteracoes": {n: "" for n, _ in CAPITULOS}}

if "projeto" not in st.session_state:
    st.session_state.projeto = inicial()

st.title("TRLP | Elaboração de Termos de Referência")
st.caption("Versão 0.2 • Editor orientado e rascunho editável. Não substitui a conferência técnica, administrativa ou jurídica.")

if MODELO.is_file():
    modelo_bytes = MODELO.read_bytes()
    blocos = capitulos_do_modelo(modelo_bytes)
    st.sidebar.success("Modelo de referência carregado do repositório.")
else:
    blocos = {n: [] for n, _ in CAPITULOS}
    st.sidebar.error("Arquivo modelo_tr.odt não encontrado na raiz do repositório.")

st.sidebar.info("O modelo é exemplificativo. Não transfira objetos, valores, prazos ou dados de contratações anteriores para o novo TR.")
importado = st.sidebar.file_uploader("Retomar trabalho salvo (.json)", type="json")
if importado is not None:
    try:
        candidato = json.loads(importado.getvalue().decode("utf-8"))
        if not isinstance(candidato, dict) or candidato.get("formato") != "TRLP-0.2":
            st.sidebar.error("Arquivo incompatível com a versão 0.2.")
        elif st.sidebar.button("Confirmar importação"):
            dados = candidato.get("dados", {})
            base = inicial()
            for k in ("unidade", "responsavel", "natureza", "regime", "objeto"):
                if isinstance(dados.get(k), str):
                    base[k] = dados[k]
            for grupo in ("capitulos", "alteracoes"):
                if isinstance(dados.get(grupo), dict):
                    for n, _ in CAPITULOS:
                        if isinstance(dados[grupo].get(n), str):
                            base[grupo][n] = dados[grupo][n]
            st.session_state.projeto = base
            st.rerun()
    except (UnicodeError, ValueError):
        st.sidebar.error("Não foi possível ler o JSON.")

p = st.session_state.projeto
aba1, aba2, aba3 = st.tabs(["1. Dados da contratação", "2. Redação por capítulo", "3. Conferência e exportação"])
with aba1:
    st.subheader("Identificação da contratação")
    p["unidade"] = st.text_input("Unidade solicitante", value=p["unidade"])
    p["responsavel"] = st.text_input("Responsável pela elaboração", value=p["responsavel"])
    p["natureza"] = st.selectbox("Natureza", ["Obra de engenharia", "Serviço de engenharia", "Outro (requer adequação do modelo)"], index=["Obra de engenharia", "Serviço de engenharia", "Outro (requer adequação do modelo)"].index(p["natureza"]) if p["natureza"] in ["Obra de engenharia", "Serviço de engenharia", "Outro (requer adequação do modelo)"] else 0)
    regimes = ["Não definido", "Empreitada por preço global", "Empreitada por preço unitário", "Tarefa"]
    p["regime"] = st.selectbox("Regime de execução", regimes, index=regimes.index(p["regime"]) if p["regime"] in regimes else 0)
    p["objeto"] = st.text_area("Descrição resumida do objeto da contratação atual", value=p["objeto"], height=100)
    st.info("A seleção de regime não insere automaticamente as cláusulas exemplificativas do item 5.1. O texto deverá ser adaptado e justificado tecnicamente.")

with aba2:
    st.subheader("Redação orientada por capítulos")
    st.write("À esquerda, consulte o texto do modelo como **referência**. À direita, redija o texto que realmente integrará o novo TR. Nenhum exemplo será copiado automaticamente.")
    escolha = st.selectbox("Capítulo", [n for n, _ in CAPITULOS], format_func=lambda n: f"{n}. {dict(CAPITULOS)[n]}")
    esq, dir = st.columns([1, 1], gap="large")
    with esq:
        st.markdown("**Modelo municipal: referência de leitura**")
        st.caption("Pode conter orientações, alternativas e exemplos de outras contratações.")
        referencia = "\n\n".join(blocos.get(escolha, []))
        st.text_area("Trechos originais (somente leitura)", value=referencia or "Capítulo não identificado no ODT.", height=460, disabled=True)
    with dir:
        st.markdown("**Texto para o TR atual**")
        p["capitulos"][escolha] = st.text_area("Redação do capítulo", value=p["capitulos"][escolha], height=360, key=f"texto_{escolha}", placeholder="Digite a redação completa com os subitens aplicáveis e sua numeração original.")
        p["alteracoes"][escolha] = st.text_area("Observações/justificativas de adaptações (não entram no TR)", value=p["alteracoes"][escolha], height=90, key=f"just_{escolha}")
        st.caption("O editor permite adaptações, mas ainda não comprova automaticamente que todos os subitens foram contemplados.")

with aba3:
    st.subheader("Conferência preliminar")
    pendencias = []
    for campo, rotulo in (("unidade", "Unidade solicitante"), ("responsavel", "Responsável"), ("objeto", "Objeto")):
        if not p[campo].strip():
            pendencias.append(f"{rotulo} não informado.")
    if p["regime"] == "Não definido":
        pendencias.append("Regime de execução não definido.")
    for num, titulo in CAPITULOS:
        if not p["capitulos"][num].strip():
            pendencias.append(f"Capítulo {num} ({titulo}) sem redação.")
    if pendencias:
        st.warning(f"{len(pendencias)} pendência(s) preliminar(es).")
        for item in pendencias:
            st.write("• " + item)
    else:
        st.success("Campos principais preenchidos. Isso não equivale à aprovação do TR.")
    st.error("Ainda não implementado: conferência integral de subitens, cláusulas condicionais, textos padronizados e conformidade legal. O arquivo é apenas um rascunho.")

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin, sec.bottom_margin = Cm(2.5), Cm(2.5)
    sec.left_margin, sec.right_margin = Cm(3), Cm(2)
    estilo = doc.styles["Normal"]
    estilo.font.name, estilo.font.size = "Liberation Serif", Pt(11)
    doc.add_heading("RASCUNHO DE TERMO DE REFERÊNCIA", 0)
    doc.add_paragraph("Documento preliminar. Conferir integralmente com o modelo municipal antes de encaminhar para licitação.")
    doc.add_paragraph("UNIDADE SOLICITANTE: " + p["unidade"])
    doc.add_paragraph("RESPONSÁVEL PELA ELABORAÇÃO: " + p["responsavel"])
    doc.add_paragraph("OBJETO: " + p["objeto"])
    for num, titulo in CAPITULOS:
        doc.add_heading(f"{num}. {titulo}", level=1)
        texto = p["capitulos"][num].strip()
        if texto:
            for linha in texto.splitlines():
                if linha.strip():
                    doc.add_paragraph(linha.strip())
        else:
            doc.add_paragraph("[PENDENTE DE PREENCHIMENTO]")
    arquivo = io.BytesIO()
    doc.save(arquivo)
    st.download_button("Baixar rascunho DOCX", arquivo.getvalue(), "TRLP_rascunho_nao_validado.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

pacote = {"formato": "TRLP-0.2", "dados": p}
st.sidebar.download_button("Salvar trabalho (.json)", json.dumps(pacote, ensure_ascii=False, indent=2).encode("utf-8"), "TRLP_trabalho.json", mime="application/json")
