"""TRLP 0.3: protótipo de formulário guiado; exige revisão antes da licitação."""
import io, json, re, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt

st.set_page_config(page_title='TRLP | Gerador orientado', page_icon='📄', layout='wide')
TITULOS=['OBJETO','JUSTIFICATIVA E OBJETIVO DA CONTRATAÇÃO','DESCRIÇÃO DA SOLUÇÃO COMO UM TODO','REQUISITOS DA CONTRATAÇÃO','FORMA DE EXECUÇÃO DO OBJETO','GESTÃO DO CONTRATO','MEDIÇÃO E PAGAMENTO','SELEÇÃO DO FORNECEDOR','ESTIMATIVA DO PREÇO','ADEQUAÇÃO ORÇAMENTÁRIA','DISPOSIÇÕES FINAIS']
MODELO=Path(__file__).with_name('modelo_tr.odt')
TEXT='{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
OFFICE='{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'

@st.cache_data
def ler_modelo(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z: root=ET.fromstring(z.read('content.xml'))
    body=root.find(f'.//{OFFICE}body/{OFFICE}text')
    if body is None:return []
    def extract(n):
        if n.tag==OFFICE+'annotation':return ''
        s=n.text or ''
        for c in n:s+=extract(c)+(c.tail or '')
        return s
    return [' '.join(extract(p).split()) for p in body.iter() if p.tag in (TEXT+'p',TEXT+'h') and extract(p).strip()]

if 'dados' not in st.session_state:st.session_state.dados={}
if 'edicoes' not in st.session_state:st.session_state.edicoes={}
d=st.session_state.dados

def sel(k,label,opcoes,help=None):
    return st.selectbox(label,opcoes,key='v_'+k,help=help)
def txt(k,label,placeholder='',required=False):
    return st.text_input(label,key='v_'+k,placeholder=placeholder)
def area(k,label,placeholder='',height=95):
    return st.text_area(label,key='v_'+k,placeholder=placeholder,height=height)
def yn(k,label):return sel(k,label,['Selecione','Sim','Não'])
def linha(n,conteudo):return f'{n}. {conteudo}' if conteudo else ''

st.title('TRLP | Elaboração orientada de Termos de Referência')
st.caption('Versão 0.3 experimental. O sistema monta um rascunho e não certifica conformidade com o modelo municipal nem com a Lei 14.133/2021.')
if MODELO.exists():
    original=ler_modelo(MODELO.read_bytes())
    st.sidebar.success('Modelo municipal carregado automaticamente')
else:
    original=[]
    st.sidebar.error('Arquivo modelo_tr.odt não encontrado na raiz do repositório.')
st.sidebar.info('O arquivo de trabalho pode ser exportado e importado. Nenhum dado é gravado no GitHub.')
imp=st.sidebar.file_uploader('Retomar trabalho (.json)',type='json')
if imp and st.sidebar.button('Importar arquivo de trabalho'):
    try:
        obj=json.load(imp)
        if obj.get('formato')!='trlp-03':raise ValueError('Arquivo incompatível com a versão 0.3')
        for k,v in obj.get('campos',{}).items():st.session_state['v_'+k]=v
        st.session_state.edicoes=obj.get('edicoes',{})
        st.rerun()
    except Exception as e:st.sidebar.error(str(e))

aba1,aba2,aba3=st.tabs(['1. Configuração','2. Perguntas por assunto','3. Prévia e exportação'])
with aba1:
    a,b=st.columns(2)
    with a:
        unidade=txt('unidade','Unidade solicitante','Secretaria de Obras e Infraestrutura')
        objeto=area('objeto','Objeto específico da contratação (informação técnica indispensável)','Ex.: execução de recapeamento asfáltico em ...')
        natureza=sel('natureza','Natureza',['Selecione','Obra de engenharia','Serviço de engenharia','Outro serviço'])
        regime=sel('regime','Regime de execução',['Selecione','Empreitada por preço global','Empreitada por preço unitário','Tarefa','Outro / requer análise'])
    with b:
        etp=yn('etp','Foi elaborado ETP?')
        memorial=yn('memorial','Existe memorial descritivo separado?')
        projetos=yn('projetos','Há projetos e documentos técnicos anexos?')
        sub=yn('sub','Será admitida subcontratação?')
        visita=yn('visita','Será prevista visita técnica facultativa?')
        garantia=yn('garantia','Será exigida garantia contratual (arts. 96 e seguintes)?')
    st.info('As escolhas devem ser tecnicamente justificadas pelo responsável. Os regimes e textos exemplificativos do modelo não são automaticamente aplicáveis a qualquer contratação.')
with aba2:
    sec=st.selectbox('Capítulo para preencher',range(1,12),format_func=lambda n:f'{n}. {TITULOS[n-1]}')
    st.subheader(f'{sec}. {TITULOS[sec-1]}')
    if sec==1:
        complemento=area('escopo','Quais serviços/etapas compõem o objeto? (apenas quando necessário)','Descreva as parcelas relevantes do escopo')
    elif sec==2:
        necessidade=area('necessidade','Qual problema ou necessidade pública a contratação atenderá?','Informação técnica indispensável')
        resultado=area('resultado','Qual resultado concreto é esperado?')
    elif sec==3:
        etapas=area('etapas','Quais etapas ou atividades integram a solução?')
        acessorios=st.multiselect('Atividades acessórias aplicáveis',['Transporte','Instalação','Montagem','Testes e ensaios','Limpeza final','Destinação de resíduos','Treinamento','Assistência técnica','Outras'],key='v_acessorios')
    elif sec==4:
        normas=area('normas','Normas técnicas específicas adicionais, se necessárias')
        exigencias=area('exigencias','Requisitos técnicos específicos indispensáveis')
        if sub=='Sim':parcelas=area('parcelas','Quais parcelas poderão ser subcontratadas?')
        if garantia=='Sim':garantia_det=area('garantia_det','Condições e percentual da garantia, conforme definição da Administração')
        if visita=='Sim':responsavel_visita=txt('responsavel_visita','Servidor responsável pela visita técnica')
    elif sec==5:
        prazo=txt('prazo','Prazo de execução (dias)','Ex.: 120')
        inicio=sel('inicio','Marco inicial do prazo',['Selecione','Emissão da Ordem de Serviço','Assinatura do contrato','Outro (descrever)'])
        if inicio.startswith('Outro'):txt('inicio_outro','Descreva o marco inicial')
        medicao=sel('medicao','Critério de medição',['Selecione','Medições periódicas por serviços executados','Por etapas concluídas','Após conclusão integral','Outro (exige justificativa)'])
        detalhes_exec=area('detalhes_exec','Particularidades de execução não cobertas pelos anexos (se houver)')
    elif sec==6:
        gestor=txt('gestor','Gestor do contrato')
        fiscal=txt('fiscal','Fiscal técnico')
        fiscaladm=txt('fiscaladm','Fiscal administrativo')
    elif sec==7:
        pagamento=sel('pagamento','Forma de pagamento',['Selecione','Conforme medições aprovadas','Após conclusão e recebimento','Por etapas previstas no cronograma','Outra (exige descrição)'])
        prazo_pg=txt('prazo_pg','Prazo de pagamento, conforme regras municipais (dias)')
        if pagamento.startswith('Outra'):area('pg_outro','Descreva a forma de pagamento')
    elif sec==8:
        criterio=sel('criterio','Critério de julgamento pretendido',['Selecione','Menor preço','Maior desconto','Outro (exige análise)'])
        habilitacao=area('habilitacao','Exigências específicas de qualificação técnica (quando cabíveis)')
    elif sec==9:
        metodo=sel('metodo','Base da estimativa',['Selecione','Tabelas de custos','Pesquisa de preços/orçamentos','Composição própria','Outra metodologia'])
        estimativa=txt('estimativa','Valor total estimado (R$)')
        fonte=area('fonte','Identificação das fontes e documentos que instruem a estimativa')
    elif sec==10:
        dotacao=txt('dotacao','Dotação orçamentária / classificação informada pelo setor competente')
        exercicio=txt('exercicio','Exercício orçamentário')
    else:
        disposicoes=area('disposicoes','Informações adicionais necessárias (se houver)')
    with st.expander('Consultar o modelo original (somente referência)'):
        if original:
            pos=next((i for i,t in enumerate(original) if re.match(rf'^{sec}\.\s+[A-ZÁÉÍÓÚÃÕÇ]',t)),None)
            end=next((i for i in range(pos+1,len(original)) if re.match(rf'^{sec+1}\.\s+[A-ZÁÉÍÓÚÃÕÇ]',original[i])),len(original)) if pos is not None else 0
            st.text('\n\n'.join(original[pos:min(end,pos+65)]) if pos is not None else 'Capítulo não localizado automaticamente.')
        else:st.warning('Modelo indisponível.')

# Geração de texto: somente cláusulas explicitamente mapeadas, não reproduz exemplos específicos.
def v(k):return st.session_state.get('v_'+k,'')
partes={i:[] for i in range(1,12)}
if v('objeto'):partes[1].append(linha('1.1',f'O presente Termo de Referência tem por objeto {v("objeto").strip().rstrip(".")}.' ))
if v('escopo'):partes[1].append(linha('1.2',v('escopo')))
if v('necessidade'):partes[2].append(v('necessidade'))
if v('resultado'):partes[2].append('Resultado esperado: '+v('resultado'))
if v('etapas'):partes[3].append(linha('3.1',v('etapas')))
if v('acessorios'):partes[3].append(linha('3.2','Atividades acessórias previstas: '+', '.join(v('acessorios'))+'.'))
if v('exigencias'):partes[4].append(linha('4.1',v('exigencias')))
if v('normas'):partes[4].append(linha('4.2','Normas técnicas adicionais: '+v('normas')))
if v('sub')=='Não':partes[4].append('4.9.1. Não é admitida a subcontratação do objeto contratual sem prévia autorização da contratante.')
if v('sub')=='Sim':partes[4].append('4.9.1. A subcontratação dependerá de prévia autorização da Administração, observadas as parcelas indicadas: '+(v('parcelas') or '[INFORMAR PARCELAS]'))
if v('garantia')=='Sim':partes[4].append('4.10.1. Será exigida garantia da contratação, nos termos dos arts. 96 e seguintes da Lei nº 14.133/2021. '+v('garantia_det'))
if v('visita')=='Sim':partes[4].append('4.11. Visita técnica facultativa, com acompanhamento de '+(v('responsavel_visita') or '[INFORMAR RESPONSÁVEL]')+'. Condições de agendamento e declaração substitutiva deverão ser conferidas.')
reg=v('regime')
if reg=='Empreitada por preço unitário':
    partes[5]+= ['5.1.1. A execução do objeto ocorrerá sob o regime de empreitada por preço unitário, nos termos do art. 46, inciso II, da Lei Federal nº 14.133/2021.','5.1.2. Os serviços serão remunerados de acordo com os quantitativos efetivamente executados e medidos pela fiscalização, aplicando-se os preços unitários constantes da proposta da contratada.','5.1.3. Eventuais variações quantitativas decorrentes das condições verificadas durante a execução deverão observar as disposições legais e contratuais aplicáveis, mediante justificativa técnica e formalização nos autos do processo.']
elif reg=='Tarefa':
    partes[5]+= ['5.1.1. A execução do objeto ocorrerá sob o regime de tarefa, nos termos do art. 46, inciso III, da Lei Federal nº 14.133/2021, mediante contratação de mão de obra para pequenos trabalhos por preço certo, com ou sem fornecimento de materiais.','5.1.2. Os serviços serão executados de acordo com as necessidades da Administração, observadas as especificações técnicas e as condições estabelecidas no contrato, sendo remunerados conforme a efetiva execução das atividades demandadas.']
elif reg=='Empreitada por preço global':
    partes[5].append('5.1.1. A execução do objeto dar-se-á de forma indireta, sob o regime de empreitada por preço global. [COMPLEMENTAR A JUSTIFICATIVA TÉCNICA ESPECÍFICA; o modelo municipal contém exemplo de projeto padrão federal que não é universal.]')
if v('prazo'):
    marco=v('inicio_outro') if v('inicio').startswith('Outro') else v('inicio')
    partes[5].append(f'5.2. Prazo de execução: {v("prazo")} dias, contados de {marco or "[DEFINIR MARCO INICIAL]"}.')
if v('detalhes_exec'):partes[5].append('Condições específicas de execução: '+v('detalhes_exec'))
if v('medicao')!='Selecione' and v('medicao'):partes[5].append('Critério de medição informado: '+v('medicao')+'. [CONFERIR ADEQUAÇÃO AO ITEM 5.15.3 DO MODELO.]')
for numero,k,t in [('6.1.1','gestor','Gestor do contrato'),('6.1.2','fiscal','Fiscal técnico do contrato'),('6.1.3','fiscaladm','Fiscal administrativo do contrato')]:
    if v(k):partes[6].append(f'{numero}. {t}: {v(k)}.')
if v('pagamento') not in ('','Selecione'):partes[7].append('Forma de pagamento proposta: '+(v('pg_outro') if v('pagamento').startswith('Outra') else v('pagamento'))+'.')
if v('prazo_pg'):partes[7].append('Prazo de pagamento informado: '+v('prazo_pg')+' dias. [CONFERIR REGRA MUNICIPAL.]')
if v('criterio') not in ('','Selecione'):partes[8].append('Critério de julgamento proposto: '+v('criterio')+'.')
if v('habilitacao'):partes[8].append('Qualificação técnica: '+v('habilitacao'))
if v('metodo') not in ('','Selecione'):partes[9].append('Metodologia de estimativa: '+v('metodo')+'.')
if v('estimativa'):partes[9].append('Valor estimado: R$ '+v('estimativa')+'.')
if v('fonte'):partes[9].append('Fontes e documentos: '+v('fonte'))
if v('dotacao'):partes[10].append('Dotação orçamentária: '+v('dotacao'))
if v('exercicio'):partes[10].append('Exercício: '+v('exercicio'))
if v('disposicoes'):partes[11].append(v('disposicoes'))

with aba3:
    pend=[]
    for k,n in [('unidade','Unidade solicitante'),('objeto','Objeto'),('natureza','Natureza'),('regime','Regime'),('necessidade','Justificativa da necessidade'),('etapas','Descrição da solução'),('prazo','Prazo de execução'),('gestor','Gestor'),('fiscal','Fiscal técnico'),('metodo','Metodologia de estimativa'),('dotacao','Dotação orçamentária')]:
        if not v(k) or v(k)=='Selecione':pend.append(n)
    for k,n in [('etp','ETP'),('memorial','Memorial'),('sub','Subcontratação'),('visita','Visita técnica'),('garantia','Garantia contratual')]:
        if v(k) in ('','Selecione'):pend.append('Decisão: '+n)
    if v('sub')=='Sim' and not v('parcelas'):pend.append('Parcelas subcontratáveis')
    if v('visita')=='Sim' and not v('responsavel_visita'):pend.append('Responsável pela visita')
    if pend:st.warning('Pendências identificadas: '+', '.join(pend))
    else:st.success('Campos principais preenchidos. Ainda é necessária conferência integral do modelo e da legislação.')
    st.info('Esta versão NÃO reproduz todas as cláusulas do modelo. Capítulos sem regras mapeadas permanecem incompletos. Exportação somente como rascunho.')
    for i,titulo in enumerate(TITULOS,1):
        with st.expander(f'{i}. {titulo}',expanded=(i==1)):
            base='\n\n'.join(partes[i])
            atual=st.text_area('Texto gerado (editável)',value=st.session_state.edicoes.get(str(i),base),height=170,key=f'edit_{i}')
            st.session_state.edicoes[str(i)]=atual
            if st.button('Restaurar texto gerado',key=f'reset_{i}'):
                st.session_state.edicoes[str(i)]=base
                st.rerun()
    campos={k[2:]:val for k,val in st.session_state.items() if k.startswith('v_')}
    pacote={'formato':'trlp-03','campos':campos,'edicoes':st.session_state.edicoes}
    st.download_button('Salvar trabalho (.json)',json.dumps(pacote,ensure_ascii=False,indent=2).encode('utf-8'),'trabalho_trlp.json','application/json')
    doc=Document(); secao=doc.sections[0];secao.top_margin=Cm(2.5);secao.bottom_margin=Cm(2.5)
    doc.add_heading('TERMO DE REFERÊNCIA — RASCUNHO INCOMPLETO',0)
    doc.add_paragraph('UNIDADE SOLICITANTE: '+v('unidade'))
    for i,titulo in enumerate(TITULOS,1):
        doc.add_heading(f'{i}. {titulo}',level=1)
        conteudo=st.session_state.edicoes.get(str(i),'\n\n'.join(partes[i]))
        for p in conteudo.split('\n'):
            if p.strip():doc.add_paragraph(p)
        if not conteudo.strip():doc.add_paragraph('[CAPÍTULO PENDENTE DE ELABORAÇÃO]')
    doc.add_paragraph('DOCUMENTO EXPERIMENTAL: revisar integralmente conforme modelo oficial antes de encaminhar à licitação.')
    buf=io.BytesIO();doc.save(buf)
    st.download_button('Baixar rascunho DOCX',buf.getvalue(),'TRLP_rascunho_v03.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document')
