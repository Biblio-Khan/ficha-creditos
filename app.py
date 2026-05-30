import streamlit as st
import pandas as pd
import io
import requests
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="BiblioKhan - Sistema de Catalogação",
    page_icon="⚖️",
    layout="wide"
)

# Estilização para manter a identidade visual roxa e fontes profissionais
st.markdown("""
    <style>
    textarea { font-family: 'Courier New', Courier, monospace !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { 
        height: 50px; 
        white-space: pre-wrap; 
        background-color: #f0f2f6; 
        border-radius: 5px 5px 0px 0px; 
        gap: 1px; 
        padding-top: 10px; 
        padding-bottom: 10px; 
    }
    .stTabs [aria-selected="true"] { background-color: #B19FFB !important; color: black !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÕES DE SUPORTE (API, DOCX E REGRAS AACR2) ---

def buscar_vcb_senado(termo_busca):
    url_api = "https://adm.senado.leg.br/vcb/vocab/services.php"
    params = {"task": "search", "arg": termo_busca, "output": "json"}
    try:
        resposta = requests.get(url_api, params=params, timeout=8, verify=False)
        if resposta.status_code == 200:
            dados = resposta.json()
            resultados_formatados = []
            bloco_result = dados.get("result", {})
            if isinstance(bloco_result, dict):
                for chave, item in bloco_result.items():
                    if isinstance(item, dict) and "string" in item:
                        resultados_formatados.append({
                            "termo": item["string"].strip(),
                            "id": f"VCB-{item.get('term_id', chave)}"
                        })
            return resultados_formatados
    except Exception: return []
    return []

def gerar_docx_lote(lista_fichas):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Courier New'
    font.size = Pt(10)
    for idx, ficha_texto in enumerate(lista_fichas):
        if idx > 0: doc.add_page_break()
        p = doc.add_paragraph(ficha_texto)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def formatar_entrada_e_corpo(tipo_autor, autores_lista, entity, titulo, tem_organizador, organizador_nome, tipo_org, tem_tradutor, tradutor_nome):
    entrada, corpo_autores, entrada_por_titulo = "", "", False
    if tem_organizador and tipo_autor == "Pessoa Física" and not any(a.strip() for a in autores_lista):
        entrada_por_titulo, corpo_autores = True, f"{tipo_org} por {organizador_nome.strip()}"
    elif tipo_autor == "Entidade (Órgão/Instituição)":
        entrada, corpo_autores = entity.strip().upper(), ""
    else:
        autores = [a.strip() for a in autores_lista if a.strip()]
        if len(autores) == 1:
            partes = autores[0].split()
            entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
            corpo_autores = autores[0]
        elif 2 <= len(autores) <= 3:
            partes = autores[0].split()
            entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
            corpo_autores = ", ".join(autores)
        elif len(autores) >= 4:
            entrada_por_titulo, corpo_autores = True, f"{autores[0]} [et al.]"
        if tem_organizador and organizador_nome.strip() and len(autores) < 4:
            corpo_autores += f" ; {tipo_org} por {organizador_nome.strip()}"
    if tem_tradutor and tradutor_nome.strip():
        corpo_autores = (corpo_autores + f" ; tradução por {tradutor_nome.strip()}") if corpo_autores else f"tradução por {tradutor_nome.strip()}"
    return entrada, corpo_autores, entrada_por_titulo

def calcular_cutter(tipo_autor, autores_lista, entidade, titulo, tem_organizador, organizador_nome):
    ref = ""
    if tipo_autor == "Entidade (Órgão/Instituição)" and entidade: ref = entidade.strip()
    elif tem_organizador and not any(a.strip() for a in autores_lista) and organizador_nome: ref = organizador_nome.strip().split()[-1]
    elif autores_lista and autores_lista[0].strip(): ref = autores_lista[0].strip().split()[-1]
    else: ref = titulo.strip() if titulo else "X"
    letra, letra_t = (ref[0].upper() if ref else "X"), (titulo.strip()[0].lower() if titulo else "x")
    return f"{letra}123{letra_t}"

# ==========================================
# ⚖️ ESTADOS DE SESSÃO E INICIALIZAÇÃO
# ==========================================
if "lote_fichas" not in st.session_state: st.session_state.lote_fichas = []
if "assuntos_selecionados" not in st.session_state: st.session_state.assuntos_selecionados = []
if "creditos_ativos" not in st.session_state: st.session_state.creditos_ativos = 0
if "token_atual" not in st.session_state: st.session_state.token_atual = ""

# CONFIGURAÇÃO DA PLANILHA (Substitua pelo seu link real do Google Sheets)
URL_PLANILHA = "https://docs.google.com/spreadsheets/d/1epaFSWFhnd2Q_ZjGq32wdL3LeWpEqmFn1JFRBCh0j_U/edit?usp=sharing"

st.title("BiblioKhan — Gestão Documental Jurídica")

with st.sidebar:
    st.markdown("### 🎓 Autoria")
    st.markdown("**Sabrina Lobeu** — *Bibliotecária*")
    st.markdown("---")
    if st.session_state.creditos_ativos > 0:
        st.success(f"💳 Saldo: {st.session_state.creditos_ativos} fichas")
    else:
        st.error("💳 Sem créditos ativos")

# ==========================================
# 🌟 CRIAÇÃO DAS ABAS GLOBAIS (ESTRUTURA PRINCIPAL)
# ==========================================
tab_gerador, tab_financeiro = st.tabs(["⚖️ Catalogação em Lote", "💳 Compra e Gestão de Créditos"])

# ---------------------------------------------------------
# ABA 1: GERADOR DE FICHAS (TRABALHO)
# ---------------------------------------------------------
with tab_gerador:
    # Aviso de bloqueio se não houver créditos
    if st.session_state.creditos_ativos <= 0:
        st.warning("🔒 O painel de salvamento está bloqueado. Valide seu Token na aba 'Gestão de Créditos' para continuar.")
    
    col_lote_1, col_lote_2 = st.columns([2, 1])
    qtd_f = len(st.session_state.lote_fichas)
    col_lote_1.subheader(f"📦 Lote Atual: {qtd_f} Ficha(s)")
    if qtd_f > 0:
        col_lote_2.download_button("📥 Baixar Word", gerar_docx_lote(st.session_state.lote_fichas), "lote_bibliokhan.docx")
        if col_lote_2.button("🗑️ Limpar Lote"):
            st.session_state.lote_fichas = []
            st.rerun()

    st.markdown("---")
    col_esq, col_dir = st.columns(2)

    with col_esq:
        st.subheader("1. Metadados")
        classif = st.text_input("Classificação", value="340.1")
        tipo_a = st.radio("Autoria", ["Pessoa Física", "Entidade"], horizontal=True)
        aut_list, ent_n = [], ""
        if tipo_a == "Pessoa Física":
            q_a = st.number_input("Qtd Autores", 0, 10, 1)
            for i in range(int(q_a)): aut_list.append(st.text_input(f"Autor {i+1}", key=f"a_{i}"))
        else: ent_n = st.text_input("Nome da Entidade")
        tit = st.text_input("Título Principal")
        
        c1, c2 = st.columns(2)
        with c1:
            tem_org = st.checkbox("Organizador?")
            org_n, t_org, a_org = "", "", ""
            if tem_org:
                pap = st.selectbox("Função", ["Organizador", "Coordenador"])
                org_n = st.text_input("Nome Org")
                t_org, a_org = ("organizado", "org.") if pap == "Organizador" else ("coordenado", "coord.")
        with c2:
            tem_tr = st.checkbox("Tradutor?")
            tr_n = st.text_input("Nome Tradutor") if tem_tr else ""

        st.subheader("2. Publicação")
        ed, edtr, cid, an = st.text_input("Edição", "1. ed."), st.text_input("Editora"), st.text_input("Cidade", "Brasília"), st.text_input("Ano", "2026")
        pag, isb = st.text_input("Páginas", "180"), st.text_input("ISBN")
        sup = st.radio("Suporte", ["Impresso", "Digital"], horizontal=True)
        url_a = st.text_input("URL/DOI") if sup == "Digital" else ""

    with col_dir:
        st.subheader("3. Indexação (Senado VCB)")
        t_busca = st.text_input("Buscar termo jurídico:")
        if t_busca:
            res = buscar_vcb_senado(t_busca)
            if res:
                opc = {i["termo"]: i for i in res}
                sel = st.selectbox("Conceito oficial:", sorted(opc.keys()))
                if st.button("➕ Vincular"):
                    if sel not in st.session_state.assuntos_selecionados:
                        st.session_state.assuntos_selecionados.append(sel)
                        st.rerun()
        
        if st.session_state.assuntos_selecionados:
            st.write("Assuntos:", ", ".join(st.session_state.assuntos_selecionados))
            if st.button("🗑️ Limpar Assuntos"):
                st.session_state.assuntos_selecionados = []; st.rerun()

        st.markdown("---")
        st.subheader("4. Pré-visualização")
        # Lógica de montagem da ficha
        ent_p, resp, e_tit = formatar_entrada_e_corpo(tipo_a, aut_list, ent_n, tit, tem_org, org_n, t_org, tem_tr, tr_n)
        cut = calcular_cutter(tipo_a, aut_list, ent_n, tit, tem_org, org_n)
        d_fis = f"1 recurso online ({pag} f.) " if sup == "Digital" else f"{pag} f"
        pub = f"{cid.strip()} : {edtr.strip()}, {an.strip()}."
        ass_s = " ".join([f"{i+1}. {ass}" for i, ass in enumerate(st.session_state.assuntos_selecionados)])
        rast = f" I. Título." if not e_tit else ""
        
        f_txt = f"{classif}\n{cut}   {ent_p if not e_tit else tit.strip()}\n            {tit.strip() if not e_tit else ''} / {resp}. – {ed} – {pub}\n            {d_fis}.\n\n            {ass_s}{rast}"
        st.text_area("Ficha AACR2", value=f_txt, height=200)
        
        # BOTÃO DE SALVAMENTO (BLOQUEADO POR CRÉDITO)
        btn_lock = st.session_state.creditos_ativos <= 0
        if st.button("💾 CONCLUIR E SALVAR NO LOTE", disabled=btn_lock):
            if tit.strip():
                st.session_state.lote_fichas.append(f_txt)
                st.session_state.creditos_ativos -= 1
                st.session_state.assuntos_selecionados = []
                st.success("Salvo! Crédito deduzido."); st.rerun()
            else: st.error("Preencha o título.")

# ---------------------------------------------------------
# ABA 2: FINANCEIRO E CRÉDITOS (GESTÃO)
# ---------------------------------------------------------
with tab_financeiro:
    st.header("💳 Gestão Financeira e Saldo")
    
    col_f1, col_f2 = st.columns(2)
    
    with col_f1:
        st.subheader("🔓 Validar Acesso")
        tk_in = st.text_input("Insira seu Token (E-mail ou Nome)", type="password")
        if st.button("Ativar Sistema"):
            try:
                # Trata a URL de forma robusta, independentemente do final (?usp=drivesdk, /edit, etc.)
                base_url = https://docs.google.com/spreadsheets/d/1epaFSWFhnd2Q_ZjGq32wdL3LeWpEqmFn1JFRBCh0j_U/edit?usp=drivesdk.split("/edit")[0]
                url_csv = f"{base_url}/gviz/tq?tqx=out:csv"
                
                # Lê a planilha usando o Pandas
                df = pd.read_csv(url_csv)
                
                # Força os cabeçalhos para letras minúsculas para evitar erros de digitação na planilha
                df.columns = df.columns.str.strip().str.lower()
                
                tk_c = tk_in.strip().upper()
                
                if 'token' in df.columns and 'creditos' in df.columns:
                    # Converte a coluna token para maiúsculas para cruzar os dados corretamente
                    df['token'] = df['token'].astype(str).str.strip().str.upper()
                    
                    if tk_c in df['token'].values:
                        st.session_state.creditos_ativos = int(df.loc[df['token'] == tk_c, 'creditos'].values[0])
                        st.session_state.token_atual = tk_c
                        st.success(f"Token Ativo! Saldo: {st.session_state.creditos_ativos}")
                        st.rerun()
                    else:
                        st.error("Token não localizado na planilha. Verifique se digitou corretamente.")
                else:
                    st.error("Erro de estrutura: A planilha precisa ter as colunas 'token' e 'creditos' na primeira linha.")
            except Exception as e:
                st.error("Erro ao conectar à planilha. Verifique o link ou se ela está configurada como 'Qualquer pessoa com o link'.")

    with col_f2:
        st.subheader("🛒 Tabela de Preços")
        st.markdown("""
        * **30 Fichas** — R$ 49,00
        * **60 Fichas** — R$ 89,00 
        * **100 Fichas** — **R$ 129,00 [RECOMENDADO]**
        * **300 Fichas** — R$ 299,00
        """)
        st.info("🔑 **PIX:** `bibliokhancontato@gmail.com`")

    st.markdown("---")
    st.subheader("📩 Envio de Comprovante")
    with st.form("pix_form"):
        n_c = st.text_input("Nome Completo")
        e_c = st.text_input("E-mail")
        arq = st.file_uploader("Anexe o comprovante", type=["jpg", "png", "pdf"])
        if st.form_submit_button("Enviar para Restauração de Saldo"):
            if n_c and arq:
                st.success("Enviado! Em breve Sabrina atualizará seu saldo na planilha e você receberá um aviso.")
            else: st.error("Preencha os campos obrigatórios.")
