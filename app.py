import streamlit as st
import pandas as pd
import io
import datetime
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

# 🌟 FUNÇÃO CORRIGIDA PARA LER O LINK DA PLANILHA DO TELEMOVEL
def carregar_creditos_planilha(url_original):
    try:
        # Se o link vier com /edit ou /usp=sharing, essa linha limpa e joga o formato correto de exportação em CSV
        if "/edit" in url_original:
            url_base = url_original.split("/edit")[0]
        else:
            url_base = url_original
        
        url_csv = f"{url_base}/gviz/tq?tqx=out:csv"
        
        # Faz a leitura usando o pandas
        df = pd.read_csv(url_csv)
        
        # Padroniza o nome das colunas para letras minúsculas e sem espaços
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        st.error(f"Erro ao processar o link da planilha: {e}")
        return None

# ==========================================
# ⚖️ ESTADOS DE SESSÃO E INICIALIZAÇÃO
# ==========================================
if "lote_fichas" not in st.session_state: st.session_state.lote_fichas = []
if "assuntos_selecionados" not in st.session_state: st.session_state.assuntos_selecionados = []
if "creditos_ativos" not in st.session_state: st.session_state.creditos_ativos = 0
if "token_atual" not in st.session_state: st.session_state.token_atual = ""

# 📊 CONFIGURAÇÕES DO SISTEMA (Preencha aqui com os seus dados)
URL_PLANILHA = "https://docs.google.com/spreadsheets/d/1epaFSWFhnd2Q_ZjGq32wdL3LeWpEqmFn1JFRBCh0j_U/edit?usp=drivesdk"

TELEGRAM_BOT_TOKEN = st.secrets["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = st.secrets["TELEGRAM_CHAT_ID"]

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
# 🌟 CRIAÇÃO DAS ABAS GLOBAIS
# ==========================================
tab_gerador, tab_financeiro = st.tabs(["⚖️ Catalogação em Lote", "💳 Compra e Gestão de Créditos"])

# ---------------------------------------------------------
# ABA 1: GERADOR DE FICHAS (TRABALHO)
# ---------------------------------------------------------
with tab_gerador:
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
        ent_p, resp, e_tit = formatar_entrada_e_corpo(tipo_a, aut_list, ent_n, tit, tem_org, org_n, t_org, tem_tr, tr_n)
        cut = calcular_cutter(tipo_a, aut_list, ent_n, tit, tem_org, org_n)
        d_fis = f"1 recurso online ({pag} f.) " if sup == "Digital" else f"{pag} f"
        pub = f"{cid.strip()} : {edtr.strip()}, {an.strip()}."
        ass_s = " ".join([f"{i+1}. {ass}" for i, ass in enumerate(st.session_state.assuntos_selecionados)])
        rast = f" I. Título." if not e_tit else ""
        
        f_txt = f"{classif}\n{cut}   {ent_p if not e_tit else tit.strip()}\n            {tit.strip() if not e_tit else ''} / {resp}. – {ed} – {pub}\n            {d_fis}.\n\n            {ass_s}{rast}"
        st.text_area("Ficha AACR2", value=f_txt, height=200)
        
       if st.button("💾 CONCLUIR E SALVAR NO LOTE", disabled=st.session_state.creditos_ativos <= 0):
            if tit.strip():
        # 1. Adiciona ao lote e atualiza estado local
        st.session_state.lote_fichas.append(f_txt)
        st.session_state.creditos_ativos -= 1
        st.session_state.assuntos_selecionados = []
        
        # 2. Salva o registro no Firestore (histórico de longo prazo)
        db.collection('historico_producao').add({
            'usuario': st.session_state["usuario_atual"], 
            'titulo_obra': tit, # Certifique-se que 'tit' é o título correto
            'data': firestore.SERVER_TIMESTAMP
        })
        
        # 3. Finaliza
        st.success("Salvo! Crédito deduzido.")
        st.rerun()
    else:
        st.error("Preencha o título.")

# ---------------------------------------------------------
# ABA 2: FINANCEIRO E CRÉDITOS (LEITURA DA PLANILHA INTEGRADA)
# ---------------------------------------------------------
with tab_financeiro:
    st.header("💳 Gestão Financeira e Saldo")
    col_f1, col_f2 = st.columns(2)
    
    with col_f1:
        st.subheader("🔓 Validar Acesso")
        tk_in = st.text_input("Insira seu Token (E-mail de Cadastro)", type="password")
        if st.button("Ativar Sistema"):
            # Chama a função inteligente que trata o link do telemóvel
            df = carregar_creditos_planilha(URL_PLANILHA)
            
            if df is not None:
                tk_c = tk_in.strip().upper()
                if 'token' in df.columns and 'creditos' in df.columns:
                    df['token'] = df['token'].astype(str).str.strip().str.upper()
                    
                    if tk_c in df['token'].values:
                        st.session_state.creditos_ativos = int(df.loc[df['token'] == tk_c, 'creditos'].values[0])
                        st.session_state.token_atual = tk_c
                        st.success(f"Token Ativo! Saldo: {st.session_state.creditos_ativos}")
                        st.rerun()
                    else:
                        st.error("E-mail/Token não cadastrado na planilha.")
                else:
                    st.error("A planilha não possui as colunas 'token' e 'creditos'. Ajuste os cabeçalhos.")

    with col_f2:
        st.subheader("🛒 Tabela de Preços")
        st.markdown("""
        * **30 Fichas** — R$ 49,00 *(R$ 1,63/un)*
        * **60 Fichas** — R$ 89,00 ⚡ *Economize R$ 9!*
        * **100 Fichas** — **R$ 129,00 [O DOBRO POR +R$40]**
        * **300 Fichas** — **R$ 299,00 [FICHA POR R$ 0,99]**
        """)
        st.info("🔑 **PIX:** `bibliokhancontato@gmail.com`")

    st.markdown("---")
    st.subheader("📩 Envio de Comprovante")
    
    with st.form("pix_form"):
        nome_cliente = st.text_input("Nome Completo")
        email_cliente = st.text_input("E-mail de Cadastro no Sistema")
        
        pacote_escolhido = st.selectbox(
            "Qual pacote de créditos você comprou?",
            options=[
                "30 Fichas (R$ 49,00)",
                "60 Fichas (R$ 89,00)",
                "100 Fichas (R$ 129,00)",
                "300 Fichas (R$ 299,00)"
            ]
        )
        
        comprovante = st.file_uploader("Anexe a imagem ou PDF do comprovante do PIX", type=["jpg", "png", "jpeg", "pdf"])
        
        if st.form_submit_button("Enviar para Restauração de Saldo"):
            if nome_cliente.strip() and email_cliente.strip() and comprovante is not None:
                with st.spinner("Enviando dados de pagamento para a Sabrina... Por favor, aguarde."):
                    try:
                        texto_notificacao = (
                            f"🔥 *NOVO COMPROVANTE RECEBIDO!*\n\n"
                            f"👤 *Cliente:* {nome_cliente.strip()}\n"
                            f"📧 *E-mail:* {email_cliente.strip()}\n"
                            f"💰 *Pacote Escolhido:* {pacote_escolhido}\n"
                            f"📅 *Data/Hora:* {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
                        )
                        
                        url_api_telegram = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
                        
                        ficheiro_envio = {
                            "photo": (comprovante.name, comprovante.getvalue(), comprovante.type)
                        }
                        
                        dados_requisicao = {
                            "chat_id": TELEGRAM_CHAT_ID,
                            "caption": texto_notificacao,
                            "parse_mode": "Markdown"
                        }
                        
                        resposta_tg = requests.post(url_api_telegram, data=dados_requisicao, files=ficheiro_envio, timeout=15)
                        
                        if resposta_tg.status_code == 200:
                            st.success("✅ Comprovante enviado com sucesso!")
                            st.info("⏳ O seu saldo será atualizado assim que a Sabrina validar o recebimento do PIX.")
                        else:
                            st.error(f"Erro na API do Telegram (Código {resposta_tg.status_code}). Verifique suas credenciais.")
                    except Exception as erro_conexao:
                        st.error(f"Erro de conexão ao tentar falar com o Telegram: {erro_conexao}")
            else:
                st.error("❌ Por favor, preencha todos os campos e anexe o ficheiro do comprovante.")

with tab_produtividade:
        st.subheader("📊 Meu Histórico de Produção")
        logs = db.collection('historico_producao').where('usuario', '==', st.session_state["usuario_atual"]).stream()
        dados = [doc.to_dict() for doc in logs]
        
        if dados:
            df_prod = pd.DataFrame(dados)
            df_prod['data'] = pd.to_datetime(df_prod['data'], unit='ms')
            col1, col2 = st.columns(2)
            col1.metric("Total de fichas geradas:", len(df_prod))
            
            st.write("### 📈 Evolução Mensal")
            df_prod['mes_ano'] = df_prod['data'].dt.to_period('M').astype(str)
            st.bar_chart(df_prod.groupby('mes_ano').size())
            
            st.write("### 📋 Detalhe Recente")
            st.dataframe(df_prod[['data', 'titulo_obra']].sort_values(by='data', ascending=False), use_container_width=True)
        else:
            st.info("Ainda não há registros de produção.")
