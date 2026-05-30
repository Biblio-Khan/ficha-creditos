import streamlit as st
import pandas as pd
import io
import requests
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Gerador de Fichas Jurídicas - VCB Senado",
    page_icon="⚖️",
    layout="wide"
)

# Força fonte monoespaçada na visualização da ficha dentro do navegador
st.markdown("""
    <style>
    textarea {
        font-family: 'Courier New', Courier, monospace !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNÇÃO DE CONEXÃO E TRATAMENTO DA API DO SENADO ---
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
                            "id": f"VCB-{item.get('term_id', chave)}",
                            "nota": "Termo oficial homologado pelo Vocabulário Controlado do Senado Federal."
                        })
            return resultados_formatados
    except Exception:
        return []
    return []

# --- FUNÇÃO PARA GERAR O ARQUIVO DO WORD (.DOCX) ---
def gerar_docx_lote(lista_fichas):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Courier New'
    font.size = Pt(10)
    for idx, ficha_texto in enumerate(lista_fichas):
        if idx > 0:
            doc.add_page_break()
        p = doc.add_paragraph(ficha_texto)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_paragraph("\n" + "-"*50 + "\n")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- REGRAS DE ENTRADA CATALOGRÁFICA (AACR2/RDA) ---
def formatar_entrada_e_corpo(tipo_autor, autores_lista, entity, titulo, tem_organizador, organizador_nome, tipo_org, tem_tradutor, tradutor_nome):
    entrada = ""
    corpo_autores = ""
    entrada_por_titulo = False
    
    if tem_organizador and tipo_autor == "Pessoa Física" and not any(a.strip() for a in autores_lista):
        entrada_por_titulo = True
        entrada = ""
        corpo_autores = f"{tipo_org} por {organizador_nome.strip()}"
    elif tipo_autor == "Entidade (Órgão/Instituição)":
        entrada = entity.strip().upper()
        corpo_autores = ""
    else:
        autores = [a.strip() for a in autores_lista if a.strip()]
        qtd = len(autores)
        
        if qtd == 1:
            partes = autores[0].split()
            entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
            corpo_autores = autores[0]
        elif qtd >= 2 and qtd <= 3:
            partes = autores[0].split()
            entrada = f"{partes[-1].upper()}, {' '.join(partes[:-1])}." if len(partes) > 1 else f"{autores[0].upper()}."
            corpo_autores = ", ".join(autores)
        elif qtd >= 4:
            entrada_por_titulo = True
            entrada = ""  
            corpo_autores = f"{autores[0]} [et al.]"
            
        if tem_organizador and organizador_nome.strip() and qtd < 4:
            corpo_autores += f" ; {tipo_org} por {organizador_nome.strip()}"

    if tem_tradutor and tradutor_nome.strip():
        if corpo_autores:
            corpo_autores += f" ; tradução por {tradutor_nome.strip()}"
        else:
            corpo_autores = f"tradução por {tradutor_nome.strip()}"
            
    return entrada, corpo_autores, entrada_por_titulo

def calcular_cutter(tipo_autor, autores_lista, entidade, titulo, tem_organizador, organizador_nome):
    referencia = ""
    if tipo_autor == "Entidade (Órgão/Instituição)" and entidade:
        referencia = entidade.strip()
    elif tem_organizador and not any(a.strip() for a in autores_lista) and organizador_nome:
        referencia = organizador_nome.strip().split()[-1]
    elif autores_lista and autores_lista[0].strip():
        referencia = autores_lista[0].strip().split()[-1]
    else:
        referencia = titulo.strip() if titulo else "X"
        
    letra = referencia[0].upper() if referencia else "X"
    letra_titulo = titulo.strip()[0].lower() if titulo else "x"
    return f"{letra}123{letra_titulo}"


# ==========================================
# ⚖️ ESTADOS DE SESSÃO (CONTROLE DE CRÉDITOS)
# ==========================================
if "lote_fichas" not in st.session_state: st.session_state.lote_fichas = []
if "assuntos_selecionados" not in st.session_state: st.session_state.assuntos_selecionados = []
if "creditos_ativos" not in st.session_state: st.session_state.creditos_ativos = 0
if "token_validado" not in st.session_state: st.session_state.token_validado = False

st.title("BiblioKhan — Sistema de Catalogação")
st.subheader("⚖️ Gerador de Fichas Jurídicas — Módulo Avançado NBR/AACR2")

with st.sidebar:
    st.markdown("### 🎓 Autoria do Sistema")
    st.markdown("Desenvolvido e idealizado por:")
    st.markdown("**Sabrina Lobeu** — *Bibliotecária*")
    st.markdown("✉️ [sabslobeu@gmail.com](mailto:sabslobeu@gmail.com)")
    st.markdown("---")
    st.caption("Ecossistema livre de campos redundantes e calibrado para literatura jurídica.")

st.markdown("---")
container_lote = st.container()
with container_lote:
    col_lote_1, col_lote_2 = st.columns([2, 1])
    qtd_fichas = len(st.session_state.lote_fichas)
    col_lote_1.subheader(f"📦 Lote de Trabalho Atual: {qtd_fichas} Ficha(s) Acumulada(s)")
    
    if qtd_fichas > 0:
        arquivo_word = gerar_docx_lote(st.session_state.lote_fichas)
        col_lote_2.download_button(
            label="📥 Baixar Lote Completo (.DOCX / Word)",
            data=arquivo_word,
            file_name="lote_fichas_aacr2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        if col_lote_2.button("🗑️ Limpar Lote"):
            st.session_state.lote_fichas = []
            st.rerun()
    else:
        col_lote_2.info("O lote está vazio. Conclua uma ficha abaixo.")

st.markdown("---")
col_esquerda, col_direita = st.columns(2)

# --- COLUNA DA ESQUERDA (DADOS DO LIVRO) ---
with col_esquerda:
    st.subheader("1. Metadados & Responsabilidade")
    classificacao = st.text_input("Número de Classificação (CDD ou CDU)", value="340.1")
    tipo_autor = st.radio("Tipo de Autoria Principal", ["Pessoa Física", "Entidade (Órgão/Instituição)"], horizontal=True)
    
    autores_lista = []
    entidade_nome = ""
    
    if tipo_autor == "Pessoa Física":
        qtd_autores_input = st.number_input("Quantidade de autores principais (0 se houver apenas Organizador)", min_value=0, max_value=10, value=1)
        for i in range(int(qtd_autores_input)):
            autores_lista.append(st.text_input(f"Autor {i+1} (Nome Sobrenome)", key=f"autor_{i}"))
    else:
        entidade_nome = st.text_input("Nome da Entidade (Ex: Brasil. Supremo Tribunal Federal)")
        
    titulo = st.text_input("Título Principal")
    
    st.markdown("---")
    col_resp_1, col_resp_2 = st.columns(2)
    
    with col_resp_1:
        tem_organizador = st.checkbox("Possui Organizador/Coordenador?")
        organizador_nome = ""
        tipo_org, abreviatura_org = "", ""
        if tem_organizador:
            papel = st.selectbox("Função:", ["Organizador", "Coordenador", "Compilador"])
            organizador_nome = st.text_input("Nome do Responsável")
            if papel == "Organizador": tipo_org, abreviatura_org = "organizado", "org."
            elif papel == "Coordenador": tipo_org, abreviatura_org = "coordenado", "coord."
            else: tipo_org, abreviatura_org = "compilado", "comp."
            
    with col_resp_2:
        tem_tradutor = st.checkbox("A obra possui Tradutor?")
        tradutor_nome = ""
        if tem_tradutor:
            tradutor_nome = st.text_input("Nome do Tradutor (Nome Sobrenome)", key="trad_nome")

    st.markdown("---")
    st.subheader("2. Publicação & Descrição Física")
    edicao = st.text_input("Edição (Ex: 2. ed., 3. ed. rev. e ampl.)", value="1. ed.")
    editora = st.text_input("Editora")
    cidade = st.text_input("Cidade de Publicação", value="Brasília")
    ano = st.text_input("Ano de Publicação", value="2026")
    paginas = st.text_input("Número de Páginas/Folhas", value="180")
    
    tem_colecao = st.checkbox("Esta obra faz parte de uma Coleção / Série?")
    colecao_nome = ""
    if tem_colecao:
        colecao_nome = st.text_input("Nome da Coleção e Volume (Ex: Biblioteca jurídica, v. 12)")
        
    isbn = st.text_input("ISBN (Ex: 978-65-0000-00-0)")
    suporte = st.radio("Suporte da Obra", ["Impresso", "Digital"], horizontal=True)
    url_acesso = st.text_input("URL de Acesso / DOI") if suporte == "Digital" else ""

# --- COLUNA DA DIREITA (CRÉDITOS NO TOPO + INDEXAÇÃO + PREVIEW) ---
with col_direita:
    st.subheader("🔑 Controle de Créditos e Acesso")
    
    # 🔒 SE O UTILIZADOR NÃO TIVER CRÉDITOS NA SESSÃO
    if st.session_state.creditos_ativos <= 0:
        st.warning("🔒 Recursos de salvamento bloqueados. Insira seu Token de Créditos para começar.")
        token_input = st.text_input("Digite seu Token de Créditos (Para teste: TESTE50)", type="password", key="token_input_topo")
        
        if st.button("🔓 Ativar Créditos", key="btn_ativar_topo"):
            if token_input.strip().upper() == "TESTE50":
                st.session_state.creditos_ativos = 50
                st.session_state.token_validado = True
                st.success("✅ Token validado! 50 créditos adicionados.")
                st.rerun()
            else:
                st.error("❌ Código inválido. Contacte bibliokhancontato@gmail.com")
                
        st.markdown("---")
        st.subheader("🛒 Adquirir Novo Lote de Fichas (Via PIX)")
        st.info("🔑 **Chave PIX (E-mail):** `bibliokhancontato@gmail.com` \n\n*Valores sugeridos:* \n* 30 fichas: R$ 49 \n* 60 fichas: R$ 89 \n* 100 fichas: R$ 129")
        
        # Formulário "Mágico de Oz" para upload de comprovantes
        with st.form(key="form_pix"):
            st.markdown("##### Envie seu comprovante para liberação dos créditos:")
            nome_comprador = st.text_input("Seu Nome Completo")
            email_comprador = st.text_input("Seu E-mail de Contato")
            comprovante_arquivo = st.file_uploader("Anexe o Comprovante do PIX", type=["pdf", "png", "jpg", "jpeg"])
            
            botao_enviar_pix = st.form_submit_button("📩 Enviar Comprovante")
            
            if botao_enviar_pix:
                if nome_comprador.strip() and email_comprador.strip() and comprovante_arquivo is not None:
                    st.success("✅ Comprovante enviado com sucesso para a nossa equipe!")
                    st.info("⏳ Em breve seus créditos serão restaurados. Você receberá um e-mail com o seu Token de Acesso assim que o pagamento for compensado manualmente.")
                else:
                    st.error("❌ Por favor, preencha todos os campos e anexe o comprovante.")
    else:
        st.markdown(f"""
        <div style="background-color: #B19FFB; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 20px;">
            <h4 style="color: black; margin: 0;">💳 Saldo Atual: <strong>{st.session_state.creditos_ativos} fichas jurídicas</strong></h4>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("3. Indexação por Assunto")
    st.markdown("##### 🏛️ Buscar no VCB do Senado Federal")
    termo_busca = st.text_input("Digite um termo jurídico para pesquisar:")
    
    if termo_busca:
        resultados_vcb = buscar_vcb_senado(termo_busca)
        if resultados_vcb:
            st.success(f"{len(resultados_vcb)} conceitos localizados no Senado!")
            mapeamento_opcoes = {item["termo"]: item for item in resultados_vcb}
            lista_opcoes = sorted(list(mapeamento_opcoes.keys()))
            termo_selecionado = st.selectbox("Selecione o conceito oficial:", lista_opcoes)
            
            if st.button("➕ Vincular Assunto do Senado"):
                if termo_selecionado not in st.session_state.assuntos_selecionados:
                    st.session_state.assuntos_selecionados.append(termo_selecionado)
                    st.rerun()
        else:
            st.warning("Nenhum termo correspondente retornado pela API do Senado.")

    st.markdown("##### ✍️ Adicionar Assunto Manualmente")
    assunto_manual = st.text_input("Digite um assunto customizado:")
    if st.button("➕ Vincular Assunto Manual"):
        if assunto_manual.strip():
            termo_limpo = assunto_manual.strip()
            if termo_limpo not in st.session_state.assuntos_selecionados:
                st.session_state.assuntos_selecionados.append(termo_limpo)
                st.rerun()

    if st.session_state.assuntos_selecionados:
        st.write("**Assuntos Vinculados à Ficha:**")
        for idx, ass in enumerate(st.session_state.assuntos_selecionados):
            st.write(f"{idx+1}. {ass}")
        if st.button("🗑️ Limpar Assuntos"):
            st.session_state.assuntos_selecionados = []
            st.rerun()

    st.markdown("---")
    st.subheader("4. Pré-visualização Normativa da Ficha")
    
    # --- MONTAGEM DA FICHA CATALOGRÁFICA ---
    entrada_principal, responsabilidade, entrada_por_titulo = formatar_entrada_e_corpo(
        tipo_autor, autores_lista, entity=entidade_nome, titulo=titulo, tem_organizador=tem_organizador, organizador_nome=organizador_nome, tipo_org=tipo_org, tem_tradutor=tem_tradutor, tradutor_nome=tradutor_nome
    )
    
    cutter = calcular_cutter(tipo_autor, autores_lista, entidade_nome, titulo, tem_organizador, organizador_nome)
    
    dgm = " [recurso eletrônico]" if suporte == "Digital" else ""
    desc_fisica = f"1 recurso online ({paginas} f.) " if suporte == "Digital" else f"{paginas} f"
    
    bloco_colecao = ""
    if tem_colecao and colecao_nome.strip():
        texto_colecao = colecao_nome.strip()
        texto_colecao = texto_colecao[0].upper() + texto_colecao[1:]
        bloco_colecao = f" ({texto_colecao})"
        
    nota_acesso = f"\n            Modo de acesso: {url_acesso}" if suporte == "Digital" and url_acesso else ""
    isbn_bloco = f"\n            ISBN {isbn}" if isbn.strip() else ""
    nota_traducao = f"\n            Traduzido de obra original." if tem_tradutor and tradutor_nome.strip() else ""
    
    ed_bloco = f"{edicao.strip()} – " if edicao.strip() else ""
    pub_bloco = f"{cidade.strip()} : {editora.strip()}, {ano.strip()}."
    
    string_assuntos = " ".join([f"{i+1}. {ass}" for i, ass in enumerate(st.session_state.assuntos_selecionados)])
    
    rastreabilidade = ""
    romanos = ["I", "II", "III", "IV", "V"]
    r_idx = 0
    
    if not entrada_por_titulo:
        rastreabilidade += f" {romanos[r_idx]}. Título."
        r_idx += 1
        
    if tem_organizador and organizador_nome.strip():
        partes_org = organizador_nome.strip().split()
        nome_invertido_org = f"{partes_org[-1].upper()}, {' '.join(partes_org[:-1])}" if len(partes_org) > 1 else organizador_nome.strip().upper()
        rastreabilidade += f" {romanos[r_idx]}. {nome_invertido_org}, {abreviatura_org}."
        r_idx += 1
        
    if tem_tradutor and tradutor_nome.strip():
        partes_trad = tradutor_nome.strip().split()
        nome_invertido_trad = f"{partes_trad[-1].upper()}, {' '.join(partes_trad[:-1])}" if len(partes_trad) > 1 else tradutor_nome.strip().upper()
        rastreabilidade += f" {romanos[r_idx]}. {nome_invertido_trad}, trad."
        r_idx += 1

    if entrada_por_titulo:
        txt_ficha = f"""{classificacao}
{cutter}   {titulo.strip()}{dgm} / {responsabilidade}. – {ed_bloco}{pub_bloco}
            {desc_fisica}.{bloco_colecao}{nota_traducao}{nota_acesso}{isbn_bloco}
            
            {string_assuntos}{rastreabilidade}"""
    else:
        txt_ficha = f"""{classificacao}
{cutter}   {entrada_principal}
            {titulo.strip()}{dgm} / {responsabilidade}. – {ed_bloco}{pub_bloco}
            {desc_fisica}.{bloco_colecao}{nota_traducao}{nota_acesso}{isbn_bloco}
            
            {string_assuntos}{rastreabilidade}"""
            
    st.text_area("Ficha Gerada (Fonte Monoespaçada)", value=txt_ficha, height=240)
    
    # CONTROLADOR DE ATIVAÇÃO DO BOTÃO
    desabilitar_botao = st.session_state.creditos_ativos <= 0
    
    if st.button("💾 CONCLUIR FICHA E ENVIAR AO LOTE", disabled=desabilitar_botao, key="btn_salvar_lote"):
        valido = True
        if tipo_autor == "Pessoa Física" and not any(a.strip() for a in autores_lista) and not tem_organizador:
            valido = False
        if tipo_autor == "Entidade (Órgão/Instituição)" and not entidade_nome.strip():
            valido = False
            
        if valido and titulo.strip():
            st.session_state.lote_fichas.append(txt_ficha)
            st.session_state.assuntos_selecionados = [] 
            
            # Desconta o crédito em tempo de execução
            st.session_state.creditos_ativos -= 1
            st.success("Ficha guardada com sucesso no lote superior e 1 crédito consumido!")
            st.rerun()
        else:
            st.error("Preencha os campos de autoria/organização e o título antes de salvar.")
