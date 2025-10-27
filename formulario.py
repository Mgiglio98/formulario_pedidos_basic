import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from openpyxl import load_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib

# --- CONFIGURAÇÃO GERAL ---
st.set_page_config(page_title="Pedido de Materiais", page_icon="📦")

# --- CSS (espaçamento + tema claro fixo) ---
st.markdown("""
<style>
[data-testid="stAppViewContainer"] .main .block-container {
    padding-top: 0.5rem;
    padding-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

# --- INICIALIZAÇÃO DE ESTADO ---
for campo, valor_padrao in {
    "insumos": [], "excel_bytes": None, "nome_arquivo": "",
    "pedido_numero": "", "solicitante": "", "executivo": "",
    "obra_selecionada": "", "cnpj": "", "endereco": "", "cep": "",
    "data_pedido": date.today()
}.items():
    st.session_state.setdefault(campo, valor_padrao)

# --- FUNÇÕES AUXILIARES ---
def limpar_formulario():
    """Reseta todos os campos e insumos."""
    for campo in ["pedido_numero", "solicitante", "executivo", "obra_selecionada",
                  "cnpj", "endereco", "cep", "excel_bytes", "nome_arquivo", "pedido_enviado"]:
        st.session_state[campo] = ""
    st.session_state.data_pedido = date.today()
    st.session_state.insumos = []

def enviar_email_pedido(assunto, arquivo_bytes, insumos_adicionados, df_insumos):
    """Envia o e-mail do pedido com o anexo Excel."""
    try:
        smtp_user = "matheus.almeida@osborne.com.br"
        smtp_pass = st.secrets["SMTP_PASSWORD"]

        basicos, especificos, sem_codigo = [], [], []

        for item in insumos_adicionados:
            qtd, desc, codigo = item["quantidade"], item["descricao"], item.get("codigo", "")
            if not codigo:
                sem_codigo.append(f"{desc} — {qtd}")
                continue

            linha = df_insumos[df_insumos["Descrição"] == desc]
            if not linha.empty and linha.iloc[0]["Basico"] and qtd <= linha.iloc[0]["Max"]:
                basicos.append(f"{desc} — {qtd}")
            else:
                especificos.append(f"{desc} — {qtd}")

        corpo = (
            "✅ Novo pedido recebido!\n\n"
            "📄 Materiais Básicos:\n" + ("\n".join(basicos) or "Nenhum") +
            "\n\n🛠️ Materiais Específicos:\n" + ("\n".join(especificos) or "Nenhum") +
            "\n\n📌 Insumos sem código:\n" + ("\n".join(sem_codigo) or "Nenhum")
        )

        msg = MIMEMultipart()
        msg["From"] = msg["To"] = smtp_user
        msg["Subject"] = assunto
        msg.attach(MIMEText(corpo, "plain"))
        msg.attach(MIMEApplication(arquivo_bytes, _subtype="xlsx", Name="Pedido.xlsx"))

        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")

def carregar_dados():
    """Carrega dados de empreendimentos e insumos."""
    df_empreend = pd.read_excel("Empreendimentos.xlsx")
    df_empreend.columns = df_empreend.columns.str.strip().str.upper()

    df_insumos = pd.read_excel("Insumos.xlsx")
    df_insumos["Min"] = pd.to_numeric(df_insumos.iloc[:, 3], errors="coerce")
    df_insumos["Max"] = pd.to_numeric(df_insumos.iloc[:, 4], errors="coerce")
    df_insumos["Basico"] = df_insumos["Min"].notna() & df_insumos["Max"].notna()
    df_insumos = df_insumos[df_insumos["Descrição"].notna() & (df_insumos["Descrição"].str.strip() != "")]
    df_insumos = pd.concat([pd.DataFrame({"Código": [""], "Descrição": [""], "Unidade": [""]}), df_insumos], ignore_index=True)

    df_empreend.loc[-1] = [""] * df_empreend.shape[1]
    df_empreend.index = df_empreend.index + 1
    return df_empreend.sort_index(), df_insumos

# --- CARREGAR BASES ---
df_empreend, df_insumos = carregar_dados()

# --- LOGO E CABEÇALHO ---
col1, col2, col3 = st.columns([1, 2, 1]) 
with col2: 
    st.image("logo.png", width=300)
st.markdown("""
    <div style='text-align: center;'>
        <h2 style='color: #000000;'>Pedido de Materiais</h2>
        <p style='font-size: 14px; color: #555;'>
            Preencha os campos com atenção. Verifique se todos os dados estão corretos antes de enviar.<br>
            Ao finalizar, o pedido será automaticamente enviado para o e-mail do setor de Suprimentos.<br>
            Você poderá baixar a planilha gerada após o envio, para registro ou controle.
        </p>
    </div>
""", unsafe_allow_html=True)

# --- DADOS DO PEDIDO ---
with st.expander("📋 Dados do Pedido", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Pedido Nº", key="pedido_numero")
        st.text_input("Solicitante", key="solicitante")
        st.selectbox("Obra", df_empreend["EMPREENDIMENTO"].unique(), key="obra_selecionada")
    with col2:
        st.date_input("Data", key="data_pedido")
        st.text_input("Executivo", key="executivo")

    if st.session_state.obra_selecionada:
        dados = df_empreend[df_empreend["EMPREENDIMENTO"] == st.session_state.obra_selecionada].iloc[0]
        st.session_state.cnpj, st.session_state.endereco, st.session_state.cep = dados["CNPJ"], dados["ENDERECO"], dados["CEP"]

    st.text_input("CNPJ/CPF", value=st.session_state.cnpj, disabled=True)
    st.text_input("Endereço", value=st.session_state.endereco, disabled=True)
    st.text_input("CEP", value=st.session_state.cep, disabled=True)

# --- ADIÇÃO DE INSUMOS ---
with st.expander("➕ Adicionar Insumo", expanded=True):
    df_insumos["opcao"] = df_insumos.apply(
        lambda x: f"{x['Descrição']} – {x['Código']} ({x['Unidade']})" if x["Código"] else x["Descrição"], axis=1
    )

    descricao_exibicao = st.selectbox("Descrição do insumo", df_insumos["opcao"], key="descricao_exibicao")
    dados_insumo = df_insumos[df_insumos["opcao"] == descricao_exibicao].iloc[0]

    codigo, unidade, descricao = dados_insumo["Código"], dados_insumo["Unidade"], dados_insumo["Descrição"]
    descricao_livre = st.text_input("Nome do insumo (livre)", key="descricao_livre", disabled=bool(codigo))
    st.text_input("Código do insumo", value=codigo, key="codigo", disabled=True)
    st.text_input("Unidade", value=unidade, key="unidade", disabled=bool(codigo))
    qtd = st.number_input("Quantidade", min_value=1, step=1, format="%d", key="quantidade")
    compl = st.text_area("Complemento (opcional)", key="complemento")

    if st.button("➕ Adicionar insumo"):
        desc_final = descricao if codigo else descricao_livre
        if desc_final:
            st.session_state.insumos.append({
                "descricao": desc_final,
                "codigo": codigo or "",
                "unidade": unidade,
                "quantidade": qtd,
                "complemento": compl
            })

            # Limpa campos após adicionar
            for campo in ["descricao_exibicao", "descricao_livre", "codigo", "unidade", "quantidade", "complemento"]:
                if campo in st.session_state:
                    del st.session_state[campo]
            st.rerun()

# --- LISTAGEM DE INSUMOS (tabela visual estilizada) ---
if st.session_state.insumos:
    st.markdown("""
        <style>
        /* Cabeçalhos */
        .tabela-header {
            font-weight: 600;
            color: #333;
            border-bottom: 2px solid #ccc;
            padding-bottom: 4px;
            margin-bottom: 4px;
            font-size: 15px;
            display: flex;
            align-items: center;
            justify-content: flex-start;
        }
        .tabela-header.center { justify-content: center; }

        /* Linhas */
        .linha-insumo {
            border-bottom: 1px solid #e6e6e6;
            padding: 3px 0;
            font-size: 14px;
            line-height: 1.4;
            display: flex;
            align-items: center;
        }
        .center { justify-content: center; text-align: center; }

        /* Botão 🗑️ */
        div[data-testid="stButton"] button {
            border: none;
            background-color: transparent;
            color: #666;
            font-size: 18px;
            padding: 0;
            line-height: 1;
            transform: translateY(-2px);
        }
        div[data-testid="stButton"] button:hover {
            color: #d9534f;
            transform: scale(1.15) translateY(-2px);
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("#### 🧾 Insumos Adicionados")

    # Cabeçalho da tabela
    col1, col2, col3, col4 = st.columns([5.8, 1.2, 1.2, 0.5])
    with col1: st.markdown("<div class='tabela-header'>Descrição</div>", unsafe_allow_html=True)
    with col2: st.markdown("<div class='tabela-header center'>Qtd</div>", unsafe_allow_html=True)
    with col3: st.markdown("<div class='tabela-header'>Unid</div>", unsafe_allow_html=True)
    with col4: st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # Linhas da tabela
    for i, insumo in enumerate(st.session_state.insumos):
        col1, col2, col3, col4 = st.columns([5.8, 1.2, 1.2, 0.5])
        with col1:
            st.markdown(f"<div class='linha-insumo'>{insumo['descricao']}</div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='linha-insumo center'>{insumo['quantidade']}</div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div class='linha-insumo'>{insumo['unidade']}</div>", unsafe_allow_html=True)
        with col4:
            if st.button("🗑️", key=f"del_{i}"):
                st.session_state.insumos.pop(i)
                st.rerun()

# --- ENVIO ---
if st.button("📤 Enviar Pedido", use_container_width=True):
    if not all([st.session_state.pedido_numero, st.session_state.solicitante, st.session_state.obra_selecionada]):
        st.warning("⚠️ Preencha os campos obrigatórios.")
        st.stop()

    try:
        wb = load_workbook("Modelo_Pedido.xlsx")
        ws = wb["Pedido"]
        ws["F2"], ws["C3"], ws["C4"], ws["C5"] = st.session_state.pedido_numero, st.session_state.data_pedido.strftime("%d/%m/%Y"), st.session_state.solicitante, st.session_state.executivo
        ws["C7"], ws["C8"], ws["C9"], ws["C10"] = st.session_state.obra_selecionada, st.session_state.cnpj, st.session_state.endereco, st.session_state.cep

        linha = 13
        for item in st.session_state.insumos:
            ws[f"B{linha}"], ws[f"C{linha}"], ws[f"D{linha}"], ws[f"E{linha}"], ws[f"F{linha}"] = item.values()
            linha += 1

        buf = BytesIO(); wb.save(buf); buf.seek(0)
        st.session_state.excel_bytes = buf.read()
        enviar_email_pedido(f"Pedido {st.session_state.pedido_numero} – {st.session_state.obra_selecionada}", st.session_state.excel_bytes, st.session_state.insumos, df_insumos)
        st.success("✅ Pedido enviado e Excel gerado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao gerar pedido: {e}")

# --- BOTÕES DE AÇÃO ---
if st.session_state.get("excel_bytes"):
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("📥 Baixar Excel", data=st.session_state.excel_bytes, file_name=f"Pedido_{st.session_state.pedido_numero}.xlsx")
    with col2:
        if st.button("🔄 Novo Pedido"): limpar_formulario(); st.rerun()

