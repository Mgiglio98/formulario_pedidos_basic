import streamlit as st
from datetime import date
from openpyxl import load_workbook
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib

st.set_page_config(page_title="Pedido de Materiais", page_icon="📦")  # sem wide

# Ajuste só do espaço superior
st.markdown("""
<style>
[data-testid="stAppViewContainer"] .main .block-container{
    padding-top: 0.5rem;   /* ajuste fino do “respiro” */
    padding-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

# --- Inicializações de sessão ---
if "insumos" not in st.session_state:
    st.session_state.insumos = []
if "resetar_insumo" not in st.session_state:
    st.session_state.resetar_insumo = False
if "resetar_pedido" not in st.session_state:
    st.session_state.resetar_pedido = False
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "nome_arquivo" not in st.session_state:
    st.session_state.nome_arquivo = ""

# --- Garantir que os campos de cabeçalho existam no session_state ---
for campo in ["pedido_numero", "solicitante", "executivo", "obra_selecionada", "cnpj", "endereco", "cep"]:
    if campo not in st.session_state:
        st.session_state[campo] = ""

# Campo de data precisa ser do tipo date, não string
if "data_pedido" not in st.session_state:
    st.session_state.data_pedido = date.today()

# --- Funções auxiliares ---
def resetar_campos_insumo():
    # Limpa apenas chaves se elas ainda existirem
    for campo in ["descricao", "descricao_livre", "codigo", "unidade", "quantidade", "complemento", "descricao_exibicao"]:
        if campo in st.session_state:
            try:
                del st.session_state[campo]
            except Exception:
                pass  # ignora caso já tenha sido removido

def resetar_formulario():
    # Marca para resetar campos de insumos
    resetar_campos_insumo()

    # Limpa outras chaves da sessão
    for campo in ["insumos", "excel_bytes", "nome_arquivo", "pedido_numero", "data_pedido", "solicitante",
                  "executivo", "obra_selecionada", "cnpj", "endereco", "cep"]:
        if campo in st.session_state:
            try:
                del st.session_state[campo]
            except Exception:
                pass

    st.session_state.resetar_pedido = False
    st.session_state.resetar_insumo = False
    
# --- Função para enviar e-mail ---
def enviar_email_pedido(assunto, arquivo_bytes, insumos_adicionados, df_insumos):
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    smtp_user = "matheus.almeida@osborne.com.br"
    smtp_password = st.secrets["SMTP_PASSWORD"]

    # Separa os tipos de insumos
    basicos = []
    especificos = []
    sem_codigo = []

    for item in insumos_adicionados:
        qtd = item["quantidade"]
        descricao = item["descricao"]
        codigo = item.get("codigo", "")

        if not codigo or str(codigo).strip() == "":
            sem_codigo.append(f"{descricao} — {qtd}")
            continue

        linha_df = df_insumos[df_insumos["Descrição"] == item["descricao"]]
        if not linha_df.empty and linha_df.iloc[0]["Basico"]:
            max_qtd = linha_df.iloc[0]["Max"]
        
            if pd.notna(max_qtd) and qtd <= max_qtd:
                basicos.append(f"{item['descricao']} — {qtd}")
            else:
                especificos.append(f"{item['descricao']} — {qtd}")
        else:
            especificos.append(f"{item['descricao']} — {qtd}")

    # Monta corpo do e-mail
    corpo = "✅ Novo pedido recebido!\n\n"
    corpo += "📄 Materiais Básicos:\n"
    corpo += "\n".join(basicos) if basicos else "Nenhum"

    corpo += "\n\n🛠️ Materiais Específicos:\n"
    corpo += "\n".join(especificos) if especificos else "Nenhum"

    corpo += "\n\n📌 Insumos sem código cadastrado:\n"
    corpo += "\n".join(sem_codigo) if sem_codigo else "Nenhum"

    # Monta o e-mail
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = smtp_user
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain"))

    # Anexa o arquivo
    anexo = MIMEApplication(arquivo_bytes, _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    anexo.add_header('Content-Disposition', 'attachment', filename="Pedido.xlsx")
    msg.attach(anexo)

    # Envia o e-mail
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
        server.quit()
        print("📨 E-mail com anexo enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# --- Carrega dados ---
def carregar_dados():
    df_empreend = pd.read_excel("Empreendimentos.xlsx")
    df_empreend.columns = df_empreend.columns.str.strip().str.upper()
    df_insumos = pd.read_excel("Insumos.xlsx")

    # Carrega min e max (colunas D e E)
    df_insumos["Min"] = pd.to_numeric(df_insumos.iloc[:, 3], errors="coerce")
    df_insumos["Max"] = pd.to_numeric(df_insumos.iloc[:, 4], errors="coerce")

    df_insumos["Basico"] = df_insumos["Min"].notna() & df_insumos["Max"].notna()

    df_insumos = df_insumos[df_insumos["Descrição"].notna() & (df_insumos["Descrição"].str.strip() != "")]

    df_empreend.loc[-1] = [""] * df_empreend.shape[1]
    df_empreend.index = df_empreend.index + 1
    df_empreend = df_empreend.sort_index()

    insumos_vazios = pd.DataFrame({"Código": [""], "Descrição": [""], "Unidade": [""]})
    df_insumos = pd.concat([insumos_vazios, df_insumos], ignore_index=True)

    return df_empreend, df_insumos

# --- Dados ---
df_empreend, df_insumos = carregar_dados()

# 🔄 Limpa os campos do cabeçalho apenas quando um pedido for enviado com sucesso
if st.session_state.get("limpar_pedido", False):
    for campo in ["pedido_numero", "solicitante", "executivo", "obra_selecionada", "cnpj", "endereco", "cep"]:
        if campo in st.session_state:
            try:
                st.session_state[campo] = ""
            except Exception:
                pass
    st.session_state.data_pedido = date.today()
    st.session_state.insumos = []
    st.session_state.limpar_pedido = False

# --- Logo e título ---
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

# --- Dados do Pedido ---
with st.expander("📋 Dados do Pedido", expanded=True):
    if st.session_state.resetar_pedido:
        st.session_state.pedido_numero = ""
        st.session_state.data_pedido = date.today()
        st.session_state.solicitante = ""
        st.session_state.executivo = ""
        st.session_state.obra_selecionada = ""
        st.session_state.cnpj = ""
        st.session_state.endereco = ""
        st.session_state.cep = ""
        st.session_state.resetar_pedido = False

    col1, col2 = st.columns(2)
    with col1:
        pedido_numero = st.text_input("Pedido Nº", key="pedido_numero")
        solicitante = st.text_input("Solicitante", key="solicitante")
        obra_selecionada = st.selectbox("Obra", df_empreend["EMPREENDIMENTO"].unique(), index=0, key="obra_selecionada")
    with col2:
        data_pedido = st.date_input(
            "Data",
            key="data_pedido",
            value=st.session_state.data_pedido if "data_pedido" in st.session_state else date.today()
        )
        executivo = st.text_input("Executivo", key="executivo")

    if obra_selecionada:
        dados_obra = df_empreend[df_empreend["EMPREENDIMENTO"] == obra_selecionada].iloc[0]
        st.session_state.cnpj = dados_obra["CNPJ"]
        st.session_state.endereco = dados_obra["ENDERECO"]
        st.session_state.cep = dados_obra["CEP"]

    st.text_input("CNPJ/CPF", value=st.session_state.get("cnpj", ""), disabled=True)
    st.text_input("Endereço", value=st.session_state.get("endereco", ""), disabled=True)
    st.text_input("CEP", value=st.session_state.get("cep", ""), disabled=True)

st.divider()

# --- Adição de Insumos ---
with st.expander("➕ Adicionar Insumo", expanded=True):

    # 🔄 Limpeza segura logo no início do ciclo
    if st.session_state.get("limpar_insumo", False):
        for campo in ["descricao", "descricao_livre", "codigo", "unidade", "quantidade", "complemento", "descricao_exibicao"]:
            if campo in st.session_state:
                try:
                    if campo == "quantidade":
                        st.session_state[campo] = 1
                    elif campo == "descricao_exibicao":
                        st.session_state[campo] = list(df_insumos.sort_values(by="Descrição", ascending=True)["Descrição"])[0]
                    else:
                        st.session_state[campo] = ""
                except Exception:
                    pass
        st.session_state.limpar_insumo = False

    # --- Lista de insumos ---
    df_insumos_lista = df_insumos.sort_values(by="Descrição", ascending=True).copy()
    df_insumos_lista["opcao_exibicao"] = df_insumos_lista.apply(
        lambda x: f"{x['Descrição']} – {x['Código']} ({x['Unidade']})"
        if pd.notna(x["Código"]) and str(x["Código"]).strip() != ""
        else x["Descrição"],
        axis=1
    )

    descricao_exibicao = st.selectbox(
        "Descrição do insumo (Digite em MAIÚSCULO)",
        df_insumos_lista["opcao_exibicao"],
        key="descricao_exibicao"
    )

    dados_insumo = df_insumos_lista[df_insumos_lista["opcao_exibicao"] == descricao_exibicao].iloc[0]
    usando_base = bool(dados_insumo["Código"]) and str(dados_insumo["Código"]).strip() != ""

    if usando_base:
        st.session_state.codigo = dados_insumo["Código"]
        st.session_state.unidade = dados_insumo["Unidade"]
        st.session_state.descricao = dados_insumo["Descrição"]
    else:
        st.session_state.codigo = ""
        st.session_state.descricao = ""
        if "unidade" not in st.session_state or not st.session_state.unidade:
            st.session_state.unidade = ""

    st.write("Ou preencha manualmente o Nome e Unidade se não estiver listado:")
    descricao_livre = st.text_input("Nome do insumo (livre)", key="descricao_livre", disabled=usando_base)
    st.text_input("Código do insumo", key="codigo", disabled=True)
    st.text_input("Unidade", key="unidade", disabled=usando_base)
    quantidade = st.number_input("Quantidade", min_value=1, step=1, format="%d", key="quantidade")
    complemento = st.text_area("Complemento, se necessário (Utilize para especificar medidas, marcas, cores e/ou tamanhos)", key="complemento")

    if st.button("➕ Adicionar insumo"):
        descricao_final = st.session_state.descricao if usando_base else descricao_livre
        if descricao_final and quantidade > 0 and (usando_base or st.session_state.unidade.strip()):
            novo_insumo = {
                "descricao": descricao_final,
                "codigo": st.session_state.codigo if usando_base else "",
                "unidade": st.session_state.unidade,
                "quantidade": quantidade,
                "complemento": complemento,
            }
            st.session_state.insumos.append(novo_insumo)
            st.session_state.limpar_insumo = True  # marca flag para próxima renderização
            st.success("Insumo adicionado com sucesso!")
            st.rerun()
        else:
            st.warning("⚠️ Preencha todos os campos obrigatórios do insumo.")

# --- Renderiza tabela de insumos ---
if st.session_state.insumos:
    st.markdown("""
        <style>
        /* Cabeçalhos */
        .tabela-header {
            font-weight: 600;
            color: #333;
            border-bottom: 2px solid #ccc;
            padding-bottom: 2px;
            margin-bottom: 2px;
            font-size: 15px;
            display: flex;
            align-items: center;
            justify-content: flex-start;
        }
        .tabela-header.center {
            justify-content: center;  /* centraliza só Qtd */
        }

        /* Linhas */
        .linha-insumo {
            border-bottom: 1px solid #e6e6e6;
            padding: 3px 0;
            font-size: 14px;
            line-height: 1.4;
            display: flex;
            align-items: center;
        }
        .center {
            justify-content: center;
            text-align: center;
        }

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

    # Cabeçalho
    col1, col2, col3, col4 = st.columns([5.8, 1.2, 1.2, 0.5])
    with col1:
        st.markdown("<div class='tabela-header'>Insumos Adicionados</div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<div class='tabela-header center'>Qtd</div>", unsafe_allow_html=True)
    with col3:
        st.markdown("<div class='tabela-header'>Unid</div>", unsafe_allow_html=True)
    with col4:
        st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    # Linhas
    for i, insumo in enumerate(st.session_state.insumos):
        col1, col2, col3, col4 = st.columns([5.8, 1.2, 1.2, 0.5])
        with col1:
            st.markdown(f"<div class='linha-insumo'>{insumo['descricao']}</div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='linha-insumo center'>{insumo['quantidade']}</div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div class='linha-insumo'>{insumo['unidade']}</div>", unsafe_allow_html=True)
        with col4:
            if st.button("🗑️", key=f"delete_{i}"):
                st.session_state.insumos.pop(i)
                st.rerun()
    
# --- Finalização do Pedido ---
if st.button("📤 Enviar Pedido", use_container_width=True):
    campos_obrigatorios = [
        st.session_state.pedido_numero,
        st.session_state.data_pedido,
        st.session_state.solicitante,
        st.session_state.executivo,
        st.session_state.obra_selecionada,
        st.session_state.cnpj,
        st.session_state.endereco,
        st.session_state.cep
    ]

    if not all(campos_obrigatorios):
        st.warning("⚠️ Preencha todos os campos obrigatórios antes de enviar o pedido.")
        st.stop()

    if not st.session_state.insumos:
        st.warning("⚠️ Adicione pelo menos um insumo antes de enviar o pedido.")
        st.stop()

    erro = None
    ok = False
    
    with st.spinner("Enviando pedido e gerando arquivo... Aguarde!"):
        try:
            caminho_modelo = "Modelo_Pedido.xlsx"
            wb = load_workbook(caminho_modelo)
            ws = wb["Pedido"]
    
            ws["F2"] = st.session_state.pedido_numero
            ws["C3"] = st.session_state.data_pedido.strftime("%d/%m/%Y")
            ws["C4"] = st.session_state.solicitante
            ws["C5"] = st.session_state.executivo
            ws["C7"] = st.session_state.obra_selecionada
            ws["C8"] = st.session_state.cnpj
            ws["C9"] = st.session_state.endereco
            ws["C10"] = st.session_state.cep
    
            linha = 13
            for insumo in st.session_state.insumos:
                ws[f"B{linha}"] = insumo["codigo"]
                ws[f"C{linha}"] = insumo["descricao"]
                ws[f"D{linha}"] = insumo["unidade"]
                ws[f"E{linha}"] = insumo["quantidade"]
                ws[f"F{linha}"] = insumo["complemento"]
                linha += 1
    
            ultima_linha_util = linha - 1
            ws.print_area = f"A1:F{ultima_linha_util}"
    
            from io import BytesIO
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            excel_bytes = buffer.read()
    
            st.session_state.excel_bytes = excel_bytes
            st.session_state.nome_arquivo = f"Pedido{st.session_state.pedido_numero} OC {st.session_state.obra_selecionada}.xlsx"
    
            # Envia e-mail
            enviar_email_pedido(
                f"Pedido{st.session_state.pedido_numero} OC {st.session_state.obra_selecionada}",
                st.session_state.excel_bytes,
                st.session_state.insumos,
                df_insumos
            )

            ok = True

        except Exception as e:
            erro = str(e)

    if ok:
        st.session_state.pedido_enviado = True
        st.success("✅ Pedido gerado e e-mail enviado com sucesso! Agora você pode baixar o arquivo Excel abaixo ⬇️")

    elif erro:
        st.error(f"❌ Erro ao gerar pedido: {erro}")

# --- Botões após envio ---
col1, col2 = st.columns(2)

with col1:
    if st.download_button(
        "📥 Baixar Excel",
        data=st.session_state.excel_bytes,
        file_name=st.session_state.nome_arquivo or "Pedido.xlsx",
        use_container_width=True
    ):
        # 🔹 Limpa imediatamente todos os campos do formulário (removendo as chaves)
        for campo in [
            "pedido_numero", "solicitante", "executivo", "obra_selecionada",
            "cnpj", "endereco", "cep", "data_pedido",
            "excel_bytes", "nome_arquivo", "pedido_enviado"
        ]:
            if campo in st.session_state:
                del st.session_state[campo]

        st.session_state.insumos = []
        st.success("🧹 Formulário limpo após download! Pronto para novo pedido.")

        st.rerun()

with col2:
    if st.button("🔄 Novo Pedido", use_container_width=True):
        # 🔹 Limpa imediatamente todos os campos (mesma lógica)
        for campo in [
            "pedido_numero", "solicitante", "executivo", "obra_selecionada",
            "cnpj", "endereco", "cep", "data_pedido",
            "excel_bytes", "nome_arquivo", "pedido_enviado"
        ]:
            if campo in st.session_state:
                del st.session_state[campo]

        st.session_state.insumos = []
        st.success("🧹 Formulário limpo e pronto para novo pedido!")

        st.rerun()
        
# --- 🔄 Keep-alive (mover para o fim do arquivo) ---
st.components.v1.html(
    """
    <script>
      setInterval(() => { fetch(window.location.pathname + '_stcore/health'); }, 120000);
    </script>
    """,
    height=0,
)
