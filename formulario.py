import streamlit as st
from datetime import date
from openpyxl import load_workbook
import pandas as pd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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

# --- Funções auxiliares ---
def resetar_campos_insumo():
    st.session_state.resetar_insumo = True

def resetar_formulario():
    st.session_state.resetar_pedido = True
    resetar_campos_insumo()
    st.session_state.insumos = []

def registrar_historico(numero, obra, data):
    historico_path = "historico_pedidos.csv"
    registro = {"numero": str(numero).strip(), "obra": str(obra).strip(), "data": data.strftime("%Y-%m-%d")}
    if os.path.exists(historico_path):
        df_hist = pd.read_csv(historico_path, dtype=str)
        if not ((df_hist["numero"] == registro["numero"]) & (df_hist["obra"] == registro["obra"])).any():
            df_hist = pd.concat([df_hist, pd.DataFrame([registro])], ignore_index=True)
            df_hist.to_csv(historico_path, index=False, encoding="utf-8")
        else:
            pass  # Já registrado, não faz nada
    else:
        df_hist = pd.DataFrame([registro])
        df_hist.to_csv(historico_path, index=False, encoding="utf-8")

# --- Função para enviar e-mail ---
def enviar_email_pedido(assunto, arquivo_bytes, insumos_adicionados, df_insumos):
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    import smtplib

    smtp_server = "smtp.office365.com"
    smtp_port = 587
    smtp_user = "matheus.almeida@osborne.com.br"
    smtp_password = "mnmhshjjvmyqnddr"

    # Separa básicos e específicos
    basicos = []
    especificos = []

    for item in insumos_adicionados:
        linha_df = df_insumos[df_insumos["Descrição"] == item["descricao"]]
        if not linha_df.empty and linha_df.iloc[0]["Basico"]:
            min_qtd = linha_df.iloc[0]["Min"]
            max_qtd = linha_df.iloc[0]["Max"]
            qtd = item["quantidade"]

            if pd.notna(min_qtd) and pd.notna(max_qtd) and min_qtd <= qtd <= max_qtd:
                basicos.append(f"{item['descricao']} — {qtd}")
            else:
                especificos.append(f"{item['descricao']} — {qtd}")
        else:
            especificos.append(f"{item['descricao']} — {qtd}")

    corpo = "✅ Novo pedido recebido!\n\n"
    corpo += "📄 Materiais Básicos:\n"
    corpo += "\n".join(basicos) if basicos else "Nenhum\n"
    corpo += "\n\n🛠️ Materiais Específicos:\n"
    corpo += "\n".join(especificos) if especificos else "Nenhum"

    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = smtp_user
    msg["Subject"] = assunto

    msg.attach(MIMEText(corpo, "plain"))

    from email.mime.base import MIMEBase
    from email import encoders

    part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    part.set_payload(arquivo_bytes)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment')
    msg.attach(part)

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
    df_insumos = pd.read_excel("Insumos.xlsx")

    # Marca os 15 primeiros como "Basico" (sem contar cabeçalho)
    df_insumos["Basico"] = False
    df_insumos.loc[:14, "Basico"] = True  # 15 primeiras linhas

    # Carrega min e max (colunas D e E)
    df_insumos["Min"] = pd.to_numeric(df_insumos.iloc[:, 3], errors="coerce")
    df_insumos["Max"] = pd.to_numeric(df_insumos.iloc[:, 4], errors="coerce")

    df_insumos = df_insumos[df_insumos["Descrição"].notna() & (df_insumos["Descrição"].str.strip() != "")]

    df_empreend.loc[-1] = ["", "", "", ""]
    df_empreend.index = df_empreend.index + 1
    df_empreend = df_empreend.sort_index()

    insumos_vazios = pd.DataFrame({"Código": [""], "Descrição": [""], "Unidade": [""]})
    df_insumos = pd.concat([insumos_vazios, df_insumos], ignore_index=True)

    return df_empreend, df_insumos

# --- Dados ---
df_empreend, df_insumos = carregar_dados()

# --- Logo e título ---
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("logo.png", width=300)

st.markdown("""
    <div style='text-align: center;'>
        <h2 style='color: #000000;'>Pedido de Materiais Básicos</h2>
        <p style='font-size: 14px; color: #555;'>Preencha os campos com atenção. Evite abreviações desnecessárias.<br>
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
        obra_selecionada = st.selectbox("Obra", df_empreend["NOME"].unique(), index=0, key="obra_selecionada")
    with col2:
        data_pedido = st.date_input("Data", value=st.session_state.get("data_pedido", date.today()), key="data_pedido")
        executivo = st.text_input("Executivo", key="executivo")

    if obra_selecionada:
        dados_obra = df_empreend[df_empreend["NOME"] == obra_selecionada].iloc[0]
        st.session_state.cnpj = dados_obra["EMPRD_CNPJFAT"]
        st.session_state.endereco = dados_obra["ENDEREÇO"]
        st.session_state.cep = dados_obra["Cep"]

    st.text_input("CNPJ/CPF", value=st.session_state.get("cnpj", ""), disabled=True)
    st.text_input("Endereço", value=st.session_state.get("endereco", ""), disabled=True)
    st.text_input("CEP", value=st.session_state.get("cep", ""), disabled=True)

st.divider()

# --- Adição de Insumos ---
with st.expander("➕ Adicionar Insumo", expanded=True):
    if st.session_state.resetar_insumo:
        st.session_state.descricao = ""
        st.session_state.descricao_livre = ""
        st.session_state.codigo = ""
        st.session_state.unidade = ""
        st.session_state.quantidade = 1  # 👈 Já inicializa com 1
        st.session_state.complemento = ""
        st.session_state.resetar_insumo = False

    # Ordena a lista de insumos
    df_insumos_lista = df_insumos.sort_values(by="Descrição", ascending=True).copy()
    lista_opcoes = df_insumos_lista["Descrição"].tolist()

    descricao = st.selectbox("Descrição do insumo (Digite em MAIÚSCULO)", lista_opcoes, key="descricao")

    usando_base = bool(descricao)

    if usando_base:
        dados_insumo = df_insumos_lista[df_insumos_lista["Descrição"] == descricao].iloc[0]
        codigo = dados_insumo["Código"]
        unidade = dados_insumo["Unidade"]
    else:
        codigo = ""
        unidade = ""

    # Campo manual para descrição livre
    st.write("Ou preencha manualmente se não estiver listado:")
    descricao_livre = st.text_input("Nome do insumo (livre)", key="descricao_livre", disabled=usando_base)

    st.text_input("Código do insumo", value=codigo, key="codigo", disabled=True)
    unidade = st.text_input("Unidade", value=unidade, key="unidade", disabled=usando_base)

    quantidade = st.number_input("Quantidade", min_value=1, step=1, format="%d", key="quantidade")
    complemento = st.text_area("Complemento", key="complemento")

    if st.button("➕ Adicionar insumo"):
        descricao_final = descricao if usando_base else descricao_livre

        if descricao_final and quantidade > 0 and (usando_base or unidade.strip()):
            novo_insumo = {
                "descricao": descricao_final,
                "codigo": codigo if usando_base else "",
                "unidade": unidade,
                "quantidade": quantidade,
                "complemento": complemento,
            }
            st.session_state.insumos.append(novo_insumo)
            st.success("Insumo adicionado com sucesso!")
            resetar_campos_insumo()
            st.rerun()
        else:
            st.warning("⚠️ Preencha todos os campos obrigatórios do insumo.")

# --- Renderiza tabela de insumos ---
if st.session_state.insumos:
    st.subheader("📦 Insumos adicionados")
    for i, insumo in enumerate(st.session_state.insumos):
        cols = st.columns([6, 1])
        with cols[0]:
            st.markdown(f"**{i+1}.** {insumo['descricao']} — {insumo['quantidade']} {insumo['unidade']}")
        with cols[1]:
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

        # Após preencher, deletar linhas extras
        ultima_linha_util = linha - 1
        total_linhas_modelo = 112  # ajuste aqui conforme teu arquivo

        if ultima_linha_util < total_linhas_modelo:
            ws.delete_rows(ultima_linha_util + 1, total_linhas_modelo - ultima_linha_util)

        nome_saida = f"Pedido{st.session_state.pedido_numero} OC {st.session_state.obra_selecionada}.xlsx"
        wb.save(nome_saida)

        with open(nome_saida, "rb") as f:
            excel_bytes = f.read()

        # Salva no estado
        st.session_state.excel_bytes = excel_bytes
        st.session_state.nome_arquivo = nome_saida

        st.success("✅ Pedido gerado e e-mail enviado com sucesso!")

        numero = st.session_state.pedido_numero
        obra = st.session_state.obra_selecionada
        data_pedido = st.session_state.data_pedido

        registrar_historico(numero, obra, data_pedido)

        # Gera assunto com o nome desejado
        assunto_email = f"Pedido{st.session_state.pedido_numero} OC {st.session_state.obra_selecionada}"
        
        # Envia e-mail com o mesmo arquivo
        enviar_email_pedido(
            assunto_email,
            st.session_state.excel_bytes,
            st.session_state.insumos,
            df_insumos
        )
    except Exception as e:
        st.error(f"Erro ao gerar pedido: {e}")

# --- Botão de download separado ---
if st.session_state.excel_bytes:
    if st.download_button("📥 Baixar Excel", data=st.session_state.excel_bytes, file_name=st.session_state.nome_arquivo, use_container_width=True):
        resetar_formulario()
        st.session_state.excel_bytes = None
        st.session_state.nome_arquivo = ""
        st.rerun()
