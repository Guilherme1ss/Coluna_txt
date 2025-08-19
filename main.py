import pandas as pd
import streamlit as st
import re
from io import StringIO, BytesIO

st.title("Conversor Excel/Text → TXT com Normalização e Limpeza")

# Escolha da fonte de dados
fonte = st.radio("Fonte dos dados:", ("Arquivo Excel", "Colar texto manualmente"))

if fonte == "Arquivo Excel":
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "xls"])
    if uploaded_file:
        # Carregar todas as planilhas do arquivo Excel
        xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_names = xl.sheet_names  # Obter nomes das planilhas
        selected_sheet = st.selectbox("Selecione a planilha:", sheet_names)
        
        # Carregar a planilha selecionada
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, dtype=str, engine="openpyxl")
        
        # Obter nomes das colunas (assumindo que a primeira linha contém os cabeçalhos)
        columns = df.columns.tolist()
        selected_columns = st.multiselect("Selecione as colunas:", columns)
        
        if selected_columns:
            # Combinar os dados das colunas selecionadas em uma única lista
            dados = []
            for col in selected_columns:
                dados.extend(df[col].dropna().astype(str).tolist())
        else:
            dados = []
            st.warning("Nenhuma coluna selecionada.")
    else:
        dados = []
        st.warning("Nenhum arquivo carregado.")
else:
    texto_colado = st.text_area("Cole os dados aqui (um valor por linha)")
    if texto_colado.strip():
        dados = [linha.strip() for linha in StringIO(texto_colado).read().splitlines() if linha.strip()]
    else:
        dados = []

prefixo = st.text_input("Texto antes", "")
sufixo = st.text_input("Texto depois", "")
separador = st.selectbox("Separador", [';', ',', '.', '/', '-'])

# Opção de normalização
normalizacao = st.radio(
    "Normalização do texto",
    ("Não alterar", "Tudo maiúsculo", "Tudo minúsculo")
)

# Opção de remover caracteres especiais
st.subheader("Remover caracteres especiais")
remove_hifen = st.checkbox("Remover '-'")
remove_virgula = st.checkbox("Remover ','")
remove_ponto = st.checkbox("Remover '.'")

filename_txt = st.text_input("Nome do arquivo .txt", "Resultado") + ".txt"
filename_excel = filename_txt.replace(".txt", ".xlsx")

if st.button("Converter"):
    if not dados:
        st.error("Nenhum dado para converter.")
    else:
        valores = pd.Series(dados, dtype=str)

        # Remover caracteres especiais
        for char, remove in [("-", remove_hifen), (",", remove_virgula), (".", remove_ponto)]:
            if remove:
                valores = valores.str.replace(re.escape(char), "", regex=True)

        # Normalização
        if normalizacao == "Tudo maiúsculo":
            valores = valores.str.upper()
        elif normalizacao == "Tudo minúsculo":
            valores = valores.str.lower()

        # Adicionar prefixo e sufixo
        valores = prefixo + valores + sufixo

        # Criar string final
        linha_unica = separador.join(valores)

        # Mostrar saída em caixa de texto
        st.subheader("Resultado")
        st.text_area("Saída:", linha_unica, height=200)

        # Botão para baixar TXT
        st.download_button(
            "Baixar como .txt",
            linha_unica,
            file_name=filename_txt,
            mime="text/plain"
        )