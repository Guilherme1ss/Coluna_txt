import pandas as pd
import streamlit as st
import re

st.title("Conversor Excel → TXT com Normalização e Limpeza")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, header=None, dtype=str, engine="openpyxl")

    coluna = st.number_input("Índice da coluna", min_value=0, value=0)
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

    filename = st.text_input("Nome do arquivo (sem extensão)", "saida") + ".txt"

    if st.button("Converter para TXT"):
        valores = df.iloc[:, coluna].dropna().astype(str)

        # Remover caracteres especiais selecionados
        for char, remove in [("-", remove_hifen), (",", remove_virgula), (".", remove_ponto)]:
            if remove:
                valores = valores.str.replace(re.escape(char), "", regex=True)

        # Normalização
        if normalizacao == "Tudo maiúsculo":
            valores = valores.str.upper()
        elif normalizacao == "Tudo minúsculo":
            valores = valores.str.lower()

        # Adicionar prefixo/sufixo
        valores = prefixo + valores + sufixo

        linha_unica = separador.join(valores)

        st.download_button(
            "Baixar TXT",
            linha_unica,
            file_name=filename,
            mime="text/plain"
        )
