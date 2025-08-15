import pandas as pd
import streamlit as st

st.title("Conversor Excel → TXT")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None, dtype=str, engine='openpyxl')
    coluna = st.number_input("Índice da coluna", min_value=0, value=0)
    prefixo = st.text_input("Texto antes", "")
    sufixo = st.text_input("Texto depois", "")
    separador = st.selectbox("Separador", [';', ',', '.', '/', '-'])
    filename = st.text_input("Nome do arquivo", "saida") + ".txt"

    if st.button("Converter para TXT"):
        valores = df.iloc[:, coluna].dropna().apply(lambda x: f'{prefixo}{str(x).strip()}{sufixo}').tolist()
        linha_unica = separador.join(valores)

        st.download_button("Baixar TXT", linha_unica, file_name=filename, mime="text/plain")
