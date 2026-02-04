import streamlit as st
import pandas as pd
import os
import zipfile
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.title("📄 Gerador de Contratos em PDF")

arquivo = st.file_uploader("Envie um CSV ou Excel", type=["csv", "xlsx"])
lote = st.selectbox("Tamanho do lote", [1, 10, 30])

def gerar_pdf(nome, cpf, caminho):
    doc = SimpleDocTemplate(caminho, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("TERMO DE CONTRATO EDUCACIONAL", styles["Title"]))
    story.append(Spacer(1, 20))

    texto = f"Eu, {nome}, inscrito no CPF nº {cpf}, declaro estar de acordo com os termos do contrato firmado com a instituição."
    story.append(Paragraph(texto, styles["Normal"]))
    story.append(Spacer(1, 30))

    story.append(Paragraph("Assinatura: ____________________________", styles["Normal"]))

    doc.build(story)

if arquivo:
    if arquivo.name.endswith(".csv"):
        df = pd.read_csv(arquivo)
    else:
        df = pd.read_excel(arquivo)

    st.dataframe(df)

    if st.button("Gerar contratos"):
        pasta = "pdfs"
        os.makedirs(pasta, exist_ok=True)

        for i in range(0, len(df), lote):
            grupo = df.iloc[i:i+lote]

            for _, row in grupo.iterrows():
                nome = row["Nome"]
                cpf = row["CPF"]
                caminho_pdf = os.path.join(pasta, f"{nome}_{cpf}.pdf")
                gerar_pdf(nome, cpf, caminho_pdf)

        zip_path = "contratos.zip"
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in os.listdir(pasta):
                zipf.write(os.path.join(pasta, file), arcname=file)

        with open(zip_path, "rb") as f:
            st.download_button("⬇️ Baixar ZIP", f, file_name="contratos.zip")
