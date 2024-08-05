import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors


file_path = r'C:\Users\samuel.grillo\Downloads\PN_Mangueira.xlsx'


df = pd.read_excel(file_path)


df.columns = df.columns.str.strip()


def create_label(os, part_number, df):
    if not part_number.isdigit():
        st.error("Part Number deve ser um número válido!")
        return
    
  
    row = df[df['Part Number'] == int(part_number)]
    
    if row.empty:
        st.error("Part Number não encontrado!")
        return


    part_number = row['Part Number'].values[0]
    diametro = row['Diâmetro'].values[0]
    comprimento = row['Comprimento'].values[0]
    dimensao_terminais = row['Dimensão terminais'].values[0]
    pressao_trabalho = row['Pressão de trabalho'].values[0]
    vedacao_terminais = row['Vedação Terminais'].values[0]
    norma = row['NORMA'].values[0]
    

    pdf_path = f"{part_number}_etiqueta.pdf"
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter


    x_margin = 20
    y_start = width - 60  
    line_height = 20

 
    headers = ["OS", "Part Number", "Diâmetro", "Comprimento", "Dimensão terminais", "Pressão de trabalho", "Vedação Terminais", "Norma"]
    values = [os, part_number, diametro, comprimento, dimensao_terminais, pressao_trabalho, vedacao_terminais, norma]


    c.setFont("Helvetica-Bold", 10)
    col_widths = []
    for i, header in enumerate(headers):
        header_width = c.stringWidth(header, "Helvetica-Bold", 10)
        value_width = c.stringWidth(str(values[i]), "Helvetica", 10)
        col_widths.append(max(header_width, value_width) + 10)  


    c.saveState()
    c.translate(0, height)  
    c.rotate(-90) 


    c.setFillColor(colors.orange)
    c.rect(x_margin, y_start - line_height, sum(col_widths), line_height, fill=1)
    c.setFillColor(colors.black)
    
    x_pos = x_margin
    for i, header in enumerate(headers):
        text_width = c.stringWidth(header, "Helvetica-Bold", 10)
        c.drawString(x_pos + (col_widths[i] - text_width) / 2, y_start - 15, header)
        x_pos += col_widths[i]


    y_start -= line_height
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.white)
    c.rect(x_margin, y_start - line_height, sum(col_widths), line_height, fill=1)
    c.setFillColor(colors.black)
    
    x_pos = x_margin
    for i, value in enumerate(values):
        text_width = c.stringWidth(str(value), "Helvetica", 10)
        c.drawString(x_pos + (col_widths[i] - text_width) / 2, y_start - 15, str(value))
        x_pos += col_widths[i]

    c.restoreState()
    c.save()
    st.success(f"Etiqueta gerada: {pdf_path}")
    st.download_button(
        label="Download Etiqueta",
        data=open(pdf_path, "rb").read(),
        file_name=pdf_path,
        mime="application/pdf"
    )

st.title("Gerador de Etiquetas de Mangueiras")
st.header("Insira as Informações")

os = st.text_input("OS")
part_number = st.text_input("Part Number")

if st.button("Gerar Etiqueta"):
    create_label(os, part_number, df)
