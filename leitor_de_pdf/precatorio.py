import re
import locale
import pandas as pd
import datetime
import pdfplumber
import streamlit as st
import base64
import io
from fpdf import FPDF  


try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'pt_BR')
    except locale.Error:
        locale.setlocale(locale.LC_ALL, '')

def process_pdf(file):
    sum_resgate = 0
    sum_aplicacao = 0
    sum_rendimento = 0
    sum_estorno_de_re = 0

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                line = re.sub(r'\s+', ' ', line)

                if re.search(r'resgate,', line, re.IGNORECASE):
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        valor = locale.atof(match.group(1))
                        sum_resgate += valor

                if re.search(r'aplicacao', line, re.IGNORECASE):
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        valor = locale.atof(match.group(1))
                        sum_aplicacao += valor

                if re.search(r'rendimento', line, re.IGNORECASE):
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        valor = locale.atof(match.group(1))
                        sum_rendimento += valor

                if re.search(r'estorno de re', line, re.IGNORECASE):
                    match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', line)
                    if match:
                        valor = locale.atof(match.group(1))
                        sum_estorno_de_re += valor

    data = {
        'Tipo': ['Resgate', 'Aplicacao', 'Rendimento', 'Estorno de RE'],
        'Total': [sum_resgate, sum_aplicacao, sum_rendimento, sum_estorno_de_re]
    }
    df = pd.DataFrame(data)
    df['Total'] = df['Total'].apply(lambda x: locale.currency(x, grouping=True))
    return df

def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    worksheet.set_column('B:B', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data


def to_pdf(df: pd.DataFrame) -> bytes:

    pdf = FPDF()  
    pdf.add_page()
    

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Relatório - Precatórios", ln=True, align="C")
    pdf.ln(10)

 
    pdf.set_font("Arial", "B", 10)
    col_widths = [50, 50]  
    headers = df.columns.tolist()
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 8, header, border=1, align="C")
    pdf.ln()


    pdf.set_font("Arial", "", 10)
    for _, row in df.iterrows():
        pdf.cell(col_widths[0], 8, str(row["Tipo"]), border=1)
        pdf.cell(col_widths[1], 8, str(row["Total"]), border=1)
        pdf.ln()

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1", "replace")
    if isinstance(out, (bytes, bytearray)):
        return bytes(out)
    # Fallback: ensure bytes
    return bytes(out)
    
# -------------------------------------------------------------------------

def main():
    st.title('Precatórios')

    file = st.file_uploader('Carregar arquivo PDF', type=['pdf'])
    if file is not None:
        filename = st.text_input("Digite o nome do arquivo para salvar: ")
        df = process_pdf(file)
        if st.button('Processar'):
            if filename != "":
                date_str = datetime.datetime.now().strftime("%d/%m/%Y")
                st.markdown(f"**Nome do arquivo:** {filename}")
                st.markdown(f"**Data de geração:** {date_str}")
                st.write(df)
                st.success(f'Arquivo gerado com sucesso: {filename}')
                
         
                excel_data = to_excel(df)
                st.download_button(
                    label="Download Excel file",
                    data=excel_data,
                    file_name=f"{filename}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
   
                pdf_data = to_pdf(df)
                st.download_button(
                    label="Download PDF file",
                    data=pdf_data,
                    file_name=f"{filename}.pdf",
                    mime='application/pdf'
                )
        

            else:
                st.error("Por favor, insira um nome para o arquivo antes de processar.")

if __name__ == '__main__':
    main()
