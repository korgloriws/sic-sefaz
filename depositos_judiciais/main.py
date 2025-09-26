import streamlit as st
import pandas as pd
import re
from fpdf import FPDF
import io

def extract_segment_b_info(line_b: str):
    DATA_TAM = 24  
    if len(line_b) < 50 or not line_b.startswith("B"):
        return None, 0.0

    line_no_seg = line_b[1:]
    data_part = line_no_seg[-DATA_TAM:]
    first_part = line_no_seg[:-DATA_TAM]

    if len(first_part) < 16:
        return None, 0.0

    codigo_pagamento = first_part[:16]
    middle_chunk = first_part[16:]

    if len(middle_chunk) <= 17:
        return codigo_pagamento, 0.0

    valor_chunk = middle_chunk[:-17]
    match_9 = re.search(r"(\d{9})$", valor_chunk)
    if not match_9:
        return codigo_pagamento, 0.0

    val_9 = match_9.group(1)
    parte_inteira = val_9[:-2].lstrip("0") or "0"
    parte_decimal = val_9[-2:]

    valor_float = float(parte_inteira + "." + parte_decimal)
    return codigo_pagamento, valor_float

def parse_ret_file(file_content: str):
    lines = file_content.splitlines()
    registros = []
    current_record = None

    for line in lines:
        line = line.strip()

        if line.startswith("A"):
            if current_record is not None:
                registros.append(current_record)

            numero_processo = line[1:21]
            idx_1633 = line.find("1633")
            creditors_str = line[idx_1633 + 4:].strip() if idx_1633 != -1 else ""

            parts = re.split(r"(\d{11,14})", creditors_str)
            pares = [(parts[i].strip(), parts[i + 1].strip()) for i in range(0, len(parts) - 1, 2)]

            credor1_nome, credor1_doc = pares[0] if len(pares) >= 1 else ("", "")
            credor2_nome, credor2_doc = pares[1] if len(pares) >= 2 else ("", "")

            current_record = {
                "numero_processo": numero_processo,
                "credor1_nome": credor1_nome,
                "credor1_doc": credor1_doc,
                "credor2_nome": credor2_nome,
                "credor2_doc": credor2_doc,
                "codigo_pagamento": "",
                "valor": 0.0
            }

        elif line.startswith("B") and current_record is not None:
            codigo_pagamento, valor_float = extract_segment_b_info(line)
            if codigo_pagamento is None:
                continue

            if current_record["codigo_pagamento"] and current_record["codigo_pagamento"] != codigo_pagamento:
                registros.append(current_record)
                current_record = {**current_record, "codigo_pagamento": codigo_pagamento, "valor": 0.0}

            current_record["codigo_pagamento"] = codigo_pagamento
            current_record["valor"] += valor_float

    if current_record is not None:
        registros.append(current_record)

    for reg in registros:
        reg["valor"] = f"{reg['valor']:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

    return registros

def renomear_colunas(df):
    return df.rename(columns={
        "numero_processo": "Nº PROCESSO TJMG",
        "credor1_nome": "POLO ATIVO",
        "credor1_doc": "CPF/CNPJ_ATIVO",
        "credor2_nome": "POLO PASSIVO",
        "credor2_doc": "CPF/CNPJ_PASSIVO",
        "codigo_pagamento": "CODIGO_PAGAMENTO",
        "valor": " VALOR PAGO "
    })

def to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Relatório RET")
        workbook = writer.book
        worksheet = writer.sheets["Relatório RET"]
        last_row = len(df) + 2
        worksheet.cell(row=last_row, column=1, value="REFERÊNCIA:")
        worksheet.cell(row=last_row + 1, column=1, value="5RESGATES CONTRA O GOVERNO")
    return output.getvalue()

def to_pdf(df: pd.DataFrame) -> bytes:
    pdf = FPDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Relatório do Arquivo Depósito Judicial", ln=True, align="C")
    pdf.ln(10)

    pdf.set_font("Arial", "B", 7)
    col_widths = [35, 60, 30, 60, 30, 40, 30]
    headers = df.columns.tolist()

    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 7, header, border=1, align='C')
    pdf.ln()

    pdf.set_font("Arial", "", 7)
    for _, row in df.iterrows():
        for i, col in enumerate(headers):
            pdf.cell(col_widths[i], 7, str(row[col]), border=1)
        pdf.ln()

    pdf.ln(5)
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 10, "REFERÊNCIA:", ln=True)
    pdf.cell(0, 5, "5RESGATES CONTRA O GOVERNO", ln=True)

    result = pdf.output(dest="S")
    if isinstance(result, (bytearray, bytes)):
        return bytes(result)
    return result.encode("latin-1")

def main():
    st.title("Depósitos Judiciais")

    uploaded_file = st.file_uploader("Selecione um arquivo .RET", type=["ret", "RET"])
    if uploaded_file:
        file_content = uploaded_file.read().decode("latin-1", errors="ignore")
        registros = parse_ret_file(file_content)

        if registros:
            df = renomear_colunas(pd.DataFrame(registros))
            df[" VALOR PAGO "] = df[" VALOR PAGO "].str.replace('.', '').str.replace(',', '.').astype(float)
            total = df[" VALOR PAGO "].sum()
            df[" VALOR PAGO "] = df[" VALOR PAGO "].apply(lambda x: f"{x:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))
            df.loc[len(df)] = ["", "TOTAL", "", "", "", "", f"{total:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")]
            st.dataframe(df)

            st.download_button("Baixar XLSX", to_excel(df), "relatorio_ret.xlsx")
            st.download_button("Baixar PDF", to_pdf(df), "relatorio_ret.pdf")
        else:
            st.warning("Não foram encontrados registros no arquivo.")

if __name__ == "__main__":
    main()
