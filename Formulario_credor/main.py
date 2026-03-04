import xlwings as xw
import pandas as pd
import os
import shutil
import streamlit as st
import tempfile
import zipfile

def brazilian_formatter(x):
    s = str(int(x * 100))
    if len(s) <= 2:
        return f"0,{s.zfill(2)}"
    else:
        int_part = s[:-2]
        decimal_part = s[-2:]
        int_part_with_dot = f"{int(int_part):,}".replace(",", ".")
        return f"{int_part_with_dot},{decimal_part}"

def sanitize_filename(name):
    """Remove caracteres inválidos do nome do arquivo (ex: / em 'ALGAR TELECOM S/A')."""
    if name is None:
        return "credor"
    s = str(name).strip()
    for char in r'/\:*?"<>|':
        s = s.replace(char, "-")
    return s.strip("-") or "credor"


def save_uploaded_file(uploaded_file):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            shutil.copyfileobj(uploaded_file, tmp_file)
            return tmp_file.name
    except Exception as e:
        st.error(f"Erro ao salvar o arquivo carregado: {e}")
        return None

def fill_and_save_form(group, name, payment_month, payment_date, template_file_path):
    try:
        safe_name = sanitize_filename(name)
        output_file_name = f'Preenchido_{safe_name}_Anexo_V - Comprovante de Retenção.xlsx'
        output_file_path = os.path.join(tempfile.gettempdir(), output_file_name)

        shutil.copy(template_file_path, output_file_path)

        wb = xw.Book(output_file_path)
        ws = wb.sheets['Plan1']

        ws.range('A9').value = group['cpf_cnpj'].iloc[0]
        ws.range('D9').value = name
        ws.range('F11').value = payment_month
        ws.range('D40').value = payment_date
        ws.range('A40').value = "Hemerson Fernandes Soares"

        start_row = 12
        group_sum = group.groupby('Código de receita').sum().reset_index()

        for i, (_, row) in enumerate(group_sum.iterrows()):
            row_number = start_row + i
            ws.range(f'C{row_number}').value = row['Código de receita']
            ws.range(f'E{row_number}').value = brazilian_formatter(row['valor_bruto'])
            ws.range(f'G{row_number}').value = brazilian_formatter(row['valor_des'])
            ws.range(f'A{row_number}').value = payment_month

        wb.save()
        wb.close()
        return output_file_path
    except Exception as e:
        st.error(f"Failed to process group {name}. Error: {e}")
        return None

def main():
    st.title("Processador de Formulários de Pagamento")

    uploaded_file = st.file_uploader("Escolha um arquivo Excel com os dados", type="xls")
    template_file = st.file_uploader("Escolha o arquivo de modelo do formulário", type="xlsx")
    payment_month = st.text_input("Insira o mês de pagamento:", "")
    payment_date = st.text_input("Insira a data do pagamento (dd/mm/aaaa):", "")
    file_name = st.text_input("Nome do arquivo para download:", "Preenchido_Formulario")

    process_button = st.button("Processar")

    if process_button and uploaded_file and template_file and payment_month and payment_date:
        template_file_path = save_uploaded_file(template_file)
        if template_file_path:
            source_df = pd.read_excel(uploaded_file)
            filtered_source_df = source_df[["credor", "cpf_cnpj", "valor_bruto", "valor_des", "Código de receita"]]
            grouped = filtered_source_df.groupby("credor")

            output_files = []
            for i, (name, group) in enumerate(grouped):
                st.write(f"Processando grupo {i+1}/{len(grouped)}: {name}")
                group = group.reset_index(drop=True)
                output_file = fill_and_save_form(group, name, payment_month, payment_date, template_file_path)
                if output_file:
                    output_files.append(output_file)

            if output_files:
                zip_file_path = os.path.join(tempfile.gettempdir(), f"{file_name}.zip")
                with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                    for file in output_files:
                        zipf.write(file, os.path.basename(file))
                        os.remove(file)  

                with open(zip_file_path, 'rb') as f:
                    st.download_button("Baixar Arquivos", f, file_name=f"{file_name}.zip")

if __name__ == "__main__":
    main()
