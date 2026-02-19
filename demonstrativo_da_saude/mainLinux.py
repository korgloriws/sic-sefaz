from openpyxl import load_workbook
import streamlit as st
import pandas as pd
from io import StringIO

def process_qgr(file):
    df = pd.read_csv(file, encoding='latin1', delimiter=';')
    df['CODIGO_RECEITA'] = df['CODIGO_RECEITA'].astype(str)
    df['VR_ARREC_ATE_MES_FONTE'] = df['VR_ARREC_ATE_MES_FONTE'].str.replace('.', '').str.replace(',', '.').astype(float)
    resultado = df.groupby('CODIGO_RECEITA')['VR_ARREC_ATE_MES_FONTE'].sum()
    return resultado.reset_index()

def process_ded(file):
    df = pd.read_csv(file, encoding='latin1', delimiter=';')
    df['fonte'] = df['fonte'].astype(str)
    df['liq_ate_mes'] = df['liq_ate_mes'].str.replace('.', '').str.replace(',', '.').astype(float)
    fontes_especificas = ['1500702', '2500702', '31500702', '51500702']
    df_filtrado = df[df['fonte'].isin(fontes_especificas)]
    resultado = df_filtrado['liq_ate_mes'].sum()
    return resultado

def fill_demostrativo_saude(file_path, qgr_data=None, ded_data=None):
    
    workbook = load_workbook(filename=file_path)
    sheet = workbook.worksheets[1]  

    soma_valores_especificos = 0

    if qgr_data is not None:
        for row in qgr_data.itertuples():
            cell_map = {
                '1112500100': 'G8', '1112530100': 'G9', '1113031100': 'G10', '1113034100': 'G11',
                '1114511100': 'G12', '1711511100': 'G17', '1711520100': 'G18', '1721500100': 'G19',
                '1721510100': 'G20', '1721520100': 'G21', '1115010000': 'G22', '1112500200': 'G28',
                '1112500300': 'G29', '1112500400': 'G30', '1112530200': 'G31', '1112530300': 'G32',
                '1112530400': 'G33', '1114511200': 'G34', '1114511300': 'G35', '1114511400': 'G36'
            }
            cell = cell_map.get(row.CODIGO_RECEITA)
            if cell:
                sheet[cell].value = row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA in ['', '9111125001', '9111125002', '9111125003', 
                                      '9111125004', '9111125301', '9111145111', '9111145112', 
                                      '9111145113', '9111145114', '9211125001', '9211125002', 
                                      '9211125003', '9211125004', '9211125301', '9211130341', 
                                      '9211145111', '9211145112', '9311125001', '9311125002', 
                                      '9311125003', '9311125004', '9311125301', '9311125302', 
                                      '9311145111', '9311145112', '9311145113', '9311145114', 
                                      '9111145111','9211130311']:
                if row.VR_ARREC_ATE_MES_FONTE != 0:
                    soma_valores_especificos += row.VR_ARREC_ATE_MES_FONTE

    sheet['G39'].value = soma_valores_especificos

    if ded_data is not None:
        sheet['G42'].value = ded_data

   
    workbook.save(filename=file_path)
    workbook.close()

def main():
    st.title("Demonstrativo de Saúde")

   
    uploaded_qgr = st.file_uploader("Carregue o arquivo QGR", type=['csv'])
    uploaded_ded = st.file_uploader("Carregue o arquivo DED", type=['csv'])
    uploaded_demostrativo_saude = st.file_uploader("Carregue o arquivo Demonstrativo Saúde", type=['xlsx'])

    if st.button('Processar Arquivos'):
        if uploaded_qgr is not None and uploaded_ded is not None and uploaded_demostrativo_saude is not None:
            try:
               
                qgr_data = process_qgr(uploaded_qgr)
                ded_data = process_ded(uploaded_ded)

                
                temp_file_path = "temp_DEMONSTRATIVO_SAUDE.xlsx"
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_demostrativo_saude.getbuffer())

                
                fill_demostrativo_saude(temp_file_path, qgr_data, ded_data)

                
                with open(temp_file_path, "rb") as f:
                    st.download_button(
                        label="Baixar Demonstrativo Saúde processado",
                        data=f,
                        file_name="DEMONSTRATIVO_SAUDE_processado.xlsx",
                        mime="application/vnd.ms-excel"
                    )

                st.success("Processamento concluído! Baixe o arquivo processado abaixo.")

            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")

        else:
            st.error("Por favor, carregue todos os arquivos necessários.")

if __name__ == "__main__":
    main()
