import streamlit as st
import pandas as pd
import openpyxl
import os

def process_qgr(file_path):
    df = pd.read_csv(file_path, encoding='latin1', delimiter=';')
    df['CODIGO_RECEITA'] = df['CODIGO_RECEITA'].astype(str)
    df['FONTE_RECURSO'] = df['FONTE_RECURSO'].apply(lambda x: str(int(float(x))) if pd.notna(x) else x)
    df['VR_ARREC_ATE_MES_FONTE'] = df['VR_ARREC_ATE_MES_FONTE'].str.replace('.', '').str.replace(',', '.').astype(float)
    resultado = df.groupby(['CODIGO_RECEITA', 'FONTE_RECURSO'])['VR_ARREC_ATE_MES_FONTE'].sum().reset_index()
    return resultado

def process_ded(file_path):
    df = pd.read_csv(file_path, encoding='latin1', delimiter=';')
    df['fonte'] = df['fonte'].astype(str)
    df['liq_ate_mes'] = df['liq_ate_mes'].str.replace('.', '').str.replace(',', '.').astype(float)
    df['pag_ate_mes'] = pd.to_numeric(df['pag_ate_mes'].str.replace('.', '').str.replace(',', '.'), errors='coerce')
    fontes_especificas = ['1500701', '2500701', '31500701', '51500701', '2500701', '21540770', '22540770']
    df_filtrado = df[df['fonte'].isin(fontes_especificas)]
    resultado = df_filtrado.groupby('fonte')[['liq_ate_mes', 'pag_ate_mes']].sum()
    return resultado.reset_index()

def process_rpnp(file_path):
    df = pd.read_csv(file_path, encoding='latin1', delimiter=';')
    df['fonte'] = df['fonte'].astype(str)
    df['valor_anu_ant'] = df['valor_anu_ant'].str.replace('.', '').str.replace(',', '.').astype(float)
    df['valor_anu_mes'] = df['valor_anu_mes'].str.replace('.', '').str.replace(',', '.').astype(float) 
    fontes_especificas = ['1500701', '31500701', '51500701', '2500701', '32500701', '52500701']
    df_filtrado = df[df['fonte'].isin(fontes_especificas)]
    resultado = df_filtrado.groupby('fonte')[['valor_anu_ant', 'valor_anu_mes']].sum().reset_index()  
    return resultado


def process_rpp_1619(file_path):
    df = pd.read_csv(file_path, encoding='latin1', delimiter=';')
    df['fonte'] = df['fonte'].astype(str)
    df['valor_anu_ant'] = df['valor_anu_ant'].str.replace('.', '').str.replace(',', '.').astype(float)
    df['valor_anu_mes'] = df['valor_anu_mes'].str.replace('.', '').str.replace(',', '.').astype(float)
    fontes_especificas = ['1500701', '31500701', '51500701', '2500701', '32500701', '52500701']
    df_filtrado = df[df['fonte'].isin(fontes_especificas)]
    resultado = df_filtrado.groupby('fonte')[['valor_anu_ant', 'valor_anu_mes']].sum().reset_index()
    return resultado

def fill_MDE(file_path, qgr_data=None, ded_data=None, rpnp_data=None, rpp_1619_data=None):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.worksheets[1]

    soma_valores_G56 = 0
    soma_valores_G57 = 0
    soma_valores_G58 = 0
    soma_valores_G59 = 0
    soma_valores_G60 = 0
    soma_valores_G61 = 0
    soma_valores_G62 = 0
    soma_valores_G63 = 0
    soma_valores_G64 = 0
    soma_valores_G65 = 0
    soma_valores_G67 = 0
    soma_valores_G71 = 0
    soma_valores_especificos = 0
    

    if qgr_data is not None:
        for row in qgr_data.itertuples():
            if row.CODIGO_RECEITA in ['1112500100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G9'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112530100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G10'].value = soma_valores
            if row.CODIGO_RECEITA in ['1113031100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G11'].value = soma_valores
            if row.CODIGO_RECEITA in ['1113034100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G12'].value = soma_valores
            if row.CODIGO_RECEITA in ['1114511100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G13'].value = soma_valores
            if row.CODIGO_RECEITA in ['1711511100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G18'].value = soma_valores
            if row.CODIGO_RECEITA in ['1711513100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G19'].value = soma_valores
            if row.CODIGO_RECEITA in ['1711512100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G20'].value = soma_valores
            if row.CODIGO_RECEITA in ['1711520100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G21'].value = soma_valores
            if row.CODIGO_RECEITA in ['1719610100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G22'].value = soma_valores
            if row.CODIGO_RECEITA in ['1721500100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G23'].value = soma_valores
            if row.CODIGO_RECEITA in ['1721510100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G24'].value = soma_valores
            if row.CODIGO_RECEITA in ['1721520100']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G25'].value = soma_valores
            if row.CODIGO_RECEITA in ['1115010000']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G26'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112500200']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G32'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112500300']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G33'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112500400']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G34'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112530200']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G35'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112530300']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G36'].value = soma_valores
            if row.CODIGO_RECEITA in ['1112530400']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G37'].value = soma_valores
            if row.CODIGO_RECEITA in ['1114511200']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G38'].value = soma_valores
            if row.CODIGO_RECEITA in ['1114511300']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G39'].value = soma_valores
            if row.CODIGO_RECEITA in ['1114511400']:
                soma_valores = qgr_data[(qgr_data['CODIGO_RECEITA'] == row.CODIGO_RECEITA)]['VR_ARREC_ATE_MES_FONTE'].sum()
                sheet['G40'].value = soma_valores
            
            elif row.CODIGO_RECEITA in [
                '9111125001','9111125002','9111125003','9111125004','9111125301','9111145111',
                '9111145112','9111145113','9111145114','9211125001','9211125002','9211125003',
                '9211125004','9211125301','9211130341','9211145111','9211145112','9311125001',
                '9311125002','9311125003','9311125004','9311125301','9311125302','9311145111',
                '9311145112','9311145113','9311145114','9111145111','9111145131','9111145121',
                '9211130311'
                
            ]:
                soma_valores_especificos += row.VR_ARREC_ATE_MES_FONTE

            elif row.CODIGO_RECEITA in [
                '9500000000','9517115111','9517115201','9517215001','9517215101','9517215201' ,]:
                soma_valores_G56 += row.VR_ARREC_ATE_MES_FONTE

            elif row.CODIGO_RECEITA in ['1751000000','1751500000','1751500100','1751500100','9817515001']:
                soma_valores_G57 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '1321010100':
                if row.FONTE_RECURSO in ['21540770','21540000','31500701','51500701']:
                    soma_valores_G58 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '1922990100':
                if row.FONTE_RECURSO in ['21540770','21540000','31500701','51500701']:
                    soma_valores_G59 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '1922510100':
                if row.FONTE_RECURSO in ['21540770','21540000','31500701','51500701']:
                    soma_valores_G60 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '9819229901':
                if row.FONTE_RECURSO in ['21540770','21540000','31500701','51500701']:
                    soma_valores_G61 += row.VR_ARREC_ATE_MES_FONTE
                    
            if row.CODIGO_RECEITA == '9819225101':
                if row.FONTE_RECURSO in ['21540770','21540000','31500701','51500701']:
                    soma_valores_G62 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '7922510100':
                if row.FONTE_RECURSO in ['21540770','21540000','31500701','51500701']:
                    soma_valores_G63 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '9217515001':
                if row.FONTE_RECURSO in ['21540770', '21540000', '31500701' ,'51500701']:
                    soma_valores_G64 += row.VR_ARREC_ATE_MES_FONTE

            if row.CODIGO_RECEITA == '1715530100':
                if row.FONTE_RECURSO in ['21546770']:
                    soma_valores_G65 += row.VR_ARREC_ATE_MES_FONTE

        sheet['G56'].value = soma_valores_G56
        sheet['G57'].value = soma_valores_G57
        sheet['G58'].value = soma_valores_G58
        sheet['G59'].value = soma_valores_G59
        sheet['G60'].value = soma_valores_G60
        sheet['G61'].value = soma_valores_G61
        sheet['G62'].value = soma_valores_G62
        sheet['G63'].value = soma_valores_G63
        sheet['G64'].value = soma_valores_G64
        sheet['G65'].value = soma_valores_G65
        sheet['G43'].value = soma_valores_especificos

    if ded_data is not None:
        for row in ded_data.itertuples():
            if row.fonte == '1500701':
                sheet['G49'].value = row.liq_ate_mes
            if row.fonte == '31500701':
                sheet['G50'].value = row.liq_ate_mes
            if row.fonte == '51500701':
                sheet['G51'].value = row.liq_ate_mes
            if row.fonte == '2500701':
                sheet['G52'].value = row.liq_ate_mes
            if row.fonte == '32500701':
                sheet['G53'].value = row.liq_ate_mes
            if row.fonte == '52500701':
                sheet['G54'].value = row.liq_ate_mes
            elif row.fonte in ['21540770', '21540000']:

                soma_valores_G67 += row.liq_ate_mes
            elif row.fonte in ['22540770', '22540000','21546770']:

                soma_valores_G71 += row.pag_ate_mes

        sheet['G71'].value = soma_valores_G71
        sheet['G67'].value = soma_valores_G67

        if rpnp_data is not None:
            for row in rpnp_data.itertuples():
                if row.fonte == '1500701':
                    sheet['G74'].value = row.valor_anu_ant
                    sheet['G82'].value = row.valor_anu_mes
                if row.fonte == '31500701':
                    sheet['G75'].value = row.valor_anu_ant
                    sheet['G83'].value = row.valor_anu_mes
                if row.fonte == '51500701':
                    sheet['G76'].value = row.valor_anu_ant
                    sheet['G84'].value = row.valor_anu_mes
                if row.fonte == '2500701':
                    sheet['G77'].value = row.valor_anu_ant
                    sheet['G85'].value = row.valor_anu_mes
                if row.fonte == '32500701':
                    sheet['G78'].value = row.valor_anu_ant
                    sheet['G86'].value = row.valor_anu_mes
                if row.fonte == '52500701':
                    sheet['G79'].value = row.valor_anu_ant
                    sheet['G87'].value = row.valor_anu_mes


    if rpp_1619_data is not None:
        for row in rpp_1619_data.itertuples():
            if row.fonte == '1500701':
                sheet['G89'].value = row.valor_anu_ant
                sheet['G97'].value = row.valor_anu_mes
            if row.fonte == '31500701':
                sheet['G90'].value = row.valor_anu_ant
                sheet['G98'].value = row.valor_anu_mes
            if row.fonte == '51500701':
                sheet['G91'].value = row.valor_anu_ant
                sheet['G99'].value = row.valor_anu_mes
            if row.fonte == '2500701':
                sheet['G92'].value = row.valor_anu_ant
                sheet['G100'].value = row.valor_anu_mes
            if row.fonte == '32500701':
                sheet['G93'].value = row.valor_anu_ant
                sheet['G101'].value = row.valor_anu_mes
            if row.fonte == '52500701':
                sheet['G94'].value = row.valor_anu_ant
                sheet['G102'].value = row.valor_anu_mes

    workbook.save(file_path)
    workbook.close()

def main():
    st.title("MDE")
    uploaded_qgr = st.file_uploader("Upload QGR file", type=["csv"])
    uploaded_ded = st.file_uploader("Upload DED file", type=["csv"])
    uploaded_rpnp = st.file_uploader("Upload RPNP file", type=["csv"])
    

    uploaded_rpp_1619 = st.file_uploader("Upload RPP_1619 file", type=["csv"])
    
    uploaded_mde = st.file_uploader("Upload MDE file", type=["xlsx"])

    if st.button("Processar Dados"):

        if uploaded_qgr and uploaded_ded and uploaded_rpnp and uploaded_rpp_1619 and uploaded_mde:
            qgr_data = process_qgr(uploaded_qgr)
            ded_data = process_ded(uploaded_ded)
            rpnp_data = process_rpnp(uploaded_rpnp)
            rpp_1619_data = process_rpp_1619(uploaded_rpp_1619) 

            temp_file_path = "temp_MDE.xlsx"
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_mde.getbuffer())

        
            fill_MDE(temp_file_path, qgr_data, ded_data, rpnp_data, rpp_1619_data)

            with open(temp_file_path, "rb") as f:
                st.download_button(
                    label="Baixar arquivo MDE processado",
                    data=f,
                    file_name="MDE_processado.xlsx",
                    mime="application/vnd.ms-excel"
                )
            os.remove(temp_file_path)
            st.success("Dados processados e MDE preenchido com sucesso!")
        else:
            st.error("Por favor, faça o upload de todos os arquivos necessários.")

if __name__ == "__main__":
    main()
