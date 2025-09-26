import streamlit as st
import pandas as pd
from io import BytesIO
import re

def format_value(value):
    number = float(value)
    formatted_value = f"{number:,.2f}"
    formatted_value = formatted_value.replace(',', 'X').replace('.', ',').replace('X', '.')
    return formatted_value

def is_significant_difference(value1, value2):
    return abs(value1 - value2) >= 0.001

def parse_brl(value):

    if pd.isna(value):
        return float('nan')
    s = str(value).strip()
    if s == '':
        return float('nan')

    s = s.replace('R$', '').replace(' ', '')
    
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1]
    if '-' in s:
        negative = True
        s = s.replace('-', '')

    s = re.sub(r'[^\d\.,]', '', s)
    if s == '':
        return float('nan')

    if ',' in s and '.' in s:
       
        if s.rfind(',') > s.rfind('.'):
            decimal_sep = ','
            thousands_sep = '.'
        else:
            decimal_sep = '.'
            thousands_sep = ','
        s_numeric = s.replace(thousands_sep, '').replace(decimal_sep, '.')
    elif ',' in s:
   
        if re.search(r',\d{1,2}$', s):
            s_numeric = s.replace('.', '').replace(',', '.')
        else:
           
            s_numeric = s.replace(',', '')
    elif '.' in s:

        last_dot = s.rfind('.')
        digits_after = len(s) - last_dot - 1
        if 1 <= digits_after <= 2:
           
            left = s[:last_dot].replace('.', '')
            right = s[last_dot+1:]
            s_numeric = left + '.' + right
        else:

            s_numeric = s.replace('.', '')
    else:
       
        s_numeric = s

    try:
        number = float(s_numeric)
    except Exception:
        return float('nan')
    return -number if negative else number

def read_file(uploaded_file):

    if uploaded_file.name.lower().endswith('.csv'):
        df = pd.read_csv(uploaded_file, dtype=str, encoding='ISO-8859-1', delimiter=';', on_bad_lines='skip')
    else:

        df = pd.read_excel(uploaded_file, dtype=str)
    return df

def main():
    st.title('Posição de bancos VS CTB')

    posicao_file = st.file_uploader("Selecione o arquivo CSV ou XLSX para 'Posição de Bancos'", type=['csv', 'xlsx'])
    ctb_file = st.file_uploader("Selecione o arquivo CSV ou XLSX para 'CTB'", type=['csv', 'xlsx'])

    if posicao_file is not None and ctb_file is not None:

        posicao_bancos = read_file(posicao_file)
        ctb = read_file(ctb_file)

 
        ctb = ctb[ctb['registro'] == '20']

   
        posicao_bancos['saldo_inicial'] = posicao_bancos['saldo_inicial'].apply(parse_brl)
        posicao_bancos['saldo_final'] = posicao_bancos['saldo_final'].apply(parse_brl)
        ctb['saldo_inicial'] = ctb['saldo_inicial'].apply(parse_brl)
        ctb['saldo_final'] = ctb['saldo_final'].apply(parse_brl)

 
        ctb_grouped = ctb.groupby('cod_reduz_banco').agg({
            'saldo_inicial': 'sum',
            'saldo_final': 'sum'
        }).reset_index()

        differences = []
        correct_records = []

        for idx, row_posicao in posicao_bancos.iterrows():
            cod_reduz = row_posicao['cod_reduz_banco']

            if cod_reduz in ctb_grouped['cod_reduz_banco'].values:
                row_ctb = ctb_grouped[ctb_grouped['cod_reduz_banco'] == cod_reduz].iloc[0]

                saldo_inicial_posicao = row_posicao['saldo_inicial']
                saldo_inicial_ctb = row_ctb['saldo_inicial']
                saldo_final_posicao = row_posicao['saldo_final']
                saldo_final_ctb = row_ctb['saldo_final']

                if (is_significant_difference(saldo_inicial_posicao, saldo_inicial_ctb) or
                    is_significant_difference(saldo_final_posicao, saldo_final_ctb)):
                    
                    differences.append({
                        'cod_reduz_banco': cod_reduz,
                        'Diferenca_saldo_inicial': saldo_inicial_posicao != saldo_inicial_ctb,
                        'Diferenca_saldo_final': saldo_final_posicao != saldo_final_ctb,
                        'saldo_inicial_posicao': format_value(saldo_inicial_posicao),
                        'saldo_inicial_ctb': format_value(saldo_inicial_ctb),
                        'saldo_final_posicao': format_value(saldo_final_posicao),
                        'saldo_final_ctb': format_value(saldo_final_ctb)
                    })
                else:
                    correct_records.append({
                        'cod_reduz_banco': cod_reduz,
                        'saldo_inicial_posicao': format_value(saldo_inicial_posicao),
                        'saldo_inicial_ctb': format_value(saldo_inicial_ctb),
                        'saldo_final_posicao': format_value(saldo_final_posicao),
                        'saldo_final_ctb': format_value(saldo_final_ctb)
                    })

        differences_df = pd.DataFrame(differences)
        correct_records_df = pd.DataFrame(correct_records)

        st.write("Diferenças encontradas:")
        st.write(differences_df)

        st.write("Registros Corretos:")
        st.write(correct_records_df)

        if st.button('Gerar Excel'):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                differences_df.to_excel(writer, index=False, sheet_name='Diferencas')
                correct_records_df.to_excel(writer, index=False, sheet_name='Corretos')
            output.seek(0)
            st.download_button(label='Baixar Excel', data=output, file_name='Diferencas_Saldos.xlsx')

if __name__ == '__main__':
    main()
