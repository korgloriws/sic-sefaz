import pandas as pd
import streamlit as st
from io import BytesIO

def load_uploaded_file(uploaded_file):
    if uploaded_file:
        return pd.read_excel(uploaded_file, dtype=str)
    return None


def normalize_col_name(name):
    """Remove espaços para comparar nomes de coluna de forma flexível."""
    return str(name).replace(' ', '').upper()


def find_column(df, expected_name):
    """Localiza coluna ignorando diferenças de espaços (ex.: 'A / B' vs 'A/B')."""
    target = normalize_col_name(expected_name)
    for col in df.columns:
        if normalize_col_name(col) == target:
            return col
    return None


def resolve_columns(df, expected_names):
    """Retorna lista de nomes reais das colunas encontradas, na ordem pedida."""
    resolved = []
    for name in expected_names:
        col = find_column(df, name)
        if col is not None:
            resolved.append(col)
    return resolved


def main():
    st.title("ICISMEP")

    uploaded_file = st.file_uploader("Selecione um arquivo Excel", type=['xlsx', 'xls'])

    if uploaded_file:
        data = load_uploaded_file(uploaded_file)
        column_name = find_column(data, 'NATUREZA DESPESA / SUBELEMENTO')

        if st.button('Filtrar e desmembrar coluna'):
            if column_name is None:
                st.error(
                    "Coluna de natureza da despesa não encontrada. "
                    "Esperado algo como 'NATUREZA DESPESA / SUBELEMENTO' "
                    "(espaços são ignorados na comparação)."
                )
                st.write("Colunas do arquivo:", list(data.columns))
                return

            filtered_data = data[
                data[column_name].astype(str).str.replace('.', '', regex=False).str.match(r'^\d{8}$')
            ].copy()

            filtered_data['NATUREZA DESPESA / SUBELEMENTO COM 6 DIGITOS'] = (
                filtered_data[column_name].str.replace('.', '', regex=False).str.slice(0, 6)
            )
            filtered_data['NATUREZA DESPESA / SUBELEMENTO COM 2 DIGITOS'] = (
                filtered_data[column_name].str.replace('.', '', regex=False).str.slice(-2)
            )

            cnpj_col = find_column(filtered_data, 'CNPJ')
            if cnpj_col:
                filtered_data[cnpj_col] = filtered_data[cnpj_col].apply(lambda x: x.zfill(14))

            for expected in ['Mês Referência', 'COD. FUNÇÃO']:
                col = find_column(filtered_data, expected)
                if col is not None:
                    filtered_data[col] = filtered_data[col].apply(
                        lambda x: x.zfill(2) if pd.notnull(x) else x
                    )

            filtered_data['CO'] = ''
            filtered_data['Registro'] = 30

            final_cols_order = resolve_columns(
                filtered_data,
                [
                    'Registro',
                    'CNPJ',
                    'Mês Referência',
                    'COD. FUNÇÃO',
                    'COD. SUBFUNÇÃO',
                    'NATUREZA DESPESA / SUBELEMENTO COM 6 DIGITOS',
                    'NATUREZA DESPESA / SUBELEMENTO COM 2 DIGITOS',
                    'FONTE',
                    'CO',
                    'valor Empenhado no mês',
                    'Valor de empenhos anulados no mês',
                    'valor Liquidado no mês',
                    'Valor de liquidações anuladas no mês',
                    'Valor pago no mês',
                    'Valor de pagamentos anulados no mês',
                ],
            )

            filtered_data = filtered_data[final_cols_order]

            value_cols = resolve_columns(
                filtered_data,
                [
                    'valor Empenhado no mês',
                    'Valor de empenhos anulados no mês',
                    'valor Liquidado no mês',
                    'Valor de liquidações anuladas no mês',
                    'Valor pago no mês',
                    'Valor de pagamentos anulados no mês',
                ],
            )
            for col in value_cols:
                filtered_data[col] = filtered_data[col].apply(
                    lambda x: f"{float(x):,.2f}".replace('.', '|').replace(',', '').replace('|', ',')
                    if pd.notnull(x) and x != '' else x
                )

            st.dataframe(filtered_data)

            csv = filtered_data.to_csv(index=False, sep=';')
            st.download_button(
                label='Clique para baixar o arquivo!',
                data=BytesIO(csv.encode()),
                file_name='dados_filtrados.csv',
                mime='text/csv'
            )

    else:
        st.write("Por favor, faça o upload de um arquivo Excel.")

if __name__ == "__main__":
    main()
