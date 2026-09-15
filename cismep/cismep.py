import re
import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st


# Colunas lógicas do processamento e possíveis nomes no Excel.
# A busca ignora espaços, acentos, pontuação e maiúsculas/minúsculas.
COLUMN_ALIASES = {
    'natureza': [
        'NATUREZA DESPESA / SUBELEMENTO',
        'NATUREZA DESPESA/SUBELEMENTO',
        'NATUREZA DA DESPESA / SUBELEMENTO',
        'NATUREZA DESPESA SUBELEMENTO',
    ],
    'cnpj': ['CNPJ'],
    'mes': ['Mês Referência', 'Mes Referencia', 'MÊS REFERÊNCIA'],
    'funcao': ['COD. FUNÇÃO', 'COD FUNCAO', 'CÓD. FUNÇÃO', 'CODIGO FUNCAO'],
    'subfuncao': ['COD. SUBFUNÇÃO', 'COD SUBFUNCAO', 'CÓD. SUBFUNÇÃO', 'CODIGO SUBFUNCAO'],
    'fonte': ['FONTE'],
    'empenhado': ['valor Empenhado no mês', 'Valor Empenhado no mês'],
    'empenhos_anulados': ['Valor de empenhos anulados no mês'],
    'liquidado': ['valor Liquidado no mês', 'Valor Liquidado no mês'],
    'liquidacoes_anuladas': ['Valor de liquidações anuladas no mês', 'Valor de liquidacoes anuladas no mês'],
    'pago': ['Valor pago no mês'],
    'pagamentos_anulados': ['Valor de pagamentos anulados no mês'],
}

OUTPUT_ORDER = [
    ('Registro', 'Registro'),
    ('CNPJ', 'cnpj'),
    ('Mês Referência', 'mes'),
    ('COD. FUNÇÃO', 'funcao'),
    ('COD. SUBFUNÇÃO', 'subfuncao'),
    ('NATUREZA DESPESA / SUBELEMENTO COM 6 DIGITOS', 'natureza_6'),
    ('NATUREZA DESPESA / SUBELEMENTO COM 2 DIGITOS', 'natureza_2'),
    ('FONTE', 'fonte'),
    ('CO', 'CO'),
    ('valor Empenhado no mês', 'empenhado'),
    ('Valor de empenhos anulados no mês', 'empenhos_anulados'),
    ('valor Liquidado no mês', 'liquidado'),
    ('Valor de liquidações anuladas no mês', 'liquidacoes_anuladas'),
    ('Valor pago no mês', 'pago'),
    ('Valor de pagamentos anulados no mês', 'pagamentos_anulados'),
]

VALUE_KEYS = [
    'empenhado',
    'empenhos_anulados',
    'liquidado',
    'liquidacoes_anuladas',
    'pago',
    'pagamentos_anulados',
]


def load_uploaded_file(uploaded_file):
    if uploaded_file:
        return pd.read_excel(uploaded_file, dtype=str)
    return None


def normalize_text(value):
    """Normaliza texto para comparação: sem acento, espaço, pontuação ou case."""
    text = str(value)
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    return re.sub(r'[^A-Z0-9]', '', text)


def find_column(df, aliases):
    """Localiza coluna real no DataFrame a partir de aliases flexíveis."""
    normalized_cols = {normalize_text(col): col for col in df.columns}
    for alias in aliases:
        key = normalize_text(alias)
        if key in normalized_cols:
            return normalized_cols[key]
    return None


def map_columns(df):
    """Mapeia colunas lógicas -> nomes reais encontrados no arquivo."""
    mapping = {}
    for key, aliases in COLUMN_ALIASES.items():
        col = find_column(df, aliases)
        if col is not None:
            mapping[key] = col
    return mapping


def only_digits(value):
    if pd.isnull(value):
        return ''
    return re.sub(r'\D', '', str(value))


def parse_number(value):
    """Aceita número em formato BR (1.234,56), US (1,234.56) ou simples."""
    if pd.isnull(value):
        return None

    text = str(value).strip()
    if text == '' or text.lower() in {'nan', 'none', '-'}:
        return None

    text = text.replace('\xa0', '').replace(' ', '')
    negative = text.startswith('(') and text.endswith(')')
    if negative:
        text = text[1:-1]
    if text.startswith('-'):
        negative = True
        text = text[1:]

    text = re.sub(r'[^\d,.\-]', '', text)
    if text in {'', '-', '.', ','}:
        return None

    if ',' in text and '.' in text:
        # Decide decimal pelo último separador.
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '').replace(',', '.')
        else:
            text = text.replace(',', '')
    elif ',' in text:
        parts = text.split(',')
        text = text.replace(',', '.') if len(parts[-1]) <= 2 else text.replace(',', '')

    try:
        number = float(text)
    except ValueError:
        return None

    return -number if negative else number


def format_br_number(value):
    number = parse_number(value)
    if number is None:
        return ''
    return f"{number:,.2f}".replace('.', '|').replace(',', '').replace('|', ',')


def pad_digits(value, width):
    digits = only_digits(value)
    return digits.zfill(width) if digits else ''


def main():
    st.title("ICISMEP")

    uploaded_file = st.file_uploader("Selecione um arquivo Excel", type=['xlsx', 'xls'])

    if not uploaded_file:
        st.write("Por favor, faça o upload de um arquivo Excel.")
        return

    data = load_uploaded_file(uploaded_file)
    columns = map_columns(data)

    if st.button('Filtrar e desmembrar coluna'):
        if 'natureza' not in columns:
            st.error(
                "Não encontrei a coluna de natureza da despesa. "
                "O programa ignora espaços, acentos e pontuação, mas precisa de algo como "
                "'NATUREZA DESPESA / SUBELEMENTO'."
            )
            st.write("Colunas do arquivo:", list(data.columns))
            return

        natureza_col = columns['natureza']
        natureza_digits = data[natureza_col].map(only_digits)
        filtered = data[natureza_digits.str.fullmatch(r'\d{8}')].copy()

        if filtered.empty:
            st.warning("Nenhuma linha com natureza da despesa de 8 dígitos foi encontrada.")
            return

        result = pd.DataFrame(index=filtered.index)
        result['Registro'] = 30
        result['CO'] = ''

        if 'cnpj' in columns:
            result['CNPJ'] = filtered[columns['cnpj']].map(lambda x: pad_digits(x, 14))
        if 'mes' in columns:
            result['Mês Referência'] = filtered[columns['mes']].map(lambda x: pad_digits(x, 2))
        if 'funcao' in columns:
            result['COD. FUNÇÃO'] = filtered[columns['funcao']].map(lambda x: pad_digits(x, 2))
        if 'subfuncao' in columns:
            result['COD. SUBFUNÇÃO'] = filtered[columns['subfuncao']].map(only_digits)
        if 'fonte' in columns:
            result['FONTE'] = filtered[columns['fonte']].fillna('').astype(str).str.strip()

        result['NATUREZA DESPESA / SUBELEMENTO COM 6 DIGITOS'] = natureza_digits.loc[filtered.index].str.slice(0, 6)
        result['NATUREZA DESPESA / SUBELEMENTO COM 2 DIGITOS'] = natureza_digits.loc[filtered.index].str.slice(-2)

        for key in VALUE_KEYS:
            if key in columns:
                out_name = next(name for name, k in OUTPUT_ORDER if k == key)
                result[out_name] = filtered[columns[key]].map(format_br_number)

        final_cols = [name for name, key in OUTPUT_ORDER if name in result.columns]
        result = result[final_cols]

        st.dataframe(result)
        st.caption(
            f"Processadas {len(result)} linhas. "
            "Colunas e números foram interpretados de forma flexível, sem alterar o arquivo original."
        )

        csv = result.to_csv(index=False, sep=';')
        st.download_button(
            label='Clique para baixar o arquivo!',
            data=BytesIO(csv.encode()),
            file_name='dados_filtrados.csv',
            mime='text/csv',
        )


if __name__ == "__main__":
    main()
