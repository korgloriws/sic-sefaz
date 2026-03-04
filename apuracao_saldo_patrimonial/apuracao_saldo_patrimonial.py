# -*- coding: utf-8 -*-
"""
Apuração de saldo patrimonial.
Processa arquivos 998 e 990 e preenche a planilha de apuração.
- main(): ponto de entrada Streamlit (como nos demais módulos).
- processar(arquivo_998, arquivo_990, arquivo_apuracao_template, ...): uso programático.
"""

import io
from pathlib import Path
from typing import Optional

import pandas as pd
import openpyxl
import streamlit as st


# --- Colunas e constantes ---
# 998: A=0 (conta), G=6 (débito), H=7 (crédito)
COLS_998 = {"conta": 0, "debito": 6, "credito": 7}
# 990: A=0 (conta), D=3 (tipo), J=9 (atributo)
COLS_990 = {"conta": 0, "tipo": 3, "atributo": 9}

TIPO_FILTRO = "5"
ATRIBUTOS_VALIDOS = ("F", "P")


def _normalizar_conta(val) -> str:
    """Garante conta como string para comparação."""
    if pd.isna(val):
        return ""
    return str(val).strip()


def carregar_998(caminho_ou_buffer) -> pd.DataFrame:
    """Carrega arquivo 998 com colunas A (conta), G (débito), H (crédito)."""
    df = pd.read_excel(
        caminho_ou_buffer,
        usecols=[COLS_998["conta"], COLS_998["debito"], COLS_998["credito"]],
        header=None,
        names=["conta", "debito", "credito"],
    )
    df["conta"] = df["conta"].apply(_normalizar_conta)
    df["debito"] = pd.to_numeric(df["debito"], errors="coerce").fillna(0)
    df["credito"] = pd.to_numeric(df["credito"], errors="coerce").fillna(0)
    return df


def carregar_990(caminho_ou_buffer) -> pd.DataFrame:
    """Carrega arquivo 990 com colunas A (conta), D (tipo), J (atributo)."""
    df = pd.read_excel(
        caminho_ou_buffer,
        usecols=[COLS_990["conta"], COLS_990["tipo"], COLS_990["atributo"]],
        header=None,
        names=["conta", "tipo", "atributo"],
    )
    df["conta"] = df["conta"].apply(_normalizar_conta)
    df["tipo"] = df["tipo"].astype(str).str.strip()
    df["atributo"] = df["atributo"].astype(str).str.strip().str.upper()
    return df


def filtrar_contas_990(
    df_990: pd.DataFrame,
    tipo: str = TIPO_FILTRO,
    atributo: Optional[str] = None,
    prefixo: Optional[str] = None,
    conta_exata: Optional[str] = None,
) -> set:
    """Retorna conjunto de contas do 990 que passam nos filtros."""
    mask = df_990["tipo"] == tipo
    if atributo is not None:
        mask = mask & (df_990["atributo"] == atributo.upper())
    if prefixo is not None:
        mask = mask & (df_990["conta"].str.startswith(prefixo, na=False))
    if conta_exata is not None:
        mask = mask & (df_990["conta"] == _normalizar_conta(conta_exata))
    return set(df_990.loc[mask, "conta"].dropna().unique())


def valor_conta_998(df_998: pd.DataFrame, conta: str) -> float:
    """Para uma conta no 998, retorna o saldo (débito - crédito)."""
    linhas = df_998[df_998["conta"] == conta]
    if linhas.empty:
        return 0.0
    total = 0.0
    for _, row in linhas.iterrows():
        d, c = row["debito"], row["credito"]
        if d != 0 or c != 0:
            total += float(d) - float(c)
    return total


def somar_contas_998(df_998: pd.DataFrame, contas: set) -> float:
    """Soma os valores (saldo) das contas dadas no 998."""
    return sum(valor_conta_998(df_998, c) for c in contas if c)


REGRAS_APURACAO = [
    ("B", 5, "111", None, "F"),
    ("B", 6, "113", None, "F"),
    ("B", 7, "114", None, "F"),
    ("B", 8, "12", None, "F"),
    ("B", 11, "11", None, "P"),
    ("B", 12, "12", None, "P"),
    ("B", 15, "21", None, "F"),
    ("B", 16, "22", None, "F"),
    ("B", 17, None, "631100000000000", None),
    ("B", 18, None, "631710000000000", None),
    ("B", 21, "21", None, "P"),
    ("B", 22, "22", None, "P"),
]


def calcular_valores_apuracao(df_998: pd.DataFrame, df_990: pd.DataFrame) -> list:
    """Aplica as regras e retorna lista de (coluna, linha, valor) para escrita."""
    resultados = []
    for col, linha, prefixo, conta_exata, atributo in REGRAS_APURACAO:
        contas = filtrar_contas_990(
            df_990,
            tipo=TIPO_FILTRO,
            atributo=atributo,
            prefixo=prefixo,
            conta_exata=conta_exata,
        )
        valor = somar_contas_998(df_998, contas)
        resultados.append((col, linha, valor))
    return resultados


def preencher_planilha_apuracao(
    caminho_template: str | Path | object,
    caminho_saida: str | Path | object,
    resultados: list,
) -> None:
    """Escreve os valores nas células da planilha de apuração. Aceita caminhos ou file-like."""
    wb = openpyxl.load_workbook(caminho_template)
    ws = wb.active
    for col, linha, valor in resultados:
        ws[f"{col}{linha}"] = abs(valor)
    wb.save(caminho_saida)


def processar(
    arquivo_998,
    arquivo_990,
    arquivo_apuracao_template,
    arquivo_apuracao_saida: Optional[str | Path] = None,
):
    """
    Processa os três arquivos e gera a apuração preenchida.
    Uso programático: from apuracao_saldo_patrimonial import processar
    """
    df_998 = carregar_998(arquivo_998)
    df_990 = carregar_990(arquivo_990)
    resultados = calcular_valores_apuracao(df_998, df_990)

    if arquivo_apuracao_saida is None:
        arquivo_apuracao_saida = arquivo_apuracao_template

    preencher_planilha_apuracao(
        arquivo_apuracao_template,
        arquivo_apuracao_saida,
        resultados,
    )
    return resultados


def main():
    """Ponto de entrada Streamlit (igual aos demais módulos). Sem set_page_config ao ser embutido no app principal."""
    st.title("Apuração de Saldo Patrimonial")
    st.markdown(
        "Envie os arquivos **998**, **990** e o modelo **Apuração de saldo patrimonial** (.xlsx). "
        "O sistema preenche a apuração com base nas contas tipo 5 e atributos F/P do 990 e nos valores do 998."
    )

    arquivo_998 = st.file_uploader("Arquivo 998", type=["xlsx"], key="998")
    arquivo_990 = st.file_uploader("Arquivo 990", type=["xlsx"], key="990")
    arquivo_apuracao = st.file_uploader(
        "Planilha Apuração de saldo patrimonial (modelo)",
        type=["xlsx"],
        key="apuracao",
    )

    if arquivo_998 and arquivo_990 and arquivo_apuracao:
        if st.button("Processar e gerar apuração"):
            try:
                arquivo_998.seek(0)
                arquivo_990.seek(0)
                arquivo_apuracao.seek(0)
                buf_saida = io.BytesIO()
                processar(
                    arquivo_998,
                    arquivo_990,
                    arquivo_apuracao,
                    arquivo_apuracao_saida=buf_saida,
                )
                buf_saida.seek(0)
                st.success("Processamento concluído.")
                st.download_button(
                    label="Baixar planilha preenchida",
                    data=buf_saida,
                    file_name="apuracao_saldo_patrimonial_preenchida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
                raise
    else:
        st.info("Envie os três arquivos .xlsx para continuar.")


if __name__ == "__main__":
    main()
