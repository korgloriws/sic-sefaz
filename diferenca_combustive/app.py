from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

import pandas as pd
import streamlit as st


@dataclass
class ComparisonResult:
    matched: pd.DataFrame
    only_abastecimento: pd.DataFrame
    only_razao: pd.DataFrame


def read_csv_with_fallback(uploaded_file) -> pd.DataFrame:
    """Read semicolon CSV trying common encodings for Brazilian files."""
    file_bytes = uploaded_file.getvalue()
    last_error: Exception | None = None

    for encoding in ("iso-8859-1", "latin1", "utf-8-sig", "utf-8"):
        try:
            return pd.read_csv(BytesIO(file_bytes), sep=";", encoding=encoding)
        except Exception as exc:  # noqa: BLE001
            last_error = exc

    raise ValueError(f"Nao foi possivel ler o CSV: {last_error}") from last_error


def normalize_date(series: pd.Series) -> pd.Series:
    """Normalize date into YYYY-MM-DD string for reliable joins."""
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return dt.dt.strftime("%Y-%m-%d")


def normalize_amount(series: pd.Series) -> pd.Series:
    """Normalize currency/decimal formats into numeric values."""
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def prepare_abastecimento(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = {"placa", "data_movimentacao", "total_por_registro"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"Colunas ausentes em abastecimento: {', '.join(sorted(missing))}")

    out = df.copy()
    out["placa"] = out["placa"].astype(str).str.strip()
    out["data_key"] = normalize_date(out["data_movimentacao"])
    out["valor_key"] = normalize_amount(out["total_por_registro"]).round(2)
    return out


def prepare_razao(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = {"data_lanc", "debito"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"Colunas ausentes em razao: {', '.join(sorted(missing))}")

    out = df.copy()
    out["data_key"] = normalize_date(out["data_lanc"])
    out["valor_key"] = normalize_amount(out["debito"]).round(2)
    return out


def compare_launches(abastecimento: pd.DataFrame, razao: pd.DataFrame) -> ComparisonResult:
    ab_key = abastecimento.reset_index(drop=False).rename(columns={"index": "ab_idx"})
    rz_key = razao.reset_index(drop=False).rename(columns={"index": "rz_idx"})

    merged = ab_key.merge(
        rz_key[["rz_idx", "data_key", "valor_key"]],
        on=["data_key", "valor_key"],
        how="outer",
        indicator=True,
    )

    only_ab = merged[merged["_merge"] == "left_only"][["ab_idx"]].dropna()
    only_rz = merged[merged["_merge"] == "right_only"][["rz_idx"]].dropna()
    matched = merged[merged["_merge"] == "both"][["ab_idx", "rz_idx"]].dropna()

    only_ab_df = (
        ab_key.loc[only_ab["ab_idx"].astype(int), ["placa", "data_movimentacao", "total_por_registro"]]
        .reset_index(drop=True)
    )
    only_rz_df = rz_key.loc[only_rz["rz_idx"].astype(int), ["data_lanc", "debito"]].reset_index(drop=True)

    matched_ab = ab_key.loc[matched["ab_idx"].astype(int), ["placa", "data_movimentacao", "total_por_registro"]]
    matched_rz = rz_key.loc[matched["rz_idx"].astype(int), ["data_lanc", "debito"]]
    matched_df = (
        pd.concat([matched_ab.reset_index(drop=True), matched_rz.reset_index(drop=True)], axis=1)
        .rename(
            columns={
                "data_movimentacao": "abastecimento_data",
                "total_por_registro": "abastecimento_valor",
                "data_lanc": "razao_data",
                "debito": "razao_valor",
            }
        )
    )

    return ComparisonResult(
        matched=matched_df,
        only_abastecimento=only_ab_df,
        only_razao=only_rz_df,
    )


def build_report_dataframes(result: ComparisonResult) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    corretos = result.matched.copy()
    corretos["status"] = "correto"

    errados_ab = result.only_abastecimento.copy().rename(
        columns={
            "data_movimentacao": "abastecimento_data",
            "total_por_registro": "abastecimento_valor",
        }
    )
    errados_ab["razao_data"] = pd.NA
    errados_ab["razao_valor"] = pd.NA
    errados_ab["origem_erro"] = "somente_abastecimento"
    errados_ab["status"] = "errado"

    errados_rz = result.only_razao.copy().rename(
        columns={
            "data_lanc": "razao_data",
            "debito": "razao_valor",
        }
    )
    errados_rz["placa"] = pd.NA
    errados_rz["abastecimento_data"] = pd.NA
    errados_rz["abastecimento_valor"] = pd.NA
    errados_rz["origem_erro"] = "somente_razao"
    errados_rz["status"] = "errado"

    base_columns = [
        "placa",
        "abastecimento_data",
        "abastecimento_valor",
        "razao_data",
        "razao_valor",
        "origem_erro",
        "status",
    ]

    corretos = corretos.assign(origem_erro="ok")[base_columns]
    errados = pd.concat(
        [errados_ab[base_columns], errados_rz[base_columns]],
        ignore_index=True,
    )
    comparacao_completa = pd.concat([corretos, errados], ignore_index=True)

    return comparacao_completa, corretos, errados


def create_excel_report(comparacao: pd.DataFrame, corretos: pd.DataFrame, errados: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output) as writer:
        comparacao.to_excel(writer, sheet_name="comparacao_completa", index=False)
        corretos.to_excel(writer, sheet_name="corretos", index=False)
        errados.to_excel(writer, sheet_name="errados", index=False)
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Comparador de Lancamentos", layout="wide")
    st.title("Comparador: Abastecimento x Razao")
    st.write(
        "Compare os lancamentos por **data** e **valor**. "
        "A coluna **placa** do abastecimento e usada apenas como identificacao."
    )

    col1, col2 = st.columns(2)
    with col1:
        file_ab = st.file_uploader("CSV Abastecimento", type=["csv"], key="ab")
    with col2:
        file_rz = st.file_uploader("CSV Razao", type=["csv"], key="rz")

    if not file_ab or not file_rz:
        st.info("Envie os dois arquivos CSV para iniciar a comparacao.")
        return

    try:
        ab_raw = read_csv_with_fallback(file_ab)
        rz_raw = read_csv_with_fallback(file_rz)

        ab = prepare_abastecimento(ab_raw)
        rz = prepare_razao(rz_raw)

        # Remove linhas sem data/valor valido antes da comparacao.
        ab = ab.dropna(subset=["data_key", "valor_key"]).reset_index(drop=True)
        rz = rz.dropna(subset=["data_key", "valor_key"]).reset_index(drop=True)

        result = compare_launches(ab, rz)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Erro ao processar arquivos: {exc}")
        return

    st.subheader("Resumo")
    c1, c2, c3 = st.columns(3)
    c1.metric("Correspondencias", len(result.matched))
    c2.metric("So em abastecimento", len(result.only_abastecimento))
    c3.metric("So em razao", len(result.only_razao))

    st.subheader("Lançamentos com correspondencia")
    st.dataframe(result.matched, use_container_width=True, hide_index=True)

    st.subheader("Lançamentos somente no abastecimento")
    st.dataframe(result.only_abastecimento, use_container_width=True, hide_index=True)

    st.subheader("Lançamentos somente no razao")
    st.dataframe(result.only_razao, use_container_width=True, hide_index=True)

    comparacao_completa, corretos, errados = build_report_dataframes(result)
    excel_bytes = create_excel_report(comparacao_completa, corretos, errados)

    st.download_button(
        label="Baixar relatorio XLSX",
        data=excel_bytes,
        file_name="relatorio_comparacao_combustivel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
