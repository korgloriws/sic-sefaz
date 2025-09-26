from __future__ import annotations

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


def _normalize_column_name(name: str) -> str:
    if name is None:
        return ""
    s = str(name).strip().lower()
    for ch in [" ", "-", "/", "\\", ".", ",", ";", ":"]:
        s = s.replace(ch, "_")
    while "__" in s:
        s = s.replace("__", "_")
    return s


def _build_normalized_map(columns):
    norm_map = {}
    for col in columns:
        norm = _normalize_column_name(col)
        if norm not in norm_map:
            norm_map[norm] = col
    return norm_map


def _resolve_columns(df: pd.DataFrame, required_to_candidates, df_label: str):
    norm_map = _build_normalized_map(df.columns)
    resolved = {}
    for required_name, candidates in required_to_candidates.items():
        found = None
        for cand in candidates:
            cand_norm = _normalize_column_name(cand)
            if cand_norm in norm_map:
                found = norm_map[cand_norm]
                break
        if not found:
            available = ", ".join([str(c) for c in df.columns])
            raise ValueError(
                f"Coluna obrigatória '{required_name}' não encontrada em '{df_label}'. Colunas disponíveis: {available}"
            )
        resolved[required_name] = found
    return resolved


def _coerce_to_string(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def _coerce_code_string(series: pd.Series) -> pd.Series:
    def fmt(v):
        if pd.isna(v):
            return ""
        if isinstance(v, (int, np.integer)):
            return str(int(v))
        if isinstance(v, float) and np.isfinite(v):
            if float(v).is_integer():
                return str(int(v))
            return str(v).strip()
        s = str(v).strip()
        if s.endswith(".0") and s[:-2].isdigit():
            return s[:-2]
        return s
    return series.apply(fmt)


def _coerce_to_number(series: pd.Series) -> pd.Series:
    nums = pd.to_numeric(series, errors="coerce")
    nums = nums.fillna(0.0)
    return nums


def _prepare_df_990(df_990: pd.DataFrame) -> pd.DataFrame:
    required = {
        "cod_contabil": ["cod_contabil", "codigo_contabil", "conta_contabil", "conta"],
        "tipo_conta": ["tipo_conta", "tipo", "classe"],
        "codigo_tce": ["codigo_tce", "cod_tce", "codigo_pcasp", "conta_tce", "pcasp"],
    }
    mapping = _resolve_columns(df_990, required, df_label="990")
    df = df_990.rename(columns={v: k for k, v in mapping.items()}).copy()
    df["cod_contabil"] = _coerce_code_string(df["cod_contabil"]).astype(str)
    df["tipo_conta"] = df["tipo_conta"].astype(str)
    df["codigo_tce"] = _coerce_code_string(df["codigo_tce"]).astype(str)
    return df[["cod_contabil", "tipo_conta", "codigo_tce"]]


def _prepare_df_998(df_998: pd.DataFrame) -> pd.DataFrame:
    required = {
        "cod_contabil": ["cod_contabil", "codigo_contabil", "conta_contabil", "conta"],
        "debito_atual": ["debito_atual", "debito", "vlr_debito", "debito_mes"],
        "credito_atual": ["credito_atual", "credito", "vlr_credito", "credito_mes"],
    }
    mapping = _resolve_columns(df_998, required, df_label="998")
    df = df_998.rename(columns={v: k for k, v in mapping.items()}).copy()
    df["cod_contabil"] = _coerce_code_string(df["cod_contabil"]).astype(str)
    df["debito_atual"] = _coerce_to_number(df["debito_atual"]) 
    df["credito_atual"] = _coerce_to_number(df["credito_atual"]) 
    return df[["cod_contabil", "debito_atual", "credito_atual"]]


def _prepare_df_pcasp(df_pcasp: pd.DataFrame) -> pd.DataFrame:
    required = {
        "conta": ["conta", "codigo", "codigo_tce", "codigo_pcasp"],
        "natureza_do_saldo": ["natureza_do_saldo", "natureza do saldo", "natureza", "nat_saldo"],
    }
    mapping = _resolve_columns(df_pcasp, required, df_label="pcasp_tce")
    df = df_pcasp.rename(columns={v: k for k, v in mapping.items()}).copy()
    df["conta"] = _coerce_code_string(df["conta"]).astype(str)
    df["natureza_do_saldo"] = df["natureza_do_saldo"].astype(str).str.upper().str.strip()
    return df[["conta", "natureza_do_saldo"]]


def _compute_lado_do_saldo(debito: pd.Series, credito: pd.Series) -> pd.Series:
    deb_pos = debito.fillna(0) > 0
    cred_pos = credito.fillna(0) > 0
    lado = np.where(
        deb_pos & ~cred_pos,
        "D",
        np.where(
            cred_pos & ~deb_pos,
            "C",
            np.where(deb_pos & cred_pos, "D/C", "SEM SALDO"),
        ),
    )
    return pd.Series(lado, index=debito.index, name="resultado_comparacao")


def _compute_erro(natureza: pd.Series, debito: pd.Series, credito: pd.Series) -> pd.Series:
    nat = natureza.fillna("").astype(str).str.upper().str.strip()
    deb_pos = debito.fillna(0) > 0
    cred_pos = credito.fillna(0) > 0
    erro = (
        (deb_pos & ~nat.isin(["D", "D/C"]))
        | (cred_pos & ~nat.isin(["C", "D/C"]))
        | ((deb_pos & cred_pos) & (nat != "D/C"))
    )
    return erro


def process_data(df_990: pd.DataFrame, df_998: pd.DataFrame, df_pcasp: pd.DataFrame):
    base_990 = _prepare_df_990(df_990)
    base_998 = _prepare_df_998(df_998)
    base_pcasp = _prepare_df_pcasp(df_pcasp)
    merged = base_998.merge(base_990, on="cod_contabil", how="inner", suffixes=("_998", "_990"))
    tipo_str = merged["tipo_conta"].astype(str).str.strip()
    merged = merged[tipo_str.str.startswith("5", na=False)]
    merged["codigo_tce"] = _coerce_code_string(merged["codigo_tce"]).astype(str)
    enriched = merged.merge(
        base_pcasp.rename(columns={"conta": "codigo_tce"}),
        on="codigo_tce",
        how="left",
    )
    enriched["resultado_comparacao"] = _compute_lado_do_saldo(
        enriched["debito_atual"], enriched["credito_atual"]
    )
    enriched["erro"] = _compute_erro(
        enriched["natureza_do_saldo"], enriched["debito_atual"], enriched["credito_atual"]
    )
    resultado = enriched[[
        "cod_contabil",
        "codigo_tce",
        "debito_atual",
        "credito_atual",
        "natureza_do_saldo",
        "resultado_comparacao",
    ]].copy()
    resultado = resultado.rename(columns={
        "natureza_do_saldo": "NATUREZA DO PCASP",
        "resultado_comparacao": "Natureza Calculada",
    })
    erros = resultado.copy()
    erros = erros[enriched["erro"] == True]
    resultado = resultado.sort_values(by=["cod_contabil", "codigo_tce"]).reset_index(drop=True)
    erros = erros.sort_values(by=["cod_contabil", "codigo_tce"]).reset_index(drop=True)
    return resultado, erros


def save_report_to_excel_bytes(resultado: pd.DataFrame, erros: pd.DataFrame, sheet_resultado: str = "resultado", sheet_erros: str = "erros") -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resultado.to_excel(writer, index=False, sheet_name=sheet_resultado)
        erros.to_excel(writer, index=False, sheet_name=sheet_erros)
    buffer.seek(0)
    return buffer.getvalue()


def _try_read_excel(file) -> pd.DataFrame:
    return pd.read_excel(file, engine="openpyxl")


def main():
    st.title("Conferência de Balancete para saldos invertidos")
    st.markdown("Faça upload dos três arquivos .xlsx (990, 998 e PCASP/TCE) para gerar o relatório.")
    file_990 = st.file_uploader("Arquivo 990 (xlsx)", type=["xlsx"], key="arq990")
    file_998 = st.file_uploader("Arquivo 998 (xlsx)", type=["xlsx"], key="arq998")
    file_pcasp = st.file_uploader("Arquivo PCASP/TCE (xlsx)", type=["xlsx"], key="arqpcasp")
    processar = st.button("Processar")

    if processar:
        if not (file_990 and file_998 and file_pcasp):
            st.error("Envie os três arquivos para continuar.")
        else:
            try:
                df_990 = _try_read_excel(file_990)
                df_998 = _try_read_excel(file_998)
                df_pcasp = _try_read_excel(file_pcasp)
                resultado, erros = process_data(df_990, df_998, df_pcasp)
                st.subheader("Resultado")
                st.dataframe(resultado, use_container_width=True)
                st.subheader("Erros")
                st.dataframe(erros, use_container_width=True)
                bytes_xlsx = save_report_to_excel_bytes(resultado, erros)
                st.download_button(
                    label="Baixar relatório (xlsx)",
                    data=bytes_xlsx,
                    file_name="conferencia_balancete.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as exc:
                st.error(f"Falha ao processar arquivos: {exc}")


if __name__ == "__main__":
    main()

