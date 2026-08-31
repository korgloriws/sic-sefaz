import streamlit as st
import pandas as pd
import io

PREFIXOS_AUSENTES = ("11111", "11131", "114")


def _codigo_com_prefixo(cod_contabil) -> bool:
    cod = str(cod_contabil).strip()
    return any(cod.startswith(prefixo) for prefixo in PREFIXOS_AUSENTES)


def _buscar_ausentes_prefixos(df1: pd.DataFrame, df2_grouped: pd.DataFrame) -> pd.DataFrame:
    """Contas ausentes em um dos arquivos (prefixos 11111, 11131 e 114)."""
    df1_filtrado = df1[df1["cod_contabil"].apply(_codigo_com_prefixo)].copy()
    df2_filtrado = df2_grouped[df2_grouped["cod_contabil"].apply(_codigo_com_prefixo)].copy()

    agg1 = df1_filtrado.groupby(["cod_contabil", "fonte"], as_index=False)["saldo_atual"].sum()
    agg2 = df2_filtrado.groupby(["cod_contabil", "fonte"], as_index=False)["saldo_matriz"].sum()

    merged = pd.merge(agg1, agg2, on=["cod_contabil", "fonte"], how="outer", indicator=True)
    ausentes = merged[merged["_merge"] != "both"].copy()

    if ausentes.empty:
        return ausentes

    ausentes["Presente_em"] = ausentes["_merge"].map(
        {"left_only": "Posição de Bancos", "right_only": "MSC"}
    )
    ausentes["Ausente_em"] = ausentes["_merge"].map(
        {"left_only": "MSC", "right_only": "Posição de Bancos"}
    )
    ausentes.drop(columns=["_merge"], inplace=True)
    ausentes["saldo_atual"] = ausentes["saldo_atual"].fillna(0)
    ausentes["saldo_matriz"] = ausentes["saldo_matriz"].fillna(0)

    colunas = [
        "cod_contabil",
        "fonte",
        "saldo_atual",
        "saldo_matriz",
        "Presente_em",
        "Ausente_em",
    ]
    return ausentes[colunas].sort_values(["cod_contabil", "fonte"]).reset_index(drop=True)


def main():
    st.title("Posição de bancos por fonte VS MSC")

    arq1 = st.file_uploader("Arquivo 1: Posição de bancos por fonte", type=["xlsx"])
    arq2 = st.file_uploader("Arquivo 2: Matriz de saldos contábeis", type=["xlsx"])

    if arq1 and arq2:
        df1 = pd.read_excel(arq1)
        df2 = pd.read_excel(arq2)

        cols1 = ["cod_contabil", "fonte", "saldo_atual"]
        cols2 = ["CONTA", "IC3", "Valor", "Tipo_valor", "Natureza_valor"]
        
        for c in cols1:
            if c not in df1.columns:
                st.error(f"Coluna '{c}' não encontrada no Arquivo 1.")
                st.stop()
        for c in cols2:
            if c not in df2.columns:
                st.error(f"Coluna '{c}' não encontrada no Arquivo 2.")
                st.stop()

        df1 = df1[cols1].copy()
        df2 = df2[cols2].copy()

        df2 = df2[df2["Tipo_valor"] == "ending_balance"].copy()

        df2["Valor_ajustado"] = df2.apply(lambda row: row["Valor"] if row["Natureza_valor"] == "D" else -row["Valor"], axis=1)

        df2_grouped = df2.groupby(["CONTA", "IC3"], as_index=False)["Valor_ajustado"].sum()
        df2_grouped.rename(columns={"CONTA": "cod_contabil", "IC3": "fonte", "Valor_ajustado": "saldo_matriz"}, inplace=True)

        df2_matriz_completa = df2_grouped.copy()

        df2_grouped = df2_grouped.merge(df1[["cod_contabil", "fonte"]].drop_duplicates(), on=["cod_contabil", "fonte"], how="inner")

        comparacao = pd.merge(df1, df2_grouped, on=["cod_contabil", "fonte"], how="outer", indicator=True)

        comparacao["saldo_atual"] = comparacao["saldo_atual"].fillna(0)
        comparacao["saldo_matriz"] = comparacao["saldo_matriz"].fillna(0)

        comparacao["correto"] = (comparacao["saldo_atual"].round(2) == comparacao["saldo_matriz"].round(2))

        aba_ausentes = _buscar_ausentes_prefixos(df1, df2_matriz_completa)

        aba_completa = comparacao.copy()
        aba_corretos = comparacao[comparacao["correto"] == True].copy()
        aba_incorretos = comparacao[comparacao["correto"] == False].copy()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            aba_completa.to_excel(writer, sheet_name="Comparacao_Completa", index=False)
            aba_corretos.to_excel(writer, sheet_name="Corretos", index=False)
            aba_incorretos.to_excel(writer, sheet_name="Incorretos", index=False)
            if not aba_ausentes.empty:
                aba_ausentes.to_excel(writer, sheet_name="Ausentes", index=False)
        output.seek(0)

        st.success("Comparação realizada!")
        st.download_button(
            label="Baixar resultado (xlsx)",
            data=output,
            file_name="comparacao_saldos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Prévia da comparação completa:")
        st.dataframe(aba_completa, use_container_width=True)

        st.subheader("Ausentes (prefixos 11111, 11131 e 114):")
        if aba_ausentes.empty:
            st.info("Nenhuma conta ausente encontrada para os prefixos informados.")
        else:
            st.warning(f"{len(aba_ausentes)} combinações conta/fonte ausentes em um dos arquivos.")
            st.dataframe(aba_ausentes, use_container_width=True)


if __name__ == "__main__":
    main() 