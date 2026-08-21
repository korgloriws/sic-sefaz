"""
Modulo de calculo de recurso disponivel a partir de arquivo MSC (CSV).

Pode ser executado com:
    streamlit run app.py

Ou importado em um sistema maior:
    from app import main, calcular_recurso_disponivel, calcular_recurso_disponivel_por_ug_fonte
"""

from __future__ import annotations

import io
from typing import BinaryIO, Union

import pandas as pd
import streamlit as st

COLUNAS_OBRIGATORIAS = [
    "CONTA",
    "IC2",
    "IC3",
    "Valor",
    "Tipo_valor",
    "Natureza_valor",
]

COLUNAS_IC_TIPO = [
    "IC1",
    "TIPO1",
    "IC2",
    "TIPO2",
    "IC3",
    "TIPO3",
    "IC4",
    "TIPO4",
    "IC5",
    "TIPO5",
    "IC6",
    "TIPO6",
]

NOMES_CONDICOES = {
    1: "ativo financeiro",
    2: "passivo financeiro",
    3: "credito empenhado a liquidar",
    4: "RP nao processados a liquidar",
    5: "Diferenca",
}


def _decodificar_bytes(raw: bytes) -> str:
    for encoding in ("iso-8859-1", "latin-1", "cp1252", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("iso-8859-1", errors="replace")


def _normalizar_conta(valor: object) -> str:
    texto = str(valor).strip() if valor is not None else ""
    if texto.lower() in {"", "nan", "none"}:
        return ""

    upper = texto.upper().replace(" ", "")
    if "E+" in upper or "E-" in upper:
        try:
            return f"{float(upper.replace(',', '.')):.0f}"
        except ValueError:
            return texto

    if "," in texto and "." not in texto:
        # numero brasileiro sem notacao cientifica
        pass

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def _normalizar_valor(valor: object) -> float:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nan", "none"}:
        return 0.0

    # 1.234.567,89 -> 1234567.89 | 1234567.89 permanece
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return 0.0


def _texto_limpo(valor: object) -> str:
    texto = str(valor).strip() if valor is not None else ""
    if texto.lower() in {"", "nan", "none"}:
        return ""
    return texto


def _extrair_fonte(row: pd.Series) -> str:
    """
    Na MSC a fonte (FR) pode estar em qualquer par IC/TIPO.
    Ex.: contas 1/2 com IC2=F usam IC3; conta 8211101 costuma trazer FR em IC2.
    """
    for indice in range(1, 7):
        tipo = _texto_limpo(row.get(f"TIPO{indice}", "")).upper()
        if tipo == "FR":
            return _texto_limpo(row.get(f"IC{indice}", ""))

    # Fallback: IC3, senao IC2 (sem usar o marcador financeiro F)
    ic3 = _texto_limpo(row.get("IC3", ""))
    if ic3:
        return ic3
    ic2 = _texto_limpo(row.get("IC2", ""))
    if ic2.upper() != "F":
        return ic2
    return ""


def carregar_csv(
    arquivo: Union[str, BinaryIO, bytes],
    remover_primeira_linha: bool = True,
) -> pd.DataFrame:
    """
    Carrega CSV ISO com separador ';'.
    Por padrao remove a primeira linha do arquivo (titulo/metadado)
    e usa a linha seguinte como cabecalho.
    """
    if isinstance(arquivo, bytes):
        raw = arquivo
    elif isinstance(arquivo, str):
        with open(arquivo, "rb") as handle:
            raw = handle.read()
    else:
        raw = arquivo.read()
        if isinstance(raw, str):
            raw = raw.encode("iso-8859-1")

    texto = _decodificar_bytes(raw)
    linhas = texto.splitlines()
    if not linhas:
        raise ValueError("Arquivo CSV vazio.")

    if remover_primeira_linha and len(linhas) >= 2:
        primeira = linhas[0].upper()
        # Se a primeira linha ja for cabecalho, nao remove.
        if "CONTA" not in primeira:
            linhas = linhas[1:]

    conteudo = "\n".join(linhas)
    df = pd.read_csv(
        io.StringIO(conteudo),
        sep=";",
        dtype=str,
        keep_default_na=False,
    )
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Colunas obrigatorias nao encontradas: " + ", ".join(faltantes)
        )

    colunas_usar = ["CONTA", "Valor", "Tipo_valor", "Natureza_valor"]
    for coluna in COLUNAS_IC_TIPO:
        if coluna in df.columns and coluna not in colunas_usar:
            colunas_usar.append(coluna)

    df = df[colunas_usar].copy()
    df["CONTA"] = df["CONTA"].map(_normalizar_conta)
    for coluna in COLUNAS_IC_TIPO:
        if coluna in df.columns:
            df[coluna] = df[coluna].map(_texto_limpo)
    df["Tipo_valor"] = df["Tipo_valor"].map(_texto_limpo)
    df["Natureza_valor"] = df["Natureza_valor"].map(_texto_limpo).str.upper()
    df["Valor"] = df["Valor"].map(_normalizar_valor)
    df["FONTE"] = df.apply(_extrair_fonte, axis=1)
    # Mantem IC3 preenchido com a fonte resolvida para compatibilidade da saida
    df["IC3"] = df["FONTE"]
    return df


def _saldo_agrupado(
    df: pd.DataFrame,
    formula: str,
    chaves_grupo: list[str],
) -> pd.Series:
    """
    Agrupa pelas chaves informadas, soma D e C, e aplica C-D ou D-C.
    """
    if df.empty:
        return pd.Series(dtype=float)

    base = df.copy()
    if "FONTE" not in base.columns:
        base["FONTE"] = base.get("IC3", "")
    if "UG" not in base.columns and "IC1" in base.columns:
        base["UG"] = base["IC1"].map(_texto_limpo)

    for chave in chaves_grupo:
        if chave not in base.columns:
            raise ValueError(f"Coluna de agrupamento ausente: {chave}")
        base = base[base[chave].map(_texto_limpo) != ""]

    if base.empty:
        return pd.Series(dtype=float)

    agrupado = (
        base.groupby(chaves_grupo + ["Natureza_valor"], dropna=False)["Valor"]
        .sum()
        .unstack(fill_value=0.0)
    )
    for natureza in ("D", "C"):
        if natureza not in agrupado.columns:
            agrupado[natureza] = 0.0

    if formula == "C-D":
        resultado = agrupado["C"] - agrupado["D"]
    elif formula == "D-C":
        resultado = agrupado["D"] - agrupado["C"]
    else:
        raise ValueError(f"Formula invalida: {formula}")

    resultado.name = "saldo"
    return resultado


def _filtrar_condicao(
    df: pd.DataFrame,
    prefixo_conta: str,
    formula: str,
    exigir_ic2_f: bool = False,
    chaves_grupo: list[str] | None = None,
) -> pd.Series:
    mascara = (
        df["CONTA"].str.startswith(prefixo_conta, na=False)
        & (df["Tipo_valor"] == "ending_balance")
    )
    # IC2 = F somente nas condicoes 1 e 2 (ativo/passivo financeiro)
    if exigir_ic2_f:
        mascara = mascara & (df["IC2"] == "F")

    if chaves_grupo is None:
        chaves_grupo = ["FONTE"]

    return _saldo_agrupado(
        df.loc[mascara],
        formula=formula,
        chaves_grupo=chaves_grupo,
    )


def _montar_resultado(
    c1: pd.Series,
    c2: pd.Series,
    c3: pd.Series,
    c4: pd.Series,
    c5: pd.Series,
    chaves_indice: list[str],
) -> pd.DataFrame:
    if len(chaves_indice) == 1:
        chaves_todas = {
            chave
            for chave in (
                set(c1.index) | set(c2.index) | set(c3.index) | set(c4.index) | set(c5.index)
            )
            if str(chave).strip()
        }
        indice = sorted(chaves_todas)
        resultado = pd.DataFrame(index=indice)
        resultado.index.name = chaves_indice[0]
    else:
        frames = []
        for serie in (c1, c2, c3, c4, c5):
            if serie.empty:
                continue
            aux = serie.reset_index()
            # reset_index cria colunas com nomes do MultiIndex ou 'index'/'level_*'
            if isinstance(serie.index, pd.MultiIndex):
                aux.columns = list(serie.index.names) + ["saldo"]
            else:
                aux.columns = [chaves_indice[0], "saldo"]
            frames.append(aux[chaves_indice].drop_duplicates())
        if not frames:
            return pd.DataFrame(columns=chaves_indice + list(NOMES_CONDICOES.values()))
        chaves_df = pd.concat(frames, ignore_index=True).drop_duplicates()
        chaves_df = chaves_df.sort_values(chaves_indice).reset_index(drop=True)
        resultado = chaves_df.set_index(chaves_indice)

    def _reindexar(serie: pd.Series) -> pd.Series:
        if serie.empty:
            return pd.Series(0.0, index=resultado.index)
        return serie.reindex(resultado.index).fillna(0.0)

    resultado[NOMES_CONDICOES[1]] = _reindexar(c1)
    resultado[NOMES_CONDICOES[2]] = _reindexar(c2)
    resultado[NOMES_CONDICOES[3]] = _reindexar(c3)
    resultado[NOMES_CONDICOES[4]] = _reindexar(c4)

    resultado["resultado_1_2_3_4"] = (
        resultado[NOMES_CONDICOES[1]]
        - resultado[NOMES_CONDICOES[2]]
        - resultado[NOMES_CONDICOES[3]]
        - resultado[NOMES_CONDICOES[4]]
    )

    valores_condicao_5 = _reindexar(c5)
    resultado["recurso_disponivel"] = (
        resultado["resultado_1_2_3_4"] - valores_condicao_5
    )
    # Diferenca (condicao 5) fica como ultima coluna
    resultado[NOMES_CONDICOES[5]] = valores_condicao_5

    saida = resultado.reset_index()
    # Compatibilidade: comparacao por fonte usa coluna IC3
    if chaves_indice == ["FONTE"] and "FONTE" in saida.columns:
        saida = saida.rename(columns={"FONTE": "IC3"})
    return saida


def calcular_recurso_disponivel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Comparacao 1: aplica as 5 condicoes por fonte e calcula:
      (1 - 2 - 3 - 4) - 5
    """
    chaves = ["FONTE"]
    c1 = _filtrar_condicao(df, "1", "D-C", exigir_ic2_f=True, chaves_grupo=chaves)
    c2 = _filtrar_condicao(df, "2", "C-D", exigir_ic2_f=True, chaves_grupo=chaves)
    c3 = _filtrar_condicao(df, "6221301", "C-D", exigir_ic2_f=False, chaves_grupo=chaves)
    c4 = _filtrar_condicao(df, "6311", "C-D", exigir_ic2_f=False, chaves_grupo=chaves)
    c5 = _filtrar_condicao(df, "8211101", "C-D", exigir_ic2_f=False, chaves_grupo=chaves)
    return _montar_resultado(c1, c2, c3, c4, c5, chaves_indice=chaves)


def calcular_recurso_disponivel_por_ug_fonte(df: pd.DataFrame) -> pd.DataFrame:
    """
    Comparacao 2: mesmas regras, detalhando por UG (IC1) e fonte.
      (1 - 2 - 3 - 4) - 5
    """
    base = df.copy()
    if "IC1" not in base.columns:
        raise ValueError("Coluna IC1 (UG) nao encontrada no arquivo.")
    base["UG"] = base["IC1"].map(_texto_limpo)

    chaves = ["UG", "FONTE"]
    c1 = _filtrar_condicao(base, "1", "D-C", exigir_ic2_f=True, chaves_grupo=chaves)
    c2 = _filtrar_condicao(base, "2", "C-D", exigir_ic2_f=True, chaves_grupo=chaves)
    c3 = _filtrar_condicao(base, "6221301", "C-D", exigir_ic2_f=False, chaves_grupo=chaves)
    c4 = _filtrar_condicao(base, "6311", "C-D", exigir_ic2_f=False, chaves_grupo=chaves)
    c5 = _filtrar_condicao(base, "8211101", "C-D", exigir_ic2_f=False, chaves_grupo=chaves)
    return _montar_resultado(c1, c2, c3, c4, c5, chaves_indice=chaves)


def _formatar_moeda_br(valor: float) -> str:
    negativo = valor < 0
    texto = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-{texto}" if negativo else texto


def _avisos_contas(df: pd.DataFrame, resultado: pd.DataFrame) -> list[str]:
    avisos = []
    if (
        resultado[NOMES_CONDICOES[3]].eq(0).all()
        and df["CONTA"].str.startswith("62213").any()
        and not df["CONTA"].str.startswith("6221301").any()
    ):
        avisos.append(
            "Nenhuma conta iniciando em 6221301 foi encontrada, mas ha contas "
            "proximas (62213...). Verifique se a coluna CONTA foi exportada "
            "como texto (sem notacao cientifica)."
        )
    if (
        resultado[NOMES_CONDICOES[5]].eq(0).all()
        and df["CONTA"].str.startswith("82111").any()
        and not df["CONTA"].str.startswith("8211101").any()
    ):
        avisos.append(
            "Nenhuma conta iniciando em 8211101 foi encontrada, mas ha contas "
            "proximas (82111...). Verifique se a coluna CONTA foi exportada "
            "como texto (sem notacao cientifica)."
        )
    return avisos


def _renderizar_resultado(
    resultado: pd.DataFrame,
    titulo: str,
    metrica_chave: str,
    metrica_label: str,
    arquivo_csv: str,
) -> None:
    st.subheader(titulo)

    col1, col2, col3 = st.columns(3)
    col1.metric(metrica_label, f"{resultado[metrica_chave].nunique()}")
    col2.metric(
        "Total resultado (1-2-3-4)",
        _formatar_moeda_br(float(resultado["resultado_1_2_3_4"].sum())),
    )
    col3.metric(
        "Total recurso disponivel",
        _formatar_moeda_br(float(resultado["recurso_disponivel"].sum())),
    )

    st.dataframe(
        resultado,
        use_container_width=True,
        hide_index=True,
    )

    csv_bytes = resultado.to_csv(index=False, sep=";", decimal=",").encode(
        "iso-8859-1",
        errors="replace",
    )
    st.download_button(
        label=f"Baixar relatorio: {titulo}",
        data=csv_bytes,
        file_name=arquivo_csv,
        mime="text/csv",
    )


def main() -> None:
    st.set_page_config(
        page_title="Calculo de Recursos Disponiveis",
        layout="wide",
    )

    st.title("Calculo de Recursos Disponiveis")
    st.caption(
        "Importacao de MSC em CSV (ISO / separador ponto e virgula). "
        "Duas comparacoes: por fonte e por UG + fonte."
    )

    with st.expander("Regras aplicadas", expanded=False):
        st.markdown(
            """
1. **ativo financeiro** — CONTA inicia com `1`, IC2 = `F`, Tipo_valor = `ending_balance`, saldo `D - C`  
2. **passivo financeiro** — CONTA inicia com `2`, IC2 = `F`, Tipo_valor = `ending_balance`, saldo `C - D`  
3. **credito empenhado a liquidar** — CONTA inicia com `6221301`, Tipo_valor = `ending_balance`, saldo `C - D`, sem filtro IC2  
4. **RP nao processados a liquidar** — CONTA inicia com `6311`, Tipo_valor = `ending_balance`, saldo `C - D`, sem filtro IC2  
5. **Diferenca** — CONTA inicia com `8211101`, Tipo_valor = `ending_balance`, saldo `C - D`, sem filtro IC2  

A fonte e lida do par IC/TIPO marcado como `FR` (pode estar em IC2 ou IC3).  
A UG e lida de `IC1`.  

**Comparacao 1 — por fonte:** totais por fonte `(1 - 2 - 3 - 4) - 5`.  
**Comparacao 2 — por UG e fonte:** mesmas regras, detalhadas por UG (`IC1`) + fonte.
            """
        )

    arquivo = st.file_uploader(
        "Selecione o arquivo CSV",
        type=["csv"],
        help="Arquivo no formato ISO com separador ';'.",
    )

    if arquivo is None:
        st.info("Envie um arquivo CSV para iniciar o calculo.")
        return

    try:
        df = carregar_csv(arquivo, remover_primeira_linha=True)
        resultado_fonte = calcular_recurso_disponivel(df)
        resultado_ug_fonte = calcular_recurso_disponivel_por_ug_fonte(df)
    except Exception as exc:
        st.error(f"Nao foi possivel processar o arquivo: {exc}")
        return

    st.success(f"Arquivo processado: {len(df)} linhas validas apos leitura.")

    for aviso in _avisos_contas(df, resultado_fonte):
        st.warning(aviso)

    aba_fonte, aba_ug = st.tabs(
        [
            "Comparacao 1 - Por fonte",
            "Comparacao 2 - Por UG e fonte",
        ]
    )

    with aba_fonte:
        st.caption(
            "Proposta original: localiza a diferenca agregando somente por fonte."
        )
        _renderizar_resultado(
            resultado_fonte,
            titulo="Resultado por fonte",
            metrica_chave="IC3",
            metrica_label="Fontes",
            arquivo_csv="recurso_disponivel_por_fonte.csv",
        )

    with aba_ug:
        st.caption(
            "Segunda proposta: mesmo calculo, detalhado por UG (IC1) e fonte, "
            "para ajudar a localizar onde esta a diferenca."
        )
        _renderizar_resultado(
            resultado_ug_fonte,
            titulo="Resultado por UG e fonte",
            metrica_chave="UG",
            metrica_label="UGs",
            arquivo_csv="recurso_disponivel_por_ug_fonte.csv",
        )


if __name__ == "__main__":
    main()
