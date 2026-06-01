import io
import re
import unicodedata
from collections import defaultdict

import pandas as pd
import pdfplumber
import streamlit as st

COLUNA_NOME = "Nome Trabalhador"

# Rodapé / textos que não são nome
_INVALID_NAME_RE = re.compile(
    r"(reais|guia|fgts|trabalhadores|estabelecimento|empregador|expressos)",
    re.IGNORECASE,
)
_CPF_RE = re.compile(r"\d{3}\.\d{3}\.\d{3}-\d{2}")
_DATE_RE = re.compile(r"^\d{2}/\d{4}$")
_TRAILING_MATRICULA_RE = re.compile(r"(\d{4,6})$")


def _pdf_bytes(pdf_source) -> bytes:
    if hasattr(pdf_source, "read"):
        pdf_source.seek(0)
        data = pdf_source.read()
        pdf_source.seek(0)
        return data
    with open(pdf_source, "rb") as f:
        return f.read()


def _normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", str(text or ""))
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def _clean_word_token(token: str) -> str:
    """Remove matrícula colada ao nome (ex.: OLIVEIRA42714 -> OLIVEIRA)."""
    token = token.strip()
    if not token:
        return ""
    return _TRAILING_MATRICULA_RE.sub("", token).strip()


def _is_valid_name(name: str) -> bool:
    if not name or len(name) < 4:
        return False
    if _INVALID_NAME_RE.search(name):
        return False
    if _CPF_RE.search(name):
        return False
    if _DATE_RE.match(name):
        return False
    if sum(c.isalpha() for c in name) < 4:
        return False
    low = name.casefold()
    if "nome" in low and "trabalhador" in low:
        return False
    return True


def _find_column_bounds(words: list[dict]) -> tuple[float, float, float] | None:
    """
    Localiza a faixa horizontal da coluna 'Nome Trabalhador'
    (entre o cabeçalho Nome/Trabalhador e a coluna Matrícula).
    Retorna (col_left, col_right, header_bottom) ou None.
    """
    nome_x0 = None
    trabalhador_x1 = None
    header_top = None
    matricula_x0 = None

    for w in words:
        text = w["text"]
        top = w["top"]
        if text == "Nome":
            for w2 in words:
                if w2["text"] == "Trabalhador" and abs(w2["top"] - top) < 8:
                    nome_x0 = w["x0"]
                    trabalhador_x1 = w2["x1"]
                    header_top = top
        if "Matr" in text and 100 < top < 140:
            matricula_x0 = w["x0"]

    if nome_x0 is None or matricula_x0 is None:
        return None

    col_left = nome_x0 - 2
    col_right = matricula_x0 - 4
    header_bottom = (header_top or 121) + 14
    return col_left, col_right, header_bottom


def _names_from_column_words(pdf_bytes: bytes) -> list[str]:
    """
    Extrai nomes juntando todas as palavras dentro da coluna Nome Trabalhador.
    Funciona mesmo quando o PDF divide o nome em várias células/palavras.
    """
    names: list[str] = []
    seen: set[str] = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words() or []
            bounds = _find_column_bounds(words)
            if bounds is None:
                continue
            col_left, col_right, header_bottom = bounds

            lines: dict[float, list[dict]] = defaultdict(list)
            for w in words:
                if w["top"] <= header_bottom:
                    continue
                if w["x0"] >= col_right or w["x1"] <= col_left:
                    continue
                line_key = round(w["top"] / 3) * 3
                lines[line_key].append(w)

            for line_key in sorted(lines.keys()):
                line_words = sorted(lines[line_key], key=lambda x: x["x0"])
                parts = [_clean_word_token(w["text"]) for w in line_words]
                parts = [p for p in parts if p]
                while parts and re.fullmatch(r"\d{4,6}", parts[-1]):
                    parts.pop()
                name = _normalize_text(" ".join(parts))
                if not _is_valid_name(name):
                    continue
                if name not in seen:
                    seen.add(name)
                    names.append(name)

    return names


def _names_from_tables(pdf_bytes: bytes) -> list[str]:
    """Fallback: tabela padrão do pdfplumber (coluna Nome Trabalhador)."""
    names: list[str] = []
    seen: set[str] = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table or not table[0]:
                    continue
                header = [_normalize_text(c) for c in table[0]]
                try:
                    col_index = header.index(COLUNA_NOME)
                except ValueError:
                    col_index = next(
                        (
                            i
                            for i, h in enumerate(header)
                            if "nome" in h.casefold() and "trabalhador" in h.casefold()
                        ),
                        None,
                    )
                if col_index is None:
                    continue
                for row in table[1:]:
                    if not row or len(row) <= col_index:
                        continue
                    name = _normalize_text(row[col_index])
                    if _is_valid_name(name) and name not in seen:
                        seen.add(name)
                        names.append(name)

    return names


def _merge_name_lists(*lists: list[str]) -> list[str]:
    """Une listas priorizando o nome mais completo (mais longo)."""
    best: dict[str, str] = {}
    order: list[str] = []
    for names in lists:
        for name in names:
            key = name.casefold()
            if key not in best:
                best[key] = name
                order.append(key)
            elif len(name) > len(best[key]):
                best[key] = name
    return [best[k] for k in order]


def extract_worker_names(pdf_source) -> list[str]:
    pdf_bytes = _pdf_bytes(pdf_source)
    # Coluna por coordenadas (junta palavras divididas) + tabela padrão
    from_words = _names_from_column_words(pdf_bytes)
    from_tables = _names_from_tables(pdf_bytes)
    return _merge_name_lists(from_words, from_tables)


@st.cache_data
def convert_df_to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output.getvalue()


def main() -> None:
    st.title("Relação de Trabalhadores FGTS")

    uploaded_file = st.file_uploader("Faça upload do arquivo PDF", type="pdf")

    if uploaded_file is None:
        return

    names = extract_worker_names(uploaded_file)

    if not names:
        st.error(
            f"Não foi possível extrair a coluna **{COLUNA_NOME}** deste PDF."
        )
        return

    df = pd.DataFrame(names, columns=[COLUNA_NOME])

    st.success(f"**{len(df)}** nomes completos extraídos da coluna **{COLUNA_NOME}**.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        label="Baixar resultado em XLSX",
        data=convert_df_to_excel(df),
        file_name="nomes_trabalhadores.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
