import io
import re
import unicodedata
from collections import defaultdict

import pandas as pd
import pdfplumber
import streamlit as st

COLUNA_NOME = "Nome Trabalhador"

_INVALID_NAME_RE = re.compile(
    r"(reais|guia|fgts|trabalhadores|estabelecimento|empregador|expressos)",
    re.IGNORECASE,
)
_CPF_RE = re.compile(r"\d{3}\.\d{3}\.\d{3}-\d{2}")
_DATE_RE = re.compile(r"^\d{2}/\d{4}$")
_TRAILING_MATRICULA_RE = re.compile(r"(\d{4,6})$")
_LONG_DIGITS_RE = re.compile(r"^\d{10,}$")


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


def _normalize_word(text: str) -> str:
    return _normalize_text(text)


def _clean_word_token(token: str) -> str:
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
    if _LONG_DIGITS_RE.match(name.replace(" ", "")):
        return False
    if sum(c.isalpha() for c in name) < 4:
        return False
    low = name.casefold()
    if "nome" in low and "trabalhador" in low:
        return False
    return True


def _is_nome_trabalhador_header(label: str) -> bool:
    norm = _normalize_text(label).casefold()
    return norm == COLUNA_NOME.casefold() or (
        "nome" in norm and "trabalhador" in norm
    )


def _find_column_bounds(words: list[dict]) -> tuple[float, float, float] | None:
    nome_x0 = None
    header_top = None
    matricula_x0 = None

    for w in words:
        text = _normalize_word(w["text"])
        top = w["top"]
        if text == "Nome":
            for w2 in words:
                if _normalize_word(w2["text"]) == "Trabalhador" and abs(w2["top"] - top) < 10:
                    nome_x0 = w["x0"]
                    header_top = top
        if "matr" in text.casefold() and 90 < top < 150:
            matricula_x0 = w["x0"]

    if nome_x0 is None or matricula_x0 is None:
        return None

    return nome_x0 - 2, matricula_x0 - 4, (header_top or 121) + 14


def _score_name_column(table: list[list], col_index: int, start_row: int) -> int:
    score = 0
    for row in table[start_row + 1 : start_row + 25]:
        if not row or len(row) <= col_index:
            continue
        name = _normalize_text(row[col_index])
        if _is_valid_name(name):
            score += 1
    return score


def _resolve_name_column_index(table: list[list]) -> tuple[int | None, int]:
    """Acha coluna do nome pelo cabeçalho; se vazia, testa colunas vizinhas (FGTS.pdf)."""
    header_row_idx = 0
    header_row = table[0] if table else []
    header_col = None

    for row_idx, row in enumerate(table[:4]):
        if not row:
            continue
        for col_idx, cell in enumerate(row):
            if _is_nome_trabalhador_header(cell):
                header_row_idx = row_idx
                header_row = row
                header_col = col_idx
                break
        if header_col is not None:
            break

    if header_col is None:
        return None, 0

    candidates = [header_col]
    for offset in (-1, 1, -2, 2):
        alt = header_col + offset
        if 0 <= alt < len(header_row) and alt not in candidates:
            candidates.append(alt)

    best_col = header_col
    best_score = -1
    for col in candidates:
        score = _score_name_column(table, col, header_row_idx)
        if score > best_score:
            best_score = score
            best_col = col

    return best_col, header_row_idx


def _extract_names_from_pdf_bytes(pdf_bytes: bytes) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()

    def add_name(name: str) -> None:
        if _is_valid_name(name) and name not in seen:
            seen.add(name)
            names.append(name)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # --- Método 1: palavras por coordenada da coluna ---
            words = page.extract_words() or []
            bounds = _find_column_bounds(words)
            if bounds:
                col_left, col_right, header_bottom = bounds
                lines: dict[float, list[dict]] = defaultdict(list)
                for w in words:
                    if w["top"] <= header_bottom:
                        continue
                    if w["x0"] >= col_right or w["x1"] <= col_left:
                        continue
                    lines[round(w["top"] / 3) * 3].append(w)

                for line_key in sorted(lines.keys()):
                    line_words = sorted(lines[line_key], key=lambda x: x["x0"])
                    parts = [_clean_word_token(w["text"]) for w in line_words]
                    parts = [p for p in parts if p]
                    while parts and re.fullmatch(r"\d{4,6}", parts[-1]):
                        parts.pop()
                    add_name(_normalize_text(" ".join(parts)))

            # --- Método 2: tabela (coluna pode estar deslocada no FGTS.pdf) ---
            for table in page.extract_tables() or []:
                if not table:
                    continue
                col_index, header_row = _resolve_name_column_index(table)
                if col_index is None:
                    continue
                for row in table[header_row + 1 :]:
                    if not row or len(row) <= col_index:
                        continue
                    add_name(_normalize_text(row[col_index]))

    return names


@st.cache_data(show_spinner=False)
def extract_worker_names_cached(pdf_bytes: bytes) -> tuple[list[str], str]:
    try:
        names = _extract_names_from_pdf_bytes(pdf_bytes)
        return names, ""
    except Exception as exc:
        return [], str(exc)


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

    pdf_bytes = _pdf_bytes(uploaded_file)

    with st.spinner("Extraindo nomes do PDF (pode levar alguns segundos em arquivos grandes)..."):
        names, error = extract_worker_names_cached(pdf_bytes)

    if error:
        st.error(f"Erro ao ler o PDF: {error}")
        return

    if not names:
        st.error(f"Não foi possível extrair a coluna **{COLUNA_NOME}** deste PDF.")
        with st.expander("Diagnóstico (1ª página)"):
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    if pdf.pages:
                        page = pdf.pages[0]
                        st.write("Palavras (amostra):", page.extract_words()[:15])
                        tables = page.extract_tables() or []
                        st.write(f"Tabelas: {len(tables)}")
                        if tables:
                            st.write("Cabeçalho:", tables[0][0])
                            st.write("1ª linha:", tables[0][1] if len(tables[0]) > 1 else None)
                            col, _ = _resolve_name_column_index(tables[0])
                            st.write(f"Coluna detectada para nomes: índice {col}")
            except Exception as diag_exc:
                st.write("Falha no diagnóstico:", diag_exc)
        return

    df = pd.DataFrame(names, columns=[COLUNA_NOME])

    st.success(f"**{len(df)}** nomes extraídos da coluna **{COLUNA_NOME}**.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        label="Baixar resultado em XLSX",
        data=convert_df_to_excel(df),
        file_name="nomes_trabalhadores.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
