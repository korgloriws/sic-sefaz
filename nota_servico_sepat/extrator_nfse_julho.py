from __future__ import annotations
import io, re, time
import streamlit as st
import pymupdf as fitz
from PIL import Image
import pytesseract
import pandas as pd


def br_to_float(s: str) -> float | None:
    if not s:
        return None
    try:
        s = re.sub(r"[^\d,.]", "", s)
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def ocr_page(pg: fitz.Page, dpi: int = 300) -> str:
    pix = pg.get_pixmap(dpi=dpi, colorspace="rgb")
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang="por")


def _value_on_line(line: str) -> float | None:
    match = re.search(r"(?:R\$\s*)?([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2}|[0-9]+[.,][0-9]{2})", line)
    if match:
        return br_to_float(match.group(1))
    return None


def _value_after_label(lines: list[str], label_idx: int) -> float | None:
    """No DANFSe v2.0 o valor costuma estar na linha seguinte ao rótulo."""
    same = _value_on_line(lines[label_idx])
    if same is not None:
        return same
    if label_idx + 1 < len(lines):
        return _value_on_line(lines[label_idx + 1])
    return None


def extract_values_from_text(text: str, page_num: int) -> dict:
    """Extrai valores do layout DANFSe v2.0 (formato julho/Senior e-Docs)."""
    values = {
        "Página": page_num,
        "VALOR TOTAL DO SERVIÇO": None,
        "Valor ISS retido": None,
        "Valor IR": None,
        "Valor INSS": None,
        "Valor líquido da NFS-e": None,
    }

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Valor bruto: "Valor da Operação / Serviço" → R$ X.XXX,XX (linha seguinte)
    for i, line in enumerate(lines):
        if re.search(r"Valor da Opera[cç][aã]o\s*/\s*Servi[cç]o", line, re.I):
            values["VALOR TOTAL DO SERVIÇO"] = _value_after_label(lines, i)
            break

    # Fallback: texto livre "Valor RS 2208.19" na discriminação
    if values["VALOR TOTAL DO SERVIÇO"] is None:
        m = re.search(r"Valor\s+RS\s*([0-9]+(?:\.[0-9]+)?)", text, re.I)
        if m:
            values["VALOR TOTAL DO SERVIÇO"] = float(m.group(1))

    # Líquido: "Valor Líquido da NFS-e" (sem "+ IBS/CBS") → linha seguinte
    for i, line in enumerate(lines):
        if re.search(r"Valor\s+[Ll][ií]quido\s+da\s+NFS-e\b", line, re.I) and "+ IBS" not in line:
            values["Valor líquido da NFS-e"] = _value_after_label(lines, i)
            break

    # Fallback: "VALOR LIQUIDO DA NOTA FISCAL RS 1793.05"
    if values["Valor líquido da NFS-e"] is None:
        m = re.search(
            r"VALOR\s+LIQUIDO\s+DA\s+NOTA\s+FISCAL\s+RS\s*([0-9]+(?:\.[0-9]+)?)",
            text,
            re.I,
        )
        if m:
            values["Valor líquido da NFS-e"] = float(m.group(1))

    if values["VALOR TOTAL DO SERVIÇO"] is not None:
        gross_value = values["VALOR TOTAL DO SERVIÇO"]
        values["Valor IR"] = gross_value * 0.048  # 4.8%
        values["Valor ISS retido"] = gross_value * 0.03  # 3%
        values["Valor INSS"] = gross_value * 0.11  # 11%

    return values


def extrair_valores(up_file) -> pd.DataFrame:
    pdf_bytes = up_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    registros: list[dict] = []
    t0 = time.perf_counter()

    for i in range(len(doc)):
        pg = doc[i]
        texto = pg.get_text("text") or ""
        if not texto.strip():
            texto = ocr_page(pg)

        # Só processa páginas do DANFSe v2.0 / nota de julho
        if not (
            "DANFSe" in texto
            or "Valor da Operação / Serviço" in texto
            or "Valor da Operacao / Servico" in texto
            or re.search(r"Valor da Opera[cç][aã]o\s*/\s*Servi[cç]o", texto, re.I)
        ):
            continue

        values = extract_values_from_text(texto, i + 1)

        if values["VALOR TOTAL DO SERVIÇO"] is not None:
            registros.append(values)

    doc.close()
    st.info(
        f"⏱️ Processadas {len(registros)} notas fiscais (formato julho/DANFSe) em "
        f"{time.perf_counter() - t0:.1f}s"
    )
    return pd.DataFrame(registros)


def format_currency(value):
    if pd.isna(value):
        return ""
    return f"R$ {value:,.2f}".replace(",", "·").replace(".", ",").replace("·", ".")


def main() -> None:
    st.title("NF-SEPAT - Serviços (formato julho / DANFSe v2.0)")

    up = st.file_uploader("Selecione o PDF", type=["pdf"])
    if not up:
        return

    with st.spinner("Processando documento…"):
        df = extrair_valores(up)

    if df.empty:
        st.warning("Nenhum valor encontrado no documento.")
        return

    st.subheader("Valores encontrados")

    display_df = df.copy()
    for col in display_df.columns:
        if col != "Página":
            display_df[col] = display_df[col].apply(format_currency)

    st.dataframe(display_df, hide_index=True, use_container_width=True)

    totals = {}
    for col in df.columns:
        if col != "Página":
            totals[col] = df[col].sum()

    st.subheader("Totais")
    for field, total in totals.items():
        st.success(f"**{field}:** {format_currency(total)}")

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Valores", index=False)

    excel_data = excel_buffer.getvalue()
    st.download_button(
        label="📥 Download Excel",
        data=excel_data,
        file_name="valores_nfse_julho.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
