import os
import re
import tempfile
import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from pathlib import Path


def extract_document_number(text):

    pattern = r"NÚMERO DO DOCUMENTO:\s*(\d+)"
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return None


def format_nap_number(document_number: str) -> str:
    """Remove 4 primeiros dígitos (data) e 2 últimos do número do documento."""
    digits = re.sub(r"\D", "", str(document_number))
    if len(digits) > 6:
        return digits[4:-2]
    return digits


def build_output_filename(document_number: str) -> str:
    nap = format_nap_number(document_number)
    return f"comprovante_NAP_{nap}.pdf"


def extract_text_from_page(pdf_path, page_number):

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_number < len(pdf.pages):
                page = pdf.pages[page_number]
                text = page.extract_text()
                return text
    except Exception as e:
        st.error(f"Erro ao extrair texto da página {page_number + 1}: {str(e)}")
        return None
    return None


def process_pdf_page(pdf_path, page_number, output_dir):

    page_text = extract_text_from_page(pdf_path, page_number)
    
    if not page_text:
        return None, f"Erro ao processar página {page_number + 1}"
    

    document_number = extract_document_number(page_text)
    
    if not document_number:
        return None, f"Número do documento não encontrado na página {page_number + 1}"
    

    output_filename = build_output_filename(document_number)
    output_path = os.path.join(output_dir, output_filename)
    

    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            
            if page_number < len(pdf_reader.pages):
                pdf_writer = PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[page_number])
                
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                
                return output_path, document_number
            else:
                return None, f"Página {page_number + 1} não existe no PDF"
                
    except Exception as e:
        return None, f"Erro ao separar página {page_number + 1}: {str(e)}"


def process_pdf_file(pdf_path, output_dir):

    results = []
    
    try:

        if not os.path.exists(pdf_path):
            return [], "Arquivo PDF não encontrado"
        

        with open(pdf_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            total_pages = len(pdf_reader.pages)
        

        for page_number in range(total_pages):
            output_path, document_number = process_pdf_page(pdf_path, page_number, output_dir)
            
            if output_path:
                results.append({
                    'page': page_number + 1,
                    'document_number': document_number,
                    'output_file': output_path,
                    'status': 'success'
                })
            else:
                results.append({
                    'page': page_number + 1,
                    'document_number': None,
                    'output_file': None,
                    'status': 'error',
                    'error': document_number  
                })
        
        return results, None
        
    except Exception as e:
        return [], f"Erro ao processar PDF: {str(e)}"


def main():


    
    st.title(" Comprovante de Remessa")
    st.markdown("---")
    

    uploaded_file = st.file_uploader(
        "Selecione o arquivo PDF para processar",
        type=['pdf'],
        help="Arraste e solte ou clique para selecionar um arquivo PDF"
    )
    
    if uploaded_file is not None:
        with tempfile.TemporaryDirectory() as temp_dir:

            temp_pdf_path = os.path.join(temp_dir, uploaded_file.name)
            with open(temp_pdf_path, 'wb') as f:
                f.write(uploaded_file.getvalue())
            
            st.success(f" Arquivo carregado: {uploaded_file.name}")
            

            if st.button(" Processar PDF", type="primary"):
                with st.spinner("Processando PDF..."):
                    results, error = process_pdf_file(temp_pdf_path, temp_dir)
                
                if error:
                    st.error(f" {error}")
                else:
                    st.success(f" Processamento concluído! {len(results)} páginas processadas.")
                    

                    st.subheader(" Resultados do Processamento")
                    

                    success_count = 0
                    error_count = 0
                    
                    for result in results:
                        if result['status'] == 'success':
                            success_count += 1
                            st.success(
                                f" Página {result['page']}:  {result['document_number']} "
                                f"- Arquivo: {os.path.basename(result['output_file'])}"
                            )
                        else:
                            error_count += 1
                            st.error(f" Página {result['page']}: {result['error']}")
                    

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total de Páginas", len(results))
                    with col2:
                        st.metric("Processadas com Sucesso", success_count)
                    with col3:
                        st.metric("Erros", error_count)
                    

                    if success_count > 0:
                        st.subheader(" Download dos Arquivos")
                        

                        import zipfile
                        import io
                        
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for result in results:
                                if result['status'] == 'success':
                                    zip_file.write(
                                        result['output_file'],
                                        os.path.basename(result['output_file'])
                                    )
                        
                        zip_buffer.seek(0)
                        st.download_button(
                            label=" Download de Todos os Arquivos (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name="documentos_processados.zip",
                            mime="application/zip"
                        )
                        

                        st.subheader(" Download Individual")
                        for result in results:
                            if result['status'] == 'success':
                                with open(result['output_file'], 'rb') as f:
                                    st.download_button(
                                        label=f" {os.path.basename(result['output_file'])}",
                                        data=f.read(),
                                        file_name=os.path.basename(result['output_file']),
                                        mime="application/pdf"
                                    )


if __name__ == "__main__":
    main() 