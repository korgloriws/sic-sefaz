
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

def ajustar_fonte_recurso(codigo):
    codigo_int = int(codigo)
    if codigo_int > 9999999:
        codigo_int = int(str(codigo_int)[1:])
    codigo_int = codigo_int // 1000 * 1000 
    return codigo_int

def ajustar_valor(valor):
    if isinstance(valor, str):
        valor = valor.replace('.', '')
        valor = valor.replace(',', '.')
    return abs(float(valor))


def somente_digitos(valor):
    if valor is None:
        return ''
    # Trata tipos numéricos evitando o efeito "1501000.0" -> "15010000"
    try:
        if isinstance(valor, (int, np.integer)):
            return re.sub(r'\D+', '', str(int(valor)))
        if isinstance(valor, (float, np.floating)) and np.isfinite(valor):
            valor_int = int(round(valor))
            if abs(valor - valor_int) < 1e-9:
                return re.sub(r'\D+', '', str(valor_int))
    except Exception:
        pass
    s = str(valor).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return re.sub(r'\D+', '', s)

def receita_norm_rec(valor):
    # REC: mantém como está (apenas dígitos). Se não começa com 9, usa 8 dígitos; se começa com 9, mantém.
    s = somente_digitos(valor)
    if not s:
        return ''
    if not s.startswith('9'):
        return s[:8]
    return s[:10]

def receita_norm_qgr(valor):

    s = somente_digitos(valor)
    if not s:
        return ''
    if s.startswith('9'):
        # manter 10 dígitos (ou o que houver, mas tipicamente 10)
        return s[:10]
    # não começa com 9: remove últimos 2 dígitos
    return s[:-2] if len(s) > 2 else ''

def fonte_nucleo_rec(valor):
    s = somente_digitos(valor)
    return s[:4] if len(s) >= 4 else ''

def fonte_nucleo_qgr(valor):
    s = somente_digitos(valor)
    # QGR: se mais de 7 dígitos, descarta o primeiro, depois considera os 4 primeiros
    if len(s) > 7:
        s = s[1:]
    return s[:4] if len(s) >= 4 else ''

def co_norm(valor):
    s = somente_digitos(valor)
    # Se não começa com '3', considerar como '0000'
    if not s or not s.startswith('3'):
        return '0000'
    # Começa com '3': manter padronizado para 4 dígitos
    if len(s) >= 4:
        return s[:4]
    return s.zfill(4)

def to_num(valor):
    if isinstance(valor, (int, float)) and np.isfinite(valor):
        return float(valor)
    s = str(valor).strip() if valor is not None else ''
    if not s:
        return 0.0
    t = re.sub(r'\s+', '', s)
    try:
        if ('.' in t) and (',' in t):
            return float(t.replace('.', '').replace(',', '.'))
        if (',' in t) and ('.' not in t):
            return float(t.replace(',', '.'))
        return float(t)
    except Exception:
        return 0.0

def approx_equal(a, b, eps=0.005):
    try:
        return abs(float(a) - float(b)) <= eps
    except Exception:
        return False

def detectar_coluna_valor(df):

    if 'VR_ARREC_MES_FONTE' in df.columns:
        return 'VR_ARREC_MES_FONTE'
    if 'VR_ARREC_MES' in df.columns:
        return 'VR_ARREC_MES'
    return None

def agregar_por_trinca(df, is_qgr=False):
   
    valor_col = detectar_coluna_valor(df)
    if valor_col is None:
        return pd.DataFrame(columns=['CODIGO_RECEITA', 'FONTE_NUCLEO', 'CO', 'VALOR'])

    tmp = pd.DataFrame({
        'CODIGO_RECEITA': df.get('CODIGO_RECEITA', ''),
        'FONTE_RECURSO': df.get('FONTE_RECURSO', ''),
        'CO': df.get('CO', ''),
        'VALOR_RAW': df.get(valor_col, 0)
    })


    if is_qgr:
        tmp['CODIGO_RECEITA'] = tmp['CODIGO_RECEITA'].apply(receita_norm_qgr)
        tmp['FONTE_NUCLEO'] = tmp['FONTE_RECURSO'].apply(fonte_nucleo_qgr)
    else:
        tmp['CODIGO_RECEITA'] = tmp['CODIGO_RECEITA'].apply(receita_norm_rec)
        tmp['FONTE_NUCLEO'] = tmp['FONTE_RECURSO'].apply(fonte_nucleo_rec)
    tmp['CO'] = tmp['CO'].apply(co_norm)
    tmp['VALOR'] = tmp['VALOR_RAW'].apply(to_num)


    if is_qgr:
        try:
            mask_nove = tmp['CODIGO_RECEITA'].astype(str).str.startswith('9')
            tmp.loc[mask_nove, 'VALOR'] = tmp.loc[mask_nove, 'VALOR'].abs()
        except Exception:
            pass

    
    mask_chaves = (tmp['CODIGO_RECEITA'] != '') & (tmp['FONTE_NUCLEO'] != '') & (tmp['CO'] != '')
    mask_valores = tmp['VALOR'] >= 0
    if is_qgr:
        mask_valores = mask_valores & (tmp['VALOR'] > 0)
    tmp = tmp[mask_chaves & mask_valores]

    agrupado = (
        tmp.groupby(['CODIGO_RECEITA', 'FONTE_NUCLEO', 'CO'], as_index=False)['VALOR']
        .sum()
        .sort_values(['CODIGO_RECEITA', 'FONTE_NUCLEO', 'CO'])
    )
    return agrupado

def formatar_ptbr(valor):
    try:
        return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return str(valor)

def validar_registros_10_11(df_rec):

    discrepancias_internas = []
    
  
    registros_10 = df_rec[df_rec.iloc[:, 0] == 10].copy()  
    registros_11 = df_rec[df_rec.iloc[:, 0] == 11].copy()  
    
    if registros_11.empty or registros_10.empty:
        return discrepancias_internas
    
    
    temp_df = pd.DataFrame({
        'CODIGO_RECEITA': registros_11.iloc[:, 1],  
        'VALOR': registros_11.iloc[:, 10]  
    })
    
  
    registros_11_agrupados = temp_df.groupby('CODIGO_RECEITA')['VALOR'].sum().reset_index()
    registros_11_agrupados.columns = ['CODIGO_RECEITA', 'VALOR_REGISTRO_11']
    
   
    for _, row in registros_11_agrupados.iterrows():
        codigo_receita = row['CODIGO_RECEITA']
        valor_soma_11 = row['VALOR_REGISTRO_11']
        
 
        registro_10 = registros_10[registros_10.iloc[:, 1] == codigo_receita]
        
        if not registro_10.empty:
            valor_registro_10 = registro_10.iloc[0, 6]  
            

            valor_soma_11 = ajustar_valor(valor_soma_11)
            valor_registro_10 = ajustar_valor(valor_registro_10)
            

            if abs(valor_soma_11) != abs(valor_registro_10):
                discrepancias_internas.append({
                    'CODIGO_RECEITA': codigo_receita,
                    'VALOR_REGISTRO_10': valor_registro_10,
                    'SOMA_REGISTROS_11': valor_soma_11,
                    'DIFERENCA': abs(valor_soma_11) - abs(valor_registro_10)
                })
    
    return discrepancias_internas

def validar_co(df_rec_original, df_geral_original):
    
  
    if 'CO' not in df_rec_original.columns or 'CO' not in df_geral_original.columns:
        return pd.DataFrame()

    excecoes = [21710010, 21749012, 21759005]

    def try_to_int(valor):
        try:
            if pd.isna(valor):
                return None
        except Exception:
            pass
        texto = str(valor).strip()
        try:
            # tenta lidar com strings tipo '21710010', '21710010.0', '21.710.010'
            texto_normalizado = texto.replace('.', '').replace(',', '.')
            return int(float(texto_normalizado))
        except Exception:
            return None

    def normalizar_numero_like(valor):
        if valor is None:
            return ''
        try:
            if isinstance(valor, float) and np.isnan(valor):
                return ''
        except Exception:
            pass
        s = str(valor).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s

    def normalizar_co(valor):
        if valor is None:
            return ''
        try:
            if isinstance(valor, float) and np.isnan(valor):
                return ''
        except Exception:
            pass
        s = str(valor).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s

    def somente_quatro_digitos(valor):
        s = normalizar_co(valor)
        s = re.sub(r'\D', '', s)
        return s if len(s) == 4 else ''

    def primeiro_valor_nao_vazio(serie):
        for v in serie:
            nv = normalizar_co(v)
            if nv != '':
                return nv
        return ''

    def ajustar_fonte_qgr_para_chave(valor):
        inteiro = try_to_int(valor)
        if inteiro is None:
            return normalizar_numero_like(valor)
        if inteiro in excecoes:
            return str(inteiro)
        try:
            return str(ajustar_fonte_recurso(inteiro))
        except Exception:
            return str(inteiro)

    # Copias de trabalho
    df_rec = df_rec_original.copy()
    df_geral = df_geral_original.copy()

    # REC: considerar apenas COD_ID == 11 (quando existir)
    if 'COD_ID' in df_rec.columns:
        df_rec = df_rec[df_rec['COD_ID'] == 11].copy()

    # Chaves normalizadas
    df_rec['CODIGO_RECEITA_KEY'] = df_rec['CODIGO_RECEITA'].apply(normalizar_numero_like)
    df_rec['FR_KEY'] = df_rec['FONTE_RECURSO'].apply(normalizar_numero_like)
    df_geral['CODIGO_RECEITA_KEY'] = df_geral['CODIGO_RECEITA'].apply(normalizar_numero_like)
    df_geral['FR_KEY'] = df_geral['FONTE_RECURSO'].apply(ajustar_fonte_qgr_para_chave)

    # Agregar CO por chave, pegando o primeiro valor não vazio
    rec_co = (
        df_rec[['CODIGO_RECEITA_KEY', 'FR_KEY', 'CO']]
        .groupby(['CODIGO_RECEITA_KEY', 'FR_KEY'], as_index=False)
        .agg({'CO': primeiro_valor_nao_vazio})
        .rename(columns={'CO': 'CO_REC'})
    )

    qgr_co = (
        df_geral[['CODIGO_RECEITA_KEY', 'FR_KEY', 'CO']]
        .groupby(['CODIGO_RECEITA_KEY', 'FR_KEY'], as_index=False)
        .agg({'CO': primeiro_valor_nao_vazio})
        .rename(columns={'CO': 'CO_QGR'})
    )


    rec_co['CO_REC_N'] = rec_co['CO_REC'].apply(somente_quatro_digitos)
    qgr_co['CO_QGR_N'] = qgr_co['CO_QGR'].apply(somente_quatro_digitos)


    combinado = pd.merge(rec_co, qgr_co, on=['CODIGO_RECEITA_KEY', 'FR_KEY'], how='outer')
    combinado['CO_REC'] = combinado['CO_REC'].fillna('')
    combinado['CO_QGR'] = combinado['CO_QGR'].fillna('')
    combinado['CO_REC_N'] = combinado['CO_REC_N'].fillna('')
    combinado['CO_QGR_N'] = combinado['CO_QGR_N'].fillna('')

    rec_tem4 = combinado['CO_REC_N'] != ''
    qgr_tem4 = combinado['CO_QGR_N'] != ''
    mask_apenas_um = rec_tem4 ^ qgr_tem4
    mask_ambos_dif = (rec_tem4 & qgr_tem4) & (combinado['CO_REC_N'] != combinado['CO_QGR_N'])
    diferencas = combinado[mask_apenas_um | mask_ambos_dif].copy()

    diferencas.rename(columns={
        'CODIGO_RECEITA_KEY': 'CODIGO_RECEITA',
        'FR_KEY': 'FONTE_RECURSO'
    }, inplace=True)

    # Ordenação opcional para melhor leitura
    if not diferencas.empty:
        diferencas = diferencas[['CODIGO_RECEITA', 'FONTE_RECURSO', 'CO_REC', 'CO_QGR']]
        diferencas.sort_values(['CODIGO_RECEITA', 'FONTE_RECURSO'], inplace=True)
    else:
        # Garantir colunas mesmo vazio
        diferencas = pd.DataFrame(columns=['CODIGO_RECEITA', 'FONTE_RECURSO', 'CO_REC', 'CO_QGR'])

    return diferencas

def main():
    st.title('Rec VS QGR')
    st.write("Por favor, faça o upload dos arquivos Rec e QGR")

    col1, col2 = st.columns(2)
    with col1:
        uploaded_file_rec = st.file_uploader("Escolha o arquivo REC", type=["xls", 'xlsx'])
    with col2:
        uploaded_file_geral = st.file_uploader("Escolha o arquivo QGR", type=['xls', 'xlsx'])

    nome_arquivo = st.text_input('Digite o nome do arquivo de saída (sem a extensão .xlsx)')

    if uploaded_file_rec is not None and uploaded_file_geral is not None:
        df_rec = pd.read_excel(uploaded_file_rec)
        df_geral = pd.read_excel(uploaded_file_geral)

        duplicated_rows = df_rec[df_rec.duplicated(subset=['CODIGO_RECEITA', 'FONTE_RECURSO', 'VR_ARREC_MES_FONTE'])]
        if not duplicated_rows.empty:
            st.write('Aviso: Linhas duplicadas encontradas no arquivo REC para os seguintes conjuntos:')
            for _, row in duplicated_rows.iterrows():
                st.write(f"CODIGO_RECEITA: {row['CODIGO_RECEITA']}, FONTE_RECURSO: {row['FONTE_RECURSO']}, VR_ARREC_MES_FONTE: {row['VR_ARREC_MES_FONTE']}")


        st.write("### Validação Interna - Registros 10 vs 11")
        discrepancias_internas = validar_registros_10_11(df_rec)
        
        if discrepancias_internas:
            df_discrepancias_internas = pd.DataFrame(discrepancias_internas)
            df_discrepancias_internas['VALOR_REGISTRO_10'] = df_discrepancias_internas['VALOR_REGISTRO_10'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')
            df_discrepancias_internas['SOMA_REGISTROS_11'] = df_discrepancias_internas['SOMA_REGISTROS_11'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')
            df_discrepancias_internas['DIFERENCA'] = df_discrepancias_internas['DIFERENCA'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')
            df_discrepancias_internas['CODIGO_RECEITA'] = df_discrepancias_internas['CODIGO_RECEITA'].astype(str)
            
            st.write("Discrepâncias encontradas entre registros 10 e 11:")
            st.dataframe(df_discrepancias_internas)
        else:
            st.write('Nenhuma discrepância encontrada entre registros 10 e 11.')

       
        # Validação isolada de CO removida, pois CO agora entra na chave de comparação principal
 
        if 'COD_ID' in df_rec.columns:
            df_rec = df_rec[df_rec['COD_ID'] == 11]

   
        agg_qgr_df = agregar_por_trinca(df_geral, is_qgr=True)
        agg_rec_df = agregar_por_trinca(df_rec, is_qgr=False)

        # Construção das categorias: Corretos, Incorretos, Ausentes
        # Mapas para comparação
        qgr_map = {f"{r['CODIGO_RECEITA']}|{r['FONTE_NUCLEO']}|{r['CO']}": float(r['VALOR']) for _, r in agg_qgr_df.iterrows()}
        rec_map = {f"{r['CODIGO_RECEITA']}|{r['FONTE_NUCLEO']}|{r['CO']}": float(r['VALOR']) for _, r in agg_rec_df.iterrows()}

        corretos = []
        incorretos = []
        ausentes_no_rec = []
        ausentes_no_qgr = []

        # QGR -> REC (inclui diferenças de valor)
        for key, vq in qgr_map.items():
            vr = rec_map.get(key)
            rec_k, fon_k, co_k = key.split('|')
            if vr is None:
                ausentes_no_rec.append({'CODIGO_RECEITA': rec_k, 'FONTE_NUCLEO': fon_k, 'CO': co_k, 'VALOR_QGR': vq})
            else:
                if approx_equal(vq, vr):
                    corretos.append({'CODIGO_RECEITA': rec_k, 'FONTE_NUCLEO': fon_k, 'CO': co_k, 'VALOR_QGR': vq, 'VALOR_REC': vr})
                else:
                    incorretos.append({'CODIGO_RECEITA': rec_k, 'FONTE_NUCLEO': fon_k, 'CO': co_k, 'VALOR_QGR': vq, 'VALOR_REC': vr, 'DIFERENCA': vq - vr})

        # REC -> QGR (ausentes no QGR)
        for key, vr in rec_map.items():
            if key not in qgr_map:
                rec_k, fon_k, co_k = key.split('|')
                ausentes_no_qgr.append({'CODIGO_RECEITA': rec_k, 'FONTE_NUCLEO': fon_k, 'CO': co_k, 'VALOR_REC': vr})

        df_corretos = pd.DataFrame(corretos)
        df_incorretos = pd.DataFrame(incorretos)
        df_ausentes_rec = pd.DataFrame(ausentes_no_rec)
        df_ausentes_qgr = pd.DataFrame(ausentes_no_qgr)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Resultado do QGR com chave tripla
            agg_qgr_df.to_excel(writer, sheet_name='Resultado_QGR', index=False)

            if not df_corretos.empty:
                df_corretos.to_excel(writer, sheet_name='Corretos', index=False)
            if not df_incorretos.empty:
                df_incorretos.to_excel(writer, sheet_name='Incorretos', index=False)
            if not df_ausentes_rec.empty:
                df_ausentes_rec.to_excel(writer, sheet_name='Ausentes_no_REC', index=False)
            if not df_ausentes_qgr.empty:
                df_ausentes_qgr.to_excel(writer, sheet_name='Ausentes_no_QGR', index=False)

            if discrepancias_internas:
                df_discrepancias_internas_excel = pd.DataFrame(discrepancias_internas)
                df_discrepancias_internas_excel.to_excel(writer, sheet_name='Discrepancias_10_11', index=False)
            # Exporta diferenças de CO entre REC e QGR
            try:
                diferencas_co = validar_co(df_rec, df_geral)
                if not diferencas_co.empty:
                    diferencas_co.to_excel(writer, sheet_name='Discrepancias_CO', index=False)
            except Exception:
                pass

        output.seek(0)
        if nome_arquivo:
            nome_arquivo += ".xlsx"
        else:
            nome_arquivo = "resultado_soma.xlsx"

        st.download_button(
            label="Baixe o arquivo de saída",
            data=output,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

   
        total_qgr = len(qgr_map)
        total_rec = len(rec_map)
        total_corretos = len(corretos)
        total_incorretos = len(incorretos)
        total_aus_rec = len(ausentes_no_rec)
        total_aus_qgr = len(ausentes_no_qgr)

        st.write("### Comparação REC × QGR (chave tripla inclui CO)")
        st.write(f"Chaves QGR: {total_qgr} · Chaves REC: {total_rec} · Corretos: {total_corretos} · Incorretos: {total_incorretos} · Ausentes no REC: {total_aus_rec} · Ausentes no QGR: {total_aus_qgr}")

        if total_incorretos:
            df_incorretos_display = df_incorretos.copy()
            df_incorretos_display['VALOR_QGR'] = df_incorretos_display['VALOR_QGR'].apply(formatar_ptbr)
            df_incorretos_display['VALOR_REC'] = df_incorretos_display['VALOR_REC'].apply(formatar_ptbr)
            df_incorretos_display['DIFERENCA'] = df_incorretos_display['DIFERENCA'].apply(formatar_ptbr)
            st.write("Registros com valores diferentes:")
            st.dataframe(df_incorretos_display)

        if total_aus_rec:
            df_ausentes_rec_display = df_ausentes_rec.copy()
            df_ausentes_rec_display['VALOR_QGR'] = df_ausentes_rec_display['VALOR_QGR'].apply(formatar_ptbr)
            st.write("Presentes apenas no QGR (ausentes no REC):")
            st.dataframe(df_ausentes_rec_display)

        if total_aus_qgr:
            df_ausentes_qgr_display = df_ausentes_qgr.copy()
            df_ausentes_qgr_display['VALOR_REC'] = df_ausentes_qgr_display['VALOR_REC'].apply(formatar_ptbr)
            st.write("Presentes apenas no REC (ausentes no QGR):")
            st.dataframe(df_ausentes_qgr_display)

if __name__ == "__main__":
    main()
