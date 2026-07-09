import streamlit as st
import pandas as pd
import re
import unicodedata
import io

st.set_page_config(page_title="Conciliador Fiscal de Notas", layout="wide")
st.title("⚖️ Conciliador Fiscal")

# --- UTILITÁRIOS ---
def formatar_moeda_br(valor):
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def normalizar(txt):
    if pd.isna(txt): return ""
    return unicodedata.normalize('NFD', str(txt)).encode('ascii', 'ignore').decode('utf-8').upper().strip()

def limpar_valor(v):
    if pd.isna(v): return 0.0
    s = str(v).replace('R$', '').replace('"', '').replace('\xa0', '').replace(' ', '').strip()
    if not s: return 0.0
    if ',' in s: s = s.replace('.', '').replace(',', '.')
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

def converter_data(d):
    if pd.isna(d): return None
    s = str(d).strip()
    if s.replace('.', '').isdigit() and len(s) >= 5:
        try: return pd.to_datetime(float(s), unit='D', origin='1899-12-30').date()
        except: pass
    if '.' in s and not s.replace('.', '').isdigit(): s = s.replace('.', '/')
    if re.match(r'^\d{4}-\d{2}-\d{2}', s):
        try: return pd.to_datetime(s, errors='coerce').date()
        except: pass
    try: return pd.to_datetime(s, dayfirst=True, errors='raise').date()
    except: return pd.to_datetime(s, errors='coerce').date()

def extrair_nota_limpa(n):
    if pd.isna(n): return ""
    s = str(n).strip()
    s = re.sub(r'\D', '', s)
    if len(s) == 44: s = s[25:34] 
    return s.lstrip('0') if s else ""

def encontrar_cabecalho(df):
    termos = ["CHAVE", "NOTA", "DATA", "VALOR", "EMISSAO", "TOTAL", "NUMERO"]
    for i in range(min(len(df), 50)):
        linha = [normalizar(str(c)) for c in df.iloc[i]]
        matches = sum(1 for c in linha if any(t in c for t in termos))
        if matches >= 2: return i
    return 0

def carregar_planilha(f):
    df = pd.read_excel(f, header=None, dtype=str)
    
    # --- NOVO: Motor de Descompactação Automática ---
    # Se o Excel tiver esmagado tudo na primeira coluna (A) devido a falhas da SEFAZ
    if len(df.columns) == 1 or df.iloc[:, 1:].isna().all().all():
        texto_esmagado = df.iloc[:, 0].dropna().astype(str).str.cat(sep='\n')
        # Descobre se o delimitador oculto é ponto e vírgula ou vírgula
        separador = ';' if ';' in texto_esmagado.split('\n')[0] else ','
        df = pd.read_csv(io.StringIO(texto_esmagado), sep=separador, dtype=str, header=None, engine='python', on_bad_lines='skip')
    # ------------------------------------------------
    
    idx_cabecalho = encontrar_cabecalho(df)
    
    nomes_colunas = [str(c).strip() if pd.notna(c) else f"Coluna_{i}" for i, c in enumerate(df.iloc[idx_cabecalho])]
    
    vistos = {}
    nomes_unicos = []
    for nome in nomes_colunas:
        if nome in vistos:
            vistos[nome] += 1
            nomes_unicos.append(f"{nome} ({vistos[nome]})")
        else:
            vistos[nome] = 0
            nomes_unicos.append(nome)
            
    df.columns = nomes_unicos
    df = df.iloc[idx_cabecalho+1:].reset_index(drop=True)
    return df

def processar_dataframe(df, col_nota, col_data, col_valor):
    res = pd.DataFrame()
    res['nota'] = df[col_nota].apply(extrair_nota_limpa)
    res['data'] = df[col_data].apply(converter_data)
    res['valor'] = df[col_valor].apply(limpar_valor)
    res = res[res['nota'] != ""].dropna(subset=['data'])
    return res.groupby(['nota', 'data'], as_index=False)['valor'].sum()

# --- INTERFACE ---
st.warning("⚠️ Salve os arquivos originais em formato **.xlsx** no Excel antes de inserir aqui.")

c1, c2 = st.columns(2)
with c1: f_arq1 = st.file_uploader("1. Arquivo de Origem (Sefaz, Sieg, etc.) - .xlsx", type=["xlsx", "xls"])
with c2: f_arq2 = st.file_uploader("2. Arquivo de Destino (Domínio, etc.) - .xlsx", type=["xlsx", "xls"])

if f_arq1 and f_arq2:
    try:
        df_1 = carregar_planilha(f_arq1)
        df_2 = carregar_planilha(f_arq2)

        st.write("---")
        st.markdown("### ⚙️ Mapeamento de Colunas")
        st.write("Selecione as colunas correspondentes em cada arquivo. **Você pode usar a Chave de Acesso (44 dígitos) ou o Número da Nota direto.**")
        
        cols_1 = list(df_1.columns)
        def_1_n = next((i for i, c in enumerate(cols_1) if "CHAVE" in normalizar(c) or "NOTA" in normalizar(c) or "NUMERO" in normalizar(c)), 0)
        def_1_d = next((i for i, c in enumerate(cols_1) if "DATA" in normalizar(c) or "EMISSAO" in normalizar(c)), 0)
        def_1_v = next((i for i, c in enumerate(cols_1) if "VALOR" in normalizar(c) or "TOTAL" in normalizar(c)), 0)

        cols_2 = list(df_2.columns)
        def_2_n = next((i for i, c in enumerate(cols_2) if "NOTA" in normalizar(c) or "DOC" in normalizar(c) or "NUMERO" in normalizar(c)), 0)
        def_2_d = next((i for i, c in enumerate(cols_2) if "DATA" in normalizar(c) or "EMISSAO" in normalizar(c)), 0)
        def_2_v = next((i for i, c in enumerate(cols_2) if "VALOR" in normalizar(c) or "CONTABIL" in normalizar(c)), 0)

        col1, col2 = st.columns(2)
        with col1:
            st.info("📄 **Sefaz (Origem)**")
            s_nota = st.selectbox("Identificador (Chave OU Número da Nota)", cols_1, index=def_1_n, key="s_nota")
            s_data = st.selectbox("Coluna da Data", cols_1, index=def_1_d, key="s_data")
            s_valor = st.selectbox("Coluna do Valor", cols_1, index=def_1_v, key="s_valor")
        
        with col2:
            st.info("📄 **Domínio (Destino)**")
            d_nota = st.selectbox("Identificador (Chave OU Número da Nota)", cols_2, index=def_2_n, key="d_nota")
            d_data = st.selectbox("Coluna da Data", cols_2, index=def_2_d, key="d_data")
            d_valor = st.selectbox("Coluna do Valor", cols_2, index=def_2_v, key="d_valor")

        if st.button("🚀 Cruzar Dados Agora", type="primary", use_container_width=True):
            with st.spinner("Processando..."):
                ds = processar_dataframe(df_1, s_nota, s_data, s_valor)
                dd = processar_dataframe(df_2, d_nota, d_data, d_valor)

                for idx, row in dd.iterrows():
                    nota_dom = row['nota']
                    data_dom = row['data']
                    match_1 = ds[ds['nota'] == nota_dom]
                    if not match_1.empty:
                        for idx_sf, row_sf in match_1.iterrows():
                            data_sf = row_sf['data']
                            if data_dom != data_sf and data_dom.day == data_sf.month and data_dom.month == data_sf.day:
                                dd.at[idx, 'data'] = data_sf

                dd = dd.groupby(['nota', 'data'], as_index=False)['valor'].sum()
                m = pd.merge(ds, dd, on=['nota', 'data'], how='outer', suffixes=('_arq1', '_arq2')).fillna(0)
                
                total_1 = m['valor_arq1'].sum()
                total_2 = m['valor_arq2'].sum()
                diferenca_global = total_1 - total_2

                st.write("---")
                st.subheader("📊 Resultado Consolidado")
                met1, met2, met3 = st.columns(3)
                with met1: st.metric("Valor Total Sefaz", formatar_moeda_br(total_1))
                with met2: st.metric("Valor Total Domínio", formatar_moeda_br(total_2))
                with met3: st.metric("Diferença Global", formatar_moeda_br(diferenca_global), delta=f"{diferenca_global:,.2f} R$" if abs(diferenca_global) > 0.01 else None, delta_color="inverse" if diferenca_global != 0 else "normal")

                divergencias = m[abs(m['valor_arq1'] - m['valor_arq2']) > 0.01].copy()
                divergencias.rename(columns={'valor_arq1': 'Valor Arq 1', 'valor_arq2': 'Valor Arq 2', 'nota': 'Identificador (Nota/Chave)', 'data': 'Data'}, inplace=True)
                divergencias = divergencias.sort_values(by=['Data', 'Identificador (Nota/Chave)'])
                divergencias['Data'] = pd.to_datetime(divergencias['Data']).dt.strftime('%d/%m/%Y')
                df_final = divergencias[['Data', 'Identificador (Nota/Chave)', 'Valor Arq 1', 'Valor Arq 2']].reset_index(drop=True)
                
                st.subheader("🔍 Detalhamento das Divergências")
                if not df_final.empty:
                    st.warning(f"Foram identificadas {len(df_final)} notas com inconsistências.")
                    df_visual = df_final.copy()
                    df_visual['Valor Arq 1'] = df_visual['Valor Arq 1'].apply(formatar_moeda_br)
                    df_visual['Valor Arq 2'] = df_visual['Valor Arq 2'].apply(formatar_moeda_br)
                    st.dataframe(df_visual, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Divergencias')
                    
                    st.download_button("📥 Baixar Planilha", data=output.getvalue(), file_name="divergencias.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.success("🎉 Excelente! Nenhuma divergência encontrada.")
    except Exception as e:
        st.error(f"Erro ao ler os arquivos. Detalhe técnico: {e}")
