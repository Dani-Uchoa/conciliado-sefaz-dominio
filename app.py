import streamlit as st
import pandas as pd
import re
import unicodedata
import io

st.set_page_config(page_title="Conciliador Fiscal", layout="wide")
st.title("⚖️ Conciliador Fiscal: SEFAZ x Domínio")

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
    termos = ["CHAVE", "NOTA", "DATA", "VALOR", "EMISSAO", "TOTAL"]
    for i in range(min(len(df), 50)):
        linha = [normalizar(str(c)) for c in df.iloc[i]]
        matches = sum(1 for c in linha if any(t in c for t in termos))
        if matches >= 2: return i
    return 0

def carregar_planilha(f):
    df = pd.read_excel(f, header=None, dtype=str)
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
with c1: f_sefaz = st.file_uploader("1. Arquivo SEFAZ (.xlsx)", type=["xlsx", "xls"])
with c2: f_dom = st.file_uploader("2. Arquivo DOMÍNIO (.xlsx)", type=["xlsx", "xls"])

if f_sefaz and f_dom:
    try:
        df_s = carregar_planilha(f_sefaz)
        df_d = carregar_planilha(f_dom)

        st.write("---")
        st.markdown("### ⚙️ Verificação de Colunas")
        st.write("O sistema tentou adivinhar quais são as colunas corretas. **Se alguma estiver errada, basta corrigir nas caixas abaixo:**")
        
        cols_s = list(df_s.columns)
        def_s_n = next((i for i, c in enumerate(cols_s) if "CHAVE" in normalizar(c) or "NOTA" in normalizar(c)), 0)
        def_s_d = next((i for i, c in enumerate(cols_s) if "DATA" in normalizar(c) or "EMISSAO" in normalizar(c)), 0)
        def_s_v = next((i for i, c in enumerate(cols_s) if "VALOR" in normalizar(c) or "TOTAL" in normalizar(c)), 0)

        cols_d = list(df_d.columns)
        def_d_n = next((i for i, c in enumerate(cols_d) if "NOTA" in normalizar(c) or "DOC" in normalizar(c)), 0)
        def_d_d = next((i for i, c in enumerate(cols_d) if "DATA" in normalizar(c) or "EMISSAO" in normalizar(c)), 0)
        def_d_v = next((i for i, c in enumerate(cols_d) if "VALOR" in normalizar(c) or "CONTABIL" in normalizar(c)), 0)

        col1, col2 = st.columns(2)
        with col1:
            st.info("📄 **Planilha SEFAZ**")
            s_nota = st.selectbox("Coluna da Chave/Nota (Sefaz)", cols_s, index=def_s_n)
            s_data = st.selectbox("Coluna da Data (Sefaz)", cols_s, index=def_s_d)
            s_valor = st.selectbox("Coluna do Valor (Sefaz)", cols_s, index=def_s_v)
        
        with col2:
            st.info("📄 **Planilha DOMÍNIO**")
            d_nota = st.selectbox("Coluna da Chave/Nota (Domínio)", cols_d, index=def_d_n)
            d_data = st.selectbox("Coluna da Data (Domínio)", cols_d, index=def_d_d)
            d_valor = st.selectbox("Coluna do Valor (Domínio)", cols_d, index=def_d_v)

        if st.button("🚀 Cruzar Dados Agora", type="primary", use_container_width=True):
            with st.spinner("Processando..."):
                ds = processar_dataframe(df_s, s_nota, s_data, s_valor)
                dd = processar_dataframe(df_d, d_nota, d_data, d_valor)

                for idx, row in dd.iterrows():
                    nota_dom = row['nota']
                    data_dom = row['data']
                    match_sefaz = ds[ds['nota'] == nota_dom]
                    if not match_sefaz.empty:
                        for idx_sf, row_sf in match_sefaz.iterrows():
                            data_sf = row_sf['data']
                            if data_dom != data_sf and data_dom.day == data_sf.month and data_dom.month == data_sf.day:
                                dd.at[idx, 'data'] = data_sf

                dd = dd.groupby(['nota', 'data'], as_index=False)['valor'].sum()
                m = pd.merge(ds, dd, on=['nota', 'data'], how='outer', suffixes=('_sefaz', '_dom')).fillna(0)
                
                total_sefaz = m['valor_sefaz'].sum()
                total_dominio = m['valor_dom'].sum()
                diferenca_global = total_sefaz - total_dominio

                st.write("---")
                st.subheader("📊 Resultado Consolidado")
                met1, met2, met3 = st.columns(3)
                with met1: st.metric("Valor Total SEFAZ", formatar_moeda_br(total_sefaz))
                with met2: st.metric("Valor Total DOMÍNIO", formatar_moeda_br(total_dominio))
                with met3: st.metric("Diferença Global", formatar_moeda_br(diferenca_global), delta=f"{diferenca_global:,.2f} R$" if abs(diferenca_global) > 0.01 else None, delta_color="inverse" if diferenca_global != 0 else "normal")

                divergencias = m[abs(m['valor_sefaz'] - m['valor_dom']) > 0.01].copy()
                divergencias.rename(columns={'valor_sefaz': 'valor sefaz', 'valor_dom': 'valor dominio'}, inplace=True)
                divergencias = divergencias.sort_values(by=['data', 'nota'])
                divergencias['data'] = pd.to_datetime(divergencias['data']).dt.strftime('%d/%m/%Y')
                df_final = divergencias[['data', 'nota', 'valor sefaz', 'valor dominio']].reset_index(drop=True)
                
                st.subheader("🔍 Detalhamento das Divergências")
                if not df_final.empty:
                    st.warning(f"Foram identificadas {len(df_final)} notas com inconsistências.")
                    df_visual = df_final.copy()
                    df_visual['valor sefaz'] = df_visual['valor sefaz'].apply(formatar_moeda_br)
                    df_visual['valor dominio'] = df_visual['valor dominio'].apply(formatar_moeda_br)
                    st.dataframe(df_visual, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Divergencias')
                    
                    st.download_button("📥 Baixar Planilha", data=output.getvalue(), file_name="divergencias.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.success("🎉 Excelente! Nenhuma divergência encontrada.")
    except Exception as e:
        st.error(f"Erro ao ler os arquivos. Detalhe técnico: {e}")
