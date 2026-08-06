import streamlit as st
import pandas as pd
import re
import unicodedata
import io

st.set_page_config(page_title="Conciliador Fiscal: SEFAZ x Domínio", layout="wide")
st.title("⚖️ Conciliador Fiscal: SEFAZ x Domínio")

# --- UTILITÁRIOS ---
def formatar_moeda_br(valor):
    s = f"{valor:_.2f}"
    s = s.replace('.', ',').replace('_', '.')
    return f"R$ {s}"

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

# --- MOTOR DE LEITURA COM BLOQUEIO DE CORRUPÇÃO ---
def carregar_planilha(f):
    f.seek(0)
    conteudo = f.read()
    df = None
    
    if f.name.lower().endswith('.csv'):
        try: texto = conteudo.decode('utf-8')
        except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
        primeira_linha = texto.split('\n')[0] if '\n' in texto else texto
        separador = ';' if ';' in primeira_linha else ','
        df = pd.read_csv(io.StringIO(texto), sep=separador, dtype=str, header=None, engine='python', on_bad_lines='skip')
        
    else:
        # Checagem de Assinatura Binária
        is_real_xls = conteudo.startswith(b'\xD0\xCF\x11\xE0')
        is_real_xlsx = conteudo.startswith(b'PK')
        
        if is_real_xls or is_real_xlsx:
            motor = 'openpyxl' if is_real_xlsx else 'xlrd'
            try:
                df = pd.read_excel(io.BytesIO(conteudo), header=None, dtype=str, engine=motor)
                
                # Motor de Descompactação SEFAZ
                if len(df.columns) == 1 or df.iloc[:, 1:].isna().all().all():
                    texto_esmagado = df.iloc[:, 0].dropna().astype(str).str.cat(sep='\n')
                    if texto_esmagado.strip():
                        separador = ';' if ';' in texto_esmagado.split('\n')[0] else ','import streamlit as st
import pandas as pd
import re
import unicodedata
import io

st.set_page_config(page_title="Conciliador Fiscal Universal", layout="wide")
st.title("⚖️ Conciliador Fiscal Universal")

# --- UTILITÁRIOS ---
def formatar_moeda_br(valor):
    s = f"{valor:_.2f}"
    s = s.replace('.', ',').replace('_', '.')
    return f"R$ {s}"

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

# --- MOTOR DE LEITURA COM BLOQUEIO DE CORRUPÇÃO ---
def carregar_planilha(f):
    f.seek(0)
    conteudo = f.read()
    df = None
    
    if f.name.lower().endswith('.csv'):
        try: texto = conteudo.decode('utf-8')
        except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
        primeira_linha = texto.split('\n')[0] if '\n' in texto else texto
        separador = ';' if ';' in primeira_linha else ','
        df = pd.read_csv(io.StringIO(texto), sep=separador, dtype=str, header=None, engine='python', on_bad_lines='skip')
        
    else:
        is_real_xls = conteudo.startswith(b'\xD0\xCF\x11\xE0')
        is_real_xlsx = conteudo.startswith(b'PK')
        
        if is_real_xls or is_real_xlsx:
            motor = 'openpyxl' if is_real_xlsx else 'xlrd'
            try:
                df = pd.read_excel(io.BytesIO(conteudo), header=None, dtype=str, engine=motor)
                
                if len(df.columns) == 1 or df.iloc[:, 1:].isna().all().all():
                    texto_esmagado = df.iloc[:, 0].dropna().astype(str).str.cat(sep='\n')
                    if texto_esmagado.strip():
                        separador = ';' if ';' in texto_esmagado.split('\n')[0] else ','
                        df = pd.read_csv(io.StringIO(texto_esmagado), sep=separador, dtype=str, header=None, engine='python', on_bad_lines='skip')
            except Exception:
                st.error(f"🛑 **ARQUIVO CORROMPIDO NA ORIGEM:** O arquivo '{f.name}' exportado pelo sistema possui defeitos na matriz binária. \n\nPara garantir a integridade da conciliação, abra este arquivo no Excel do seu computador e selecione **'Salvar Como -> Pasta de Trabalho do Excel (.xlsx)'** antes de enviar.")
                return pd.DataFrame()
        else:
            try:
                dfs = pd.read_html(io.BytesIO(conteudo))
                df = dfs[0].astype(str)
            except Exception:
                try: texto = conteudo.decode('utf-8')
                except UnicodeDecodeError: texto = conteudo.decode('latin1', errors='replace')
                
                if '\t' in texto: sep = '\t'
                elif ';' in texto: sep = ';'
                else: sep = ','
                df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, header=None, engine='python', on_bad_lines='skip')
    
    if df is None or df.empty:
        return pd.DataFrame()
        
    idx_cabecalho = encontrar_cabecalho(df)
    
    if idx_cabecalho >= len(df):
        return df

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

# --- INTERFACE GRÁFICA ---
st.info("💡 **Atenção:** Arquivos defeituosos de alguns sistemas exigem reparo no Excel (Salvar Como .xlsx).")

c1, c2 = st.columns(2)
with c1: f_origem = st.file_uploader("1. Planilha de ORIGEM (Sieg, Prefeitura, SEFAZ...)", type=["xlsx", "xls", "csv"])
with c2: f_dom = st.file_uploader("2. Planilha da DOMÍNIO", type=["xlsx", "xls", "csv"])

if f_origem and f_dom:
    try:
        df_o = carregar_planilha(f_origem)
        df_d = carregar_planilha(f_dom)

        if not df_o.empty and not df_d.empty:
            st.write("---")
            st.markdown("### ⚙️ Identificação das Colunas")
            
            cols_o = list(df_o.columns)
            def_o_n = next((i for i, c in enumerate(cols_o) if "CHAVE" in normalizar(c) or "NOTA" in normalizar(c) or "NUMERO" in normalizar(c)), 0)
            def_o_d = next((i for i, c in enumerate(cols_o) if "DATA" in normalizar(c) or "EMISSAO" in normalizar(c)), 0)
            def_o_v = next((i for i, c in enumerate(cols_o) if "VALOR" in normalizar(c) or "TOTAL" in normalizar(c)), 0)

            cols_d = list(df_d.columns)
            def_d_n = next((i for i, c in enumerate(cols_d) if "NOTA" in normalizar(c) or "DOC" in normalizar(c) or "NUMERO" in normalizar(c)), 0)
            def_d_d = next((i for i, c in enumerate(cols_d) if "DATA" in normalizar(c) or "EMISSAO" in normalizar(c)), 0)
            def_d_v = next((i for i, c in enumerate(cols_d) if "VALOR" in normalizar(c) or "CONTABIL" in normalizar(c)), 0)

            col1, col2 = st.columns(2)
            with col1:
                st.info("📊 **Colunas da ORIGEM**")
                o_nota = st.selectbox("Coluna do Número da Nota ou Chave (Origem)", cols_o, index=def_o_n, key="o_nota")
                o_data = st.selectbox("Coluna da Data de Emissão (Origem)", cols_o, index=def_o_d, key="o_data")
                o_valor = st.selectbox("Coluna do Valor da Nota (Origem)", cols_o, index=def_o_v, key="o_valor")
            
            with col2:
                st.info("📊 **Colunas da DOMÍNIO**")
                d_nota = st.selectbox("Coluna do Número da Nota (Domínio)", cols_d, index=def_d_n, key="d_nota")
                d_data = st.selectbox("Coluna da Data (Domínio)", cols_d, index=def_d_d, key="d_data")
                d_valor = st.selectbox("Coluna do Valor Contábil (Domínio)", cols_d, index=def_d_v, key="d_valor")

            if st.button("🚀 Cruzar Dados e Buscar Divergências", type="primary", use_container_width=True):
                with st.spinner("Cruzando informações fiscais..."):
                    ds = processar_dataframe(df_o, o_nota, o_data, o_valor)
                    dd = processar_dataframe(df_d, d_nota, d_data, d_valor)

                    for idx, row in dd.iterrows():
                        nota_dom = row['nota']
                        data_dom = row['data']
                        match_s = ds[ds['nota'] == nota_dom]
                        if not match_s.empty:
                            for idx_sf, row_sf in match_s.iterrows():
                                data_sf = row_sf['data']
                                if data_dom != data_sf and data_dom.day == data_sf.month and data_dom.month == data_sf.day:
                                    dd.at[idx, 'data'] = data_sf

                    dd = dd.groupby(['nota', 'data'], as_index=False)['valor'].sum()
                    m = pd.merge(ds, dd, on=['nota', 'data'], how='outer', suffixes=('_origem', '_dom')).fillna(0)
                    
                    total_origem = m['valor_origem'].sum()
                    total_dominio = m['valor_dom'].sum()
                    diferenca_global = total_origem - total_dominio

                    st.write("---")
                    st.subheader("📊 Totais Consolidados")
                    met1, met2, met3 = st.columns(3)
                    with met1: st.metric("Soma Total ORIGEM", formatar_moeda_br(total_origem))
                    with met2: st.metric("Soma Total DOMÍNIO", formatar_moeda_br(total_dominio))
                    with met3: st.metric("Diferença Global", formatar_moeda_br(diferenca_global), delta=f"{diferenca_global:,.2f} R$" if abs(diferenca_global) > 0.01 else None, delta_color="inverse" if diferenca_global != 0 else "normal")

                    divergencias = m[abs(m['valor_origem'] - m['valor_dom']) > 0.01].copy()
                    
                    divergencias.rename(columns={
                        'nota': 'Número da Nota',
                        'data': 'Data',
                        'valor_origem': 'Valor Origem',
                        'valor_dom': 'Valor Domínio'
                    }, inplace=True)
                    
                    divergencias = divergencias.sort_values(by=['Data', 'Número da Nota'])
                    divergencias['Data'] = pd.to_datetime(divergencias['Data']).dt.strftime('%d/%m/%Y')
                    df_final = divergencias[['Data', 'Número da Nota', 'Valor Origem', 'Valor Domínio']].reset_index(drop=True)
                    
                    st.subheader("🔍 Detalhamento das Divergências Encontradas")
                    if not df_final.empty:
                        st.warning(f"Foram identificadas {len(df_final)} notas com inconsistências de valores ou de lançamento.")
                        df_visual = df_final.copy()
                        df_visual['Valor Origem'] = df_visual['Valor Origem'].apply(formatar_moeda_br)
                        df_visual['Valor Domínio'] = df_visual['Valor Domínio'].apply(formatar_moeda_br)
                        st.dataframe(df_visual, use_container_width=True)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='Divergências')
                        
                        st.download_button(
                            label="📥 Baixar Planilha de Divergências",
                            data=output.getvalue(),
                            file_name="divergencias_fiscais.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.success("🎉 Excelente! Nenhuma divergência individual foi encontrada entre os arquivos.")
    except Exception as e:
        st.error(f"Erro processual: {e}")
