import streamlit as st
import pandas as pd
import re
import unicodedata
import io

st.set_page_config(page_title="Conciliador Fiscal - Sefaz x Dominio", layout="wide")
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
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(re.sub(r'[^\d.]', '', s))
    except:
        return 0.0

def converter_data(d):
    if pd.isna(d): return None
    s = str(d).strip()
    
    if s.replace('.', '').isdigit() and len(s) >= 5:
        try: return pd.to_datetime(float(s), unit='D', origin='1899-12-30').date()
        except: pass
        
    if '.' in s and not s.replace('.', '').isdigit():
        s = s.replace('.', '/')
        
    if re.match(r'^\d{4}-\d{2}-\d{2}', s):
        try: return pd.to_datetime(s, errors='coerce').date()
        except: pass

    try:
        return pd.to_datetime(s, dayfirst=True, errors='raise').date()
    except:
        return pd.to_datetime(s, errors='coerce').date()

def extrair_nota_limpa(n):
    if pd.isna(n): return ""
    s = str(n).strip()
    s = re.sub(r'\D', '', s)
    if len(s) == 44:
        s = s[25:34] 
    return s.lstrip('0') if s else ""

# --- MOTOR DE LEITURA MULTI-SEPARADOR ULTRA RESILIENTE ---
def ler_planilha(file, sistema):
    if file is None: return None, "Arquivo não detectado."
    try:
        file.seek(0)
        if file.name.lower().endswith('.csv'):
            encodings = ['utf-8', 'iso-8859-1', 'latin1', 'cp1252']
            conteudo = None
            
            for enc in encodings:
                try:
                    file.seek(0)
                    conteudo = file.read().decode(enc)
                    break
                except:
                    continue
            
            if conteudo is None:
                return None, "Não foi possível identificar a codificação do arquivo CSV."
            
            primeira_linha = conteudo.splitlines()[0] if conteudo.splitlines() else ""
            if '\t' in conteudo:
                delimitador = '\t'
            elif ';' in primeira_linha or ';' in conteudo[:1000]:
                delimitador = ';'
            else:
                delimitador = ','
            
            linhas_brutas = []
            for line in conteudo.splitlines():
                if not line.strip():
                    continue
                if delimitador == '\t':
                    partes = line.split('\t')
                else:
                    partes = re.split(rf'{delimitador}(?=(?:[^"]*"[^"]*")*[^"]*$)', line)
                linhas_brutas.append([p.strip('" ') for p in partes])
                
            df = pd.DataFrame(linhas_brutas)
        else:
            df = pd.read_excel(file, header=None, dtype=str)

        col_n, col_d, col_v = -1, -1, -1
        idx_inicio = -1

        # Dicionário forçado com os nomes exatos passados
        termos_n = ["CHAVE DE ACESSO", "NUMERO DO DOCUMENTO FISCAL", "N NF-E", "NUM NFE", "NUMERO", "DOC", "N NF", "NOTA", "CHAVE"]
        termos_d = ["DATA DE EMISSAO", "DATA DA EMISSAO", "DATA EMISSAO", "EMISSAO", "DATA", "DT."]
        termos_v = ["VALOR R$", "VALOR TOTAL", "VALOR NF-E", "VALOR", "VLR CONTABIL", "VALOR CONTABIL", "TOTAL", "CONTABIL"]

        for i in range(min(len(df), 400)):
            linha = [normalizar(c) for c in df.iloc[i]]
            
            t_n = next((idx for idx, c in enumerate(linha) if any(t in c for t in termos_n)), -1)
            t_d = next((idx for idx, c in enumerate(linha) if any(t in c for t in termos_d)), -1)
            t_v = next((idx for idx, c in enumerate(linha) if any(t in c for t in termos_v)), -1)

            if t_d != -1 and t_v != -1:
                if t_n == -1:
                    for idx, c in enumerate(df.iloc[i+1]):
                        if pd.notna(c) and re.sub(r'\D', '', str(c)).isdigit() and idx != t_d:
                            t_n = idx; break
                
                if t_n != -1:
                    # Trava de segurança para impedir colunas sobrepostas na leitura "esmagada" do Excel
                    if len(set([t_n, t_d, t_v])) == 3:
                        col_n, col_d, col_v = t_n, t_d, t_v
                        idx_inicio = i; break

        if idx_inicio == -1:
            amostra = " | ".join([str(c) for c in df.iloc[0].tolist()[:3]]) if len(df) > 0 else "Planilha Vazia"
            return None, f"Colunas não mapeadas. O Excel agrupou tudo em uma só coluna? O programa leu a linha 1 assim: {amostra}"

        dados = df.iloc[idx_inicio + 1:].copy()
        res = pd.DataFrame()
        res['nota'] = dados[col_n].apply(extrair_nota_limpa)
        res['data'] = dados[col_d].apply(converter_data)
        res['valor'] = dados[col_v].apply(limpar_valor)
        
        res = res[res['nota'] != ""].dropna(subset=['data'])
        return res.groupby(['nota', 'data'], as_index=False)['valor'].sum(), None

    except Exception as e:
        return None, f"Erro estrutural na leitura: {str(e)}"

# --- INTERFACE GRÁFICA ---
st.warning(
    "⚠️ **INSTRUÇÃO IMPORTANTE**\n\n"
    "Para evitar erros de leitura, prefira inserir arquivos no formato Excel verdadeiro. Se um arquivo falhar:\n"
    "1. Abra a planilha original no **Excel**.\n"
    "2. Selecione a primeira coluna e vá em Dados > Texto para Colunas (se os dados estiverem grudados).\n"
    "3. Clique em **Arquivo > Salvar Como** no tipo **Pasta de Trabalho do Excel (*.xlsx)**.\n"
    "4. Faça o upload do novo arquivo gerado."
)

st.write("Insira os arquivos abaixo para cruzar as informações. O sistema gerará a planilha de divergências automaticamente.")

c1, c2 = st.columns(2)
with c1: f_sefaz = st.file_uploader("1. Arquivo SEFAZ (Excel ou CSV)", type=["csv", "xlsx", "xls"])
with c2: f_dom = st.file_uploader("2. Arquivo DOMÍNIO (Excel ou CSV)", type=["csv", "xlsx", "xls"])

if f_sefaz and f_dom:
    with st.spinner("Processando e cruzando dados fiscais..."):
        ds, err_s = ler_planilha(f_sefaz, "SEFAZ")
        dd, err_plan = ler_planilha(f_dom, "DOMINIO")

        if err_s: st.error(f"🚨 Erro no arquivo SEFAZ: {err_s}")
        if err_plan: st.error(f"🚨 Erro no arquivo DOMÍNIO: {err_plan}")

        if ds is not None and dd is not None:
            
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
            st.subheader("📊 Totais Consolidados dos Arquivos")
            met1, met2, met3 = st.columns(3)
            
            with met1:
                st.metric(label="Valor Total SEFAZ", value=formatar_moeda_br(total_sefaz))
            with met2:
                st.metric(label="Valor Total DOMÍNIO", value=formatar_moeda_br(total_dominio))
            with met3:
                st.metric(
                    label="Diferença Global", 
                    value=formatar_moeda_br(diferenca_global),
                    delta=f"{diferenca_global:,.2f} R$" if abs(diferenca_global) > 0.01 else None,
                    delta_color="inverse" if diferenca_global != 0 else "normal"
                )
            st.write("---")

            divergencias = m[abs(m['valor_sefaz'] - m['valor_dom']) > 0.01].copy()
            
            divergencias.rename(columns={
                'valor_sefaz': 'valor sefaz',
                'valor_dom': 'valor dominio'
            }, inplace=True)
            
            divergencias = divergencias.sort_values(by=['data', 'nota'])
            divergencias['data'] = pd.to_datetime(divergencias['data']).dt.strftime('%d/%m/%Y')
            
            df_final = divergencias[['data', 'nota', 'valor sefaz', 'valor dominio']].reset_index(drop=True)
            
            st.subheader("🔍 Detalhamento das Divergências Encontradas")
            
            if not df_final.empty:
                st.warning(f"Foram identificadas {len(df_final)} notas com inconsistências de valores ou datas.")
                
                df_visual = df_final.copy()
                df_visual['valor sefaz'] = df_visual['valor sefaz'].apply(formatar_moeda_br)
                df_visual['valor dominio'] = df_visual['valor dominio'].apply(formatar_moeda_br)
                st.dataframe(df_visual, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Divergências')
                dados_excel = output.getvalue()
                
                st.write("")
                st.download_button(
                    label="📥 Clique aqui para baixar a Planilha de Divergências",
                    data=dados_excel,
                    file_name="divergencias_sefaz_dominio.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("🎉 Excelente! Nenhuma divergência individual foi encontrada entre os arquivos.")
