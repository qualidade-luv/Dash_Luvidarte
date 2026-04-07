import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# ======================
# CONFIGURAÇÕES
# ======================
CAMINHO_CREDENCIAIS = r'\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\4-TRS\dashboard-gerencial-492613-042470f98e27.json'
ID_PLANILHA = '1Hjy4UGtgwIPJgqmcv46LyXNWOrYk_oeJWWV5vlfKF2k'
ABA = 'TRS_INDUSTRIAL'

st.set_page_config(page_title="Dashboard TRS", layout="wide")

# ======================
# FUNÇÃO PARA CONVERTER NÚMERO (CORRIGIDA)
# ======================
def converter_numero_br(valor):
    """Converte string com formato brasileiro (1.234,56) para float"""
    if valor is None or pd.isna(valor):
        return 0.0
    
    try:
        if isinstance(valor, (int, float)):
            # Se o número for muito grande (> 1e9), provavelmente está errado
            if valor > 1e9:
                return 0.0
            return float(valor)
        
        valor_str = str(valor).strip()
        if not valor_str:
            return 0.0
        
        # Remover % se houver
        if '%' in valor_str:
            valor_str = valor_str.replace('%', '')
        
        # Para números como "1.705" (milhar) ou "1.234,56"
        # Contar quantos pontos e vírgulas
        num_pontos = valor_str.count('.')
        num_virgulas = valor_str.count(',')
        
        # Caso: tem vírgula (formato brasileiro com decimal)
        if num_virgulas > 0:
            # Remove pontos de milhar e troca vírgula por ponto
            valor_str = valor_str.replace('.', '').replace(',', '.')
        elif num_pontos > 0:
            # Verifica se é ponto de milhar (ex: 1.234) ou decimal (ex: 1234.56)
            partes = valor_str.split('.')
            # Se tem mais de 2 partes ou a última parte tem 3 dígitos, é milhar
            if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
                # Remove pontos de milhar
                valor_str = valor_str.replace('.', '')
            # Se não, mantém (já é decimal)
        
        # Remover qualquer caractere não numérico exceto ponto
        valor_str = re.sub(r'[^\d.-]', '', valor_str)
        
        if not valor_str or valor_str == '.':
            return 0.0
        
        resultado = float(valor_str)
        # Limitar valores absurdos (Meta não deve ser maior que 10 milhões)
        if resultado > 10_000_000:
            return 0.0
        
        return resultado
    except:
        return 0.0

# ======================
# FUNÇÃO PARA CONVERTER DATA (CORRIGIDA - SEM DATAS FUTURAS)
# ======================
def converter_data_br(data_str, data_referencia=None):
    """Converte data no formato dd/mm/aaaa, retorna None se for inválida ou futura"""
    if data_str is None or pd.isna(data_str):
        return None
    
    try:
        if isinstance(data_str, (datetime, pd.Timestamp)):
            # Se já for datetime, verificar se não é futura
            if data_referencia and data_str > data_referencia:
                return None
            if data_str > datetime.now():
                return None
            return data_str
        
        data_str = str(data_str).strip()
        if not data_str:
            return None
        
        # Tentar extrair data no formato dd/mm/aaaa
        if '/' in data_str:
            partes = data_str.split('/')
            if len(partes) == 3:
                dia, mes, ano = int(partes[0]), int(partes[1]), int(partes[2])
                if ano < 100:
                    ano = 2000 + ano
                
                # Validar se é uma data real
                data_obj = datetime(ano, mes, dia)
                
                # Verificar se não é data futura (depois de hoje)
                hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                if data_obj > hoje:
                    return None
                
                # Verificar se a data não é muito antiga (opcional: antes de 2020)
                if data_obj.year < 2020:
                    return None
                
                return data_obj
        
        # Tentar converter com pandas
        data_obj = pd.to_datetime(data_str, errors='coerce', dayfirst=True)
        if pd.notna(data_obj):
            # Verificar se não é data futura
            hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            if data_obj > hoje:
                return None
            return data_obj
        
        return None
    except:
        return None

# ======================
# FUNÇÃO PARA CARREGAR DADOS DO GOOGLE SHEETS
# ======================
@st.cache_data(ttl=300)
def carregar_dados():
    """Carrega dados do Google Sheets com cabeçalho na linha 2"""
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_key(ID_PLANILHA).worksheet(ABA)
        
        # Carregar TODOS os dados
        todos_dados = sheet.get_all_values()
        
        if len(todos_dados) < 2:
            st.error("Planilha vazia ou com dados insuficientes")
            return pd.DataFrame()
        
        # CABEÇALHO na LINHA 2 (índice 1)
        cabecalho = todos_dados[1]
        
        # DADOS a partir da LINHA 3 (índice 2)
        valores = todos_dados[2:]
        
        # Criar DataFrame
        df = pd.DataFrame(valores, columns=cabecalho)
        df.columns = df.columns.str.strip().str.upper()
        
        # Converter DATA (ignorando datas inválidas e futuras)
        if 'DATA' in df.columns:
            df['DATA'] = df['DATA'].apply(converter_data_br)
            # Remover linhas sem data válida
            df = df.dropna(subset=['DATA'])
        
        # Renomear APROVADO FINAL para EMBALADO
        if 'APROVADO FINAL' in df.columns:
            df = df.rename(columns={'APROVADO FINAL': 'EMBALADO'})
        
        # Converter colunas numéricas
        colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO', 'BOQUETA']
        for col in colunas_numericas:
            if col in df.columns:
                df[col] = df[col].apply(converter_numero_br)
        
        return df
        
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

# ======================
# CARREGAR DADOS
# ======================
with st.spinner("Carregando dados do Google Sheets..."):
    df_base = carregar_dados()

if df_base.empty:
    st.warning("Não foi possível carregar os dados. Verifique a conexão.")
    st.stop()

# ======================
# CALCULAR MELHORES TRS DA BASE COMPLETA
# ======================
df_base_calc = df_base.copy()

colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO']
for col in colunas_numericas:
    if col in df_base_calc.columns:
        df_base_calc[col] = pd.to_numeric(df_base_calc[col], errors='coerce').fillna(0)

# Calcular TRS FINAL para toda a base
if 'EMBALADO' in df_base_calc.columns and 'META LIQUIDA' in df_base_calc.columns:
    df_base_calc['TRS FINAL (%)'] = df_base_calc.apply(
        lambda row: (row['EMBALADO'] / row['META LIQUIDA'] * 100) if row['META LIQUIDA'] != 0 else 0, 
        axis=1
    )
    df_base_calc['TRS FINAL (%)'] = df_base_calc['TRS FINAL (%)'].round(2)
else:
    df_base_calc['TRS FINAL (%)'] = 0

# Encontrar os melhores TRS FINAL históricos
melhores_trs_historico = {}
if 'REFERÊNCIA' in df_base_calc.columns:
    for ref in df_base_calc['REFERÊNCIA'].unique():
        ref_df = df_base_calc[df_base_calc['REFERÊNCIA'] == ref]
        if not ref_df.empty:
            max_trs = ref_df['TRS FINAL (%)'].max()
            if max_trs > 0:
                melhores_trs_historico[ref] = max_trs

# ======================
# SIDEBAR (FILTROS) - SEM DEBUG
# ======================
st.sidebar.title("Filtros")

filtro_melhores_trs = st.sidebar.checkbox("Apenas Melhores TRS por Referência", value=False)

data_ini = st.sidebar.date_input("Data inicial", value=None)
data_fim = st.sidebar.date_input("Data final", value=None)
turno = st.sidebar.selectbox("Turno", options=["(Todos)", "M", "T", "N"])
referencia = st.sidebar.text_input("Referência (parte do código)")
prensa_tipo = st.sidebar.selectbox("Tipo de prensa", ["(Todos)", "Semi-Automática", "Automática"])
mostrar_defeitos = st.sidebar.checkbox("Exibir Somatório de Defeitos", value=True)

qtd = st.sidebar.number_input(
    "Qtd de linhas na tabela (0 = mostrar todas)",
    min_value=0,
    max_value=5000,
    value=20,
    step=10
)

# ======================
# APLICAÇÃO DOS FILTROS
# ======================
df = df_base.copy()

if data_ini:
    df = df[df['DATA'] >= pd.to_datetime(data_ini)]
if data_fim:
    df = df[df['DATA'] <= pd.to_datetime(data_fim)]

if turno != "(Todos)" and 'TURNO' in df.columns:
    df = df[df['TURNO'].fillna('').str.upper() == turno.upper()]

if referencia and 'REFERÊNCIA' in df.columns:
    df = df[df['REFERÊNCIA'].fillna('').str.lower().str.contains(referencia.lower())]

if prensa_tipo != "(Todos)" and 'BOQUETA' in df.columns:
    if "Semi" in prensa_tipo:
        df = df[df['BOQUETA'] == 1]
    elif "Auto" in prensa_tipo:
        df = df[df['BOQUETA'] == 2]

# ======================
# KPIs (após filtros)
# ======================
if not df.empty:
    colunas_para_converter = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO']
    for col in colunas_para_converter:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    total_prod = int(df['PRODUZIDO'].sum())
    total_apro = int(df['APROVADO'].sum())
    total_embal = int(df['EMBALADO'].sum()) if 'EMBALADO' in df.columns else 0
    total_meta = int(df['META LIQUIDA'].sum())
    
    trs_total = (total_apro / total_meta * 100) if total_meta else 0
    trs_final_total = (total_embal / total_meta * 100) if total_meta else 0
else:
    total_prod = total_apro = total_embal = total_meta = trs_total = trs_final_total = 0

# Mostrar 4 KPIs
col1, col2, col3, col4 = st.columns(4)
col1.metric("Produzido", f"{total_prod:,}".replace(",", "."))
col2.metric("Aprovado", f"{total_apro:,}".replace(",", "."))
col3.metric("Embalado", f"{total_embal:,}".replace(",", "."))
col4.metric("TRS Final (%)", f"{trs_final_total:.2f}%")

# ======================
# TABELA PRINCIPAL
# ======================
st.subheader("📋 Tabela de Produção")

if not df.empty:
    df['TRS (%)'] = df.apply(
        lambda row: (row['APROVADO'] / row['META LIQUIDA'] * 100) if row['META LIQUIDA'] != 0 else 0, 
        axis=1
    )
    
    if 'EMBALADO' in df.columns:
        df['TRS FINAL (%)'] = df.apply(
            lambda row: (row['EMBALADO'] / row['META LIQUIDA'] * 100) if row['META LIQUIDA'] != 0 else 0, 
            axis=1
        )
    else:
        df['TRS FINAL (%)'] = 0
    
    df['TRS (%)'] = df['TRS (%)'].round(2)
    df['TRS FINAL (%)'] = df['TRS FINAL (%)'].round(2)

# Ordena da data mais recente para a mais antiga
df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

# Aplicar filtro de melhores TRS se selecionado
if filtro_melhores_trs and not df_sorted.empty and 'REFERÊNCIA' in df_sorted.columns:
    df_sorted = df_sorted[df_sorted.apply(
        lambda row: row['REFERÊNCIA'] in melhores_trs_historico and 
                    abs(row['TRS FINAL (%)'] - melhores_trs_historico[row['REFERÊNCIA']]) < 0.01,
        axis=1
    )].reset_index(drop=True)
    
    if not df_sorted.empty:
        st.info(f"Mostrando {len(df_sorted)} registro(s) com Melhor TRS Final Histórico por referência")
    else:
        st.warning("Nenhum registro encontrado com Melhor TRS Final Histórico")

df_view = df_sorted if qtd == 0 else df_sorted.head(qtd)

if not df_view.empty:
    df_display = df_view.copy()
    
    df_display['DATA'] = pd.to_datetime(df_display['DATA']).dt.strftime('%d/%m/%Y')
    
    colunas_inteiras = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'REFUGADO', 'META LIQUIDA']
    for col in colunas_inteiras:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: int(round(x, 0)) if pd.notnull(x) else 0)
            df_display[col] = df_display[col].apply(lambda x: f"{x:,}".replace(",", "."))
    
    if 'TRS (%)' in df_display.columns:
        df_display['TRS (%)'] = df_display['TRS (%)'].apply(lambda x: f"{x:.2f}%")
    
    if 'TRS FINAL (%)' in df_display.columns:
        df_display['TRS FINAL (%)'] = df_display['TRS FINAL (%)'].apply(lambda x: f"{x:.2f}%")
    
    colunas_exibir = ['DATA', 'REFERÊNCIA', 'TURNO', 'PRODUZIDO', 'APROVADO', 'EMBALADO', 
                      'REFUGADO', 'META LIQUIDA', 'TRS (%)', 'TRS FINAL (%)']
    colunas_exibir = [col for col in colunas_exibir if col in df_display.columns]
    
    def highlight_max_trs_final(row):
        ref = row['REFERÊNCIA']
        if ref in melhores_trs_historico:
            trs_final_str = row['TRS FINAL (%)']
            if isinstance(trs_final_str, str):
                trs_final_valor = float(trs_final_str.replace('%', '').replace(',', '.'))
            else:
                trs_final_valor = float(trs_final_str)
            
            if abs(trs_final_valor - melhores_trs_historico[ref]) < 0.01:
                return ['background-color: #FFA500; color: black; font-weight: bold'] * len(row)
        return [''] * len(row)
    
    styled_df = df_display[colunas_exibir].style.apply(highlight_max_trs_final, axis=1)
    
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=400
    )
    
    if not filtro_melhores_trs:
        st.caption("🎯 Linhas em laranja: Melhor TRS Final Histórico para cada referência")
else:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")

# ======================
# GRÁFICO DE TRS POR TURNO
# ======================
if not df.empty and 'TURNO' in df.columns and 'EMBALADO' in df.columns:
    st.subheader("📊 TRS Final por Turno")

    turno_data = []
    for t in df['TURNO'].unique():
        df_turno = df[df['TURNO'] == t]
        total_embal_turno = df_turno['EMBALADO'].sum()
        total_meta_turno = df_turno['META LIQUIDA'].sum()
        trs_final_turno = (total_embal_turno / total_meta_turno * 100) if total_meta_turno > 0 else 0
        turno_data.append({'Turno': t, 'TRS Final': trs_final_turno})
    
    df_turno_trs = pd.DataFrame(turno_data)
    
    if not df_turno_trs.empty:
        fig, ax = plt.subplots()
        cores = {'M': '#1f77b4', 'T': '#ff7f0e', 'N': '#2ca02c'}
        bar_colors = [cores.get(t, '#808080') for t in df_turno_trs['Turno']]
        
        df_turno_trs.plot(x='Turno', y='TRS Final', kind='bar', ax=ax, 
                         color=bar_colors, edgecolor="black")
        ax.set_ylabel("TRS Final (%)")
        ax.set_ylim(0, 110)
        ax.set_title("TRS Final por Turno")
        
        for i, v in enumerate(df_turno_trs['TRS Final']):
            ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom', fontweight='bold')
        
        st.pyplot(fig)
    else:
        st.info("Nenhum turno com META válida para calcular TRS.")

# ======================
# DEFEITOS
# ======================
if mostrar_defeitos:
    st.subheader("❌ Estratificação de Defeitos")

    try:
        colunas_basicas = ['DATA', 'REFERÊNCIA', 'TURNO', 'PRODUZIDO', 'APROVADO', 
                          'EMBALADO', 'REFUGADO', 'META LIQUIDA', 'BOQUETA']
        colunas_basicas = [c for c in colunas_basicas if c in df.columns]
        
        idx_inicio = len(colunas_basicas)
        idx_fim = min(idx_inicio + 20, len(df.columns))
        
        if idx_inicio < idx_fim:
            defeitos_cols = df.columns[idx_inicio:idx_fim]
            df_defeitos = df[defeitos_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_defeitos_sum = df_defeitos.sum().sort_values(ascending=False)
            df_defeitos_sum = df_defeitos_sum[df_defeitos_sum > 0]

            if not df_defeitos_sum.empty:
                st.bar_chart(df_defeitos_sum)
                
                df_defeitos_display = df_defeitos_sum.reset_index().rename(
                    columns={'index':'Defeito', 0:'Quantidade'}
                ).sort_values('Quantidade', ascending=False)
                
                df_defeitos_display['Quantidade'] = df_defeitos_display['Quantidade'].apply(
                    lambda x: f"{int(x):,}".replace(",", ".")
                )
                
                st.dataframe(df_defeitos_display, use_container_width=True)
            else:
                st.info("Nenhum defeito encontrado nos filtros aplicados.")
        else:
            st.info("Colunas de defeitos não identificadas.")
    except Exception as e:
        st.warning(f"Não foi possível calcular defeitos: {e}")