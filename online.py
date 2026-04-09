import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import os
import numpy as np

# ======================
# CONFIGURAÇÕES
# ======================
CAMINHO_CREDENCIAIS = r'\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\4-TRS\dashboard-gerencial-492613-042470f98e27.json'
ID_PLANILHA = '1Hjy4UGtgwIPJgqmcv46LyXNWOrYk_oeJWWV5vlfKF2k'

# Praças que NÃO são SOPRO (devem ser EXCLUÍDAS da aba SOPRO)
PRACAS_NAO_SOPRO = ['GIL', 'GILSIMAR', 'ED CARLOS', 'EDI CARLOS', 'ROBÔ 2', 'ROBÔ-2', 'ROBÔ', 'ROBO']

# Abas disponíveis - APENAS DUAS
ABAS = {
    'PRENSADOS': 'TRS_INDUSTRIAL',
    'SOPRO': 'TRS_SOPRO'
}

st.set_page_config(page_title="Dashboard TRS", layout="wide")

# ======================
# SELEÇÃO DA ABA (SIDEBAR)
# ======================
st.sidebar.title("Navegação")
aba_selecionada = st.sidebar.radio("Selecione o setor:", list(ABAS.keys()))

# ======================
# FUNÇÃO PARA CARREGAR CREDENCIAIS
# ======================
def get_gspread_client():
    """Cria cliente gspread tentando primeiro secrets, depois arquivo local"""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        if 'gcp_service_account' in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            return client
    except Exception as e:
        pass
    
    try:
        if not os.path.exists(CAMINHO_CREDENCIAIS):
            st.error(f"Arquivo de credenciais não encontrado em: {CAMINHO_CREDENCIAIS}")
            return None
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(CAMINHO_CREDENCIAIS, scope)
        client = gspread.authorize(creds)
        return client
        
    except Exception as e:
        st.error(f"Erro ao autenticar com Google Sheets (local): {e}")
        return None

# ======================
# FUNÇÕES AUXILIARES COMUNS
# ======================
def converter_numero_br(valor):
    if valor is None or pd.isna(valor):
        return 0.0
    
    try:
        if isinstance(valor, (int, float)):
            if valor > 1e9:
                return 0.0
            return float(valor)
        
        valor_str = str(valor).strip()
        if not valor_str:
            return 0.0
        
        if '%' in valor_str:
            valor_str = valor_str.replace('%', '')
        
        num_pontos = valor_str.count('.')
        num_virgulas = valor_str.count(',')
        
        if num_virgulas > 0:
            valor_str = valor_str.replace('.', '').replace(',', '.')
        elif num_pontos > 0:
            partes = valor_str.split('.')
            if len(partes) > 2 or (len(partes) == 2 and len(partes[1]) == 3):
                valor_str = valor_str.replace('.', '')
        
        valor_str = re.sub(r'[^\d.-]', '', valor_str)
        
        if not valor_str or valor_str == '.':
            return 0.0
        
        resultado = float(valor_str)
        if resultado > 10_000_000:
            return 0.0
        
        return resultado
    except:
        return 0.0

def converter_data_br(data_str):
    if data_str is None or pd.isna(data_str):
        return None
    
    try:
        if isinstance(data_str, (datetime, pd.Timestamp)):
            if data_str > datetime.now():
                return None
            return data_str
        
        data_str = str(data_str).strip()
        if not data_str:
            return None
        
        if '/' in data_str:
            partes = data_str.split('/')
            if len(partes) == 3:
                dia, mes, ano = int(partes[0]), int(partes[1]), int(partes[2])
                if ano < 100:
                    ano = 2000 + ano
                
                data_obj = datetime(ano, mes, dia)
                hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                if data_obj > hoje:
                    return None
                if data_obj.year < 2020:
                    return None
                
                return data_obj
        
        data_obj = pd.to_datetime(data_str, errors='coerce', dayfirst=True)
        if pd.notna(data_obj):
            hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            if data_obj > hoje:
                return None
            return data_obj
        
        return None
    except:
        return None

def minutos_para_horas_str(minutos):
    if pd.isna(minutos) or minutos is None or minutos == 0:
        return "00:00"
    horas = int(minutos) // 60
    mins = int(minutos) % 60
    return f"{horas:02d}:{mins:02d}"

def converter_tempo_para_minutos(valor):
    if pd.isna(valor) or valor is None or valor == '':
        return 0
    
    if hasattr(valor, 'hour') and hasattr(valor, 'minute'):
        try:
            return valor.hour * 60 + valor.minute + (valor.second // 60 if hasattr(valor, 'second') else 0)
        except:
            pass
    
    if isinstance(valor, str):
        valor = valor.strip()
        if not valor:
            return 0
        
        if ':' in valor:
            partes = valor.split(':')
            try:
                if len(partes) == 3:
                    h, m, s = map(int, partes)
                    return h * 60 + m + s // 60
                elif len(partes) == 2:
                    h, m = map(int, partes)
                    return h * 60 + m
            except:
                pass
        
        try:
            num = float(valor.replace(',', '.'))
            if num > 0:
                if num < 24:
                    return int(num * 60)
                else:
                    return int(num)
        except:
            pass
        
        return 0
    
    elif isinstance(valor, (int, float)):
        if valor > 0:
            if valor < 24:
                return int(valor * 60)
            elif valor > 100:
                return int(valor)
            else:
                return int(valor * 60)
    
    return 0

# ==================================================================================================
# SE FOR PRENSADOS (MANTIDO EXATAMENTE COMO ESTAVA)
# ==================================================================================================
if aba_selecionada == 'PRENSADOS':
    ABA = 'TRS_INDUSTRIAL'
    
    @st.cache_data(ttl=300)
    def carregar_dados_prensados():
        try:
            client = get_gspread_client()
            if client is None:
                return pd.DataFrame()

            sheet = client.open_by_key(ID_PLANILHA).worksheet(ABA)
            todos_dados = sheet.get_all_values()
            
            if len(todos_dados) < 2:
                st.error("Planilha vazia ou com dados insuficientes")
                return pd.DataFrame()
            
            cabecalho = todos_dados[1]
            valores = todos_dados[2:]
            
            df = pd.DataFrame(valores, columns=cabecalho)
            df.columns = df.columns.str.strip().str.upper()
            
            if 'DATA' in df.columns:
                df['DATA'] = df['DATA'].apply(converter_data_br)
                df = df.dropna(subset=['DATA'])
            
            if 'APROVADO FINAL' in df.columns:
                df = df.rename(columns={'APROVADO FINAL': 'EMBALADO'})
            
            colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO', 'BOQUETA']
            for col in colunas_numericas:
                if col in df.columns:
                    df[col] = df[col].apply(converter_numero_br)
            
            # Adicionar colunas auxiliares para gráficos de parada
            df['ANO_MES'] = df['DATA'].dt.to_period('M').astype(str)
            df['DIA_SEMANA'] = df['DATA'].dt.day_name()
            df['SEMANA'] = df['DATA'].dt.isocalendar().week
            df['IS_SABADO'] = df['DATA'].dt.dayofweek == 5
            
            for col in df.columns:
                col_upper = str(col).upper()
                if 'ACERTO' in col_upper and 'MIN' not in col_upper:
                    df['ACERTOS_MIN'] = df[col].apply(converter_tempo_para_minutos)
                if 'MANUT' in col_upper and 'MIN' not in col_upper:
                    df['MANUT_MIN'] = df[col].apply(converter_tempo_para_minutos)
                if 'HORAS TOTAIS' in col_upper or 'HORA TOTAL' in col_upper:
                    df['HORAS_TOTAIS_MIN'] = df[col].apply(converter_tempo_para_minutos)
            
            if 'ACERTOS_MIN' in df.columns:
                df['ACERTOS_MIN_AJUSTADO'] = df.apply(
                    lambda row: max(0, row['ACERTOS_MIN'] - 165) if row['IS_SABADO'] else row['ACERTOS_MIN'],
                    axis=1
                )
            
            return df
            
        except Exception as e:
            st.error(f"Erro ao carregar dados: {e}")
            return pd.DataFrame()

    with st.spinner("Carregando dados do setor PRENSADOS..."):
        df_base = carregar_dados_prensados()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados. Verifique a conexão.")
        st.stop()

    # Calcular TRS
    df_base_calc = df_base.copy()
    colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO']
    for col in colunas_numericas:
        if col in df_base_calc.columns:
            df_base_calc[col] = pd.to_numeric(df_base_calc[col], errors='coerce').fillna(0)

    if 'EMBALADO' in df_base_calc.columns and 'META LIQUIDA' in df_base_calc.columns:
        df_base_calc['TRS FINAL (%)'] = df_base_calc.apply(
            lambda row: (row['EMBALADO'] / row['META LIQUIDA'] * 100) if row['META LIQUIDA'] != 0 else 0, 
            axis=1
        )
        df_base_calc['TRS FINAL (%)'] = df_base_calc['TRS FINAL (%)'].round(2)
    else:
        df_base_calc['TRS FINAL (%)'] = 0

    melhores_trs_historico = {}
    if 'REFERÊNCIA' in df_base_calc.columns:
        for ref in df_base_calc['REFERÊNCIA'].unique():
            ref_df = df_base_calc[df_base_calc['REFERÊNCIA'] == ref]
            if not ref_df.empty:
                max_trs = ref_df['TRS FINAL (%)'].max()
                if max_trs > 0:
                    melhores_trs_historico[ref] = max_trs

    # Sidebar filtros
    st.sidebar.title("Filtros - PRENSADOS")
    filtro_melhores_trs = st.sidebar.checkbox("Apenas Melhores TRS por Referência", value=False)
    data_ini = st.sidebar.date_input("Data inicial", value=None, key="prensados_data_ini")
    data_fim = st.sidebar.date_input("Data final", value=None, key="prensados_data_fim")
    turno = st.sidebar.selectbox("Turno", options=["(Todos)", "M", "T", "N"], key="prensados_turno")
    referencia = st.sidebar.text_input("Referência (parte do código)", key="prensados_ref")
    prensa_tipo = st.sidebar.selectbox("Tipo de prensa", ["(Todos)", "Semi-Automática", "Automática"], key="prensados_tipo")
    mostrar_defeitos = st.sidebar.checkbox("Exibir Somatório de Defeitos", value=True, key="prensados_defeitos")
    qtd = st.sidebar.number_input("Qtd de linhas na tabela (0 = mostrar todas)", min_value=0, max_value=5000, value=20, step=10, key="prensados_qtd")

    # Aplicar filtros
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

    # KPIs
    if not df.empty:
        colunas_para_converter = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO']
        for col in colunas_para_converter:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        total_prod = int(df['PRODUZIDO'].sum())
        total_apro = int(df['APROVADO'].sum())
        total_embal = int(df['EMBALADO'].sum()) if 'EMBALADO' in df.columns else 0
        total_meta = int(df['META LIQUIDA'].sum())
        trs_final_total = (total_embal / total_meta * 100) if total_meta else 0
    else:
        total_prod = total_apro = total_embal = total_meta = trs_final_total = 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Produzido", f"{total_prod:,}".replace(",", "."))
    col2.metric("Aprovado", f"{total_apro:,}".replace(",", "."))
    col3.metric("Embalado", f"{total_embal:,}".replace(",", "."))
    col4.metric("TRS Final (%)", f"{trs_final_total:.2f}%")

    # Tabela principal
    st.subheader("📋 Tabela de Produção - PRENSADOS")

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

    df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

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
        
        def destacar_melhor_trs(row):
            ref = row['REFERÊNCIA']
            if ref in melhores_trs_historico:
                trs_val = row['TRS FINAL (%)']
                if isinstance(trs_val, str):
                    trs_val = float(trs_val.replace('%', '').replace(',', '.'))
                if abs(trs_val - melhores_trs_historico[ref]) < 0.01:
                    return ['background-color: #FFA500; color: black; font-weight: bold'] * len(row)
            return [''] * len(row)
        
        styled_df = df_display[colunas_exibir].style.apply(destacar_melhor_trs, axis=1)
        st.dataframe(styled_df, use_container_width=True, height=400)
        
        if not filtro_melhores_trs:
            st.caption("🎯 Linhas em laranja: Melhor TRS Final Histórico para cada referência")

    # Gráfico TRS Diário
    st.subheader("📈 TRS - Evolução Diária")
    
    if not df.empty:
        resumo_dia = df.groupby(df['DATA'].dt.date).agg({
            'PRODUZIDO': 'sum',
            'APROVADO': 'sum',
            'META LIQUIDA': 'sum'
        }).reset_index()
        
        resumo_dia['DATA'] = pd.to_datetime(resumo_dia['DATA'])
        resumo_dia['TRS (%)'] = (resumo_dia['APROVADO'] / resumo_dia['META LIQUIDA'].replace(0, 1) * 100).fillna(0)
        resumo_dia = resumo_dia.sort_values('DATA')
        
        if not resumo_dia.empty:
            fig, ax = plt.subplots(figsize=(14, 6), facecolor='black')
            ax.set_facecolor('black')
            
            ax.plot(resumo_dia['DATA'], resumo_dia['TRS (%)'], 
                    marker='o', markersize=8, linewidth=3, 
                    color='cyan', alpha=0.8, label='TRS Diário')
            
            if len(resumo_dia) > 1:
                media_movel = resumo_dia['TRS (%)'].rolling(window=min(3, len(resumo_dia)), min_periods=1).mean()
                ax.plot(resumo_dia['DATA'], media_movel, 
                        color='yellow', alpha=0.7, linewidth=2, 
                        linestyle='--', label='Média 3 Dias')
            
            ax.axhline(y=85, color='red', linestyle=':', alpha=0.7, linewidth=2, label='Meta (85%)')
            ax.fill_between(resumo_dia['DATA'], 0, resumo_dia['TRS (%)'], alpha=0.2, color='cyan')
            
            ax.set_ylabel("TRS (%)", fontsize=12, color='white')
            ax.set_xlabel("Data", fontsize=12, color='white')
            ax.set_title("Evolução do TRS - Período Selecionado", fontsize=16, fontweight='bold', color='white')
            ax.tick_params(colors='white')
            ax.grid(True, alpha=0.3, color='#444444')
            ax.legend(facecolor='black', edgecolor='white', labelcolor='white')
            
            for spine in ax.spines.values():
                spine.set_edgecolor('white')
            
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    st.markdown("---")
    
    # Média Manual vs Automática
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🔧 Média TRS - Manual")
        if 'BOQUETA' in df.columns:
            df_manual = df[df['BOQUETA'] == 1]
            if not df_manual.empty:
                total_aprovado_manual = df_manual['APROVADO'].sum()
                total_meta_manual = df_manual['META LIQUIDA'].replace(0, 1).sum()
                media_manual = (total_aprovado_manual / total_meta_manual * 100) if total_meta_manual > 0 else 0
                prod_manual = df_manual['PRODUZIDO'].sum()
                
                st.markdown(f"""
                <div style="background: #111111; padding: 20px; border-radius: 10px; border: 2px solid cyan; text-align: center;">
                    <p style="font-size: 48px; font-weight: bold; color: cyan; margin: 10px 0;">{media_manual:.1f}%</p>
                    <p style="font-size: 16px; color: white;">Produção: {prod_manual:,.0f} un</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("Sem dados para Prensa Manual")
    
    with col2:
        st.subheader("⚙️ Média TRS - Automática")
        if 'BOQUETA' in df.columns:
            df_auto = df[df['BOQUETA'] == 2]
            if not df_auto.empty:
                total_aprovado_auto = df_auto['APROVADO'].sum()
                total_meta_auto = df_auto['META LIQUIDA'].replace(0, 1).sum()
                media_auto = (total_aprovado_auto / total_meta_auto * 100) if total_meta_auto > 0 else 0
                prod_auto = df_auto['PRODUZIDO'].sum()
                
                st.markdown(f"""
                <div style="background: #111111; padding: 20px; border-radius: 10px; border: 2px solid lime; text-align: center;">
                    <p style="font-size: 48px; font-weight: bold; color: lime; margin: 10px 0;">{media_auto:.1f}%</p>
                    <p style="font-size: 16px; color: white;">Produção: {prod_auto:,.0f} un</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("Sem dados para Prensa Automática")
    
    st.markdown("---")
    
    # Análise de Paradas
    st.subheader("⏱️ Análise de Paradas")
    
    horas_trabalhadas = 0
    total_acertos = 0
    total_manut = 0
    
    if 'HORAS_TOTAIS_MIN' in df.columns:
        horas_trabalhadas = df['HORAS_TOTAIS_MIN'].sum()
    else:
        dias_uteis = df[~df['IS_SABADO']]['DATA'].nunique() if 'IS_SABADO' in df.columns else 0
        dias_sabado = df[df['IS_SABADO']]['DATA'].nunique() if 'IS_SABADO' in df.columns else 0
        horas_trabalhadas = (dias_uteis * 8 * 60) + (dias_sabado * 6 * 60)
    
    total_acertos = df['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
    total_manut = df['MANUT_MIN'].sum() if 'MANUT_MIN' in df.columns else 0
    total_paradas = total_acertos + total_manut
    horas_produtivas = max(0, horas_trabalhadas - total_paradas)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if 'BOQUETA' in df.columns:
            df_manual_paradas = df[df['BOQUETA'] == 1]
            df_auto_paradas = df[df['BOQUETA'] == 2]
            
            acertos_manual = df_manual_paradas['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
            manut_manual = df_manual_paradas['MANUT_MIN'].sum() if 'MANUT_MIN' in df.columns else 0
            acertos_auto = df_auto_paradas['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
            manut_auto = df_auto_paradas['MANUT_MIN'].sum() if 'MANUT_MIN' in df.columns else 0
            
            categorias = ['Manual', 'Automática']
            acertos = [acertos_manual, acertos_auto]
            manut = [manut_manual, manut_auto]
            
            fig, ax = plt.subplots(figsize=(8, 6), facecolor='black')
            ax.set_facecolor('black')
            
            x = np.arange(len(categorias))
            width = 0.6
            
            bars_acertos = ax.bar(x, acertos, width, label='ACERTOS', color='#f39c12', alpha=0.8, edgecolor='white')
            bars_manut = ax.bar(x, manut, width, bottom=acertos, label='MANUTENÇÃO', color='#e74c3c', alpha=0.8, edgecolor='white')
            
            for i, (a, m) in enumerate(zip(acertos, manut)):
                total = a + m
                if a > 0:
                    ax.text(i, a/2, minutos_para_horas_str(a), ha='center', va='center', color='white', fontweight='bold')
                if m > 0:
                    ax.text(i, a + m/2, minutos_para_horas_str(m), ha='center', va='center', color='white', fontweight='bold')
                if total > 0:
                    ax.text(i, total + (max(acertos + manut) * 0.03), f'TOTAL: {minutos_para_horas_str(total)}', 
                           ha='center', va='bottom', color='white', fontweight='bold', fontsize=11)
            
            ax.set_ylabel("Minutos", color='white')
            ax.set_xlabel("Tipo de Prensa", color='white')
            ax.set_title("Composição das Paradas", fontsize=14, fontweight='bold', color='white')
            ax.set_xticks(x)
            ax.set_xticklabels(categorias)
            ax.tick_params(colors='white')
            ax.legend(facecolor='black', edgecolor='white', labelcolor='white')
            ax.grid(True, alpha=0.3, color='#444444')
            
            for spine in ax.spines.values():
                spine.set_edgecolor('white')
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    with col2:
        if horas_trabalhadas > 0:
            fig, ax = plt.subplots(figsize=(8, 6), facecolor='black')
            ax.set_facecolor('black')
            
            labels = ['Horas Produtivas', 'Acertos', 'Manutenção']
            valores = [horas_produtivas, total_acertos, total_manut]
            cores_pizza = ['#2ecc71', '#f39c12', '#e74c3c']
            
            labels_filtrados = []
            valores_filtrados = []
            cores_filtradas = []
            for l, v, c in zip(labels, valores, cores_pizza):
                if v > 0:
                    labels_filtrados.append(l)
                    valores_filtrados.append(v)
                    cores_filtradas.append(c)
            
            if valores_filtrados:
                wedges, texts, autotexts = ax.pie(valores_filtrados, labels=labels_filtrados, colors=cores_filtradas,
                                                   autopct='%1.1f%%', startangle=90, textprops={'color': 'white'})
                
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontweight('bold')
                
                ax.set_title(f"Distribuição do Tempo Total\n{minutos_para_horas_str(horas_trabalhadas)} trabalhadas", 
                            fontsize=14, fontweight='bold', color='white')
                
                legenda_texto = f"Horas Produtivas: {minutos_para_horas_str(horas_produtivas)}\n"
                legenda_texto += f"Acertos: {minutos_para_horas_str(total_acertos)}\n"
                legenda_texto += f"Manutenção: {minutos_para_horas_str(total_manut)}"
                
                ax.text(1.3, 0.5, legenda_texto, transform=ax.transAxes, fontsize=11, color='white',
                       bbox=dict(boxstyle='round', facecolor='#222222', edgecolor='white', alpha=0.8),
                       verticalalignment='center')
                
                fig.tight_layout()
                st.pyplot(fig)
                plt.close(fig)
        else:
            st.info("Sem dados de tempo para exibir")
    
    # Resumo de Paradas
    st.markdown("---")
    st.subheader("📊 Resumo de Paradas")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Horas Trabalhadas", minutos_para_horas_str(horas_trabalhadas))
    col2.metric("Total Acertos", minutos_para_horas_str(total_acertos))
    col3.metric("Total Manutenção", minutos_para_horas_str(total_manut))
    col4.metric("Horas Produtivas", minutos_para_horas_str(horas_produtivas))
    
    st.markdown("---")
    
    # Gráfico TRS por Turno
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
            fig, ax = plt.subplots(figsize=(10, 5))
            cores = {'M': '#1f77b4', 'T': '#ff7f0e', 'N': '#2ca02c'}
            bar_colors = [cores.get(t, '#808080') for t in df_turno_trs['Turno']]
            df_turno_trs.plot(x='Turno', y='TRS Final', kind='bar', ax=ax, color=bar_colors, edgecolor="black")
            ax.set_ylabel("TRS Final (%)")
            ax.set_ylim(0, 110)
            ax.set_title("TRS Final por Turno")
            for i, v in enumerate(df_turno_trs['TRS Final']):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom', fontweight='bold')
            st.pyplot(fig)

    # Defeitos
    if mostrar_defeitos:
        st.subheader("❌ Estratificação de Defeitos - PRENSADOS")
        colunas_defeitos = [
            'BOLHA', 'PEDRA', 'TRINCA', 'RUGAS', 'CORTE TESOURA', 'DOBRA', 
            'SUJEIRA', 'QUEBRA', 'ARREADO', 'VIDRO GRUDADO', 'CONTRA-PEÇA', 
            'FALHAS', 'CHUPADO', 'ÓLEO TESOURA', 'CROMO', 'MACHO', 'BARRO', 
            'EMPENO', 'OUTROS'
        ]
        defeitos_existentes = []
        for defeito in colunas_defeitos:
            for col in df.columns:
                if col.upper() == defeito.upper():
                    defeitos_existentes.append(col)
                    break
        
        if defeitos_existentes:
            df_defeitos = df[defeitos_existentes].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_defeitos_sum = df_defeitos.sum().sort_values(ascending=False)
            df_defeitos_sum = df_defeitos_sum[df_defeitos_sum > 0]
            if not df_defeitos_sum.empty:
                st.bar_chart(df_defeitos_sum)
                total_defeitos = df_defeitos_sum.sum()
                st.metric("Total de Defeitos", f"{int(total_defeitos):,}".replace(",", "."))

# ==================================================================================================
# SE FOR SOPRO (AGORA COM GRÁFICOS GERENCIAIS INCORPORADOS)
# ==================================================================================================
elif aba_selecionada == 'SOPRO':
    ABA = 'TRS_SOPRO'
    
    @st.cache_data(ttl=300)
    def carregar_dados_sopro():
        try:
            client = get_gspread_client()
            if client is None:
                return pd.DataFrame()

            sheet = client.open_by_key(ID_PLANILHA).worksheet(ABA)
            todos_dados = sheet.get_all_values()
            
            if len(todos_dados) < 2:
                st.error("Planilha vazia ou com dados insuficientes")
                return pd.DataFrame()
            
            cabecalho = todos_dados[0]
            valores = todos_dados[1:]
            
            df = pd.DataFrame(valores, columns=cabecalho)
            df.columns = df.columns.str.strip().str.upper()
            
            # FILTRAR PARA EXCLUIR AS PRAÇAS QUE NÃO SÃO SOPRO
            if 'PRAÇA' in df.columns:
                df['PRAÇA_NORM'] = df['PRAÇA'].fillna('').astype(str).str.upper().str.strip()
                mascara_sopro = ~df['PRAÇA_NORM'].apply(
                    lambda x: any(praca in x for praca in [p.upper() for p in PRACAS_NAO_SOPRO])
                )
                df = df[mascara_sopro].copy()
                df = df.drop(columns=['PRAÇA_NORM'])
                
                st.sidebar.info(f"✅ Filtrado: excluídas praças de PRENSADO ({len(df)} registros SOPRO)")
            
            if 'DATA' in df.columns:
                df['DATA'] = df['DATA'].apply(converter_data_br)
                df = df.dropna(subset=['DATA'])
            
            if 'PRODUZIDO' in df.columns:
                df['PRODUZIDO'] = df['PRODUZIDO'].apply(converter_numero_br)
            if 'APROVADO' in df.columns:
                df['APROVADO'] = df['APROVADO'].apply(converter_numero_br)
            if 'TRS_BRUTO' in df.columns:
                df['TRS_BRUTO'] = df['TRS_BRUTO'].apply(converter_numero_br)
            
            if 'PRODUZIDO' in df.columns and 'APROVADO' in df.columns:
                df['REFUGADO'] = df['PRODUZIDO'] - df['APROVADO']
                df['REFUGADO'] = df['REFUGADO'].apply(lambda x: max(0, x))
            else:
                df['REFUGADO'] = 0
            
            df['ANO_MES'] = df['DATA'].dt.to_period('M').astype(str)
            
            return df
            
        except Exception as e:
            st.error(f"Erro ao carregar dados: {e}")
            return pd.DataFrame()

    with st.spinner("Carregando dados do setor SOPRO..."):
        df_base = carregar_dados_sopro()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados. Verifique a conexão.")
        st.stop()

    st.info(f"📊 Dados carregados: {len(df_base)} registros (apenas SOPRO)")

    # Calcular TRS LÍQUIDO
    df_base_calc = df_base.copy()
    if 'TRS_BRUTO' in df_base_calc.columns:
        df_base_calc['TRS LÍQUIDO (%)'] = df_base_calc['TRS_BRUTO'] * 100
        df_base_calc['TRS LÍQUIDO (%)'] = df_base_calc['TRS LÍQUIDO (%)'].round(2)
    else:
        df_base_calc['TRS LÍQUIDO (%)'] = 0

    melhores_trs_historico = {}
    if 'REFERÊNCIA' in df_base_calc.columns:
        for ref in df_base_calc['REFERÊNCIA'].unique():
            ref_df = df_base_calc[df_base_calc['REFERÊNCIA'] == ref]
            if not ref_df.empty:
                max_trs = ref_df['TRS LÍQUIDO (%)'].max()
                if max_trs > 0:
                    melhores_trs_historico[ref] = max_trs

    # Sidebar filtros
    st.sidebar.title("Filtros - SOPRO")
    filtro_melhores_trs = st.sidebar.checkbox("Apenas Melhores TRS Líquido por Referência", value=False, key="sopro_melhores")
    data_ini = st.sidebar.date_input("Data inicial", value=None, key="sopro_data_ini")
    data_fim = st.sidebar.date_input("Data final", value=None, key="sopro_data_fim")
    
    if 'TURNO' in df_base.columns:
        turnos_disponiveis = ["(Todos)"] + sorted(df_base['TURNO'].dropna().unique().tolist())
        turno = st.sidebar.selectbox("Turno", options=turnos_disponiveis, key="sopro_turno")
    else:
        turno = "(Todos)"
    
    referencia = st.sidebar.text_input("Referência (parte do código)", key="sopro_ref")
    
    if 'PRAÇA' in df_base.columns:
        pracas_disponiveis = ["(Todas)"] + sorted(df_base['PRAÇA'].dropna().unique().tolist())
        praca = st.sidebar.selectbox("Praça", options=pracas_disponiveis, key="sopro_praca")
    else:
        praca = "(Todas)"
    
    mostrar_defeitos = st.sidebar.checkbox("Exibir Somatório de Defeitos", value=True, key="sopro_defeitos")
    qtd = st.sidebar.number_input("Qtd de linhas na tabela (0 = mostrar todas)", min_value=0, max_value=5000, value=20, step=10, key="sopro_qtd")

    # Aplicar filtros
    df = df_base.copy()
    if data_ini:
        df = df[df['DATA'] >= pd.to_datetime(data_ini)]
    if data_fim:
        df = df[df['DATA'] <= pd.to_datetime(data_fim)]
    if turno != "(Todos)" and 'TURNO' in df.columns:
        df = df[df['TURNO'].fillna('').str.upper() == turno.upper()]
    if referencia and 'REFERÊNCIA' in df.columns:
        df = df[df['REFERÊNCIA'].fillna('').str.lower().str.contains(referencia.lower())]
    if praca != "(Todas)" and 'PRAÇA' in df.columns:
        df = df[df['PRAÇA'].fillna('').str.upper() == praca.upper()]

    # ====================== KPIs PRINCIPAIS ======================
    if not df.empty:
        colunas_para_converter = ['PRODUZIDO', 'REFUGADO', 'APROVADO', 'TRS_BRUTO']
        for col in colunas_para_converter:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        total_prod = int(df['PRODUZIDO'].sum()) if 'PRODUZIDO' in df.columns else 0
        total_refugo = int(df['REFUGADO'].sum()) if 'REFUGADO' in df.columns else 0
        total_apro = int(df['APROVADO'].sum()) if 'APROVADO' in df.columns else 0
        
        if 'TRS_BRUTO' in df.columns:
            trs_liquido_medio = df['TRS_BRUTO'].mean() * 100
        else:
            trs_liquido_medio = 0
    else:
        total_prod = total_refugo = total_apro = trs_liquido_medio = 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Produzido", f"{total_prod:,}".replace(",", "."))
    col2.metric("Aprovado", f"{total_apro:,}".replace(",", "."))
    col3.metric("Refugo", f"{total_refugo:,}".replace(",", "."))
    col4.metric("TRS Líquido Médio (%)", f"{trs_liquido_medio:.2f}%")

    st.markdown("---")

    # ====================== TABELA PRINCIPAL ======================
    st.subheader("📋 Tabela de Produção - SOPRO")

    if not df.empty:
        if 'TRS_BRUTO' in df.columns:
            df['TRS LÍQUIDO (%)'] = df['TRS_BRUTO'] * 100
            df['TRS LÍQUIDO (%)'] = df['TRS LÍQUIDO (%)'].round(2)

    df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

    if filtro_melhores_trs and not df_sorted.empty and 'REFERÊNCIA' in df_sorted.columns:
        if 'TRS LÍQUIDO (%)' in df_sorted.columns:
            df_sorted = df_sorted[df_sorted.apply(
                lambda row: row['REFERÊNCIA'] in melhores_trs_historico and 
                            abs(row['TRS LÍQUIDO (%)'] - melhores_trs_historico[row['REFERÊNCIA']]) < 0.01,
                axis=1
            )].reset_index(drop=True)
            
            if not df_sorted.empty:
                st.info(f"Mostrando {len(df_sorted)} registro(s) com Melhor TRS Líquido Histórico por referência")
            else:
                st.warning("Nenhum registro encontrado com Melhor TRS Líquido Histórico")

    df_view = df_sorted if qtd == 0 else df_sorted.head(qtd)

    if not df_view.empty:
        df_display = df_view.copy()
        df_display['DATA'] = pd.to_datetime(df_display['DATA']).dt.strftime('%d/%m/%Y')
        
        colunas_inteiras = ['PRODUZIDO', 'APROVADO', 'REFUGADO']
        for col in colunas_inteiras:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: int(round(x, 0)) if pd.notnull(x) else 0)
                df_display[col] = df_display[col].apply(lambda x: f"{x:,}".replace(",", "."))
        
        if 'TRS LÍQUIDO (%)' in df_display.columns:
            df_display['TRS LÍQUIDO (%)'] = df_display['TRS LÍQUIDO (%)'].apply(lambda x: f"{x:.2f}%")
        
        colunas_exibir = ['DATA', 'TURNO', 'PRAÇA', 'REFERÊNCIA', 'PRODUZIDO', 'REFUGADO', 'APROVADO', 'TRS LÍQUIDO (%)']
        colunas_exibir = [col for col in colunas_exibir if col in df_display.columns]
        
        def destacar_melhor_trs_sopro(row):
            ref = row['REFERÊNCIA']
            if ref in melhores_trs_historico:
                trs_val = row['TRS LÍQUIDO (%)']
                if isinstance(trs_val, str):
                    trs_val = float(trs_val.replace('%', '').replace(',', '.'))
                if abs(trs_val - melhores_trs_historico[ref]) < 0.01:
                    return ['background-color: #FFA500; color: black; font-weight: bold'] * len(row)
            return [''] * len(row)
        
        styled_df = df_display[colunas_exibir].style.apply(destacar_melhor_trs_sopro, axis=1)
        st.dataframe(styled_df, use_container_width=True, height=400)
        
        if not filtro_melhores_trs:
            st.caption("🎯 Linhas em laranja: Melhor TRS Líquido Histórico para cada referência")
    else:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")

    st.markdown("---")

    # ====================== GRÁFICO TRS DIÁRIO ======================
    st.subheader("📈 TRS Líquido - Evolução Diária")
    
    if not df.empty and 'TRS_BRUTO' in df.columns:
        resumo_dia = df.groupby(df['DATA'].dt.date).agg({
            'TRS_BRUTO': 'mean',
            'PRODUZIDO': 'sum',
            'APROVADO': 'sum'
        }).reset_index()
        
        resumo_dia['DATA'] = pd.to_datetime(resumo_dia['DATA'])
        resumo_dia['TRS Líquido (%)'] = resumo_dia['TRS_BRUTO'] * 100
        resumo_dia = resumo_dia.sort_values('DATA')
        
        if not resumo_dia.empty:
            fig, ax = plt.subplots(figsize=(14, 6), facecolor='black')
            ax.set_facecolor('black')
            
            ax.plot(resumo_dia['DATA'], resumo_dia['TRS Líquido (%)'], 
                    marker='o', markersize=8, linewidth=3, 
                    color='lime', alpha=0.8, label='TRS Líquido Diário')
            
            if len(resumo_dia) > 1:
                media_movel = resumo_dia['TRS Líquido (%)'].rolling(window=min(3, len(resumo_dia)), min_periods=1).mean()
                ax.plot(resumo_dia['DATA'], media_movel, 
                        color='yellow', alpha=0.7, linewidth=2, 
                        linestyle='--', label='Média 3 Dias')
            
            ax.axhline(y=85, color='red', linestyle=':', alpha=0.7, linewidth=2, label='Meta (85%)')
            ax.fill_between(resumo_dia['DATA'], 0, resumo_dia['TRS Líquido (%)'], alpha=0.2, color='lime')
            
            ax.set_ylabel("TRS Líquido (%)", fontsize=12, color='white')
            ax.set_xlabel("Data", fontsize=12, color='white')
            ax.set_title("Evolução do TRS Líquido - Período Selecionado", fontsize=16, fontweight='bold', color='white')
            ax.tick_params(colors='white')
            ax.grid(True, alpha=0.3, color='#444444')
            ax.legend(facecolor='black', edgecolor='white', labelcolor='white')
            
            for spine in ax.spines.values():
                spine.set_edgecolor('white')
            
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45)
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    st.markdown("---")

    # ====================== TRS POR PRAÇA ======================
    if 'PRAÇA' in df.columns and not df.empty:
        st.subheader("🏭 TRS Líquido por Praça")
        
        resumo_praca = df.groupby('PRAÇA').agg({
            'TRS_BRUTO': 'mean',
            'PRODUZIDO': 'sum'
        }).reset_index()
        
        resumo_praca['TRS Líquido (%)'] = resumo_praca['TRS_BRUTO'] * 100
        resumo_praca = resumo_praca.sort_values('TRS Líquido (%)', ascending=False)
        
        if not resumo_praca.empty:
            fig, ax = plt.subplots(figsize=(10, 6), facecolor='black')
            ax.set_facecolor('black')
            
            bars = ax.bar(range(len(resumo_praca)), resumo_praca['TRS Líquido (%)'], 
                         color='lime', alpha=0.8, edgecolor='white')
            
            for bar, valor in zip(bars, resumo_praca['TRS Líquido (%)']):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                       f'{valor:.1f}%', ha='center', va='bottom', 
                       fontsize=10, color='white', fontweight='bold')
            
            ax.set_ylabel("TRS Líquido (%)", color='white')
            ax.set_xlabel("Praça", color='white')
            ax.set_title("TRS Líquido Médio por Praça", fontsize=14, fontweight='bold', color='white')
            ax.set_xticks(range(len(resumo_praca)))
            ax.set_xticklabels(resumo_praca['PRAÇA'], rotation=45, ha='right')
            ax.tick_params(colors='white')
            ax.grid(True, alpha=0.3, color='#444444')
            ax.axhline(y=85, color='red', linestyle='--', alpha=0.5)
            
            for spine in ax.spines.values():
                spine.set_edgecolor('white')
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    st.markdown("---")

    # ====================== TRS MENSAL ======================
    st.subheader("📊 TRS Líquido Mensal")
    
    if not df.empty and 'ANO_MES' in df.columns:
        resumo_mensal = df.groupby('ANO_MES').agg({
            'TRS_BRUTO': 'mean',
            'PRODUZIDO': 'sum',
            'APROVADO': 'sum'
        }).reset_index()
        
        resumo_mensal['TRS Líquido (%)'] = resumo_mensal['TRS_BRUTO'] * 100
        resumo_mensal = resumo_mensal.sort_values('ANO_MES')
        
        if not resumo_mensal.empty:
            fig, ax = plt.subplots(figsize=(12, 6), facecolor='black')
            ax.set_facecolor('black')
            
            x_pos = range(len(resumo_mensal))
            bars = ax.bar(x_pos, resumo_mensal['TRS Líquido (%)'], 
                         color='cyan', alpha=0.8, edgecolor='white')
            
            for bar, valor in zip(bars, resumo_mensal['TRS Líquido (%)']):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                       f'{valor:.1f}%', ha='center', va='bottom', 
                       fontsize=10, color='white', fontweight='bold')
            
            ax.set_ylabel("TRS Líquido (%)", color='white')
            ax.set_xlabel("Mês", color='white')
            ax.set_title("TRS Líquido Mensal", fontsize=14, fontweight='bold', color='white')
            ax.tick_params(colors='white')
            ax.grid(True, alpha=0.3, color='#444444')
            ax.axhline(y=85, color='red', linestyle='--', alpha=0.5)
            
            for spine in ax.spines.values():
                spine.set_edgecolor('white')
            
            ax.set_xticks(x_pos)
            meses_formatados = []
            for m in resumo_mensal['ANO_MES']:
                m_str = str(m)
                if len(m_str) > 6:
                    meses_formatados.append(m_str[5:7] + '/' + m_str[:4])
                else:
                    meses_formatados.append(m_str)
            ax.set_xticklabels(meses_formatados, rotation=45)
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    st.markdown("---")

    # ====================== GRÁFICO TRS POR TURNO ======================
    if not df.empty and 'TURNO' in df.columns and 'TRS_BRUTO' in df.columns:
        st.subheader("📊 TRS Líquido por Turno")
        turno_data = []
        for t in df['TURNO'].unique():
            df_turno = df[df['TURNO'] == t]
            trs_medio = df_turno['TRS_BRUTO'].mean() * 100
            turno_data.append({'Turno': t, 'TRS Líquido': trs_medio})
        
        df_turno_trs = pd.DataFrame(turno_data)
        if not df_turno_trs.empty:
            fig, ax = plt.subplots(figsize=(10, 5))
            cores = {'M': '#1f77b4', 'T': '#ff7f0e', 'N': '#2ca02c', 'A': '#1f77b4', 'B': '#ff7f0e', 'C': '#2ca02c'}
            bar_colors = [cores.get(t, '#808080') for t in df_turno_trs['Turno']]
            df_turno_trs.plot(x='Turno', y='TRS Líquido', kind='bar', ax=ax, color=bar_colors, edgecolor="black")
            ax.set_ylabel("TRS Líquido (%)")
            ax.set_ylim(0, 110)
            ax.set_title("TRS Líquido por Turno")
            for i, v in enumerate(df_turno_trs['TRS Líquido']):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom', fontweight='bold')
            st.pyplot(fig)

    # ====================== DEFEITOS ======================
    if mostrar_defeitos:
        st.subheader("❌ Estratificação de Defeitos - SOPRO")
        colunas_defeitos = [
            'BOLHA', 'PEDRA', 'CALCINADO', 'BALANÇANDO', 'AMASSADO', 'OVAL', 
            'CORTE', 'QUEBRADA', 'VIDRO GRUDADO', 'CORDA', 'FORMA', 'RISCO', 
            'TORTO', 'RUGA', 'GABARITO', 'SUJEIRA', 'EMPENO', 'MARCAS', 
            'FALHADA', 'DOBRA', 'CHUPADO', 'ARREADO', 'GOSMA', 'BARRO', 
            'CROMO', 'MACHO'
        ]
        defeitos_existentes = []
        for defeito in colunas_defeitos:
            for col in df.columns:
                if col.upper() == defeito.upper():
                    defeitos_existentes.append(col)
                    break
        
        if defeitos_existentes:
            df_defeitos = df[defeitos_existentes].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_defeitos_sum = df_defeitos.sum().sort_values(ascending=False)
            df_defeitos_sum = df_defeitos_sum[df_defeitos_sum > 0]
            if not df_defeitos_sum.empty:
                st.bar_chart(df_defeitos_sum)
                total_defeitos = df_defeitos_sum.sum()
                st.metric("Total de Defeitos", f"{int(total_defeitos):,}".replace(",", "."))

    st.markdown("---")
    st.caption(f"Dashboard SOPRO - Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
