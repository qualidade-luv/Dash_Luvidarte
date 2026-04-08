import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import os

# ======================
# CONFIGURAÇÕES
# ======================
CAMINHO_CREDENCIAIS = r'\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\4-TRS\dashboard-gerencial-492613-042470f98e27.json'
ID_PLANILHA = '1Hjy4UGtgwIPJgqmcv46LyXNWOrYk_oeJWWV5vlfKF2k'

# Abas disponíveis
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
# FUNÇÃO PARA CARREGAR CREDENCIAIS (SECRETS PRIMEIRO, DEPOIS LOCAL)
# ======================
def get_gspread_client():
    """Cria cliente gspread tentando primeiro secrets, depois arquivo local"""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # Tentativa 1: Usar secrets do Streamlit Cloud
    try:
        if 'gcp_service_account' in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            return client
    except Exception as e:
        st.warning(f"Não foi possível usar secrets do Streamlit Cloud. Tentando arquivo local...")
    
    # Tentativa 2: Usar arquivo de credenciais local
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
# SE FOR PRENSADOS
# ======================
if aba_selecionada == 'PRENSADOS':
    ABA = 'TRS_INDUSTRIAL'
    
    # ======================
    # FUNÇÃO PARA CONVERTER NÚMERO
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

    # ======================
    # FUNÇÃO PARA CONVERTER DATA
    # ======================
    def converter_data_br(data_str, data_referencia=None):
        if data_str is None or pd.isna(data_str):
            return None
        
        try:
            if isinstance(data_str, (datetime, pd.Timestamp)):
                if data_referencia and data_str > data_referencia:
                    return None
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

    # ======================
    # FUNÇÃO PARA CARREGAR DADOS
    # ======================
    @st.cache_data(ttl=300)
    def carregar_dados():
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
    # CALCULAR MELHORES TRS
    # ======================
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

    # ======================
    # SIDEBAR (FILTROS)
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
    # KPIs
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
                    
                    df_defeitos_display = df_defeitos_sum.reset_index().rename(
                        columns={'index':'Defeito', 0:'Quantidade'}
                    ).sort_values('Quantidade', ascending=False)
                    
                    df_defeitos_display['Quantidade'] = df_defeitos_display['Quantidade'].apply(
                        lambda x: f"{int(x):,}".replace(",", ".")
                    )
                    
                    st.dataframe(df_defeitos_display, use_container_width=True)
                    
                    total_defeitos = df_defeitos_sum.sum()
                    st.metric("Total de Defeitos", f"{int(total_defeitos):,}".replace(",", "."))
                else:
                    st.info("Nenhum defeito encontrado nos filtros aplicados.")
            else:
                st.warning("Nenhuma das colunas de defeito foi encontrada na planilha.")
                st.info(f"Colunas disponíveis na planilha: {', '.join(df.columns[:20])}...")
                
        except Exception as e:
            st.warning(f"Erro ao processar defeitos: {e}")

# ======================
# SE FOR SOPRO
# ======================
else:
    ABA = 'TRS_SOPRO'
    
    # ======================
    # FUNÇÕES AUXILIARES PARA SOPRO
    # ======================
    def converter_numero_br_sopro(valor):
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

    def converter_data_br_sopro(data_str):
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
            
            if 'DATA' in df.columns:
                df['DATA'] = df['DATA'].apply(converter_data_br_sopro)
                df = df.dropna(subset=['DATA'])
            
            # Converter PRODUZIDO e APROVADO primeiro
            if 'PRODUZIDO' in df.columns:
                df['PRODUZIDO'] = df['PRODUZIDO'].apply(converter_numero_br_sopro)
            if 'APROVADO' in df.columns:
                df['APROVADO'] = df['APROVADO'].apply(converter_numero_br_sopro)
            if 'TRS_BRUTO' in df.columns:
                df['TRS_BRUTO'] = df['TRS_BRUTO'].apply(converter_numero_br_sopro)
            
            # CALCULAR REFUGADO = PRODUZIDO - APROVADO
            if 'PRODUZIDO' in df.columns and 'APROVADO' in df.columns:
                df['REFUGADO'] = df['PRODUZIDO'] - df['APROVADO']
                # Garantir que não tenha valores negativos
                df['REFUGADO'] = df['REFUGADO'].apply(lambda x: max(0, x))
            else:
                df['REFUGADO'] = 0
            
            return df
            
        except Exception as e:
            st.error(f"Erro ao carregar dados: {e}")
            return pd.DataFrame()

    # ======================
    # CARREGAR DADOS SOPRO
    # ======================
    with st.spinner("Carregando dados do setor SOPRO..."):
        df_base = carregar_dados_sopro()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados. Verifique a conexão.")
        st.stop()

    # ======================
    # CALCULAR MELHORES TRS LÍQUIDO
    # ======================
    df_base_calc = df_base.copy()

    colunas_numericas = ['PRODUZIDO', 'REFUGADO', 'APROVADO', 'TRS_BRUTO']
    for col in colunas_numericas:
        if col in df_base_calc.columns:
            df_base_calc[col] = pd.to_numeric(df_base_calc[col], errors='coerce').fillna(0)

    # Calcular TRS LÍQUIDO (multiplicado por 100)
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

    # ======================
    # SIDEBAR (FILTROS) - SOPRO
    # ======================
    st.sidebar.title("Filtros - SOPRO")

    filtro_melhores_trs = st.sidebar.checkbox("Apenas Melhores TRS Líquido por Referência", value=False)

    data_ini = st.sidebar.date_input("Data inicial", value=None)
    data_fim = st.sidebar.date_input("Data final", value=None)
    
    if 'TURNO' in df_base.columns:
        turnos_disponiveis = ["(Todos)"] + sorted(df_base['TURNO'].dropna().unique().tolist())
        turno = st.sidebar.selectbox("Turno", options=turnos_disponiveis)
    else:
        turno = "(Todos)"
    
    referencia = st.sidebar.text_input("Referência (parte do código)")
    
    if 'PRAÇA' in df_base.columns:
        pracas_disponiveis = ["(Todas)"] + sorted(df_base['PRAÇA'].dropna().unique().tolist())
        praca = st.sidebar.selectbox("Praça", options=pracas_disponiveis)
    else:
        praca = "(Todas)"
    
    mostrar_defeitos = st.sidebar.checkbox("Exibir Somatório de Defeitos", value=True)

    qtd = st.sidebar.number_input(
        "Qtd de linhas na tabela (0 = mostrar todas)",
        min_value=0,
        max_value=5000,
        value=20,
        step=10
    )

    # ======================
    # APLICAÇÃO DOS FILTROS - SOPRO
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

    if praca != "(Todas)" and 'PRAÇA' in df.columns:
        df = df[df['PRAÇA'].fillna('').str.upper() == praca.upper()]

    # ======================
    # KPIs (após filtros) - SOPRO
    # ======================
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

    # ======================
    # TABELA PRINCIPAL - SOPRO
    # ======================
    st.subheader("📋 Tabela de Produção - SOPRO")

    if not df.empty:
        if 'TRS_BRUTO' in df.columns:
            df['TRS LÍQUIDO (%)'] = df['TRS_BRUTO'] * 100
            df['TRS LÍQUIDO (%)'] = df['TRS LÍQUIDO (%)'].round(2)

    df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

    if filtro_melhores_trs and not df_sorted.empty and 'REFERÊNCIA' in df_sorted.columns:
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
        
        colunas_exibir = ['DATA', 'TURNO', 'PRAÇA', 'REFERÊNCIA', 'PRODUZIDO', 'REFUGADO', 
                          'APROVADO', 'TRS LÍQUIDO (%)']
        colunas_exibir = [col for col in colunas_exibir if col in df_display.columns]
        
        def highlight_max_trs_liquido(row):
            if 'REFERÊNCIA' in row and row['REFERÊNCIA'] in melhores_trs_historico:
                trs_str = row['TRS LÍQUIDO (%)']
                if isinstance(trs_str, str):
                    trs_valor = float(trs_str.replace('%', '').replace(',', '.'))
                else:
                    trs_valor = float(trs_str)
                
                if abs(trs_valor - melhores_trs_historico[row['REFERÊNCIA']]) < 0.01:
                    return ['background-color: #FFA500; color: black; font-weight: bold'] * len(row)
            return [''] * len(row)
        
        styled_df = df_display[colunas_exibir].style.apply(highlight_max_trs_liquido, axis=1)
        
        st.dataframe(
            styled_df,
            use_container_width=True,
            height=400
        )
        
        if not filtro_melhores_trs:
            st.caption("🎯 Linhas em laranja: Melhor TRS Líquido Histórico para cada referência")
    else:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")

    # ======================
    # GRÁFICO DE TRS POR TURNO - SOPRO
    # ======================
    if not df.empty and 'TURNO' in df.columns and 'TRS_BRUTO' in df.columns:
        st.subheader("📊 TRS Líquido por Turno")

        turno_data = []
        for t in df['TURNO'].unique():
            df_turno = df[df['TURNO'] == t]
            trs_medio = df_turno['TRS_BRUTO'].mean() * 100
            turno_data.append({'Turno': t, 'TRS Líquido': trs_medio})
        
        df_turno_trs = pd.DataFrame(turno_data)
        
        if not df_turno_trs.empty:
            fig, ax = plt.subplots()
            cores = {'M': '#1f77b4', 'T': '#ff7f0e', 'N': '#2ca02c', 'A': '#1f77b4', 'B': '#ff7f0e', 'C': '#2ca02c'}
            bar_colors = [cores.get(t, '#808080') for t in df_turno_trs['Turno']]
            
            df_turno_trs.plot(x='Turno', y='TRS Líquido', kind='bar', ax=ax, 
                             color=bar_colors, edgecolor="black")
            ax.set_ylabel("TRS Líquido (%)")
            ax.set_ylim(0, 110)
            ax.set_title("TRS Líquido por Turno")
            
            for i, v in enumerate(df_turno_trs['TRS Líquido']):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom', fontweight='bold')
            
            st.pyplot(fig)
        else:
            st.info("Nenhum turno com dados válidos para calcular TRS.")

    # ======================
    # DEFEITOS - SOPRO
    # ======================
    if mostrar_defeitos:
        st.subheader("❌ Estratificação de Defeitos - SOPRO")

        try:
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
                    
                    df_defeitos_display = df_defeitos_sum.reset_index().rename(
                        columns={'index':'Defeito', 0:'Quantidade'}
                    ).sort_values('Quantidade', ascending=False)
                    
                    df_defeitos_display['Quantidade'] = df_defeitos_display['Quantidade'].apply(
                        lambda x: f"{int(x):,}".replace(",", ".")
                    )
                    
                    st.dataframe(df_defeitos_display, use_container_width=True)
                    
                    total_defeitos = df_defeitos_sum.sum()
                    st.metric("Total de Defeitos", f"{int(total_defeitos):,}".replace(",", "."))
                else:
                    st.info("Nenhum defeito encontrado nos filtros aplicados.")
            else:
                st.warning("Nenhuma das colunas de defeito foi encontrada na planilha.")
                
        except Exception as e:
            st.warning(f"Erro ao processar defeitos: {e}")
