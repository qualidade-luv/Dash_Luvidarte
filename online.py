import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime, timedelta
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import os
import numpy as np

# ======================
# CONFIGURAÇÕES
# ======================
ID_PLANILHA = '1Hjy4UGtgwIPJgqmcv46LyXNWOrYk_oeJWWV5vlfKF2k'

PRACAS_NAO_SOPRO = ['GIL', 'GILSIMAR', 'ED CARLOS', 'EDI CARLOS', 'ROBÔ 2', 'ROBÔ-2', 'ROBÔ', 'ROBO']

ABAS = {
    'PRENSADOS': 'TRS_INDUSTRIAL',
    'SOPRO': 'TRS_SOPRO'
}

# ======================
# TEMA VISUAL
# ======================
THEME = {
    'bg_primary':     '#0A0C10',
    'bg_card':        '#111520',
    'bg_card2':       '#161B27',
    'accent_cyan':    '#00D4FF',
    'accent_lime':    '#39FF14',
    'accent_orange':  '#FF6B35',
    'accent_yellow':  '#FFD700',
    'accent_red':     '#FF3860',
    'text_primary':   '#E8EAF0',
    'text_muted':     '#6B7280',
    'border':         '#1E2535',
    'border_bright':  '#2A3550',
    'grid':           '#1A2030',
}

st.set_page_config(
    page_title="TRS Dashboard",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================
# FUNÇÃO PARA HORÁRIO BRASÍLIA
# ======================
def get_horario_brasilia():
    """Retorna o horário atual de Brasília (GMT-3)"""
    fuso_brasilia = pytz.timezone('America/Sao_Paulo')
    agora_brasilia = datetime.now(fuso_brasilia)
    return agora_brasilia.strftime('%d/%m/%Y %H:%M')

# ======================
# CSS GLOBAL COM ÊNFASE NOS RADIOBUTTONS BRANCO NEGRITO
# ======================
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&family=Barlow:wght@300;400;500;600&display=swap');

  /* ── Reset & Base ── */
  html, body, [class*="css"] {{
      font-family: 'Barlow', sans-serif;
      background-color: {THEME['bg_primary']} !important;
      color: {THEME['text_primary']} !important;
  }}
  .stApp {{ background-color: {THEME['bg_primary']} !important; }}

  /* ── Sidebar ── */
  [data-testid="stSidebar"] {{
      background: linear-gradient(180deg, #0D1018 0%, #0A0C14 100%) !important;
      border-right: 1px solid {THEME['border_bright']} !important;
  }}
  
  /* ============================================= */
  /* FORÇA RADIOBUTTONS (PRENSADOS/SOPRO) BRANCO NEGRITO */
  /* ============================================= */
  
  /* Camada 1 - Seletor direto do label do radio */
  [data-testid="stSidebar"] .stRadio label {{
      color: #FFFFFF !important;
      font-weight: bold !important;
      font-family: 'Rajdhani', sans-serif !important;
      font-size: 15px !important;
      letter-spacing: 0.08em;
      text-shadow: 0 0 2px rgba(0,0,0,0.5) !important;
  }}
  
  /* Camada 2 - Força todos os radios dentro do radiogroup */
  [data-testid="stSidebar"] div[role="radiogroup"] label {{
      color: #FFFFFF !important;
      font-weight: bold !important;
      font-family: 'Rajdhani', sans-serif !important;
      font-size: 15px !important;
  }}
  
  /* Camada 3 - Força os span internos do radio */
  [data-testid="stSidebar"] .stRadio label span {{
      color: #FFFFFF !important;
      font-weight: bold !important;
  }}
  
  /* Camada 4 - Força qualquer elemento filho do radio */
  [data-testid="stSidebar"] .stRadio div {{
      color: #FFFFFF !important;
  }}
  
  /* Camada 5 - Seletor universal para qualquer texto dentro do radio */
  [data-testid="stSidebar"] .stRadio * {{
      color: #FFFFFF !important;
      font-weight: bold !important;
  }}
  
  /* Camada 6 - Força especificamente os textos PRENSADOS e SOPRO */
  [data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] + div label,
  [data-testid="stSidebar"] .stRadio label p,
  [data-testid="stSidebar"] .stRadio label span {{
      color: #FFFFFF !important;
      font-weight: bold !important;
  }}
  
  /* Camada 7 - Hover também mantém branco negrito */
  [data-testid="stSidebar"] .stRadio label:hover {{
      color: #FFFFFF !important;
      font-weight: bold !important;
      opacity: 0.9;
  }}
  
  /* ── Sidebar Filters - BRANCO NEGRITO (melhor leitura) ── */
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stTextInput label,
  [data-testid="stSidebar"] .stDateInput label,
  [data-testid="stSidebar"] .stNumberInput label,
  [data-testid="stSidebar"] .stCheckbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stSlider label,
  [data-testid="stSidebar"] .stTimeInput label,
  [data-testid="stSidebar"] .stTextArea label,
  [data-testid="stSidebar"] .stColorPicker label,
  [data-testid="stSidebar"] .stFileUploader label,
  [data-testid="stSidebar"] .stForm label {{
      color: #FFFFFF !important;
      font-weight: bold !important;
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 11px !important;
      text-transform: uppercase;
      letter-spacing: 0.12em;
  }}
  
  /* Força todos os labels da sidebar */
  [data-testid="stSidebar"] label {{
      color: #FFFFFF !important;
      font-weight: bold !important;
  }}
  
  [data-testid="stSidebar"] h1 {{
      font-family: 'Rajdhani', sans-serif !important;
      color: {THEME['accent_cyan']} !important;
      font-size: 20px !important;
      font-weight: 700 !important;
      letter-spacing: 0.15em;
      text-transform: uppercase;
      border-bottom: 1px solid {THEME['border_bright']};
      padding-bottom: 8px;
      margin-bottom: 12px;
  }}

  /* ── Input widgets ── */
  .stSelectbox div[data-baseweb="select"] > div,
  .stTextInput input,
  .stNumberInput input,
  .stDateInput input {{
      background-color: {THEME['bg_card2']} !important;
      border: 1px solid {THEME['border_bright']} !important;
      color: {THEME['text_primary']} !important;
      border-radius: 4px !important;
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 12px !important;
  }}
  .stSelectbox div[data-baseweb="select"] > div:hover {{
      border-color: {THEME['accent_cyan']} !important;
  }}

  /* ── Checkbox ── */
  .stCheckbox > label > div[data-testid="stCheckbox"] {{
      border-color: {THEME['accent_cyan']} !important;
  }}

  /* ── Metric cards ── */
  [data-testid="stMetric"] {{
      background: linear-gradient(135deg, {THEME['bg_card']} 0%, {THEME['bg_card2']} 100%) !important;
      border: 1px solid {THEME['border_bright']} !important;
      border-radius: 8px !important;
      padding: 16px 20px !important;
      position: relative;
      overflow: hidden;
  }}
  [data-testid="stMetric"]::before {{
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 2px;
      background: linear-gradient(90deg, {THEME['accent_cyan']}, transparent);
  }}
  [data-testid="stMetricLabel"] {{
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 10px !important;
      font-weight: 500 !important;
      letter-spacing: 0.15em !important;
      text-transform: uppercase !important;
      color: {THEME['text_muted']} !important;
  }}
  [data-testid="stMetricValue"] {{
      font-family: 'Rajdhani', sans-serif !important;
      font-size: 32px !important;
      font-weight: 700 !important;
      color: {THEME['accent_cyan']} !important;
      letter-spacing: 0.05em;
  }}

  /* ── Section headers ── */
  h1 {{
      font-family: 'Rajdhani', sans-serif !important;
      font-weight: 700 !important;
      font-size: 26px !important;
      letter-spacing: 0.1em !important;
      text-transform: uppercase !important;
      color: {THEME['text_primary']} !important;
  }}
  h2, h3 {{
      font-family: 'Rajdhani', sans-serif !important;
      font-weight: 600 !important;
      letter-spacing: 0.08em !important;
      text-transform: uppercase !important;
      color: {THEME['text_primary']} !important;
  }}
  .stSubheader, [data-testid="stMarkdownContainer"] h3 {{
      border-left: 3px solid {THEME['accent_cyan']};
      padding-left: 10px;
  }}

  /* ── Dataframe ── */
  .stDataFrame {{
      border: 1px solid {THEME['border_bright']} !important;
      border-radius: 6px !important;
      overflow: hidden;
  }}
  .stDataFrame thead th {{
      background-color: {THEME['bg_card']} !important;
      color: {THEME['accent_cyan']} !important;
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 11px !important;
      text-transform: uppercase !important;
      letter-spacing: 0.1em !important;
      border-bottom: 1px solid {THEME['border_bright']} !important;
  }}
  .stDataFrame tbody tr:hover td {{
      background-color: rgba(0, 212, 255, 0.06) !important;
  }}
  .stDataFrame tbody td {{
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 12px !important;
      border-color: {THEME['border']} !important;
  }}

  /* ── Divider ── */
  hr {{
      border: none !important;
      border-top: 1px solid {THEME['border_bright']} !important;
      margin: 24px 0 !important;
  }}

  /* ── Info/Warning/Success banners ── */
  .stInfo {{ background-color: rgba(0,212,255,0.08) !important; border-left: 3px solid {THEME['accent_cyan']} !important; border-radius: 4px !important; }}
  .stWarning {{ background-color: rgba(255,107,53,0.08) !important; border-left: 3px solid {THEME['accent_orange']} !important; border-radius: 4px !important; }}
  .stSuccess {{ background-color: rgba(57,255,20,0.08) !important; border-left: 3px solid {THEME['accent_lime']} !important; border-radius: 4px !important; }}
  .stError {{ background-color: rgba(255,56,96,0.08) !important; border-left: 3px solid {THEME['accent_red']} !important; border-radius: 4px !important; }}

  /* ── Spinner ── */
  .stSpinner > div {{ border-top-color: {THEME['accent_cyan']} !important; }}

  /* ── Caption ── */
  .stCaption {{ color: {THEME['text_muted']} !important; font-family: 'JetBrains Mono', monospace !important; font-size: 10px !important; }}

  /* ── Bar chart ── */
  [data-testid="stArrowVegaLiteChart"] {{ background: {THEME['bg_card']} !important; border-radius: 8px; }}

  /* ── Scrollbar ── */
  ::-webkit-scrollbar {{ width: 4px; height: 4px; }}
  ::-webkit-scrollbar-track {{ background: {THEME['bg_primary']}; }}
  ::-webkit-scrollbar-thumb {{ background: {THEME['border_bright']}; border-radius: 2px; }}
  ::-webkit-scrollbar-thumb:hover {{ background: {THEME['accent_cyan']}; }}
  
  /* ── Remove qualquer mensagem de conexão ── */
  .stAlert, .element-container:has(.stAlert), [data-testid="stNotification"] {{
      display: none !important;
  }}
  
  /* ── Estilo para linha de análise ── */
  .analise-row td {{
      background-color: #0D2A1A !important;
      color: #5EF08A !important;
      font-style: italic !important;
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 11px !important;
  }}
</style>
""", unsafe_allow_html=True)


def render_page_header(title: str, subtitle: str, accent: str = THEME['accent_cyan']):
    """Renderiza cabeçalho da página com estilo industrial."""
    st.markdown(f"""
    <div style="
        padding: 28px 0 20px 0;
        border-bottom: 1px solid {THEME['border_bright']};
        margin-bottom: 28px;
        display: flex;
        align-items: center;
        gap: 16px;
    ">
        <div style="
            width: 4px;
            height: 48px;
            background: linear-gradient(180deg, {accent}, transparent);
            border-radius: 2px;
            flex-shrink: 0;
        "></div>
        <div>
            <div style="
                font-family: 'JetBrains Mono', monospace;
                font-size: 10px;
                letter-spacing: 0.25em;
                color: {accent};
                text-transform: uppercase;
                margin-bottom: 4px;
                opacity: 0.8;
            ">LUVIDARTE / TRS DASHBOARD</div>
            <div style="
                font-family: 'Rajdhani', sans-serif;
                font-size: 36px;
                font-weight: 700;
                color: {THEME['text_primary']};
                letter-spacing: 0.1em;
                text-transform: uppercase;
                line-height: 1;
            ">{title}</div>
            <div style="
                font-family: 'Barlow', sans-serif;
                font-size: 13px;
                color: {THEME['text_muted']};
                margin-top: 4px;
                letter-spacing: 0.05em;
            ">{subtitle}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_section_header(title: str, icon: str = "▸", accent: str = THEME['accent_cyan']):
    st.markdown(f"""
    <div style="
        display: flex;
        align-items: center;
        gap: 10px;
        margin: 28px 0 14px 0;
        padding-bottom: 8px;
        border-bottom: 1px solid {THEME['border']};
    ">
        <span style="color: {accent}; font-size: 16px; font-family: 'Rajdhani', sans-serif;">{icon}</span>
        <span style="
            font-family: 'Rajdhani', sans-serif;
            font-size: 18px;
            font-weight: 600;
            letter-spacing: 0.1em;
            text-transform: uppercase;
            color: {THEME['text_primary']};
        ">{title}</span>
    </div>
    """, unsafe_allow_html=True)


def render_kpi_card(label: str, value: str, accent: str = THEME['accent_cyan'], icon: str = ""):
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, {THEME['bg_card']} 0%, {THEME['bg_card2']} 100%);
        border: 1px solid {THEME['border_bright']};
        border-radius: 8px;
        padding: 18px 22px;
        position: relative;
        overflow: hidden;
        height: 100%;
    ">
        <div style="
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 2px;
            background: linear-gradient(90deg, {accent}, transparent);
        "></div>
        <div style="
            font-family: 'JetBrains Mono', monospace;
            font-size: 9px;
            letter-spacing: 0.2em;
            text-transform: uppercase;
            color: {THEME['text_muted']};
            margin-bottom: 8px;
        ">{icon} {label}</div>
        <div style="
            font-family: 'Rajdhani', sans-serif;
            font-size: 34px;
            font-weight: 700;
            color: {accent};
            letter-spacing: 0.03em;
            line-height: 1;
        ">{value}</div>
    </div>
    """, unsafe_allow_html=True)


def apply_chart_style(ax, fig, title: str, xlabel: str = "", ylabel: str = "",
                      accent: str = THEME['accent_cyan']):
    """Aplica estilo dark/industrial consistente em todos os gráficos matplotlib."""
    fig.patch.set_facecolor(THEME['bg_card'])
    ax.set_facecolor(THEME['bg_card'])

    ax.set_title(title, fontsize=14, fontweight='bold',
                 color=THEME['text_primary'], fontfamily='sans-serif',
                 pad=16, loc='left')
    if xlabel:
        ax.set_xlabel(xlabel, fontsize=10, color=THEME['text_muted'], labelpad=8)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10, color=THEME['text_muted'], labelpad=8)

    ax.tick_params(colors=THEME['text_muted'], labelsize=9)
    ax.grid(True, alpha=0.15, color=THEME['grid'], linewidth=0.8, linestyle='--')
    ax.set_axisbelow(True)

    for spine in ax.spines.values():
        spine.set_edgecolor(THEME['border_bright'])
        spine.set_linewidth(0.8)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)


# ======================
# SIDEBAR - navegação
# ======================
with st.sidebar:
    st.markdown(f"""
    <div style="
        text-align: center;
        padding: 20px 0 16px;
        border-bottom: 1px solid {THEME['border_bright']};
        margin-bottom: 20px;
    ">
        <div style="
            font-family: 'Rajdhani', sans-serif;
            font-size: 24px;
            font-weight: 700;
            color: {THEME['accent_cyan']};
            letter-spacing: 0.2em;
        ">⚙ TRS</div>
        <div style="
            font-family: 'JetBrains Mono', monospace;
            font-size: 9px;
            color: {THEME['text_muted']};
            letter-spacing: 0.2em;
            text-transform: uppercase;
            margin-top: 2px;
        ">Industrial Dashboard</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"<div style='font-family:JetBrains Mono,monospace;font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:{THEME['accent_cyan']};margin-bottom:8px'>▸ Setor</div>", unsafe_allow_html=True)
    aba_selecionada = st.radio("", list(ABAS.keys()), label_visibility="collapsed")

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


def get_gspread_client():
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

    caminhos_credenciais = [
        r'C:\Users\elton\OneDrive\Documentos\dashboard-gerencial-492613-042470f98e27.json',
        r'C:\Users\elton\OneDrive\Desktop\dashboard-gerencial-492613-042470f98e27.json',
        r'C:\Users\elton\Desktop\dashboard-gerencial-492613-042470f98e27.json',
        r'C:\Users\elton\Downloads\dashboard-gerencial-492613-042470f98e27.json',
        r'dashboard-gerencial-492613-042470f98e27.json',
        r'\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\4-TRS\dashboard-gerencial-492613-042470f98e27.json',
    ]

    for caminho in caminhos_credenciais:
        try:
            if os.path.exists(caminho):
                creds = ServiceAccountCredentials.from_json_keyfile_name(caminho, scope)
                client = gspread.authorize(creds)
                return client
        except Exception:
            pass

    return None


# ==================================================================================================
# PRENSADOS
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
                    lambda row: max(0, row['ACERTOS_MIN'] - 165) if row['IS_SABADO'] else row['ACERTOS_MIN'], axis=1
                )
            return df
        except Exception:
            return pd.DataFrame()

    with st.spinner("Carregando dados..."):
        df_base = carregar_dados_prensados()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados.")
        st.stop()

    df_base_calc = df_base.copy()
    colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO']
    for col in colunas_numericas:
        if col in df_base_calc.columns:
            df_base_calc[col] = pd.to_numeric(df_base_calc[col], errors='coerce').fillna(0)

    if 'EMBALADO' in df_base_calc.columns and 'META LIQUIDA' in df_base_calc.columns:
        df_base_calc['TRS FINAL (%)'] = df_base_calc.apply(
            lambda row: (row['EMBALADO'] / row['META LIQUIDA'] * 100) if row['META LIQUIDA'] != 0 else 0, axis=1
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

    # ── Sidebar filtros PRENSADOS ──
    with st.sidebar:
        st.markdown(f"<div style='font-family:JetBrains Mono,monospace;font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:{THEME['accent_cyan']};margin:20px 0 10px;border-top:1px solid {THEME['border_bright']};padding-top:16px'>▸ Filtros · Prensados</div>", unsafe_allow_html=True)
        filtro_melhores_trs = st.checkbox("Melhores TRS por Referência", value=False)
        data_ini = st.date_input("Data inicial", value=None, key="prensados_data_ini")
        data_fim = st.date_input("Data final", value=None, key="prensados_data_fim")
        turno = st.selectbox("Turno", options=["(Todos)", "M", "T", "N"], key="prensados_turno")
        referencia = st.text_input("Referência (parte do código)", key="prensados_ref")
        prensa_tipo = st.selectbox("Tipo de prensa", ["(Todos)", "Semi-Automática", "Automática"], key="prensados_tipo")
        mostrar_defeitos = st.checkbox("Somatório de Defeitos", value=True, key="prensados_defeitos")
        qtd = st.number_input("Linhas na tabela (0 = todas)", min_value=0, max_value=5000, value=0, step=10, key="prensados_qtd")

    # ── Aplicar filtros ──
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

    # ── KPIs ──
    if not df.empty:
        for col in ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'META LIQUIDA', 'REFUGADO']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        total_prod  = int(df['PRODUZIDO'].sum())
        total_apro  = int(df['APROVADO'].sum())
        total_embal = int(df['EMBALADO'].sum()) if 'EMBALADO' in df.columns else 0
        total_meta  = int(df['META LIQUIDA'].sum())
        trs_final_total = (total_embal / total_meta * 100) if total_meta else 0
    else:
        total_prod = total_apro = total_embal = total_meta = trs_final_total = 0

    # ── Page header ──
    render_page_header(
        "PRENSADOS",
        f"Industrial · {len(df):,} registros carregados · Atualizado {get_horario_brasilia()}",
        THEME['accent_cyan']
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1: render_kpi_card("Produzido", f"{total_prod:,}".replace(",","."), THEME['accent_cyan'], "◈")
    with c2: render_kpi_card("Aprovado",  f"{total_apro:,}".replace(",","."), THEME['accent_lime'], "◈")
    with c3: render_kpi_card("Embalado",  f"{total_embal:,}".replace(",","."), THEME['accent_yellow'], "◈")
    with c4:
        trs_color = THEME['accent_lime'] if trs_final_total >= 85 else THEME['accent_orange'] if trs_final_total >= 70 else THEME['accent_red']
        render_kpi_card("TRS Final", f"{trs_final_total:.1f}%", trs_color, "◎")

    # ── Tabela de produção ──
    render_section_header("Tabela de Produção", "▸")

    if not df.empty:
        df['TRS (%)'] = df.apply(
            lambda r: (r['APROVADO'] / r['META LIQUIDA'] * 100) if r['META LIQUIDA'] != 0 else 0, axis=1
        )
        if 'EMBALADO' in df.columns:
            df['TRS FINAL (%)'] = df.apply(
                lambda r: (r['EMBALADO'] / r['META LIQUIDA'] * 100) if r['META LIQUIDA'] != 0 else 0, axis=1
            )
        else:
            df['TRS FINAL (%)'] = 0
        df['TRS (%)'] = df['TRS (%)'].round(2)
        df['TRS FINAL (%)'] = df['TRS FINAL (%)'].round(2)

    df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

    if filtro_melhores_trs and not df_sorted.empty and 'REFERÊNCIA' in df_sorted.columns:
        df_sorted = df_sorted[df_sorted.apply(
            lambda row: row['REFERÊNCIA'] in melhores_trs_historico and
                        abs(row['TRS FINAL (%)'] - melhores_trs_historico[row['REFERÊNCIA']]) < 0.01, axis=1
        )].reset_index(drop=True)
        if not df_sorted.empty:
            st.info(f"Exibindo {len(df_sorted)} registro(s) — Melhor TRS Final Histórico por referência")
        else:
            st.warning("Nenhum registro encontrado com Melhor TRS Final Histórico")

    df_view = df_sorted if qtd == 0 else df_sorted.head(qtd)

    if not df_view.empty:
        df_display = df_view.copy()
        df_display['DATA'] = pd.to_datetime(df_display['DATA']).dt.strftime('%d/%m/%Y')

        for col in ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'REFUGADO', 'META LIQUIDA']:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: int(round(x)) if pd.notnull(x) else 0)
                df_display[col] = df_display[col].apply(lambda x: f"{x:,}".replace(",", "."))
        if 'TRS (%)' in df_display.columns:
            df_display['TRS (%)'] = df_display['TRS (%)'].apply(lambda x: f"{x:.2f}%")
        if 'TRS FINAL (%)' in df_display.columns:
            df_display['TRS FINAL (%)'] = df_display['TRS FINAL (%)'].apply(lambda x: f"{x:.2f}%")

        colunas_exibir = ['DATA', 'REFERÊNCIA', 'TURNO', 'PRODUZIDO', 'APROVADO', 'EMBALADO',
                          'REFUGADO', 'META LIQUIDA', 'TRS (%)', 'TRS FINAL (%)']
        if 'ANALISE' in df_display.columns:
            colunas_exibir.append('ANALISE')
        colunas_exibir = [col for col in colunas_exibir if col in df_display.columns]

        # FUNÇÃO: linha de análise sempre abaixo da linha correspondente
        def render_tabela_com_analise(df_display, colunas_exibir, melhores_trs_historico):
            cols_sem_analise = [c for c in colunas_exibir if c != 'ANALISE']
            
            # Constrói as linhas na ordem correta: linha normal, depois análise (abaixo)
            linhas_finais = []
            for idx, row in df_display.iterrows():
                # 1. Linha normal de dados
                linha_normal = {col: row[col] for col in cols_sem_analise if col in row.index}
                linha_normal['_tipo'] = 'dado'
                linhas_finais.append(linha_normal)
                
                # 2. Linha de análise (se existir) - sempre ABAIXO da linha correspondente
                if 'ANALISE' in df_display.columns:
                    v = row.get('ANALISE', '')
                    if pd.notna(v) and str(v).strip():
                        linha_analise = {}
                        for i, col in enumerate(cols_sem_analise):
                            if i == 0:
                                linha_analise[col] = '📋 ANÁLISE:'
                            elif i == 1:
                                linha_analise[col] = str(v)
                            else:
                                linha_analise[col] = ''
                        linha_analise['_tipo'] = 'analise'
                        linhas_finais.append(linha_analise)
            
            df_final = pd.DataFrame(linhas_finais)
            
            # Função de estilo
            def destacar_linhas(row):
                if row.get('_tipo') == 'analise':
                    return [f'background-color: #0D2A1A; color: #5EF08A; font-style: italic; font-family: JetBrains Mono, monospace; font-size: 11px;'] * len(row)
                ref = row.get('REFERÊNCIA', '')
                if ref in melhores_trs_historico:
                    trs_v = row.get('TRS FINAL (%)', '0')
                    if isinstance(trs_v, str):
                        trs_v = float(trs_v.replace('%', '').replace(',', '.'))
                    if abs(trs_v - melhores_trs_historico[ref]) < 0.01:
                        return ['background-color: #2A1A00; color: #FFD700; font-weight: 600;'] * len(row)
                return [''] * len(row)
            
            styled_df = df_final[cols_sem_analise].style.apply(destacar_linhas, axis=1)
            st.dataframe(styled_df, use_container_width=True, height=500)
        
        render_tabela_com_analise(df_display, colunas_exibir, melhores_trs_historico)

        if not filtro_melhores_trs:
            st.caption("▸ Dourado: Melhor TRS Final Histórico por referência   ▸ Verde: Análise registrada")

    # ── Gráfico TRS Diário ──
    render_section_header("Evolução Diária do TRS", "▸")

    if not df.empty:
        resumo_dia = df.groupby(df['DATA'].dt.date).agg(
            PRODUZIDO=('PRODUZIDO','sum'),
            APROVADO=('APROVADO','sum'),
            META_LIQUIDA=('META LIQUIDA','sum')
        ).reset_index()
        resumo_dia['DATA'] = pd.to_datetime(resumo_dia['DATA'])
        resumo_dia['TRS (%)'] = (resumo_dia['APROVADO'] / resumo_dia['META_LIQUIDA'].replace(0,1) * 100).fillna(0)
        resumo_dia = resumo_dia.sort_values('DATA')

        if not resumo_dia.empty:
            fig, ax = plt.subplots(figsize=(14, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS Diário — Período Selecionado", ylabel="TRS (%)")

            ax.fill_between(resumo_dia['DATA'], 0, resumo_dia['TRS (%)'],
                            alpha=0.12, color=THEME['accent_cyan'])

            ax.plot(resumo_dia['DATA'], resumo_dia['TRS (%)'],
                    marker='o', markersize=6, linewidth=2.5,
                    color=THEME['accent_cyan'], alpha=0.95, label='TRS Diário',
                    markerfacecolor=THEME['bg_card'], markeredgecolor=THEME['accent_cyan'], markeredgewidth=2)

            if len(resumo_dia) > 1:
                mm = resumo_dia['TRS (%)'].rolling(window=min(3,len(resumo_dia)), min_periods=1).mean()
                ax.plot(resumo_dia['DATA'], mm, color=THEME['accent_yellow'],
                        alpha=0.8, linewidth=1.8, linestyle='--', label='Média 3 dias')

            ax.axhline(y=85, color=THEME['accent_red'], linestyle=':', alpha=0.7, linewidth=1.5, label='Meta 85%')

            ax.legend(framealpha=0.15, facecolor=THEME['bg_card'], edgecolor=THEME['border_bright'],
                      labelcolor=THEME['text_primary'], fontsize=9)
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha='right', fontsize=8,
                     color=THEME['text_muted'])
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Manual vs Automática ──
    render_section_header("Desempenho por Tipo de Prensa", "▸")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"""<div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.2em;
            text-transform:uppercase;color:{THEME['text_muted']};margin-bottom:8px">▸ Semi-Automática (Manual)</div>""",
            unsafe_allow_html=True)
        if 'BOQUETA' in df.columns:
            df_manual = df[df['BOQUETA'] == 1]
            if not df_manual.empty:
                t_ap_m = df_manual['APROVADO'].sum()
                t_mt_m = df_manual['META LIQUIDA'].replace(0,1).sum()
                med_m  = (t_ap_m / t_mt_m * 100) if t_mt_m > 0 else 0
                prod_m = df_manual['PRODUZIDO'].sum()
                trs_color_m = THEME['accent_lime'] if med_m >= 85 else THEME['accent_orange'] if med_m >= 70 else THEME['accent_red']
                render_kpi_card("TRS Médio — Manual", f"{med_m:.1f}%", trs_color_m)
                st.caption(f"Produção: {prod_m:,.0f} un".replace(",","."))
            else:
                st.info("Sem dados para Prensa Manual")

    with col2:
        st.markdown(f"""<div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.2em;
            text-transform:uppercase;color:{THEME['text_muted']};margin-bottom:8px">▸ Automática</div>""",
            unsafe_allow_html=True)
        if 'BOQUETA' in df.columns:
            df_auto = df[df['BOQUETA'] == 2]
            if not df_auto.empty:
                t_ap_a = df_auto['APROVADO'].sum()
                t_mt_a = df_auto['META LIQUIDA'].replace(0,1).sum()
                med_a  = (t_ap_a / t_mt_a * 100) if t_mt_a > 0 else 0
                prod_a = df_auto['PRODUZIDO'].sum()
                trs_color_a = THEME['accent_lime'] if med_a >= 85 else THEME['accent_orange'] if med_a >= 70 else THEME['accent_red']
                render_kpi_card("TRS Médio — Automática", f"{med_a:.1f}%", trs_color_a)
                st.caption(f"Produção: {prod_a:,.0f} un".replace(",","."))
            else:
                st.info("Sem dados para Prensa Automática")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Análise de Paradas ──
    render_section_header("Análise de Paradas", "▸")

    horas_trabalhadas = 0
    total_acertos = 0
    total_manut = 0

    if 'HORAS_TOTAIS_MIN' in df.columns:
        horas_trabalhadas = df['HORAS_TOTAIS_MIN'].sum()
    else:
        dias_uteis  = df[~df['IS_SABADO']]['DATA'].nunique() if 'IS_SABADO' in df.columns else 0
        dias_sabado = df[df['IS_SABADO']]['DATA'].nunique()  if 'IS_SABADO' in df.columns else 0
        horas_trabalhadas = (dias_uteis * 8 * 60) + (dias_sabado * 6 * 60)

    total_acertos = df['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
    total_manut   = df['MANUT_MIN'].sum()            if 'MANUT_MIN'            in df.columns else 0
    total_paradas = total_acertos + total_manut
    horas_produtivas = max(0, horas_trabalhadas - total_paradas)

    p1, p2, p3, p4 = st.columns(4)
    with p1: render_kpi_card("Horas Trabalhadas",  minutos_para_horas_str(horas_trabalhadas),  THEME['accent_cyan'])
    with p2: render_kpi_card("Acertos",             minutos_para_horas_str(total_acertos),       THEME['accent_yellow'])
    with p3: render_kpi_card("Manutenção",          minutos_para_horas_str(total_manut),          THEME['accent_red'])
    with p4: render_kpi_card("Horas Produtivas",   minutos_para_horas_str(horas_produtivas),    THEME['accent_lime'])

    col1, col2 = st.columns(2)

    with col1:
        if 'BOQUETA' in df.columns:
            df_manual_p = df[df['BOQUETA'] == 1]
            df_auto_p   = df[df['BOQUETA'] == 2]
            acertos_m = df_manual_p['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
            manut_m   = df_manual_p['MANUT_MIN'].sum()            if 'MANUT_MIN'            in df.columns else 0
            acertos_a = df_auto_p['ACERTOS_MIN_AJUSTADO'].sum()   if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
            manut_a   = df_auto_p['MANUT_MIN'].sum()              if 'MANUT_MIN'            in df.columns else 0

            categorias = ['Manual', 'Automática']
            acertos_v  = [acertos_m, acertos_a]
            manut_v    = [manut_m, manut_a]

            fig, ax = plt.subplots(figsize=(7, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "Composição das Paradas", ylabel="Minutos")

            x = np.arange(len(categorias))
            w = 0.55
            ax.bar(x, acertos_v, w, label='Acertos',    color=THEME['accent_yellow'], alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5)
            ax.bar(x, manut_v,   w, bottom=acertos_v, label='Manutenção', color=THEME['accent_red'],    alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5)

            for i, (a, m) in enumerate(zip(acertos_v, manut_v)):
                total = a + m
                if a > 0: ax.text(i, a/2, minutos_para_horas_str(a), ha='center', va='center', color='white', fontweight='bold', fontsize=10)
                if m > 0: ax.text(i, a + m/2, minutos_para_horas_str(m), ha='center', va='center', color='white', fontweight='bold', fontsize=10)
                if total > 0:
                    ax.text(i, total * 1.05, minutos_para_horas_str(total),
                            ha='center', va='bottom', color=THEME['text_primary'], fontweight='bold', fontsize=11)

            ax.set_xticks(x)
            ax.set_xticklabels(categorias, fontsize=10, color=THEME['text_muted'])
            ax.legend(framealpha=0.15, facecolor=THEME['bg_card'], edgecolor=THEME['border_bright'],
                      labelcolor=THEME['text_primary'], fontsize=9)
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    with col2:
        if horas_trabalhadas > 0:
            labels_p = ['Produtivas', 'Acertos', 'Manutenção']
            vals_p   = [horas_produtivas, total_acertos, total_manut]
            cores_p  = [THEME['accent_lime'], THEME['accent_yellow'], THEME['accent_red']]
            lf, vf, cf = zip(*[(l, v, c) for l, v, c in zip(labels_p, vals_p, cores_p) if v > 0]) if any(v > 0 for v in vals_p) else ([], [], [])

            if vf:
                fig, ax = plt.subplots(figsize=(7, 5), facecolor=THEME['bg_card'])
                fig.patch.set_facecolor(THEME['bg_card'])
                ax.set_facecolor(THEME['bg_card'])

                wedges, texts, autotexts = ax.pie(
                    vf, labels=lf, colors=cf,
                    autopct='%1.1f%%', startangle=90,
                    textprops={'color': THEME['text_primary'], 'fontsize': 10},
                    wedgeprops={'edgecolor': THEME['bg_card'], 'linewidth': 2}
                )
                for at in autotexts:
                    at.set_color('white')
                    at.set_fontweight('bold')
                    at.set_fontsize(10)

                ax.set_title(
                    f"Distribuição do Tempo\n{minutos_para_horas_str(horas_trabalhadas)} trabalhadas",
                    fontsize=13, fontweight='bold', color=THEME['text_primary'], pad=14
                )
                fig.tight_layout(pad=1.5)
                st.pyplot(fig)
                plt.close(fig)
        else:
            st.info("Sem dados de tempo para exibir")

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── TRS por Turno ──
    if not df.empty and 'TURNO' in df.columns and 'EMBALADO' in df.columns:
        render_section_header("TRS Final por Turno", "▸")
        turno_data = []
        for t in df['TURNO'].unique():
            df_t = df[df['TURNO'] == t]
            te = df_t['EMBALADO'].sum()
            tm = df_t['META LIQUIDA'].sum()
            turno_data.append({'Turno': t, 'TRS Final': (te/tm*100) if tm > 0 else 0})
        df_tt = pd.DataFrame(turno_data)
        if not df_tt.empty:
            fig, ax = plt.subplots(figsize=(8, 4), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS Final por Turno", ylabel="TRS Final (%)")
            cores_turno = {'M': THEME['accent_cyan'], 'T': THEME['accent_orange'], 'N': THEME['accent_lime']}
            bar_colors  = [cores_turno.get(t, THEME['text_muted']) for t in df_tt['Turno']]
            bars = ax.bar(range(len(df_tt)), df_tt['TRS Final'], color=bar_colors,
                          alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5, width=0.55)
            ax.axhline(y=85, color=THEME['accent_red'], linestyle='--', alpha=0.5, linewidth=1.5)
            for i, v in enumerate(df_tt['TRS Final']):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom',
                        fontweight='bold', fontsize=11, color=THEME['text_primary'])
            ax.set_xticks(range(len(df_tt)))
            ax.set_xticklabels(df_tt['Turno'], fontsize=11)
            ax.set_ylim(0, 115)
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    # ── Defeitos ──
    if mostrar_defeitos:
        render_section_header("Estratificação de Defeitos", "▸")
        colunas_defeitos = [
            'BOLHA','PEDRA','TRINCA','RUGAS','CORTE TESOURA','DOBRA','SUJEIRA','QUEBRA',
            'ARREADO','VIDRO GRUDADO','CONTRA-PEÇA','FALHAS','CHUPADO','ÓLEO TESOURA',
            'CROMO','MACHO','BARRO','EMPENO','OUTROS'
        ]
        defeitos_existentes = []
        for defeito in colunas_defeitos:
            for col in df.columns:
                if col.upper() == defeito.upper():
                    defeitos_existentes.append(col)
                    break
        if defeitos_existentes:
            df_def = df[defeitos_existentes].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_def_sum = df_def.sum().sort_values(ascending=False)
            df_def_sum = df_def_sum[df_def_sum > 0]
            if not df_def_sum.empty:
                fig, ax = plt.subplots(figsize=(12, 4), facecolor=THEME['bg_card'])
                apply_chart_style(ax, fig, "Defeitos — Somatório", ylabel="Quantidade")
                bars = ax.bar(range(len(df_def_sum)), df_def_sum.values,
                              color=THEME['accent_red'], alpha=0.8,
                              edgecolor=THEME['bg_card'], linewidth=1.2)
                ax.set_xticks(range(len(df_def_sum)))
                ax.set_xticklabels(df_def_sum.index, rotation=40, ha='right',
                                   fontsize=9, color=THEME['text_muted'])
                for bar, val in zip(bars, df_def_sum.values):
                    if val > 0:
                        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() * 1.02,
                                f"{int(val):,}".replace(",","."), ha='center', va='bottom',
                                fontsize=8, color=THEME['text_primary'])
                fig.tight_layout(pad=1.5)
                st.pyplot(fig)
                plt.close(fig)
                total_def = df_def_sum.sum()
                st.caption(f"Total de defeitos: {int(total_def):,}".replace(",","."))

    st.markdown(f"""
    <div style="text-align:right;padding:16px 0 8px;
        font-family:'JetBrains Mono',monospace;font-size:10px;
        color:{THEME['text_muted']};letter-spacing:.1em;">
        TRS DASHBOARD · PRENSADOS · {get_horario_brasilia()}
    </div>
    """, unsafe_allow_html=True)


# ==================================================================================================
# SOPRO
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
                return pd.DataFrame()
            cabecalho = todos_dados[0]
            valores   = todos_dados[1:]
            df = pd.DataFrame(valores, columns=cabecalho)
            df.columns = df.columns.str.strip().str.upper()
            if 'PRAÇA' in df.columns:
                df['PRAÇA_NORM'] = df['PRAÇA'].fillna('').astype(str).str.upper().str.strip()
                mascara = ~df['PRAÇA_NORM'].apply(
                    lambda x: any(p in x for p in [p.upper() for p in PRACAS_NAO_SOPRO])
                )
                df = df[mascara].copy()
                df = df.drop(columns=['PRAÇA_NORM'])
            if 'DATA' in df.columns:
                df['DATA'] = df['DATA'].apply(converter_data_br)
                df = df.dropna(subset=['DATA'])
            for col in ['PRODUZIDO', 'APROVADO', 'TRS_BRUTO']:
                if col in df.columns:
                    df[col] = df[col].apply(converter_numero_br)
            if 'PRODUZIDO' in df.columns and 'APROVADO' in df.columns:
                df['REFUGADO'] = (df['PRODUZIDO'] - df['APROVADO']).clip(lower=0)
            else:
                df['REFUGADO'] = 0
            df['ANO_MES'] = df['DATA'].dt.to_period('M').astype(str)
            return df
        except Exception:
            return pd.DataFrame()

    with st.spinner("Carregando dados..."):
        df_base = carregar_dados_sopro()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados.")
        st.stop()

    df_base_calc = df_base.copy()
    if 'TRS_BRUTO' in df_base_calc.columns:
        df_base_calc['TRS LÍQUIDO (%)'] = (df_base_calc['TRS_BRUTO'] * 100).round(2)
    else:
        df_base_calc['TRS LÍQUIDO (%)'] = 0

    melhores_trs_historico = {}
    if 'REFERÊNCIA' in df_base_calc.columns:
        for ref in df_base_calc['REFERÊNCIA'].unique():
            ref_df = df_base_calc[df_base_calc['REFERÊNCIA'] == ref]
            if not ref_df.empty:
                mt = ref_df['TRS LÍQUIDO (%)'].max()
                if mt > 0:
                    melhores_trs_historico[ref] = mt

    # ── Sidebar filtros SOPRO ──
    with st.sidebar:
        st.markdown(f"<div style='font-family:JetBrains Mono,monospace;font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:{THEME['accent_lime']};margin:20px 0 10px;border-top:1px solid {THEME['border_bright']};padding-top:16px'>▸ Filtros · Sopro</div>", unsafe_allow_html=True)
        filtro_melhores_trs = st.checkbox("Melhores TRS Líquido por Referência", value=False, key="sopro_melhores")
        data_ini = st.date_input("Data inicial", value=None, key="sopro_data_ini")
        data_fim = st.date_input("Data final",   value=None, key="sopro_data_fim")
        if 'TURNO' in df_base.columns:
            turnos_disp = ["(Todos)"] + sorted(df_base['TURNO'].dropna().unique().tolist())
            turno = st.selectbox("Turno", options=turnos_disp, key="sopro_turno")
        else:
            turno = "(Todos)"
        referencia = st.text_input("Referência (parte do código)", key="sopro_ref")
        if 'PRAÇA' in df_base.columns:
            pracas_disp = ["(Todas)"] + sorted(df_base['PRAÇA'].dropna().unique().tolist())
            praca = st.selectbox("Praça", options=pracas_disp, key="sopro_praca")
        else:
            praca = "(Todas)"
        mostrar_defeitos = st.checkbox("Somatório de Defeitos", value=True, key="sopro_defeitos")
        qtd = st.number_input("Linhas na tabela (0 = todas)", min_value=0, max_value=5000, value=0, step=10, key="sopro_qtd")

    # ── Filtros ──
    df = df_base.copy()
    if data_ini: df = df[df['DATA'] >= pd.to_datetime(data_ini)]
    if data_fim: df = df[df['DATA'] <= pd.to_datetime(data_fim)]
    if turno != "(Todos)" and 'TURNO' in df.columns:
        df = df[df['TURNO'].fillna('').str.upper() == turno.upper()]
    if referencia and 'REFERÊNCIA' in df.columns:
        df = df[df['REFERÊNCIA'].fillna('').str.lower().str.contains(referencia.lower())]
    if praca != "(Todas)" and 'PRAÇA' in df.columns:
        df = df[df['PRAÇA'].fillna('').str.upper() == praca.upper()]

    for col in ['PRODUZIDO', 'REFUGADO', 'APROVADO', 'TRS_BRUTO']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    total_prod   = int(df['PRODUZIDO'].sum()) if 'PRODUZIDO' in df.columns else 0
    total_refugo = int(df['REFUGADO'].sum())  if 'REFUGADO'  in df.columns else 0
    total_apro   = int(df['APROVADO'].sum())  if 'APROVADO'  in df.columns else 0
    trs_liq_med  = df['TRS_BRUTO'].mean() * 100 if 'TRS_BRUTO' in df.columns and not df.empty else 0

    # ── Page header ──
    render_page_header(
        "SOPRO",
        f"Industrial · {len(df):,} registros carregados · Atualizado {get_horario_brasilia()}",
        THEME['accent_lime']
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1: render_kpi_card("Produzido",         f"{total_prod:,}".replace(",","."), THEME['accent_cyan'],   "◈")
    with c2: render_kpi_card("Aprovado",           f"{total_apro:,}".replace(",","."), THEME['accent_lime'],   "◈")
    with c3: render_kpi_card("Refugo",             f"{total_refugo:,}".replace(",","."), THEME['accent_orange'], "◈")
    with c4:
        trs_c = THEME['accent_lime'] if trs_liq_med >= 85 else THEME['accent_orange'] if trs_liq_med >= 70 else THEME['accent_red']
        render_kpi_card("TRS Líquido Médio", f"{trs_liq_med:.1f}%", trs_c, "◎")

    # ── Tabela ──
    render_section_header("Tabela de Produção", "▸", THEME['accent_lime'])

    if not df.empty and 'TRS_BRUTO' in df.columns:
        df['TRS LÍQUIDO (%)'] = (df['TRS_BRUTO'] * 100).round(2)

    df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

    if filtro_melhores_trs and not df_sorted.empty and 'REFERÊNCIA' in df_sorted.columns and 'TRS LÍQUIDO (%)' in df_sorted.columns:
        df_sorted = df_sorted[df_sorted.apply(
            lambda r: r['REFERÊNCIA'] in melhores_trs_historico and
                      abs(r['TRS LÍQUIDO (%)'] - melhores_trs_historico[r['REFERÊNCIA']]) < 0.01, axis=1
        )].reset_index(drop=True)
        if not df_sorted.empty:
            st.info(f"Exibindo {len(df_sorted)} registro(s) — Melhor TRS Líquido Histórico por referência")
        else:
            st.warning("Nenhum registro encontrado")

    df_view = df_sorted if qtd == 0 else df_sorted.head(qtd)

    if not df_view.empty:
        df_display = df_view.copy()
        df_display['DATA'] = pd.to_datetime(df_display['DATA']).dt.strftime('%d/%m/%Y')
        for col in ['PRODUZIDO', 'APROVADO', 'REFUGADO']:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: int(round(x)) if pd.notnull(x) else 0)
                df_display[col] = df_display[col].apply(lambda x: f"{x:,}".replace(",","."))
        if 'TRS LÍQUIDO (%)' in df_display.columns:
            df_display['TRS LÍQUIDO (%)'] = df_display['TRS LÍQUIDO (%)'].apply(lambda x: f"{x:.2f}%")

        colunas_exibir = ['DATA','TURNO','PRAÇA','REFERÊNCIA','PRODUZIDO','REFUGADO','APROVADO','TRS LÍQUIDO (%)']
        colunas_exibir = [c for c in colunas_exibir if c in df_display.columns]

        def destacar_sopro(row):
            ref = row.get('REFERÊNCIA', '')
            if ref in melhores_trs_historico:
                tv = row.get('TRS LÍQUIDO (%)', '0')
                if isinstance(tv, str): tv = float(tv.replace('%','').replace(',','.'))
                if abs(tv - melhores_trs_historico[ref]) < 0.01:
                    return ['background-color: #2A1A00; color: #FFD700; font-weight: 600;'] * len(row)
            return [''] * len(row)

        styled = df_display[colunas_exibir].style.apply(destacar_sopro, axis=1)
        st.dataframe(styled, use_container_width=True, height=400)
        if not filtro_melhores_trs:
            st.caption("▸ Dourado: Melhor TRS Líquido Histórico por referência")
    else:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")

    # ── TRS Líquido Diário ──
    render_section_header("Evolução Diária do TRS Líquido", "▸", THEME['accent_lime'])
    if not df.empty and 'TRS_BRUTO' in df.columns:
        res_dia = df.groupby(df['DATA'].dt.date).agg(TRS_BRUTO=('TRS_BRUTO','mean'), PRODUZIDO=('PRODUZIDO','sum'), APROVADO=('APROVADO','sum')).reset_index()
        res_dia['DATA'] = pd.to_datetime(res_dia['DATA'])
        res_dia['TRS Líquido (%)'] = res_dia['TRS_BRUTO'] * 100
        res_dia = res_dia.sort_values('DATA')
        if not res_dia.empty:
            fig, ax = plt.subplots(figsize=(14, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS Líquido Diário — Período Selecionado", ylabel="TRS Líquido (%)")
            ax.fill_between(res_dia['DATA'], 0, res_dia['TRS Líquido (%)'], alpha=0.12, color=THEME['accent_lime'])
            ax.plot(res_dia['DATA'], res_dia['TRS Líquido (%)'],
                    marker='o', markersize=6, linewidth=2.5, color=THEME['accent_lime'], alpha=0.95,
                    label='TRS Líquido', markerfacecolor=THEME['bg_card'],
                    markeredgecolor=THEME['accent_lime'], markeredgewidth=2)
            if len(res_dia) > 1:
                mm = res_dia['TRS Líquido (%)'].rolling(window=min(3, len(res_dia)), min_periods=1).mean()
                ax.plot(res_dia['DATA'], mm, color=THEME['accent_yellow'], alpha=0.8,
                        linewidth=1.8, linestyle='--', label='Média 3 dias')
            ax.axhline(y=85, color=THEME['accent_red'], linestyle=':', alpha=0.7, linewidth=1.5, label='Meta 85%')
            ax.legend(framealpha=0.15, facecolor=THEME['bg_card'], edgecolor=THEME['border_bright'],
                      labelcolor=THEME['text_primary'], fontsize=9)
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha='right', fontsize=8, color=THEME['text_muted'])
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Por Praça ──
    if 'PRAÇA' in df.columns and not df.empty and 'TRS_BRUTO' in df.columns:
        render_section_header("TRS Líquido por Praça", "▸", THEME['accent_lime'])
        res_praca = df.groupby('PRAÇA').agg(TRS_BRUTO=('TRS_BRUTO','mean'), PRODUZIDO=('PRODUZIDO','sum')).reset_index()
        res_praca['TRS Líquido (%)'] = res_praca['TRS_BRUTO'] * 100
        res_praca = res_praca.sort_values('TRS Líquido (%)', ascending=False)
        if not res_praca.empty:
            fig, ax = plt.subplots(figsize=(10, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS Líquido Médio por Praça", ylabel="TRS Líquido (%)")
            bar_cols = [THEME['accent_lime'] if v >= 85 else THEME['accent_orange'] if v >= 70 else THEME['accent_red']
                        for v in res_praca['TRS Líquido (%)']]
            bars = ax.bar(range(len(res_praca)), res_praca['TRS Líquido (%)'], color=bar_cols,
                          alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5, width=0.6)
            ax.axhline(y=85, color=THEME['accent_red'], linestyle='--', alpha=0.4, linewidth=1.5)
            for bar, v in zip(bars, res_praca['TRS Líquido (%)']):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                        f'{v:.1f}%', ha='center', va='bottom', fontsize=9,
                        color=THEME['text_primary'], fontweight='bold')
            ax.set_xticks(range(len(res_praca)))
            ax.set_xticklabels(res_praca['PRAÇA'], rotation=40, ha='right', fontsize=9)
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Mensal ──
    render_section_header("TRS Líquido Mensal", "▸", THEME['accent_lime'])
    if not df.empty and 'ANO_MES' in df.columns and 'TRS_BRUTO' in df.columns:
        res_mes = df.groupby('ANO_MES').agg(TRS_BRUTO=('TRS_BRUTO','mean'), PRODUZIDO=('PRODUZIDO','sum'), APROVADO=('APROVADO','sum')).reset_index()
        res_mes['TRS Líquido (%)'] = res_mes['TRS_BRUTO'] * 100
        res_mes = res_mes.sort_values('ANO_MES')
        if not res_mes.empty:
            fig, ax = plt.subplots(figsize=(12, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS Líquido Mensal", ylabel="TRS Líquido (%)")
            bar_cols = [THEME['accent_lime'] if v >= 85 else THEME['accent_orange'] if v >= 70 else THEME['accent_red']
                        for v in res_mes['TRS Líquido (%)']]
            bars = ax.bar(range(len(res_mes)), res_mes['TRS Líquido (%)'], color=bar_cols,
                          alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5, width=0.65)
            ax.axhline(y=85, color=THEME['accent_red'], linestyle='--', alpha=0.4, linewidth=1.5)
            for bar, v in zip(bars, res_mes['TRS Líquido (%)']):
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                        f'{v:.1f}%', ha='center', va='bottom', fontsize=10,
                        color=THEME['text_primary'], fontweight='bold')
            meses_fmt = []
            for m in res_mes['ANO_MES']:
                ms = str(m)
                meses_fmt.append(ms[5:7]+'/'+ms[:4] if len(ms) > 6 else ms)
            ax.set_xticks(range(len(res_mes)))
            ax.set_xticklabels(meses_fmt, rotation=40, ha='right', fontsize=9)
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Por Turno ──
    if not df.empty and 'TURNO' in df.columns and 'TRS_BRUTO' in df.columns:
        render_section_header("TRS Líquido por Turno", "▸", THEME['accent_lime'])
        turno_data = [{'Turno': t, 'TRS Líquido': df[df['TURNO'] == t]['TRS_BRUTO'].mean() * 100} for t in df['TURNO'].unique()]
        df_tt = pd.DataFrame(turno_data)
        if not df_tt.empty:
            fig, ax = plt.subplots(figsize=(8, 4), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS Líquido por Turno", ylabel="TRS Líquido (%)")
            cores_t = {'M': THEME['accent_cyan'], 'T': THEME['accent_orange'], 'N': THEME['accent_lime'],
                       'A': THEME['accent_cyan'], 'B': THEME['accent_orange'], 'C': THEME['accent_lime']}
            bc = [cores_t.get(t, THEME['text_muted']) for t in df_tt['Turno']]
            ax.bar(range(len(df_tt)), df_tt['TRS Líquido'], color=bc, alpha=0.88,
                   edgecolor=THEME['bg_card'], linewidth=1.5, width=0.55)
            ax.axhline(y=85, color=THEME['accent_red'], linestyle='--', alpha=0.5, linewidth=1.5)
            for i, v in enumerate(df_tt['TRS Líquido']):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom', fontweight='bold', fontsize=11, color=THEME['text_primary'])
            ax.set_xticks(range(len(df_tt)))
            ax.set_xticklabels(df_tt['Turno'], fontsize=11)
            ax.set_ylim(0, 115)
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    # ── Defeitos ──
    if mostrar_defeitos:
        render_section_header("Estratificação de Defeitos", "▸", THEME['accent_lime'])
        colunas_defeitos = [
            'BOLHA','PEDRA','CALCINADO','BALANÇANDO','AMASSADO','OVAL','CORTE','QUEBRADA',
            'VIDRO GRUDADO','CORDA','FORMA','RISCO','TORTO','RUGA','GABARITO','SUJEIRA',
            'EMPENO','MARCAS','FALHADA','DOBRA','CHUPADO','ARREADO','GOSMA','BARRO','CROMO','MACHO'
        ]
        def_exist = []
        for defeito in colunas_defeitos:
            for col in df.columns:
                if col.upper() == defeito.upper():
                    def_exist.append(col)
                    break
        if def_exist:
            df_def = df[def_exist].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_def_s = df_def.sum().sort_values(ascending=False)
            df_def_s = df_def_s[df_def_s > 0]
            if not df_def_s.empty:
                fig, ax = plt.subplots(figsize=(12, 4), facecolor=THEME['bg_card'])
                apply_chart_style(ax, fig, "Defeitos — Somatório", ylabel="Quantidade")
                bars = ax.bar(range(len(df_def_s)), df_def_s.values,
                              color=THEME['accent_red'], alpha=0.8,
                              edgecolor=THEME['bg_card'], linewidth=1.2)
                ax.set_xticks(range(len(df_def_s)))
                ax.set_xticklabels(df_def_s.index, rotation=40, ha='right', fontsize=9, color=THEME['text_muted'])
                for bar, val in zip(bars, df_def_s.values):
                    if val > 0:
                        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() * 1.02,
                                f"{int(val):,}".replace(",","."), ha='center', va='bottom',
                                fontsize=8, color=THEME['text_primary'])
                fig.tight_layout(pad=1.5)
                st.pyplot(fig)
                plt.close(fig)
                st.caption(f"Total de defeitos: {int(df_def_s.sum()):,}".replace(",","."))

    st.markdown(f"""
    <div style="text-align:right;padding:16px 0 8px;
        font-family:'JetBrains Mono',monospace;font-size:10px;
        color:{THEME['text_muted']};letter-spacing:.1em;">
        TRS DASHBOARD · SOPRO · {get_horario_brasilia()}
    </div>
    """, unsafe_allow_html=True)
