import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
import os
import numpy as np
import io
import unicodedata
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import traceback
from typing import List, Optional, Dict, Any
from dataclasses import dataclass
import time
import json
from datetime import datetime, timedelta, date, time

# ======================
# SISTEMA DE NOTIFICAÇÕES - POPUP SIMPLES (APENAS REGISTROS NOVOS DO DIA ATUAL)
# ======================

# Arquivo para armazenar IDs dos registros já notificados
ARQUIVO_NOTIFICACOES = "notificacoes_enviadas.json"

class SistemaNotificacao:
    """Sistema simples de notificações - só mostra registros novos do dia atual"""
    
    def __init__(self):
        self.notificacoes_enviadas = self.carregar_notificacoes()
    
    def carregar_notificacoes(self):
        """Carrega lista de registros já notificados"""
        try:
            if os.path.exists(ARQUIVO_NOTIFICACOES):
                with open(ARQUIVO_NOTIFICACOES, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    if isinstance(dados, dict) and 'ar' in dados and 'rm' in dados:
                        return dados
        except:
            pass
        return {"ar": [], "rm": [], "data_ultima_limpeza": ""}
    
    def salvar_notificacoes(self):
        """Salva lista de registros notificados"""
        try:
            with open(ARQUIVO_NOTIFICACOES, 'w', encoding='utf-8') as f:
                json.dump(self.notificacoes_enviadas, f, ensure_ascii=False)
        except:
            pass
    
    def limpar_notificacoes_antigas(self):
        """Remove notificações de dias anteriores (executa uma vez por dia)"""
        hoje = datetime.now().strftime("%Y-%m-%d")
        if self.notificacoes_enviadas.get("data_ultima_limpeza") != hoje:
            self.notificacoes_enviadas["ar"] = []
            self.notificacoes_enviadas["rm"] = []
            self.notificacoes_enviadas["data_ultima_limpeza"] = hoje
            self.salvar_notificacoes()
    
    def verificar_novos_registros(self):
        """
        Verifica apenas registros do dia atual que ainda NÃO foram notificados.
        Retorna listas de novos ARs e RMs.
        """
        hoje = datetime.now().date()
        novos_ar = []
        novos_rm = []
        
        # Limpar registros antigos (uma vez por dia)
        self.limpar_notificacoes_antigas()
        
        # ===== VERIFICAR NOVOS AVISOS DE REJEIÇÃO (AR) =====
        try:
            registros_ar = carregar_registros_ar()
            if registros_ar:
                for registro in registros_ar:
                    if registro.data and registro.data.date() == hoje:
                        if str(registro.numero) not in self.notificacoes_enviadas["ar"]:
                            novos_ar.append({
                                "numero": registro.numero,
                                "data": registro.data.strftime("%d/%m/%Y"),
                                "hora": registro.hora,
                                "referencia": registro.referencia[:35] + "..." if len(registro.referencia) > 35 else registro.referencia,
                                "emissor": registro.emissor,
                                "tipo": "AR"
                            })
                            self.notificacoes_enviadas["ar"].append(str(registro.numero))
        except Exception as e:
            pass
        
        # ===== VERIFICAR NOVAS REQUISIÇÕES DE MANUTENÇÃO (RM) =====
        try:
            registros_rm = carregar_registros_rm()
            if registros_rm:
                for registro in registros_rm:
                    if registro.data and registro.data.date() == hoje:
                        if str(registro.id) not in self.notificacoes_enviadas["rm"]:
                            novos_rm.append({
                                "id": registro.id,
                                "data": registro.data.strftime("%d/%m/%Y"),
                                "hora": registro.hora,
                                "equipamento": registro.equipamento[:35] + "..." if len(registro.equipamento) > 35 else registro.equipamento,
                                "emissor": registro.emissor,
                                "tipo": "RM"
                            })
                            self.notificacoes_enviadas["rm"].append(str(registro.id))
        except Exception as e:
            pass
        
        if novos_ar or novos_rm:
            self.salvar_notificacoes()
        
        return novos_ar, novos_rm

# Instância global do sistema de notificações
sistema_notificacao = SistemaNotificacao()

# ======================
# CSS PARA POPUP SIMPLES (BALÃO INFERIOR DIREITO)
# ======================
NOTIFICACAO_CSS = """
<style>
/* Container de popups - canto inferior direito */
.popup-container {
    position: fixed;
    bottom: 20px;
    right: 20px;
    z-index: 99999;
    display: flex;
    flex-direction: column;
    gap: 10px;
    pointer-events: none;
}

/* Estilo do popup */
.simple-popup {
    pointer-events: auto;
    background: white;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    width: 280px;
    overflow: hidden;
    animation: slideIn 0.3s ease-out;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

@keyframes slideIn {
    from {
        transform: translateX(100%);
        opacity: 0;
    }
    to {
        transform: translateX(0);
        opacity: 1;
    }
}

@keyframes fadeOut {
    from {
        opacity: 1;
        transform: translateX(0);
    }
    to {
        opacity: 0;
        transform: translateX(100%);
        visibility: hidden;
    }
}

.simple-popup.fade-out {
    animation: fadeOut 0.3s ease-out forwards;
}

/* Header do popup */
.popup-header {
    padding: 8px 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid;
}

.popup-title {
    font-weight: 600;
    font-size: 12px;
    display: flex;
    align-items: center;
    gap: 6px;
}

.popup-close {
    cursor: pointer;
    font-size: 16px;
    font-weight: bold;
    color: #9ca3af;
    background: none;
    border: none;
    padding: 0 4px;
    line-height: 1;
}

.popup-close:hover {
    color: #ef4444;
}

/* Corpo do popup */
.popup-body {
    padding: 10px 12px;
    font-size: 11px;
}

.popup-line {
    margin-bottom: 6px;
    display: flex;
    gap: 8px;
}

.popup-label {
    font-weight: 600;
    color: #6b7280;
    min-width: 55px;
    font-size: 10px;
}

.popup-value {
    color: #1f2937;
    word-break: break-word;
    flex: 1;
    font-size: 11px;
}

/* Footer */
.popup-footer {
    background: #f9fafb;
    padding: 5px 12px;
    font-size: 9px;
    color: #9ca3af;
    text-align: right;
    border-top: 1px solid #e5e7eb;
}
</style>

<script>
function fecharPopup(elementId) {
    const element = document.getElementById(elementId);
    if (element) {
        element.classList.add('fade-out');
        setTimeout(() => {
            if (element.parentNode) element.remove();
        }, 300);
    }
}
</script>
"""

def gerar_popup_html(notificacao):
    """Gera o HTML do popup para uma notificação"""
    tipo = notificacao["tipo"]
    timestamp = datetime.now().strftime("%H:%M:%S")
    
    if tipo == "AR":
        cor = "#dc2626"
        icone = "📋"
        titulo = "NOVO AVISO DE REJEIÇÃO"
        popup_id = f"popup_ar_{notificacao['numero']}_{timestamp.replace(':', '')}"
        
        html = f'''
        <div id="{popup_id}" class="simple-popup">
            <div class="popup-header" style="border-bottom-color: {cor};">
                <div class="popup-title" style="color: {cor};">
                    <span>{icone}</span> {titulo}
                </div>
                <button class="popup-close" onclick="fecharPopup('{popup_id}')">✕</button>
            </div>
            <div class="popup-body">
                <div class="popup-line">
                    <span class="popup-label">Nº:</span>
                    <span class="popup-value">{notificacao['numero']}</span>
                </div>
                <div class="popup-line">
                    <span class="popup-label">Ref:</span>
                    <span class="popup-value">{notificacao['referencia']}</span>
                </div>
                <div class="popup-line">
                    <span class="popup-label">Emissor:</span>
                    <span class="popup-value">{notificacao['emissor']}</span>
                </div>
                <div class="popup-line">
                    <span class="popup-label">Hora:</span>
                    <span class="popup-value">{notificacao['hora']}</span>
                </div>
            </div>
            <div class="popup-footer">
                {timestamp}
            </div>
        </div>
        '''
    else:
        cor = "#10b981"
        icone = "🔧"
        titulo = "NOVA REQUISIÇÃO DE MANUTENÇÃO"
        popup_id = f"popup_rm_{notificacao['id']}_{timestamp.replace(':', '')}"
        
        html = f'''
        <div id="{popup_id}" class="simple-popup">
            <div class="popup-header" style="border-bottom-color: {cor};">
                <div class="popup-title" style="color: {cor};">
                    <span>{icone}</span> {titulo}
                </div>
                <button class="popup-close" onclick="fecharPopup('{popup_id}')">✕</button>
            </div>
            <div class="popup-body">
                <div class="popup-line">
                    <span class="popup-label">ID:</span>
                    <span class="popup-value">{notificacao['id']}</span>
                </div>
                <div class="popup-line">
                    <span class="popup-label">Equip:</span>
                    <span class="popup-value">{notificacao['equipamento']}</span>
                </div>
                <div class="popup-line">
                    <span class="popup-label">Emissor:</span>
                    <span class="popup-value">{notificacao['emissor']}</span>
                </div>
                <div class="popup-line">
                    <span class="popup-label">Hora:</span>
                    <span class="popup-value">{notificacao['hora']}</span>
                </div>
            </div>
            <div class="popup-footer">
                {timestamp}
            </div>
        </div>
        '''
    
    return html, popup_id

def verificar_e_exibir_popups():
    """Função principal que verifica novos registros e exibe popups"""
    aba_atual = st.session_state.get("aba_selecionada", "")
    if aba_atual in ["AVISO DE REJEIÇÃO", "REQUISIÇÃO MANUTENÇÃO"]:
        return
    
    novos_ar, novos_rm = sistema_notificacao.verificar_novos_registros()
    
    todas_notificacoes = []
    for notif in novos_ar:
        todas_notificacoes.append(notif)
    for notif in novos_rm:
        todas_notificacoes.append(notif)
    
    if todas_notificacoes:
        if "popups_para_exibir" not in st.session_state:
            st.session_state.popups_para_exibir = []
        
        for notif in todas_notificacoes:
            chave = f"{notif['tipo']}_{notif.get('numero', notif.get('id'))}"
            if chave not in [p.get("chave") for p in st.session_state.popups_para_exibir]:
                notif["chave"] = chave
                st.session_state.popups_para_exibir.append(notif)
        
        st.rerun()

def renderizar_popups_pendentes():
    """Renderiza todos os popups pendentes no container inferior direito"""
    if "popups_para_exibir" in st.session_state and st.session_state.popups_para_exibir:
        st.markdown('<div class="popup-container" id="popup-container">', unsafe_allow_html=True)
        
        for notif in st.session_state.popups_para_exibir.copy():
            html_popup, _ = gerar_popup_html(notif)
            st.markdown(html_popup, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        st.session_state.popups_para_exibir = []

# ======================
# CONFIGURAÇÕES
# ======================
ID_PLANILHA_PRENSADOS_SOPRO = '1Hjy4UGtgwIPJgqmcv46LyXNWOrYk_oeJWWV5vlfKF2k'
ID_PLANILHA_TEMPERA = '1GJegUHosaQLEJVMCH6QVuKjSjuaxrWkgzNEr9vM5Yio'

ID_PLANILHA_AR = '12pz6EE1KDo41szDyGEyTyK27mAOb9F1FyU77_M1kL0o'
ABA_AR = 'AR'
ABA_RM = 'RM'
EMAIL_SERVICE_ACCOUNT = 'script-atualizacao@dashboard-gerencial-492613.iam.gserviceaccount.com'

PRACAS_NAO_SOPRO = ['GIL', 'GILSIMAR', 'ED CARLOS', 'EDI CARLOS', 'ROBÔ 2', 'ROBÔ-2', 'ROBÔ', 'ROBO']

ABAS = {
    'PRENSADOS': 'TRS_INDUSTRIAL',
    'SOPRO': 'TRS_SOPRO',
    'TÊMPERA': 'TRS_TEMPERA',
    'AVISO DE REJEIÇÃO': 'AR',
    'REQUISIÇÃO MANUTENÇÃO': 'RM',
    'FECHAMENTO TURNO': 'FT'  # NOVA LINHA
}

CAMINHO_PDF_AR = r"\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\0-AVISO DE REJEIÇÃO\1-PDF"
CAMINHO_PDF_RELATORIO_AR = r"\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\0-AVISO DE REJEIÇÃO\2-PDF"

EMAIL_CONFIG_AR = {
    "usuario": "erp@luvidarte.com.br",
    "senha": "Qualidade123#",
    "destinatarios": ["producao@luvidarte.com.br", "engenharia@luvidarte.com.br", 
                     "qualidade@luvidarte.com.br", "qualidade2@luvidarte.com.br"],
    "smtp_server": "email-ssl.com.br",
    "smtp_port": 465
}

OPCOES_DECISAO_AR = ["APROVADO CONDICIONAL", "REPROVADO", "EM ANÁLISE"]
OPCOES_STATUS_AR = ["ABERTO", "FINALIZADO"]
OPCOES_TURNO_AR = ["Manhã", "Tarde", "Noite"]

for caminho in [CAMINHO_PDF_AR, CAMINHO_PDF_RELATORIO_AR]:
    try:
        os.makedirs(caminho, exist_ok=True)
    except:
        pass

# ======================
# TEMA VISUAL - CLARO (LIGHT MODE)
# ======================
THEME = {
    'bg_primary':     '#F5F7FA',
    'bg_card':        '#FFFFFF',
    'bg_card2':       '#F8F9FC',
    'accent_cyan':    '#0078D4',
    'accent_lime':    '#107C10',
    'accent_orange':  '#E86C2C',
    'accent_yellow':  '#FFB900',
    'accent_red':     '#E81123',
    'accent_purple':  '#6B46C1',
    'text_primary':   '#1E1E1E',
    'text_muted':     '#605E5C',
    'border':         '#D1D1D1',
    'border_bright':  '#C0C0C0',
    'grid':           '#E0E0E0',
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
    from datetime import timezone, timedelta
    utc_now = datetime.now(timezone.utc)
    brasilia_offset = timezone(timedelta(hours=-3))
    agora_brasilia = utc_now.astimezone(brasilia_offset)
    return agora_brasilia.strftime('%d/%m/%Y %H:%M')
    
def get_horario_brasilia_obj():
    from datetime import timezone, timedelta
    utc_now = datetime.now(timezone.utc)
    brasilia_offset = timezone(timedelta(hours=-3))
    return utc_now.astimezone(brasilia_offset)    

# ======================
# FUNÇÕES AUXILIARES GLOBAIS
# ======================
def safe_float_tempera(val):
    if val is None or pd.isna(val):
        return 0.0
    try:
        val_str = str(val).strip()
        if val_str == '' or val_str == 'nan':
            return 0.0
        val_str = val_str.replace(',', '.')
        return float(val_str)
    except:
        return 0.0

def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if 'gcp_service_account' in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            return gspread.authorize(creds)
    except:
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
                return gspread.authorize(creds)
        except:
            pass
    return None

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

# ======================
# FUNÇÕES DE CARREGAMENTO DE DADOS
# ======================
@st.cache_data(ttl=300)
def carregar_dados_prensados():
    try:
        client = get_gspread_client()
        if client is None:
            return pd.DataFrame()
        sheet = client.open_by_key(ID_PLANILHA_PRENSADOS_SOPRO).worksheet('TRS_INDUSTRIAL')
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
        colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'TRS 100%', 'REFUGADO', 'BOQUETA']
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
    except Exception as e:
        st.error(f"Erro ao carregar dados de Prensados: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def carregar_dados_sopro():
    try:
        client = get_gspread_client()
        if client is None:
            return pd.DataFrame()
        sheet = client.open_by_key(ID_PLANILHA_PRENSADOS_SOPRO).worksheet('TRS_SOPRO')
        todos_dados = sheet.get_all_values()
        if len(todos_dados) < 2:
            return pd.DataFrame()
        cabecalho = todos_dados[0]
        valores = todos_dados[1:]
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

@st.cache_data(ttl=300)
def carregar_dados_tempera():
    try:
        client = get_gspread_client()
        if client is None:
            return pd.DataFrame()
        sheet = client.open_by_key(ID_PLANILHA_TEMPERA).worksheet('TRS_TEMPERA')
        todos_dados = sheet.get_all_values()
        
        if len(todos_dados) < 2:
            return pd.DataFrame()
        
        cabecalho = todos_dados[0]
        valores = todos_dados[1:]
        df = pd.DataFrame(valores, columns=cabecalho)
        colunas = list(df.columns)
        
        if len(colunas) >= 5:
            df = df.rename(columns={
                colunas[0]: 'PRODUCAO',
                colunas[1]: 'DATA_TEMP',
                colunas[2]: 'TURNO_TEMP',
                colunas[3]: 'PRODUTO',
                colunas[4]: 'GANCHEIRA'
            })
        
        if len(colunas) >= 8:
            df = df.rename(columns={
                colunas[5]: 'SUPERIOR',
                colunas[6]: 'MEIO',
                colunas[7]: 'INFERIOR'
            })
        
        if len(colunas) >= 11:
            df = df.rename(columns={
                colunas[8]: 'A1',
                colunas[9]: 'C1',
                colunas[10]: 'A2'
            })
        
        if len(colunas) >= 14:
            df = df.rename(columns={
                colunas[11]: 'C2',
                colunas[12]: 'A3',
                colunas[13]: 'C3'
            })
        
        if len(colunas) >= 17:
            df = df.rename(columns={
                colunas[14]: 'A4',
                colunas[15]: 'C4',
                colunas[16]: 'A5'
            })
        
        if len(colunas) >= 20:
            df = df.rename(columns={
                colunas[17]: 'C5',
                colunas[18]: 'A e B'
            })
        
        if 'DATA_TEMP' in df.columns:
            df['DATA'] = df['DATA_TEMP'].apply(converter_data_br)
        elif 'PRODUCAO' in df.columns:
            df['DATA'] = df['PRODUCAO'].apply(converter_data_br)
        
        if 'DATA' in df.columns:
            df = df.dropna(subset=['DATA'])
        
        colunas_numericas = ['SUPERIOR', 'MEIO', 'INFERIOR', 'A1', 'C1', 'A2', 'C2', 'A3', 'C3', 'A4', 'C4', 'A5', 'C5', 'A e B']
        
        for col in colunas_numericas:
            if col in df.columns:
                df[col] = df[col].apply(safe_float_tempera)
        
        if 'C2' in df.columns:
            def converter_tempo_c2(val):
                if pd.isna(val) or val == 0:
                    return 0
                if val <= 1:
                    return val * 100
                elif val <= 10:
                    return val * 10
                else:
                    return val
            df['C2'] = df['C2'].apply(converter_tempo_c2)
        
        colunas_posicoes_validas = []
        for col in df.columns:
            try:
                num = int(str(col).strip())
                if 19 <= num <= 70:
                    colunas_posicoes_validas.append(col)
            except:
                pass
        
        df['TOTAL_PECAS'] = 40
        df['APROVADO'] = 40
        df['TOTAL_DEFEITOS'] = 0
        df['IS_CRITICO'] = False
        
        MAPEAMENTO_DEFEITOS = {
            1: 'Estourou após furar',
            2: 'Quebra no resfriamento',
            3: 'Quebra teste impacto',
            4: 'Furada e não fraturou',
            5: 'Quebra de quarentena',
            6: 'Ovalizada'
        }
        CODIGOS_DEFEITO_REAIS = [2, 3, 4, 5, 6]
        
        for codigo, nome in MAPEAMENTO_DEFEITOS.items():
            nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
            df[f'QTD_{nome_clean}'] = 0
        
        for idx, row in df.iterrows():
            defeitos_contagem = {codigo: 0 for codigo in MAPEAMENTO_DEFEITOS.keys()}
            
            for col in colunas_posicoes_validas:
                try:
                    val = row[col]
                    if pd.notna(val) and str(val).strip():
                        codigo = int(float(str(val).strip()))
                        if codigo in MAPEAMENTO_DEFEITOS:
                            defeitos_contagem[codigo] += 1
                except:
                    pass
            
            total_defeitos_reais = sum(defeitos_contagem.get(cod, 0) for cod in CODIGOS_DEFEITO_REAIS)
            aprovadas = 40 - total_defeitos_reais
            
            df.at[idx, 'APROVADO'] = aprovadas
            df.at[idx, 'TOTAL_DEFEITOS'] = total_defeitos_reais
            df.at[idx, 'TRS (%)'] = (aprovadas / 40 * 100) if 40 > 0 else 0
            
            is_critico = False
            if defeitos_contagem.get(4, 0) >= 1:
                is_critico = True
            if defeitos_contagem.get(3, 0) > 2:
                is_critico = True
            df.at[idx, 'IS_CRITICO'] = is_critico
            
            for codigo, nome in MAPEAMENTO_DEFEITOS.items():
                nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
                col_nome = f'QTD_{nome_clean}'
                if col_nome in df.columns:
                    df.at[idx, col_nome] = defeitos_contagem.get(codigo, 0)
        
        return df
        
    except Exception as e:
        st.error(f"Erro ao carregar dados: {str(e)}")
        return pd.DataFrame()

# ======================
# FUNÇÕES DO SISTEMA AR (AVISO DE REJEIÇÃO)
# ======================
@dataclass
class RegistroAR:
    numero: Optional[int] = None
    data: Optional[datetime] = None
    hora: str = ""
    codigo: str = ""
    emissor: str = ""
    referencia: str = ""
    decisao: str = ""
    descricao: str = ""
    status: str = "ABERTO"
    disposicao: str = ""
    data_finalizacao: Optional[datetime] = None
    turno: str = ""

def sanitize_filename_ar(filename: str) -> str:
    filename = unicodedata.normalize("NFKD", filename).encode("ASCII", "ignore").decode("ASCII")
    filename = re.sub(r'[^a-zA-Z0-9_-]', '_', filename)
    return filename[:50]

def obter_proximo_numero_ar():
    try:
        client = get_gspread_client()
        if client is None:
            return 1
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_AR)
        todos_dados = sheet.get_all_values()
        if len(todos_dados) < 2:
            return 1
        numeros = []
        for row in todos_dados[1:]:
            if len(row) > 0 and row[0]:
                try:
                    num = int(float(row[0]))
                    numeros.append(num)
                except:
                    pass
        if not numeros:
            return 1
        return max(numeros) + 1
    except:
        return 1

def carregar_registros_ar(filtros: Dict[str, Any] = None) -> List[RegistroAR]:
    registros = []
    try:
        client = get_gspread_client()
        if client is None:
            return registros
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_AR)
        todos_dados = sheet.get_all_values()
        if len(todos_dados) < 2:
            return registros
        for row in todos_dados[1:]:
            if len(row) < 12:
                continue
            try:
                registro = RegistroAR()
                registro.numero = int(float(row[0])) if row[0].strip() else None
                registro.data = converter_data_br(row[1])
                registro.hora = row[2] if len(row) > 2 else ""
                registro.codigo = row[3] if len(row) > 3 else ""
                registro.emissor = row[4] if len(row) > 4 else ""
                registro.referencia = row[5] if len(row) > 5 else ""
                registro.decisao = row[6] if len(row) > 6 else ""
                registro.descricao = row[7] if len(row) > 7 else ""
                registro.status = row[8] if len(row) > 8 else "ABERTO"
                registro.disposicao = row[9] if len(row) > 9 else ""
                registro.data_finalizacao = converter_data_br(row[10]) if len(row) > 10 else None
                registro.turno = row[11] if len(row) > 11 else ""
                
                if filtros:
                    incluir = True
                    if filtros.get('numero') and filtros['numero'] != registro.numero:
                        incluir = False
                    if not incluir or (filtros.get('status') and filtros['status'].upper() != registro.status.upper()):
                        incluir = False
                    if not incluir or (filtros.get('decisao') and filtros['decisao'].upper() != registro.decisao.upper()):
                        incluir = False
                
                if not filtros or incluir:
                    registros.append(registro)
            except:
                continue
        registros.sort(key=lambda x: x.data if x.data else datetime.min, reverse=True)
    except:
        pass
    return registros

def salvar_registro_ar(registro: RegistroAR, eh_alteracao: bool = False) -> bool:
    try:
        client = get_gspread_client()
        if client is None:
            return False
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_AR)
        dados = [
            str(registro.numero), registro.data.strftime("%d/%m/%Y") if registro.data else "",
            registro.hora, registro.codigo, registro.emissor, registro.referencia,
            registro.decisao, registro.descricao, registro.status, registro.disposicao,
            registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else "",
            registro.turno
        ]
        if eh_alteracao:
            cell = sheet.find(str(registro.numero), in_column=1)
            if cell:
                for col, valor in enumerate(dados, start=1):
                    sheet.update_cell(cell.row, col, valor)
            else:
                sheet.append_row(dados)
        else:
            sheet.insert_row(dados, index=2)
        return True
    except:
        return False

def excluir_registro_ar(numero: int) -> bool:
    try:
        client = get_gspread_client()
        if client is None:
            return False
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_AR)
        cell = sheet.find(str(numero), in_column=1)
        if cell:
            sheet.delete_rows(cell.row)
            return True
        return False
    except:
        return False

def gerar_pdf_ar(registro: RegistroAR) -> Optional[bytes]:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elementos = []
        styles = getSampleStyleSheet()
        styleN = styles["Normal"]
        style_grande = ParagraphStyle('style_grande', parent=styleN, fontSize=12, leading=16)
        
        elementos.append(Paragraph("<b>AVISO DE REJEIÇÃO</b>", ParagraphStyle(name='Titulo', parent=styleN, fontSize=16, alignment=1, spaceAfter=12)))
        elementos.append(Paragraph("<b>CQ-018 REV004 - Luvidarte</b>", ParagraphStyle(name='Subtitulo', parent=styleN, fontSize=12, alignment=1, spaceAfter=24)))
        
        data_str = registro.data.strftime("%d/%m/%Y") if registro.data else ""
        data_fim_str = registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else ""
        
        tabela_cabecalho = Table([
            ["Nº Controle:", registro.numero, "Data:", data_str],
            ["Hora:", registro.hora, "Turno:", registro.turno],
            ["Código:", registro.codigo, "Status:", registro.status],
            ["Emissor:", registro.emissor, "", ""],
            ["Referência:", registro.referencia, "", ""],
            ["Situação:", registro.decisao, "Data Finalização:", data_fim_str]
        ], colWidths=[4*cm, 6*cm, 4*cm, 6*cm])
        
        tabela_cabecalho.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.5, colors.black),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,-1), "LEFT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("PADDING", (0,0), (-1,-1), 6),
            ("SPAN", (2,3), (3,3)),
            ("SPAN", (2,4), (3,4)),
        ]))
        elementos.append(tabela_cabecalho)
        elementos.append(Spacer(1, 24))
        
        elementos.append(Paragraph("<b>DESCRIÇÃO DO PROBLEMA:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        elementos.append(Paragraph(registro.descricao or "-", style_grande))
        elementos.append(Spacer(1, 24))
        
        elementos.append(Paragraph("<b>DISPOSIÇÃO / AÇÕES TOMADAS:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        if registro.disposicao and registro.disposicao.strip():
            elementos.append(Paragraph(registro.disposicao, style_grande))
            elementos.append(Spacer(1, 12))
        
        elementos.append(Paragraph("<b>ASSINATURAS:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        
        tabela_assinatura = Table([
            ["Responsável:", "________________________________________"],
            ["Cargo:", "________________________________________"],
            ["Visto:", "________________________________________"],
            ["Cargo:", "________________________________________"],
            ["Data:", "________________________________________"]
        ], colWidths=[4*cm, 14*cm])
        
        tabela_assinatura.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
            ("ALIGN", (0,0), (0,-1), "LEFT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("PADDING", (0,0), (-1,-1), 10),
            ("BACKGROUND", (0,0), (0,-1), colors.lightgrey),
        ]))
        elementos.append(tabela_assinatura)
        
        elementos.append(Spacer(1, 36))
        elementos.append(Paragraph(f"Documento gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ParagraphStyle(name='Rodape', parent=styleN, fontSize=8, alignment=2)))
        elementos.append(Paragraph("* Espaços em branco devem ser preenchidos manualmente após impressão", ParagraphStyle(name='RodapeInstrucao', parent=styleN, fontSize=8, alignment=2, textColor=colors.gray)))
        
        doc.build(elementos)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        return None

def enviar_email_ar(destinatarios, assunto, corpo, anexo_bytes=None, nome_anexo=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG_AR["usuario"]
        msg['To'] = ", ".join(destinatarios)
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain'))
        if anexo_bytes and nome_anexo:
            anexo = MIMEApplication(anexo_bytes, _subtype='pdf')
            anexo.add_header('Content-Disposition', 'attachment', filename=nome_anexo)
            msg.attach(anexo)
        with smtplib.SMTP_SSL(EMAIL_CONFIG_AR["smtp_server"], EMAIL_CONFIG_AR["smtp_port"], timeout=30) as server:
            server.login(EMAIL_CONFIG_AR["usuario"], EMAIL_CONFIG_AR["senha"])
            server.send_message(msg)
        return True
    except:
        return False

# ======================
# FUNÇÕES DO SISTEMA RM (REQUISIÇÃO MANUTENÇÃO)
# ======================
@dataclass
class RegistroRM:
    id: Optional[int] = None
    data: Optional[datetime] = None
    hora: str = ""
    emissor: str = ""
    equipamento: str = ""
    setor: str = ""
    caracter: str = ""
    setor2: str = ""
    problema: str = ""
    trabalho: str = ""
    analise: str = ""
    status: str = "ABERTO"
    data_finalizacao: Optional[datetime] = None
    emissor2: str = ""

OPCOES_CARATER_RM = ["1 - Risco Físico/Segurança", "2 - Impacto Imediato na Produção", "3 - Impacto a Longo Prazo", "4 - Melhoria/Preventiva"]
OPCOES_SETORES_RM = ["Produção", "Corte", "Vidraria", "Rodaria", "Embalagem", "Expedição", "Qualidade", "Ferramentaria", "Manutenção", "Outros"]
OPCOES_SETORES2_RM = ["Elétrica", "Mecânica", "Informática", "Ferramentaria", "Manutenção Geral"]
OPCOES_STATUS_RM = ["ABERTO", "EM ANDAMENTO", "FINALIZADO", "CANCELADO"]

def obter_proximo_id_rm():
    try:
        client = get_gspread_client()
        if client is None:
            return 1
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_RM)
        todos_dados = sheet.get_all_values()
        if len(todos_dados) < 2:
            return 1
        ids = []
        for row in todos_dados[1:]:
            if row and row[0].strip():
                try:
                    ids.append(int(row[0]))
                except:
                    pass
        return max(ids) + 1 if ids else 1
    except:
        return 1

def carregar_registros_rm(filtros: Dict[str, Any] = None) -> List[RegistroRM]:
    registros = []
    try:
        client = get_gspread_client()
        if client is None:
            return registros
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_RM)
        todos_dados = sheet.get_all_values()
        if len(todos_dados) < 2:
            return registros
        for row in todos_dados[1:]:
            if len(row) < 14:
                continue
            try:
                registro = RegistroRM()
                registro.id = int(row[0]) if row[0].strip() else None
                registro.data = converter_data_br(row[1])
                registro.hora = row[2] if len(row) > 2 else ""
                registro.emissor = row[3] if len(row) > 3 else ""
                registro.equipamento = row[4] if len(row) > 4 else ""
                registro.setor = row[5] if len(row) > 5 else ""
                registro.caracter = row[6] if len(row) > 6 else ""
                registro.setor2 = row[7] if len(row) > 7 else ""
                registro.problema = row[8] if len(row) > 8 else ""
                registro.trabalho = row[9] if len(row) > 9 else ""
                registro.analise = row[10] if len(row) > 10 else ""
                registro.status = row[11] if len(row) > 11 else "ABERTO"
                registro.data_finalizacao = converter_data_br(row[12]) if len(row) > 12 else None
                registro.emissor2 = row[13] if len(row) > 13 else ""
                
                if filtros:
                    incluir = True
                    if filtros.get('id') and filtros['id'] != registro.id:
                        incluir = False
                    if filtros.get('equipamento') and filtros['equipamento'].lower() not in registro.equipamento.lower():
                        incluir = False
                    if filtros.get('status') and filtros['status'] != registro.status:
                        incluir = False
                else:
                    incluir = True
                
                if incluir and registro.id is not None:
                    registros.append(registro)
            except:
                continue
        registros.sort(key=lambda x: x.id if x.id else 0, reverse=True)
    except:
        pass
    return registros

def salvar_registro_rm(registro: RegistroRM, eh_alteracao: bool = False) -> bool:
    try:
        client = get_gspread_client()
        if client is None:
            return False
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_RM)
        dados = [
            str(registro.id) if registro.id else "",
            registro.data.strftime("%d/%m/%Y") if registro.data else "",
            registro.hora, registro.emissor, registro.equipamento, registro.setor,
            registro.caracter, registro.setor2, registro.problema, registro.trabalho,
            registro.analise, registro.status,
            registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else "",
            registro.emissor2
        ]
        if eh_alteracao:
            cell = sheet.find(str(registro.id), in_column=1)
            if cell:
                for col, valor in enumerate(dados, start=1):
                    sheet.update_cell(cell.row, col, valor)
            else:
                sheet.append_row(dados)
        else:
            sheet.insert_row(dados, index=2)
        return True
    except:
        return False

def excluir_registro_rm(id: int) -> bool:
    try:
        client = get_gspread_client()
        if client is None:
            return False
        sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_RM)
        cell = sheet.find(str(id), in_column=1)
        if cell:
            sheet.delete_rows(cell.row)
            return True
        return False
    except:
        return False

def gerar_pdf_rm(registro: RegistroRM) -> Optional[bytes]:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elementos = []
        styles = getSampleStyleSheet()
        styleN = styles["Normal"]
        
        elementos.append(Paragraph("<b>REQUISIÇÃO DE MANUTENÇÃO</b>", ParagraphStyle(name='Titulo', parent=styles["Heading1"], fontSize=16, alignment=1, spaceAfter=12)))
        elementos.append(Paragraph("<b>MF-001 - Luvidarte</b>", ParagraphStyle(name='Subtitulo', parent=styles["Heading2"], fontSize=12, alignment=1, spaceAfter=24)))
        
        data_str = registro.data.strftime("%d/%m/%Y") if registro.data else ""
        data_fim_str = registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else ""
        
        tabela_dados = Table([
            ["ID:", registro.id, "Data:", data_str, "Hora:", registro.hora],
            ["Emissor:", registro.emissor, "Equipamento:", registro.equipamento, "Setor:", registro.setor],
            ["Caráter:", registro.caracter, "Setor Destino:", registro.setor2, "Status:", registro.status],
            ["Emissor Técnico:", registro.emissor2, "Data Finalização:", data_fim_str, "", ""],
        ], colWidths=[2.5*cm, 4*cm, 2.5*cm, 4*cm, 2*cm, 2.5*cm])
        
        tabela_dados.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.5, colors.black),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (0,0), (-1,-1), "LEFT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("PADDING", (0,0), (-1,-1), 6),
        ]))
        elementos.append(tabela_dados)
        elementos.append(Spacer(1, 24))
        
        elementos.append(Paragraph("<b>DESCRIÇÃO DO PROBLEMA:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        elementos.append(Paragraph(registro.problema or "-", styleN))
        elementos.append(Spacer(1, 24))
        
        elementos.append(Paragraph("<b>TRABALHO REALIZADO:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        elementos.append(Paragraph(registro.trabalho or "_________________________", styleN))
        elementos.append(Spacer(1, 24))
        
        elementos.append(Paragraph("<b>ANÁLISE DO SERVIÇO:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        elementos.append(Paragraph(registro.analise or "_________________________", styleN))
        elementos.append(Spacer(1, 24))
        
        elementos.append(Paragraph("<b>ASSINATURAS:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
        tabela_assinatura = Table([
            ["Solicitante", "Responsável Técnico", "Conferência Qualidade"],
            ["_________________________", "_________________________", "_________________________"],
            ["Data: __/__/____", "Data: __/__/____", "Data: __/__/____"]
        ], colWidths=[5.5*cm, 5.5*cm, 5.5*cm])
        tabela_assinatura.setStyle(TableStyle([
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("PADDING", (0,0), (-1,-1), 10),
            ("FONTNAME", (0,0), (0,0), "Helvetica-Bold"),
            ("FONTNAME", (1,0), (1,0), "Helvetica-Bold"),
            ("FONTNAME", (2,0), (2,0), "Helvetica-Bold"),
        ]))
        elementos.append(tabela_assinatura)
        
        elementos.append(Spacer(1, 36))
        elementos.append(Paragraph(f"Documento gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ParagraphStyle(name='Rodape', parent=styleN, fontSize=8, alignment=2)))
        
        doc.build(elementos)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        return None

# ======================
# FUNÇÕES DE RENDERIZAÇÃO
# ======================
def render_page_header(title: str, subtitle: str, accent: str = None):
    if accent is None:
        accent = THEME['accent_cyan']
    st.markdown(f"""
    <div style="padding: 28px 0 20px 0; border-bottom: 1px solid {THEME['border_bright']}; margin-bottom: 28px; display: flex; align-items: center; gap: 16px;">
        <div style="width: 4px; height: 48px; background: linear-gradient(180deg, {accent}, transparent); border-radius: 2px;"></div>
        <div>
            <div style="font-family: 'JetBrains Mono', monospace; font-size: 10px; letter-spacing: 0.25em; color: {accent}; text-transform: uppercase;">LUVIDARTE / TRS DASHBOARD</div>
            <div style="font-family: 'Rajdhani', sans-serif; font-size: 36px; font-weight: 700; color: {THEME['text_primary']}; letter-spacing: 0.1em; text-transform: uppercase;">{title}</div>
            <div style="font-family: 'Barlow', sans-serif; font-size: 13px; color: {THEME['text_muted']}; margin-top: 4px;">{subtitle}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_section_header(title: str, icon: str = "▸", accent: str = None):
    if accent is None:
        accent = THEME['accent_cyan']
    st.markdown(f"""
    <div style="display: flex; align-items: center; gap: 10px; margin: 28px 0 14px 0; padding-bottom: 8px; border-bottom: 1px solid {THEME['border']};">
        <span style="color: {accent}; font-size: 16px;">{icon}</span>
        <span style="font-family: 'Rajdhani', sans-serif; font-size: 18px; font-weight: 600; letter-spacing: 0.1em; text-transform: uppercase; color: {THEME['text_primary']};">{title}</span>
    </div>
    """, unsafe_allow_html=True)

def render_kpi_card(label: str, value: str, accent: str = None, icon: str = ""):
    if accent is None:
        accent = THEME['accent_cyan']
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, {THEME['bg_card']} 0%, {THEME['bg_card2']} 100%); border: 1px solid {THEME['border_bright']}; border-radius: 8px; padding: 18px 22px; position: relative; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <div style="position: absolute; top: 0; left: 0; right: 0; height: 2px; background: linear-gradient(90deg, {accent}, transparent);"></div>
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 9px; letter-spacing: 0.2em; text-transform: uppercase; color: {THEME['text_muted']}; margin-bottom: 8px;">{icon} {label}</div>
        <div style="font-family: 'Rajdhani', sans-serif; font-size: 34px; font-weight: 700; color: {accent}; letter-spacing: 0.03em; line-height: 1;">{value}</div>
    </div>
    """, unsafe_allow_html=True)

def apply_chart_style(ax, fig, title: str, xlabel: str = "", ylabel: str = "", accent: str = None):
    if accent is None:
        accent = THEME['accent_cyan']
    fig.patch.set_facecolor(THEME['bg_card'])
    ax.set_facecolor(THEME['bg_card'])
    ax.set_title(title, fontsize=14, fontweight='bold', color=THEME['text_primary'], pad=16, loc='left')
    if xlabel:
        ax.set_xlabel(xlabel, fontsize=10, color=THEME['text_muted'], labelpad=8)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10, color=THEME['text_muted'], labelpad=8)
    ax.tick_params(colors=THEME['text_muted'], labelsize=9)
    ax.grid(True, alpha=0.3, color=THEME['grid'], linewidth=0.8, linestyle='--')
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_edgecolor(THEME['border_bright'])
        spine.set_linewidth(0.8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

# ======================
# CSS GLOBAL
# ======================
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&family=Barlow:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] {{
      font-family: 'Barlow', sans-serif;
      background-color: {THEME['bg_primary']} !important;
      color: {THEME['text_primary']} !important;
  }}
  .stApp {{ background-color: {THEME['bg_primary']} !important; }}

  [data-testid="stSidebar"] {{
      background: linear-gradient(180deg, #FFFFFF 0%, #F0F2F5 100%) !important;
      border-right: 1px solid {THEME['border_bright']} !important;
  }}
  
  [data-testid="stSidebar"] .stRadio label {{
      color: #000000 !important;
      font-weight: bold !important;
      font-family: 'Rajdhani', sans-serif !important;
      font-size: 15px !important;
      letter-spacing: 0.08em;
  }}
  
  [data-testid="stSidebar"] .stRadio * {{
      color: #000000 !important;
      font-weight: bold !important;
  }}
  
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stTextInput label,
  [data-testid="stSidebar"] .stDateInput label,
  [data-testid="stSidebar"] .stNumberInput label,
  [data-testid="stSidebar"] .stCheckbox label {{
      color: #000000 !important;
      font-weight: bold !important;
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 11px !important;
      text-transform: uppercase;
      letter-spacing: 0.12em;
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
  }}

  .stSelectbox div[data-baseweb="select"] > div,
  .stTextInput input,
  .stNumberInput input,
  .stDateInput input {{
      background-color: {THEME['bg_card']} !important;
      border: 1px solid {THEME['border_bright']} !important;
      color: {THEME['text_primary']} !important;
      border-radius: 4px !important;
      font-family: 'JetBrains Mono', monospace !important;
      font-size: 12px !important;
  }}

  [data-testid="stMetric"] {{
      background: linear-gradient(135deg, {THEME['bg_card']} 0%, {THEME['bg_card2']} 100%) !important;
      border: 1px solid {THEME['border_bright']} !important;
      border-radius: 8px !important;
      padding: 16px 20px !important;
      position: relative;
      overflow: hidden;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
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
  }}

  hr {{
      border: none !important;
      border-top: 1px solid {THEME['border_bright']} !important;
      margin: 24px 0 !important;
  }}

  .stInfo {{ background-color: rgba(0,120,212,0.08) !important; border-left: 3px solid {THEME['accent_cyan']} !important; }}
  .stWarning {{ background-color: rgba(232,108,44,0.08) !important; border-left: 3px solid {THEME['accent_orange']} !important; }}
  .stSuccess {{ background-color: rgba(16,124,16,0.08) !important; border-left: 3px solid {THEME['accent_lime']} !important; }}
  .stError {{ background-color: rgba(232,17,35,0.08) !important; border-left: 3px solid {THEME['accent_red']} !important; }}
</style>
{NOTIFICACAO_CSS}
""", unsafe_allow_html=True)

# ======================
# SIDEBAR - navegação
# ======================
with st.sidebar:
    st.markdown(f"""
    <div style="text-align: center; padding: 20px 0 16px; border-bottom: 1px solid {THEME['border_bright']}; margin-bottom: 20px;">
        <div style="font-family: 'Rajdhani', sans-serif; font-size: 24px; font-weight: 700; color: {THEME['accent_cyan']}; letter-spacing: 0.2em;">⚙ TRS</div>
        <div style="font-family: 'JetBrains Mono', monospace; font-size: 9px; color: {THEME['text_muted']}; letter-spacing: 0.2em; text-transform: uppercase;">Industrial Dashboard</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"<div style='font-family:JetBrains Mono,monospace;font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:{THEME['accent_cyan']};margin-bottom:8px'>▸ Setor</div>", unsafe_allow_html=True)
    aba_selecionada = st.radio("", list(ABAS.keys()), label_visibility="collapsed")
    st.session_state.aba_selecionada = aba_selecionada

# ======================
# RENDERIZAR POPUPS PENDENTES E VERIFICAR NOVOS REGISTROS
# ======================
renderizar_popups_pendentes()
verificar_e_exibir_popups()


# ==================================================================================================
# PRENSADOS
# ==================================================================================================
if aba_selecionada == 'PRENSADOS':
    with st.spinner("Carregando dados..."):
        df_base = carregar_dados_prensados()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados.")
        st.stop()

    df_base_calc = df_base.copy()
    colunas_numericas = ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'TRS 100%', 'REFUGADO']
    for col in colunas_numericas:
        if col in df_base_calc.columns:
            df_base_calc[col] = pd.to_numeric(df_base_calc[col], errors='coerce').fillna(0)

    if 'TRS 100%' in df_base_calc.columns:
        df_base_calc['TRS 1ª ESCOLHA (%)'] = df_base_calc.apply(
            lambda row: (row['APROVADO'] / row['TRS 100%'] * 100) if row['TRS 100%'] != 0 else 0, axis=1
        )
        df_base_calc['TRS FINAL (%)'] = df_base_calc.apply(
            lambda row: (row['EMBALADO'] / row['TRS 100%'] * 100) if row['TRS 100%'] != 0 else 0, axis=1
        )
    else:
        df_base_calc['TRS 1ª ESCOLHA (%)'] = 0
        df_base_calc['TRS FINAL (%)'] = 0

    melhores_trs_historico = {}
    if 'REFERÊNCIA' in df_base_calc.columns:
        for ref in df_base_calc['REFERÊNCIA'].unique():
            ref_df = df_base_calc[df_base_calc['REFERÊNCIA'] == ref]
            if not ref_df.empty:
                max_trs = ref_df['TRS FINAL (%)'].max()
                if max_trs > 0:
                    melhores_trs_historico[ref] = max_trs

    # Sidebar filtros PRENSADOS
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
        for col in ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'TRS 100%', 'REFUGADO']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        total_prod = int(df['PRODUZIDO'].sum())
        total_apro = int(df['APROVADO'].sum())
        total_embal = int(df['EMBALADO'].sum()) if 'EMBALADO' in df.columns else 0
        total_meta = int(df['TRS 100%'].sum()) if 'TRS 100%' in df.columns else 0
        
        trs_primeira_escolha = (total_apro / total_meta * 100) if total_meta else 0
        trs_final_total = (total_embal / total_meta * 100) if total_meta else 0
    else:
        total_prod = total_apro = total_embal = total_meta = trs_primeira_escolha = trs_final_total = 0

    # Page header
    render_page_header("PRENSADOS", f"Industrial · {len(df):,} registros carregados · Atualizado {get_horario_brasilia()}", THEME['accent_cyan'])

    # KPIs (6 cards)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: render_kpi_card("Produzido", f"{total_prod:,}".replace(",","."), THEME['accent_cyan'], "◈")
    with c2: render_kpi_card("Aprovado", f"{total_apro:,}".replace(",","."), THEME['accent_lime'], "◈")
    with c3: render_kpi_card("Meta Líquida", f"{total_meta:,}".replace(",","."), THEME['accent_purple'], "◈")
    with c4: render_kpi_card("Embalado", f"{total_embal:,}".replace(",","."), THEME['accent_yellow'], "◈")
    with c5:
        trs_primeira_cor = THEME['accent_lime'] if trs_primeira_escolha >= 85 else THEME['accent_orange'] if trs_primeira_escolha >= 70 else THEME['accent_red']
        render_kpi_card("TRS 1ª Escolha", f"{trs_primeira_escolha:.1f}%", trs_primeira_cor, "◎")
    with c6:
        trs_final_cor = THEME['accent_lime'] if trs_final_total >= 85 else THEME['accent_orange'] if trs_final_total >= 70 else THEME['accent_red']
        render_kpi_card("TRS Final", f"{trs_final_total:.1f}%", trs_final_cor, "◎")

    # Tabela de produção
    render_section_header("Tabela de Produção", "▸")

    if not df.empty:
        if 'TRS 100%' in df.columns:
            df['TRS 1ª ESCOLHA (%)'] = df.apply(lambda r: (r['APROVADO'] / r['TRS 100%'] * 100) if r['TRS 100%'] != 0 else 0, axis=1)
            df['TRS FINAL (%)'] = df.apply(lambda r: (r['EMBALADO'] / r['TRS 100%'] * 100) if r['TRS 100%'] != 0 else 0, axis=1)
        df['TRS 1ª ESCOLHA (%)'] = df['TRS 1ª ESCOLHA (%)'].round(2)
        df['TRS FINAL (%)'] = df['TRS FINAL (%)'].round(2)

    df_sorted = df.sort_values(by="DATA", ascending=False).reset_index(drop=True)

    if filtro_melhores_trs and not df_sorted.empty and 'REFERÊNCIA' in df_sorted.columns:
        df_sorted = df_sorted[df_sorted.apply(lambda row: row['REFERÊNCIA'] in melhores_trs_historico and abs(row['TRS FINAL (%)'] - melhores_trs_historico[row['REFERÊNCIA']]) < 0.01, axis=1)].reset_index(drop=True)
        if not df_sorted.empty:
            st.info(f"Exibindo {len(df_sorted)} registro(s) — Melhor TRS Final Histórico por referência")
        else:
            st.warning("Nenhum registro encontrado com Melhor TRS Final Histórico")

    df_view = df_sorted if qtd == 0 else df_sorted.head(qtd)

    if not df_view.empty:
        df_display = df_view.copy()
        df_display['DATA'] = pd.to_datetime(df_display['DATA']).dt.strftime('%d/%m/%Y')

        for col in ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'REFUGADO', 'TRS 100%']:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: int(round(x)) if pd.notnull(x) else 0)
                df_display[col] = df_display[col].apply(lambda x: f"{x:,}".replace(",", "."))
        
        if 'TRS 1ª ESCOLHA (%)' in df_display.columns:
            df_display['TRS 1ª ESCOLHA (%)'] = df_display['TRS 1ª ESCOLHA (%)'].apply(lambda x: f"{x:.2f}%")
        if 'TRS FINAL (%)' in df_display.columns:
            df_display['TRS FINAL (%)'] = df_display['TRS FINAL (%)'].apply(lambda x: f"{x:.2f}%")

        colunas_exibir = ['DATA', 'REFERÊNCIA', 'TURNO', 'PRODUZIDO', 'APROVADO', 'TRS 100%', 'EMBALADO', 'REFUGADO', 'TRS 1ª ESCOLHA (%)', 'TRS FINAL (%)']
        if 'ANALISE' in df_display.columns:
            colunas_exibir.append('ANALISE')
        colunas_exibir = [col for col in colunas_exibir if col in df_display.columns]

        st.dataframe(df_display[colunas_exibir], use_container_width=True, height=400)

        if not filtro_melhores_trs:
            st.caption("▸ Dourado: Melhor TRS Final Histórico por referência   ▸ Verde: Análise registrada")

    # Gráfico TRS Diário
    render_section_header("Evolução Diária do TRS", "▸")

    if not df.empty and 'TRS 100%' in df.columns:
        colunas_agg = {}
        for col in ['PRODUZIDO', 'APROVADO', 'EMBALADO', 'TRS 100%']:
            if col in df.columns:
                colunas_agg[col] = 'sum'
        
        if colunas_agg:
            resumo_dia = df.groupby(df['DATA'].dt.date).agg(colunas_agg).reset_index()
            resumo_dia['DATA'] = pd.to_datetime(resumo_dia['DATA'])
            
            if 'APROVADO' in resumo_dia.columns and 'TRS 100%' in resumo_dia.columns:
                resumo_dia['TRS 1ª ESCOLHA (%)'] = (resumo_dia['APROVADO'] / resumo_dia['TRS 100%'].replace(0, 1) * 100).fillna(0)
            
            if 'EMBALADO' in resumo_dia.columns and 'TRS 100%' in resumo_dia.columns:
                resumo_dia['TRS FINAL (%)'] = (resumo_dia['EMBALADO'] / resumo_dia['TRS 100%'].replace(0, 1) * 100).fillna(0)
            
            resumo_dia = resumo_dia.sort_values('DATA')

            if not resumo_dia.empty and ('TRS 1ª ESCOLHA (%)' in resumo_dia.columns or 'TRS FINAL (%)' in resumo_dia.columns):
                fig, ax = plt.subplots(figsize=(14, 5), facecolor=THEME['bg_card'])
                apply_chart_style(ax, fig, "TRS Diário — Período Selecionado", ylabel="TRS (%)")

                if 'TRS 1ª ESCOLHA (%)' in resumo_dia.columns:
                    ax.plot(resumo_dia['DATA'], resumo_dia['TRS 1ª ESCOLHA (%)'],
                            marker='o', markersize=6, linewidth=2.5,
                            color=THEME['accent_cyan'], alpha=0.95, label='TRS 1ª Escolha',
                            markerfacecolor=THEME['bg_card'], markeredgecolor=THEME['accent_cyan'], markeredgewidth=2)

                if 'TRS FINAL (%)' in resumo_dia.columns:
                    ax.plot(resumo_dia['DATA'], resumo_dia['TRS FINAL (%)'],
                            marker='s', markersize=6, linewidth=2.5,
                            color=THEME['accent_orange'], alpha=0.95, label='TRS Final',
                            markerfacecolor=THEME['bg_card'], markeredgecolor=THEME['accent_orange'], markeredgewidth=2)

                ax.axhline(y=85, color=THEME['accent_red'], linestyle=':', alpha=0.7, linewidth=1.5, label='Meta 85%')

                ax.legend(framealpha=0.15, facecolor=THEME['bg_card'], edgecolor=THEME['border_bright'],
                          labelcolor=THEME['text_primary'], fontsize=9)
                plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha='right', fontsize=8,
                         color=THEME['text_muted'])
                fig.tight_layout(pad=1.5)
                st.pyplot(fig)
                plt.close(fig)

    st.markdown("<hr>", unsafe_allow_html=True)

    # Manual vs Automática
    render_section_header("Desempenho por Tipo de Prensa", "▸")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"""<div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.2em;
            text-transform:uppercase;color:{THEME['text_muted']};margin-bottom:8px">▸ Semi-Automática (Manual)</div>""", unsafe_allow_html=True)
        if 'BOQUETA' in df.columns:
            df_manual = df[df['BOQUETA'] == 1]
            if not df_manual.empty:
                t_ap_m = df_manual['APROVADO'].sum()
                t_emb_m = df_manual['EMBALADO'].sum()
                t_mt_m = df_manual['TRS 100%'].sum() if 'TRS 100%' in df_manual.columns else 1
                trs1_m = (t_ap_m / t_mt_m * 100) if t_mt_m > 0 else 0
                trs_final_m = (t_emb_m / t_mt_m * 100) if t_mt_m > 0 else 0
                prod_m = df_manual['PRODUZIDO'].sum()
                
                trs_color_m = THEME['accent_lime'] if trs1_m >= 85 else THEME['accent_orange'] if trs1_m >= 70 else THEME['accent_red']
                render_kpi_card("TRS 1ª Escolha — Manual", f"{trs1_m:.1f}%", trs_color_m)
                st.caption(f"TRS Final: {trs_final_m:.1f}% | Produção: {prod_m:,.0f} un".replace(",","."))
            else:
                st.info("Sem dados para Prensa Manual")

    with col2:
        st.markdown(f"""<div style="font-family:'JetBrains Mono',monospace;font-size:10px;letter-spacing:.2em;
            text-transform:uppercase;color:{THEME['text_muted']};margin-bottom:8px">▸ Automática</div>""", unsafe_allow_html=True)
        if 'BOQUETA' in df.columns:
            df_auto = df[df['BOQUETA'] == 2]
            if not df_auto.empty:
                t_ap_a = df_auto['APROVADO'].sum()
                t_emb_a = df_auto['EMBALADO'].sum()
                t_mt_a = df_auto['TRS 100%'].sum() if 'TRS 100%' in df_auto.columns else 1
                trs1_a = (t_ap_a / t_mt_a * 100) if t_mt_a > 0 else 0
                trs_final_a = (t_emb_a / t_mt_a * 100) if t_mt_a > 0 else 0
                prod_a = df_auto['PRODUZIDO'].sum()
                
                trs_color_a = THEME['accent_lime'] if trs1_a >= 85 else THEME['accent_orange'] if trs1_a >= 70 else THEME['accent_red']
                render_kpi_card("TRS 1ª Escolha — Automática", f"{trs1_a:.1f}%", trs_color_a)
                st.caption(f"TRS Final: {trs_final_a:.1f}% | Produção: {prod_a:,.0f} un".replace(",","."))
            else:
                st.info("Sem dados para Prensa Automática")

    st.markdown("<hr>", unsafe_allow_html=True)

    # Análise de Paradas
    render_section_header("Análise de Paradas", "▸")

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

    p1, p2, p3, p4 = st.columns(4)
    with p1: render_kpi_card("Horas Trabalhadas", minutos_para_horas_str(horas_trabalhadas), THEME['accent_cyan'])
    with p2: render_kpi_card("Acertos", minutos_para_horas_str(total_acertos), THEME['accent_yellow'])
    with p3: render_kpi_card("Manutenção", minutos_para_horas_str(total_manut), THEME['accent_red'])
    with p4: render_kpi_card("Horas Produtivas", minutos_para_horas_str(horas_produtivas), THEME['accent_lime'])

    col1, col2 = st.columns(2)

    with col1:
        if 'BOQUETA' in df.columns:
            df_manual_p = df[df['BOQUETA'] == 1]
            df_auto_p = df[df['BOQUETA'] == 2]
            acertos_m = df_manual_p['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
            manut_m = df_manual_p['MANUT_MIN'].sum() if 'MANUT_MIN' in df.columns else 0
            acertos_a = df_auto_p['ACERTOS_MIN_AJUSTADO'].sum() if 'ACERTOS_MIN_AJUSTADO' in df.columns else 0
            manut_a = df_auto_p['MANUT_MIN'].sum() if 'MANUT_MIN' in df.columns else 0

            categorias = ['Manual', 'Automática']
            acertos_v = [acertos_m, acertos_a]
            manut_v = [manut_m, manut_a]

            fig, ax = plt.subplots(figsize=(7, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "Composição das Paradas", ylabel="Minutos")

            x = np.arange(len(categorias))
            w = 0.55
            ax.bar(x, acertos_v, w, label='Acertos', color=THEME['accent_yellow'], alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5)
            ax.bar(x, manut_v, w, bottom=acertos_v, label='Manutenção', color=THEME['accent_red'], alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5)

            for i, (a, m) in enumerate(zip(acertos_v, manut_v)):
                total = a + m
                if a > 0: ax.text(i, a/2, minutos_para_horas_str(a), ha='center', va='center', color='black', fontweight='bold', fontsize=10)
                if m > 0: ax.text(i, a + m/2, minutos_para_horas_str(m), ha='center', va='center', color='black', fontweight='bold', fontsize=10)
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
            vals_p = [horas_produtivas, total_acertos, total_manut]
            cores_p = [THEME['accent_lime'], THEME['accent_yellow'], THEME['accent_red']]
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
                    at.set_color('black')
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

    # TRS por Turno
    if not df.empty and 'TURNO' in df.columns and 'TRS 100%' in df.columns:
        render_section_header("TRS por Turno", "▸")
        turno_data = []
        for t in df['TURNO'].unique():
            df_t = df[df['TURNO'] == t]
            te = df_t['EMBALADO'].sum() if 'EMBALADO' in df_t.columns else 0
            ta = df_t['APROVADO'].sum()
            tm = df_t['TRS 100%'].sum()
            turno_data.append({
                'Turno': t, 
                'TRS 1ª Escolha': (ta/tm*100) if tm > 0 else 0,
                'TRS Final': (te/tm*100) if tm > 0 else 0
            })
        df_tt = pd.DataFrame(turno_data)
        if not df_tt.empty:
            fig, ax = plt.subplots(figsize=(10, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "TRS por Turno", ylabel="TRS (%)")
            
            x = np.arange(len(df_tt))
            width = 0.35
            
            bars1 = ax.bar(x - width/2, df_tt['TRS 1ª Escolha'], width, label='TRS 1ª Escolha', color=THEME['accent_cyan'], alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5)
            bars2 = ax.bar(x + width/2, df_tt['TRS Final'], width, label='TRS Final', color=THEME['accent_orange'], alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5)
            
            ax.axhline(y=85, color=THEME['accent_red'], linestyle='--', alpha=0.5, linewidth=1.5, label='Meta 85%')
            
            for bar in bars1:
                h = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2, h + 1, f'{h:.1f}%', ha='center', va='bottom', fontsize=8, color=THEME['accent_cyan'])
            
            for bar in bars2:
                h = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2, h + 1, f'{h:.1f}%', ha='center', va='bottom', fontsize=8, color=THEME['accent_orange'])
            
            ax.set_xticks(x)
            ax.set_xticklabels(df_tt['Turno'], fontsize=11)
            ax.legend(loc='upper right', fontsize=9)
            ax.set_ylim(0, 105)
            
            fig.tight_layout(pad=1.5)
            st.pyplot(fig)
            plt.close(fig)

    # Defeitos de Prensados
    if mostrar_defeitos:
        render_section_header("Estratificação de Defeitos - Prensados", "▸")
        colunas_defeitos_prensados = [
            'TRINCA', 'RUGAS', 'DOBRA', 'SUJEIRA', 'FALHAS', 'CHUPADO',
            'CROMO', 'BARRO', 'EMPENO', 'OUTROS', 'REMANEJAMENTO'
        ]
        defeitos_existentes = []
        for defeito in colunas_defeitos_prensados:
            for col in df.columns:
                if col.upper() == defeito.upper():
                    defeitos_existentes.append(col)
                    break
        
        if not defeitos_existentes:
            for col in df.columns:
                col_upper = col.upper()
                if col_upper in ['TRINCA', 'RUGAS', 'DOBRA', 'SUJEIRA', 'FALHAS', 'CHUPADO', 'CROMO', 'BARRO', 'EMPENO', 'OUTROS']:
                    defeitos_existentes.append(col)
        
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
            else:
                st.info("Nenhum defeito registrado no período selecionado")
        else:
            st.info("Colunas de defeitos não encontradas na planilha de Prensados")

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
    with st.spinner("Carregando dados..."):
        df_base = carregar_dados_sopro()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados.")
        st.stop()

    df_base_calc = df_base.copy()
    if 'TRS_BRUTO' in df_base_calc.columns:
        df_base_calc['TRS LÍQUIDO (%)'] = (df_base_calc['TRS_BRUTO'] * 100).round(2)
        df_base_calc['META'] = df_base_calc.apply(
            lambda row: (row['APROVADO'] / (row['TRS LÍQUIDO (%)'] / 100)) if row['TRS LÍQUIDO (%)'] > 0 else row['APROVADO'], axis=1
        ).round(0)
    else:
        df_base_calc['TRS LÍQUIDO (%)'] = 0
        df_base_calc['META'] = 0

    melhores_trs_historico = {}
    if 'REFERÊNCIA' in df_base_calc.columns:
        for ref in df_base_calc['REFERÊNCIA'].unique():
            ref_df = df_base_calc[df_base_calc['REFERÊNCIA'] == ref]
            if not ref_df.empty:
                mt = ref_df['TRS LÍQUIDO (%)'].max()
                if mt > 0:
                    melhores_trs_historico[ref] = mt

    # Sidebar filtros SOPRO
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

    # Filtros
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
    
    if not df.empty and 'TRS_BRUTO' in df.columns and 'APROVADO' in df.columns:
        df['TRS_LIQUIDO_PCT'] = df['TRS_BRUTO'] * 100
        df['META_CALC'] = df.apply(
            lambda row: (row['APROVADO'] / (row['TRS_LIQUIDO_PCT'] / 100)) if row['TRS_LIQUIDO_PCT'] > 0 else row['APROVADO'], axis=1
        )
        total_meta = int(df['META_CALC'].sum())
    else:
        total_meta = 0

    # Page header
    render_page_header("SOPRO", f"Industrial · {len(df):,} registros carregados · Atualizado {get_horario_brasilia()}", THEME['accent_lime'])

    # KPIs (5 cards)
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1: render_kpi_card("Produzido", f"{total_prod:,}".replace(",","."), THEME['accent_cyan'], "◈")
    with c2: render_kpi_card("Aprovado", f"{total_apro:,}".replace(",","."), THEME['accent_lime'], "◈")
    with c3: render_kpi_card("Meta (TRS 100%)", f"{total_meta:,}".replace(",","."), THEME['accent_purple'], "◈")
    with c4: render_kpi_card("Refugo", f"{total_refugo:,}".replace(",","."), THEME['accent_orange'], "◈")
    with c5:
        trs_c = THEME['accent_lime'] if trs_liq_med >= 85 else THEME['accent_orange'] if trs_liq_med >= 70 else THEME['accent_red']
        render_kpi_card("TRS Líquido Médio", f"{trs_liq_med:.1f}%", trs_c, "◎")

    # Tabela
    render_section_header("Tabela de Produção", "▸", THEME['accent_lime'])

    if not df.empty and 'TRS_BRUTO' in df.columns:
        df['TRS LÍQUIDO (%)'] = (df['TRS_BRUTO'] * 100).round(2)
        df['META'] = df.apply(
            lambda row: int(round(row['APROVADO'] / (row['TRS LÍQUIDO (%)'] / 100))) if row['TRS LÍQUIDO (%)'] > 0 else int(row['APROVADO']), axis=1
        )

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
        
        for col in ['PRODUZIDO', 'APROVADO', 'REFUGADO', 'META']:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: int(round(x)) if pd.notnull(x) else 0)
                df_display[col] = df_display[col].apply(lambda x: f"{x:,}".replace(",", "."))
        
        if 'TRS LÍQUIDO (%)' in df_display.columns:
            df_display['TRS LÍQUIDO (%)'] = df_display['TRS LÍQUIDO (%)'].apply(lambda x: f"{x:.2f}%")

        colunas_exibir = ['DATA','TURNO','PRAÇA','REFERÊNCIA','PRODUZIDO','META','APROVADO','REFUGADO','TRS LÍQUIDO (%)']
        colunas_exibir = [c for c in colunas_exibir if c in df_display.columns]

        st.dataframe(df_display[colunas_exibir], use_container_width=True, height=400)
        if not filtro_melhores_trs:
            st.caption("▸ Dourado: Melhor TRS Líquido Histórico por referência")
    else:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")

    # TRS Líquido Diário
    render_section_header("Evolução Diária do TRS Líquido", "▸", THEME['accent_lime'])
    if not df.empty and 'TRS_BRUTO' in df.columns:
        res_dia = df.groupby(df['DATA'].dt.date).agg({
            'TRS_BRUTO': 'mean', 
            'PRODUZIDO': 'sum', 
            'APROVADO': 'sum'
        }).reset_index()
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

    # Por Praça
    if 'PRAÇA' in df.columns and not df.empty and 'TRS_BRUTO' in df.columns:
        render_section_header("TRS Líquido por Praça", "▸", THEME['accent_lime'])
        res_praca = df.groupby('PRAÇA').agg({
            'TRS_BRUTO': 'mean', 
            'PRODUZIDO': 'sum',
            'APROVADO': 'sum'
        }).reset_index()
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

    # Mensal
    render_section_header("TRS Líquido Mensal", "▸", THEME['accent_lime'])
    if not df.empty and 'ANO_MES' in df.columns and 'TRS_BRUTO' in df.columns:
        res_mes = df.groupby('ANO_MES').agg({
            'TRS_BRUTO': 'mean', 
            'PRODUZIDO': 'sum', 
            'APROVADO': 'sum'
        }).reset_index()
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

    # Por Turno
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

    # Defeitos
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


# ==================================================================================================
# TÊMPERA
# ==================================================================================================
elif aba_selecionada == 'TÊMPERA':
    ABA = 'TRS_TEMPERA'

    # Mapeamento dos códigos de defeito
    MAPEAMENTO_DEFEITOS = {
        1: 'Estourou após furar',
        2: 'Quebra no resfriamento',
        3: 'Quebra teste impacto',
        4: 'Furada e não fraturou',
        5: 'Quebra de quarentena',
        6: 'Ovalizada'
    }
    
    CODIGOS_DEFEITO_REAIS = [2, 3, 4, 5, 6]

    def safe_float(val):
        """Converte valor para float de forma segura"""
        if val is None or pd.isna(val):
            return 0.0
        try:
            val_str = str(val).strip()
            if val_str == '' or val_str == 'nan':
                return 0.0
            val_str = val_str.replace(',', '.')
            return float(val_str)
        except:
            return 0.0

    @st.cache_data(ttl=300)
    def carregar_dados_tempera():
        try:
            client = get_gspread_client()
            if client is None:
                return pd.DataFrame()
            sheet = client.open_by_key(ID_PLANILHA_TEMPERA).worksheet(ABA)
            todos_dados = sheet.get_all_values()
            
            if len(todos_dados) < 2:
                return pd.DataFrame()
            
            cabecalho = todos_dados[0]
            valores = todos_dados[1:]
            df = pd.DataFrame(valores, columns=cabecalho)
            colunas = list(df.columns)
            
            # Mapear colunas por posição (baseado nos nomes reais)
            if len(colunas) >= 5:
                df = df.rename(columns={
                    colunas[0]: 'PRODUCAO',
                    colunas[1]: 'DATA_TEMP',
                    colunas[2]: 'TURNO_TEMP',
                    colunas[3]: 'PRODUTO',
                    colunas[4]: 'GANCHEIRA'
                })
            
            if len(colunas) >= 8:
                df = df.rename(columns={
                    colunas[5]: 'SUPERIOR',
                    colunas[6]: 'MEIO',
                    colunas[7]: 'INFERIOR'
                })
            
            if len(colunas) >= 11:
                df = df.rename(columns={
                    colunas[8]: 'A1',
                    colunas[9]: 'C1',
                    colunas[10]: 'A2'
                })
            
            if len(colunas) >= 14:
                df = df.rename(columns={
                    colunas[11]: 'C2',
                    colunas[12]: 'A3',
                    colunas[13]: 'C3'
                })
            
            if len(colunas) >= 17:
                df = df.rename(columns={
                    colunas[14]: 'A4',
                    colunas[15]: 'C4',      # HUMIDADE
                    colunas[16]: 'A5'
                })
            
            if len(colunas) >= 20:
                df = df.rename(columns={
                    colunas[17]: 'C5',
                    colunas[18]: 'A e B'    # PRESSÃO AR
                })
            
            # Converter datas
            if 'DATA_TEMP' in df.columns:
                df['DATA'] = df['DATA_TEMP'].apply(converter_data_br)
            elif 'PRODUCAO' in df.columns:
                df['DATA'] = df['PRODUCAO'].apply(converter_data_br)
            
            if 'DATA' in df.columns:
                df = df.dropna(subset=['DATA'])
            
            # Converter colunas numéricas
            colunas_numericas = ['SUPERIOR', 'MEIO', 'INFERIOR', 'A1', 'C1', 'A2', 'C2', 'A3', 'C3', 'A4', 'C4', 'A5', 'C5', 'A e B']
            
            for col in colunas_numericas:
                if col in df.columns:
                    df[col] = df[col].apply(safe_float)
            
            # CORREÇÃO: Tempo C2 - converter para segundos
            if 'C2' in df.columns:
                def converter_tempo_c2(val):
                    if pd.isna(val) or val == 0:
                        return 0
                    if val <= 1:
                        return val * 100
                    elif val <= 10:
                        return val * 10
                    else:
                        return val
                df['C2'] = df['C2'].apply(converter_tempo_c2)
            
            # Identificar colunas de posições (19 a 70)
            colunas_posicoes_validas = []
            for col in df.columns:
                try:
                    num = int(str(col).strip())
                    if 19 <= num <= 70:
                        colunas_posicoes_validas.append(col)
                except:
                    pass
            
            # Inicializar colunas
            df['TOTAL_PECAS'] = 40
            df['APROVADO'] = 40
            df['TOTAL_DEFEITOS'] = 0
            df['IS_CRITICO'] = False
            
            for codigo, nome in MAPEAMENTO_DEFEITOS.items():
                nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
                df[f'QTD_{nome_clean}'] = 0
            
            for idx, row in df.iterrows():
                defeitos_contagem = {codigo: 0 for codigo in MAPEAMENTO_DEFEITOS.keys()}
                
                for col in colunas_posicoes_validas:
                    try:
                        val = row[col]
                        if pd.notna(val) and str(val).strip():
                            codigo = int(float(str(val).strip()))
                            if codigo in MAPEAMENTO_DEFEITOS:
                                defeitos_contagem[codigo] += 1
                    except:
                        pass
                
                total_defeitos_reais = sum(defeitos_contagem.get(cod, 0) for cod in CODIGOS_DEFEITO_REAIS)
                aprovadas = 40 - total_defeitos_reais
                
                df.at[idx, 'APROVADO'] = aprovadas
                df.at[idx, 'TOTAL_DEFEITOS'] = total_defeitos_reais
                df.at[idx, 'TRS (%)'] = (aprovadas / 40 * 100) if 40 > 0 else 0
                
                is_critico = False
                if defeitos_contagem.get(4, 0) >= 1:
                    is_critico = True
                if defeitos_contagem.get(3, 0) > 2:
                    is_critico = True
                df.at[idx, 'IS_CRITICO'] = is_critico
                
                for codigo, nome in MAPEAMENTO_DEFEITOS.items():
                    nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
                    col_nome = f'QTD_{nome_clean}'
                    if col_nome in df.columns:
                        df.at[idx, col_nome] = defeitos_contagem.get(codigo, 0)
            
            return df
            
        except Exception as e:
            st.error(f"Erro ao carregar dados: {str(e)}")
            return pd.DataFrame()

    with st.spinner("Carregando dados da Têmpera..."):
        df_base = carregar_dados_tempera()

    if df_base.empty:
        st.warning("Não foi possível carregar os dados da Têmpera.")
        st.stop()

    # ── Sidebar filtros ──
    with st.sidebar:
        st.markdown(f"<div style='font-family:JetBrains Mono,monospace;font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:{THEME['accent_purple']};margin:20px 0 10px;border-top:1px solid {THEME['border_bright']};padding-top:16px'>▸ Filtros · Têmpera</div>", unsafe_allow_html=True)
        
        data_ini = st.date_input("Data inicial", value=None, key="tempera_data_ini")
        data_fim = st.date_input("Data final", value=None, key="tempera_data_fim")
        
        if 'TURNO_TEMP' in df_base.columns:
            turnos_disp = ["(Todos)"] + sorted([str(t) for t in df_base['TURNO_TEMP'].dropna().unique()])
            turno = st.selectbox("Turno", options=turnos_disp, key="tempera_turno")
        else:
            turno = "(Todos)"
        
        if 'PRODUTO' in df_base.columns:
            produtos_disp = ["(Todos)"] + sorted([str(p) for p in df_base['PRODUTO'].dropna().unique()])
            produto = st.selectbox("Produto", options=produtos_disp, key="tempera_produto")
        else:
            produto = "(Todos)"
        
        excluir_criticos = st.checkbox("Excluir registros críticos", value=False, key="tempera_excluir_criticos")
        qtd = st.number_input("Linhas na tabela", min_value=0, max_value=5000, value=20, step=10, key="tempera_qtd")

    # ── Aplicar filtros ──
    df = df_base.copy()
    
    if data_ini:
        df = df[df['DATA'] >= pd.to_datetime(data_ini)]
    if data_fim:
        df = df[df['DATA'] <= pd.to_datetime(data_fim)]
    if turno != "(Todos)" and 'TURNO_TEMP' in df.columns:
        df = df[df['TURNO_TEMP'].astype(str).str.upper() == turno.upper()]
    if produto != "(Todos)" and 'PRODUTO' in df.columns:
        df = df[df['PRODUTO'].astype(str) == produto]
    
    if excluir_criticos and 'IS_CRITICO' in df.columns:
        df = df[~df['IS_CRITICO']].copy()

    if df.empty:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")
        st.stop()

    # ── Função para média ignorando zeros ──
    def safe_mean(col):
        """Calcula média ignorando valores zero e NaN"""
        if col in df.columns:
            valores = df[col][(df[col] > 0) & (pd.notna(df[col]))]
            if len(valores) > 0:
                return valores.mean()
        return 0

    # ── KPIs (médias ignorando zeros) ──
    total_registros = len(df)
    total_pecas = total_registros * 40
    total_aprovado = int(df['APROVADO'].sum())
    total_defeitos = int(df['TOTAL_DEFEITOS'].sum())
    trs_medio = (total_aprovado / total_pecas * 100) if total_pecas > 0 else 0
    
    temp_sup = safe_mean('SUPERIOR')
    temp_meio = safe_mean('MEIO')
    temp_inf = safe_mean('INFERIOR')
    temp_entrada = safe_mean('A1')
    tempo_c2 = safe_mean('C2')
    humidade = safe_mean('C4')
    pressao_ar = safe_mean('A e B')

    # ── Page header ──
    render_page_header(
        "TÊMPERA",
        f"Industrial · {total_registros:,} registros · Atualizado {get_horario_brasilia()}",
        THEME['accent_purple']
    )

    # KPIs principais
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_kpi_card("Total Peças", f"{total_pecas:,}".replace(",","."), THEME['accent_cyan'])
    with c2:
        render_kpi_card("Aprovadas", f"{total_aprovado:,}".replace(",","."), THEME['accent_lime'])
    with c3:
        render_kpi_card("Defeitos", f"{total_defeitos:,}".replace(",","."), THEME['accent_red'])
    with c4:
        trs_color = THEME['accent_lime'] if trs_medio >= 80 else THEME['accent_orange'] if trs_medio >= 70 else THEME['accent_red']
        render_kpi_card("TRS Médio", f"{trs_medio:.1f}%", trs_color)

    # Temperaturas Forno (médias ignorando zeros)
    c1, c2, c3 = st.columns(3)
    with c1:
        render_kpi_card("Temp. Superior", f"{temp_sup:.0f}°C" if temp_sup > 0 else "N/A", THEME['accent_orange'])
    with c2:
        render_kpi_card("Temp. Meio", f"{temp_meio:.0f}°C" if temp_meio > 0 else "N/A", THEME['accent_orange'])
    with c3:
        render_kpi_card("Temp. Inferior", f"{temp_inf:.0f}°C" if temp_inf > 0 else "N/A", THEME['accent_orange'])

    # Processo - Médias (ignorando zeros)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-family:JetBrains Mono,monospace;font-size:10px;color:{THEME['accent_purple']}';>▸ PROCESSO - MÉDIAS DO PERÍODO (ignorando zeros)</div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_kpi_card("Temp. Entrada (A1)", f"{temp_entrada:.0f}°C" if temp_entrada > 0 else "N/A", THEME['accent_cyan'])
    with c2:
        render_kpi_card("Tempo (C2)", f"{tempo_c2:.0f}s" if tempo_c2 > 0 else "N/A", THEME['accent_lime'])
    with c3:
        render_kpi_card("Humidade (C4)", f"{humidade:.1f}%" if humidade > 0 else "N/A", THEME['accent_orange'])
    with c4:
        render_kpi_card("Pressão Ar (A e B)", f"{pressao_ar:.1f}" if pressao_ar > 0 else "N/A", THEME['accent_purple'])

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Tabela ──
    render_section_header("Registros de Têmpera", "▸", THEME['accent_purple'])
    
    df_display = df.sort_values(by="DATA", ascending=False).head(qtd if qtd > 0 else 100).copy()
    df_display['DATA'] = pd.to_datetime(df_display['DATA']).dt.strftime('%d/%m/%Y')
    df_display['TRS (%)'] = df_display['TRS (%)'].round(1).astype(str) + '%'
    
    colunas = ['DATA', 'TURNO_TEMP', 'PRODUTO', 'GANCHEIRA', 'APROVADO', 'TOTAL_DEFEITOS', 'TRS (%)']
    colunas = [c for c in colunas if c in df_display.columns]
    
    st.dataframe(df_display[colunas], use_container_width=True, height=400)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Gráfico TRS Diário ──
    render_section_header("Evolução Diária do TRS", "▸", THEME['accent_purple'])
    
    resumo_dia = df.groupby(df['DATA'].dt.date).agg({'APROVADO': 'sum'}).reset_index()
    resumo_dia['DATA'] = pd.to_datetime(resumo_dia['DATA'])
    counts = df.groupby(df['DATA'].dt.date).size().values
    resumo_dia['TRS (%)'] = (resumo_dia['APROVADO'] / (counts * 40) * 100)
    resumo_dia = resumo_dia.sort_values('DATA')
    
    if not resumo_dia.empty:
        fig, ax = plt.subplots(figsize=(12, 4), facecolor=THEME['bg_card'])
        apply_chart_style(ax, fig, "TRS Diário", ylabel="TRS (%)", accent=THEME['accent_purple'])
        ax.fill_between(resumo_dia['DATA'], 0, resumo_dia['TRS (%)'], alpha=0.12, color=THEME['accent_purple'])
        ax.plot(resumo_dia['DATA'], resumo_dia['TRS (%)'], marker='o', markersize=5, linewidth=2, color=THEME['accent_purple'])
        ax.axhline(y=80, color=THEME['accent_red'], linestyle=':', linewidth=1.5, label='Meta 80%')
        ax.legend(loc='upper right', fontsize=9)
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha='right', fontsize=8)
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # ── PADRÃO DE EXCELÊNCIA (Média das TOP 15) ──
    st.markdown("<hr>", unsafe_allow_html=True)
    render_section_header("🏆 PADRÃO DE EXCELÊNCIA (Média Top 15)", "▸", THEME['accent_purple'])
    
    TOP_N = 15
    
    # Filtra produções válidas (com dados de pressão, se disponível)
    if 'A e B' in df.columns:
        df_validos = df[df['A e B'] > 0].copy()
    else:
        df_validos = df.copy()
    
    if len(df_validos) >= TOP_N:
        # Seleciona as TOP N produções com base no maior número de peças aprovadas
        df_top = df_validos.nlargest(TOP_N, ['APROVADO', 'TRS (%)'])
        
        # Calcula as médias
        media_aprovadas = df_top['APROVADO'].mean()
        media_trs = df_top['TRS (%)'].mean()
        media_defeitos = df_top['TOTAL_DEFEITOS'].mean()
        
        # Médias dos parâmetros de processo
        media_temp_sup = df_top['SUPERIOR'].mean()
        media_temp_meio = df_top['MEIO'].mean()
        media_temp_inf = df_top['INFERIOR'].mean()
        media_temp_entrada = df_top['A1'].mean()
        media_tempo_c2 = df_top['C2'].mean()
        media_humidade = df_top['C4'].mean()
        media_pressao_ar = df_top['A e B'].mean()
        
        # Desvio padrão para noção de estabilidade
        std_trs = df_top['TRS (%)'].std()
        criticas_top = df_top['IS_CRITICO'].sum()
        
        st.markdown(f"""
        <div style="background: {THEME['bg_card2']}; padding: 20px; border-radius: 10px; border-left: 4px solid {THEME['accent_lime']};">
            <h4 style="margin:0 0 5px 0; color:{THEME['accent_lime']};">🎯 Referência de Excelência</h4>
            <p style="margin:0 0 15px 0; font-size:12px; color:{THEME['text_muted']};">Média calculada com base nas {TOP_N} melhores produções do período</p>
        """, unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 TRS Médio (Top 15)", f"{media_trs:.1f}%", f"±{std_trs:.1f}%" if std_trs > 0 else "estável")
        with col2:
            st.metric("✅ Aprovadas (Média)", f"{media_aprovadas:.1f}/40")
        with col3:
            st.metric("❌ Defeitos (Média)", f"{media_defeitos:.1f}")
        with col4:
            st.metric("⚠️ Produções Críticas", f"{criticas_top}/{TOP_N}")
        
        st.markdown("<hr style='margin:15px 0; opacity:0.3;'>", unsafe_allow_html=True)
        
        st.markdown("#### 🔥 Temperaturas do Forno (Média)")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Superior", f"{media_temp_sup:.0f}°C" if pd.notna(media_temp_sup) else "N/A")
        with col2:
            st.metric("Meio", f"{media_temp_meio:.0f}°C" if pd.notna(media_temp_meio) else "N/A")
        with col3:
            st.metric("Inferior", f"{media_temp_inf:.0f}°C" if pd.notna(media_temp_inf) else "N/A")
        
        st.markdown("#### 📍 Principais Indicadores (Média)")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Temp. Entrada (A1)", f"{media_temp_entrada:.0f}°C" if pd.notna(media_temp_entrada) else "N/A")
        with col2:
            st.metric("Tempo (C2)", f"{media_tempo_c2:.0f}s" if pd.notna(media_tempo_c2) else "N/A")
        with col3:
            st.metric("Humidade (C4)", f"{media_humidade:.1f}%" if pd.notna(media_humidade) else "N/A")
        with col4:
            st.metric("Pressão Ar (A e B)", f"{media_pressao_ar:.1f}" if pd.notna(media_pressao_ar) else "N/A")
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Expander com detalhes das Top 15
        with st.expander(f"🔍 Ver detalhes das {TOP_N} melhores produções"):
            st.markdown(f"**As {TOP_N} produções que definem o padrão de excelência:**")
            df_top_display = df_top[['DATA', 'TURNO_TEMP', 'PRODUTO', 'GANCHEIRA', 'APROVADO', 'TRS (%)', 'TOTAL_DEFEITOS']].copy()
            df_top_display['DATA'] = pd.to_datetime(df_top_display['DATA']).dt.strftime('%d/%m/%Y')
            df_top_display['TRS (%)'] = df_top_display['TRS (%)'].round(1).astype(str) + '%'
            st.dataframe(df_top_display, use_container_width=True, height=300)
            
            st.markdown("**📊 Estatísticas das Top 15 vs. Média Geral:**")
            media_geral_trs = (df_validos['APROVADO'].sum() / (len(df_validos) * 40) * 100) if len(df_validos) > 0 else 0
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Média Geral TRS", f"{media_geral_trs:.1f}%")
            with col2:
                ganho = media_trs - media_geral_trs
                st.metric("Ganho Potencial", f"+{ganho:.1f}%", delta_color="normal")
                
    elif len(df_validos) > 0:
        st.warning(f"Dados insuficientes para calcular a média das {TOP_N} melhores produções. Apenas {len(df_validos)} registros encontrados. Exibindo a melhor produção individual:")
        
        # Fallback: mostra a melhor produção individual
        idx_melhor = df_validos['APROVADO'].idxmax()
        melhor = df_validos.loc[idx_melhor]
        
        st.markdown(f"""
        <div style="background: {THEME['bg_card2']}; padding: 20px; border-radius: 10px; border-left: 4px solid {THEME['accent_lime']};">
            <h4 style="margin:0 0 15px 0; color:{THEME['accent_lime']};">🎯 Melhor Resultado Individual</h4>
        """, unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            data_str = pd.to_datetime(melhor['DATA']).strftime('%d/%m/%Y') if pd.notna(melhor['DATA']) else 'N/A'
            st.metric("📅 Data", data_str)
        with col2:
            st.metric("⏰ Turno", str(melhor.get('TURNO_TEMP', 'N/A')))
        with col3:
            st.metric("📦 Produto", str(melhor.get('PRODUTO', 'N/A'))[:15])
        with col4:
            st.metric("🔧 Gancheira", str(melhor.get('GANCHEIRA', 'N/A')))
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("✅ Aprovadas", f"{int(melhor['APROVADO'])}/40")
        with col2:
            st.metric("📊 TRS", f"{melhor['TRS (%)']:.1f}%")
        with col3:
            st.metric("❌ Defeitos", int(melhor['TOTAL_DEFEITOS']))
        with col4:
            criticos = "Sim" if melhor.get('IS_CRITICO', False) else "Não"
            st.metric("⚠️ Crítico", criticos)
        
        st.markdown("<hr style='margin:15px 0; opacity:0.3;'>", unsafe_allow_html=True)
        
        st.markdown("#### 🔥 Temperaturas do Forno")
        col1, col2, col3 = st.columns(3)
        with col1:
            val = melhor.get('SUPERIOR', 0)
            st.metric("Superior", f"{val:.0f}°C" if pd.notna(val) and val > 0 else "N/A")
        with col2:
            val = melhor.get('MEIO', 0)
            st.metric("Meio", f"{val:.0f}°C" if pd.notna(val) and val > 0 else "N/A")
        with col3:
            val = melhor.get('INFERIOR', 0)
            st.metric("Inferior", f"{val:.0f}°C" if pd.notna(val) and val > 0 else "N/A")
        
        st.markdown("#### 📍 Principais Indicadores")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            val = melhor.get('A1', 0)
            st.metric("Temp. Entrada (A1)", f"{val:.0f}°C" if pd.notna(val) and val > 0 else "N/A")
        with col2:
            val = melhor.get('C2', 0)
            st.metric("Tempo (C2)", f"{val:.0f}s" if pd.notna(val) and val > 0 else "N/A")
        with col3:
            val = melhor.get('C4', 0)
            st.metric("Humidade (C4)", f"{val:.1f}%" if pd.notna(val) and val > 0 else "N/A")
        with col4:
            val = melhor.get('A e B', 0)
            st.metric("Pressão Ar (A e B)", f"{val:.1f}" if pd.notna(val) and val > 0 else "N/A")
        
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.info("Nenhum registro válido encontrado (Pressão Ar > 0).")

    # ── Ranking de Gancheiras (Pior → Melhor) ──
    st.markdown("<hr>", unsafe_allow_html=True)
    render_section_header("🏭 Ranking de Gancheiras (Pior → Melhor)", "▸", THEME['accent_purple'])
    
    if not df.empty and 'GANCHEIRA' in df.columns:
        ranking_gancheiras = []
        for gancheira in df['GANCHEIRA'].dropna().unique():
            df_g = df[df['GANCHEIRA'] == gancheira]
            total_registros_g = len(df_g)
            total_aprovado_g = int(df_g['APROVADO'].sum())
            total_defeitos_g = int(df_g['TOTAL_DEFEITOS'].sum())
            total_pecas_g = total_registros_g * 40
            trs_g = (total_aprovado_g / total_pecas_g * 100) if total_pecas_g > 0 else 0
            
            defeitos_gancheira = {}
            for codigo, nome in MAPEAMENTO_DEFEITOS.items():
                if codigo in CODIGOS_DEFEITO_REAIS:
                    nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
                    col = f'QTD_{nome_clean}'
                    if col in df_g.columns:
                        qtd = int(df_g[col].sum())
                        if qtd > 0:
                            defeitos_gancheira[nome] = qtd
            
            media_defeitos = total_defeitos_g / total_registros_g if total_registros_g > 0 else 0
            det_str = ' | '.join([f"{k}: {v}" for k, v in defeitos_gancheira.items()]) if defeitos_gancheira else "-"
            
            ranking_gancheiras.append({
                'Pos': 0,
                'Gancheira': str(gancheira),
                'Reg': total_registros_g,
                'Defeitos': total_defeitos_g,
                'Média': media_defeitos,
                'TRS_num': trs_g,
                'TRS': f"{trs_g:.1f}%",
                'Detalhamento': det_str
            })
        
        df_ranking = pd.DataFrame(ranking_gancheiras)
        df_ranking = df_ranking.sort_values('Defeitos', ascending=False)
        df_ranking['Pos'] = range(1, len(df_ranking) + 1)
        
        if not df_ranking.empty:
            piores = df_ranking.head(3)
            st.warning(f"⚠️ **Piores gancheiras:** {', '.join(piores['Gancheira'].tolist())}")
            
            df_tabela = df_ranking[['Pos', 'Gancheira', 'Reg', 'Defeitos', 'Média', 'TRS']].copy()
            df_tabela['Média'] = df_tabela['Média'].round(1)
            
            def estilo_ranking(row):
                styles = [''] * len(row)
                pos = row['Pos']
                if pos <= 3:
                    styles[0] = 'color: #E81123; font-weight: bold;'
                elif pos > len(df_ranking) - 3:
                    styles[0] = 'color: #107C10; font-weight: bold;'
                styles[3] = 'color: #E81123; font-weight: bold;'
                try:
                    trs_val = float(row['TRS'].replace('%', ''))
                    if trs_val >= 80:
                        styles[5] = 'color: #107C10; font-weight: bold;'
                    elif trs_val >= 70:
                        styles[5] = 'color: #E86C2C; font-weight: bold;'
                except:
                    pass
                return styles
            
            styled = df_tabela.style.apply(estilo_ranking, axis=1)
            st.dataframe(styled, use_container_width=True, height=250)
            
            with st.expander("🔍 Detalhamento das 5 Piores Gancheiras", expanded=False):
                for i, (_, row) in enumerate(df_ranking.head(5).iterrows(), 1):
                    st.markdown(f"""
                    **{i}º - Gancheira {row['Gancheira']}** | ❌ {row['Defeitos']} defeitos | TRS: {row['TRS']}  
                    📋 {row['Detalhamento']}
                    """)
            
            fig, ax = plt.subplots(figsize=(10, 3), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "Top 10 Piores Gancheiras", accent=THEME['accent_purple'])
            top10 = df_ranking.head(10)
            colors = [THEME['accent_red'] if i < 3 else THEME['accent_orange'] for i in range(len(top10))]
            ax.barh(range(len(top10)), top10['Defeitos'], color=colors, alpha=0.8)
            ax.set_yticks(range(len(top10)))
            ax.set_yticklabels(top10['Gancheira'])
            ax.invert_yaxis()
            ax.set_xlabel('Defeitos')
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        else:
            st.info("Sem dados de gancheiras disponíveis.")
    else:
        st.info("Coluna GANCHEIRA não encontrada.")

    # ── Análise de Posições da Pior Gancheira ──
    st.markdown("<hr>", unsafe_allow_html=True)
    render_section_header("🔧 Análise de Posições - Pior Gancheira", "▸", THEME['accent_purple'])
    
    if not df.empty and 'GANCHEIRA' in df.columns:
        ranking = []
        for g in df['GANCHEIRA'].dropna().unique():
            df_g = df[df['GANCHEIRA'] == g]
            ranking.append({'Gancheira': str(g), 'Defeitos': int(df_g['TOTAL_DEFEITOS'].sum()), 'Reg': len(df_g)})
        
        if ranking:
            df_rank = pd.DataFrame(ranking).sort_values('Defeitos', ascending=False)
            pior = df_rank.iloc[0]
            
            st.markdown(f"""
            <div style="background: {THEME['bg_card2']}; padding: 10px 15px; margin-bottom: 15px; border-left: 4px solid {THEME['accent_red']};">
                <span style="font-weight: bold; color: {THEME['accent_red']};">🔴 PIOR GANCHEIRA: {pior['Gancheira']}</span> | 
                {pior['Defeitos']} defeitos em {pior['Reg']} registros
            </div>
            """, unsafe_allow_html=True)
            
            df_pior = df[df['GANCHEIRA'] == pior['Gancheira']]
            posicoes_dados = []
            for col in df_base.columns:
                try:
                    num = int(str(col).strip())
                    if 11 <= num <= 78 and col in df_pior.columns:
                        contagem = {c: 0 for c in CODIGOS_DEFEITO_REAIS}
                        total = 0
                        for val in df_pior[col].dropna():
                            try:
                                cod = int(float(str(val).strip()))
                                if cod in CODIGOS_DEFEITO_REAIS:
                                    contagem[cod] += 1
                                    total += 1
                            except:
                                pass
                        if total > 0:
                            principal_cod = max(contagem, key=contagem.get)
                            principal_nome = MAPEAMENTO_DEFEITOS.get(principal_cod, '?')
                            posicoes_dados.append({
                                'Posição': num,
                                'Defeitos': total,
                                'Principal': f"{principal_nome[:20]} ({contagem[principal_cod]})"
                            })
                except:
                    pass
            
            if posicoes_dados:
                df_pos = pd.DataFrame(posicoes_dados).sort_values('Defeitos', ascending=False)
                total_def = df_pos['Defeitos'].sum()
                df_pos['%'] = (df_pos['Defeitos'] / total_def * 100).round(1)
                
                st.metric("Posições afetadas", len(df_pos))
                
                df_tabela_pos = df_pos[['Posição', 'Defeitos', '%', 'Principal']].head(15).copy()
                df_tabela_pos['%'] = df_tabela_pos['%'].astype(str) + '%'
                
                def estilo_pos(row):
                    styles = [''] * len(row)
                    defeitos = row['Defeitos']
                    if defeitos > 5:
                        styles[1] = 'color: #E81123; font-weight: bold;'
                    elif defeitos > 2:
                        styles[1] = 'color: #E86C2C; font-weight: bold;'
                    return styles
                
                styled_pos = df_tabela_pos.style.apply(estilo_pos, axis=1)
                st.dataframe(styled_pos, use_container_width=True, height=300)
                
                criticas = df_pos[df_pos['Defeitos'] > 5]['Posição'].tolist()
                alerta = df_pos[(df_pos['Defeitos'] >= 2) & (df_pos['Defeitos'] <= 5)]['Posição'].tolist()
                
                if criticas:
                    st.error(f"🚨 **Críticas (>5):** {', '.join(map(str, criticas))}")
                if alerta:
                    st.warning(f"⚠️ **Alerta (2-5):** {', '.join(map(str, alerta))}")
                if not criticas and not alerta:
                    st.success("✅ Nenhuma posição crítica ou em alerta")
                
                if len(df_pos) > 0:
                    fig, ax = plt.subplots(figsize=(10, 3), facecolor=THEME['bg_card'])
                    apply_chart_style(ax, fig, f"Posições com defeitos - Gancheira {pior['Gancheira']}", accent=THEME['accent_red'])
                    top = df_pos.head(20)
                    colors = [THEME['accent_red'] if d > 5 else THEME['accent_orange'] if d > 2 else THEME['accent_yellow'] for d in top['Defeitos']]
                    ax.barh(range(len(top)), top['Defeitos'], color=colors, alpha=0.8)
                    ax.set_yticks(range(len(top)))
                    ax.set_yticklabels(top['Posição'].astype(str))
                    ax.invert_yaxis()
                    ax.set_xlabel('Defeitos')
                    fig.tight_layout()
                    st.pyplot(fig)
                    plt.close(fig)
            else:
                st.info(f"Nenhum defeito nas posições da gancheira {pior['Gancheira']}.")
        else:
            st.info("Sem dados de gancheiras.")
    else:
        st.info("Coluna GANCHEIRA não encontrada.")

    # ── Comparativo por Turnos ──
    st.markdown("<hr>", unsafe_allow_html=True)
    render_section_header("📊 Comparativo por Turno", "▸", THEME['accent_purple'])
    
    if not df.empty and 'TURNO_TEMP' in df.columns:
        turnos = df['TURNO_TEMP'].dropna().unique()
        
        if len(turnos) > 0:
            dados_turnos = []
            for t in turnos:
                df_t = df[df['TURNO_TEMP'] == t]
                total_gancheiras = len(df_t)
                total_aprovado = int(df_t['APROVADO'].sum())
                total_defeitos = int(df_t['TOTAL_DEFEITOS'].sum())
                trs_medio_t = (total_aprovado / (total_gancheiras * 40) * 100) if total_gancheiras > 0 else 0
                
                dados_turnos.append({
                    'Turno': str(t),
                    'Gancheiras': total_gancheiras,
                    'Aprovados': total_aprovado,
                    'Defeitos': total_defeitos,
                    'TRS_num': trs_medio_t,
                    'TRS': f"{trs_medio_t:.1f}%"
                })
            
            df_turnos = pd.DataFrame(dados_turnos)
            df_turnos = df_turnos.sort_values('TRS_num', ascending=False)
            
            df_tabela_turno = df_turnos[['Turno', 'Gancheiras', 'Aprovados', 'Defeitos', 'TRS']].copy()
            
            def estilo_turno(row):
                styles = [''] * len(row)
                try:
                    trs_val = float(row['TRS'].replace('%', ''))
                    if trs_val >= 80:
                        styles[4] = 'color: #107C10; font-weight: bold;'
                    elif trs_val >= 70:
                        styles[4] = 'color: #E86C2C; font-weight: bold;'
                    else:
                        styles[4] = 'color: #E81123; font-weight: bold;'
                except:
                    pass
                styles[2] = 'color: #107C10;'
                styles[3] = 'color: #E81123;'
                return styles
            
            styled_turno = df_tabela_turno.style.apply(estilo_turno, axis=1)
            st.dataframe(styled_turno, use_container_width=True, height=150)
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig, ax = plt.subplots(figsize=(5, 3.5), facecolor=THEME['bg_card'])
                apply_chart_style(ax, fig, "Aprovados vs Defeitos", accent=THEME['accent_purple'])
                x = range(len(df_turnos))
                width = 0.35
                bars1 = ax.bar([i - width/2 for i in x], df_turnos['Aprovados'], width, label='Aprovados', color=THEME['accent_lime'], alpha=0.8)
                bars2 = ax.bar([i + width/2 for i in x], df_turnos['Defeitos'], width, label='Defeitos', color=THEME['accent_red'], alpha=0.8)
                ax.set_xticks(x)
                ax.set_xticklabels(df_turnos['Turno'], fontsize=9)
                ax.legend(loc='upper right', fontsize=8)
                for bar in bars1:
                    h = bar.get_height()
                    if h > 0:
                        ax.text(bar.get_x() + bar.get_width()/2, h + 5, f'{int(h)}', ha='center', va='bottom', fontsize=7, color=THEME['accent_lime'])
                for bar in bars2:
                    h = bar.get_height()
                    if h > 0:
                        ax.text(bar.get_x() + bar.get_width()/2, h + 2, f'{int(h)}', ha='center', va='bottom', fontsize=7, color=THEME['accent_red'])
                fig.tight_layout()
                st.pyplot(fig)
                plt.close(fig)
            
            with col2:
                fig2, ax2 = plt.subplots(figsize=(5, 3.5), facecolor=THEME['bg_card'])
                apply_chart_style(ax2, fig2, "TRS por Turno", ylabel="TRS (%)", accent=THEME['accent_purple'])
                cores_turno = {'M': THEME['accent_cyan'], 'T': THEME['accent_orange'], 'N': THEME['accent_lime']}
                bar_colors = [cores_turno.get(str(t), THEME['accent_purple']) for t in df_turnos['Turno']]
                bars = ax2.bar(range(len(df_turnos)), df_turnos['TRS_num'], color=bar_colors, alpha=0.88, edgecolor=THEME['bg_card'], linewidth=1.5, width=0.55)
                ax2.axhline(y=80, color=THEME['accent_red'], linestyle='--', alpha=0.5, linewidth=1.5, label='Meta 80%')
                for i, (_, row) in enumerate(df_turnos.iterrows()):
                    ax2.text(i, row['TRS_num'] + 1, f"{row['TRS_num']:.1f}%", ha='center', va='bottom', fontweight='bold', fontsize=9, color=THEME['text_primary'])
                ax2.set_xticks(range(len(df_turnos)))
                ax2.set_xticklabels(df_turnos['Turno'], fontsize=10)
                ax2.set_ylim(0, 105)
                ax2.legend(loc='upper right', fontsize=8)
                fig2.tight_layout()
                st.pyplot(fig2)
                plt.close(fig2)
            
            melhor_turno = df_turnos.iloc[0]
            pior_turno = df_turnos.iloc[-1]
            st.info(f"🏆 **Melhor turno:** {melhor_turno['Turno']} ({melhor_turno['TRS']}) | ⚠️ **Pior turno:** {pior_turno['Turno']} ({pior_turno['TRS']})")
        else:
            st.info("Sem dados de turno disponíveis.")
    else:
        st.info("Coluna TURNO_TEMP não encontrada.")

    # ── Defeitos por Tipo ──
    st.markdown("<hr>", unsafe_allow_html=True)
    render_section_header("Defeitos por Tipo", "▸", THEME['accent_purple'])
    
    defeitos_totais = {}
    for codigo, nome in MAPEAMENTO_DEFEITOS.items():
        nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
        col = f'QTD_{nome_clean}'
        if col in df.columns:
            total = int(df[col].sum())
            if total > 0:
                defeitos_totais[nome] = total
    
    if defeitos_totais:
        df_def = pd.DataFrame(list(defeitos_totais.items()), columns=['Defeito', 'Qtd']).sort_values('Qtd', ascending=False)
        
        fig, ax = plt.subplots(figsize=(10, 3), facecolor=THEME['bg_card'])
        apply_chart_style(ax, fig, "", accent=THEME['accent_purple'])
        
        colors = [THEME['accent_lime'] if 'Estourou' in d else THEME['accent_red'] for d in df_def['Defeito']]
        bars = ax.bar(range(len(df_def)), df_def['Qtd'], color=colors, alpha=0.8)
        ax.set_xticks(range(len(df_def)))
        ax.set_xticklabels(df_def['Defeito'], rotation=30, ha='right', fontsize=8)
        
        for bar, v in zip(bars, df_def['Qtd']):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, str(v), ha='center', fontsize=9)
        
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)
    else:
        st.info("Nenhum defeito registrado.")

    # ── Defeitos x Parâmetros ── (CORRIGIDO - agora dentro do bloco TÊMPERA)
    st.markdown("<hr>", unsafe_allow_html=True)
    render_section_header("📈 Defeitos x Parâmetros", "▸", THEME['accent_purple'])

    if not df.empty and 'DATA' in df.columns:
        # Agrupar por data
        df_diario = df.groupby(df['DATA'].dt.date).agg({
            'APROVADO': 'sum',
            'TOTAL_DEFEITOS': 'sum',
            'MEIO': 'mean',
            'C4': 'mean',
            'C2': 'mean'
        }).reset_index()
        
        df_diario['DATA'] = pd.to_datetime(df_diario['DATA'])
        df_diario = df_diario.sort_values('DATA')
        
        # Adicionar contagem de cada defeito por dia
        for codigo in CODIGOS_DEFEITO_REAIS:
            nome = MAPEAMENTO_DEFEITOS[codigo]
            nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
            col_nome = f'QTD_{nome_clean}'
            if col_nome in df.columns:
                df_diario[col_nome] = df.groupby(df['DATA'].dt.date)[col_nome].sum().values
        
        if len(df_diario) >= 2:
            # Opções de defeitos que possuem dados
            opcoes_defeitos = []
            for codigo in CODIGOS_DEFEITO_REAIS:
                nome = MAPEAMENTO_DEFEITOS[codigo]
                nome_clean = nome.upper().replace(' ', '_').replace('Ç', 'C').replace('Ã', 'A').replace('Á', 'A').replace('Ó', 'O')
                col_nome = f'QTD_{nome_clean}'
                if col_nome in df_diario.columns and df_diario[col_nome].sum() > 0:
                    opcoes_defeitos.append((nome, col_nome))
            
            # Adicionar opção "Todos os Defeitos"
            opcoes_defeitos.insert(0, ("📊 TODOS OS DEFEITOS", "TOTAL_DEFEITOS"))
            
            if opcoes_defeitos:
                defeito_selecionado = st.selectbox(
                    "Selecione o defeito para análise:",
                    options=[n for n, _ in opcoes_defeitos],
                    key="defeito_param"
                )
                
                # Encontrar a coluna do defeito selecionado
                col_defeito = next(c for n, c in opcoes_defeitos if n == defeito_selecionado)
                
                # Criar gráfico
                fig, ax1 = plt.subplots(figsize=(14, 6), facecolor=THEME['bg_card'])
                apply_chart_style(ax1, fig, f"{defeito_selecionado} vs Parâmetros de Processo", 
                                  xlabel="Data", ylabel="Quantidade de Defeitos", accent=THEME['accent_red'])
                
                # Barras - defeitos
                bars = ax1.bar(df_diario['DATA'], df_diario[col_defeito], 
                               color=THEME['accent_red'], alpha=0.5, width=0.8, label=defeito_selecionado)
                
                # Adicionar valores nas barras
                for bar, val in zip(bars, df_diario[col_defeito]):
                    if val > 0:
                        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3, 
                                str(int(val)), ha='center', va='bottom', fontsize=7, color=THEME['accent_red'])
                
                ax1.set_ylabel('Quantidade de Defeitos', color=THEME['accent_red'])
                ax1.tick_params(axis='y', labelcolor=THEME['accent_red'])
                
                # Eixo secundário - parâmetros
                ax2 = ax1.twinx()
                
                # Linha 1: Temperatura Meio
                if 'MEIO' in df_diario.columns:
                    ax2.plot(df_diario['DATA'], df_diario['MEIO'], 
                            color=THEME['accent_orange'], marker='o', markersize=4, linewidth=2, 
                            linestyle='-', label='🌡️ Temp. Meio (°C)')
                
                # Linha 2: Humidade C4
                if 'C4' in df_diario.columns:
                    ax2.plot(df_diario['DATA'], df_diario['C4'], 
                            color=THEME['accent_cyan'], marker='s', markersize=4, linewidth=2, 
                            linestyle='--', label='💧 Humidade C4 (%)')
                
                # Linha 3: Tempo C2
                if 'C2' in df_diario.columns:
                    ax2.plot(df_diario['DATA'], df_diario['C2'], 
                            color=THEME['accent_lime'], marker='^', markersize=4, linewidth=2, 
                            linestyle='-.', label='⏱️ Tempo C2 (s)')
                
                ax2.set_ylabel('Valores dos Parâmetros', color=THEME['text_primary'])
                ax2.tick_params(axis='y', labelcolor=THEME['text_primary'])
                
                # Legendas combinadas
                lines1, labels1 = ax1.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=8, 
                          framealpha=0.9, facecolor=THEME['bg_card'])
                
                plt.setp(ax1.xaxis.get_majorticklabels(), rotation=35, ha='right', fontsize=8)
                fig.tight_layout()
                st.pyplot(fig)
                plt.close(fig)
                
                # ── ESTATÍSTICAS DE CORRELAÇÃO COM ANÁLISE INTELIGENTE ──
                st.markdown("#### 📊 Correlações com Análise Contextual")
                
                # Dicionário de observações específicas por tipo de defeito e parâmetro
                def get_observacao(defeito_nome, parametro, corr_valor):
                    """Retorna observação contextual baseada no defeito e correlação"""
                    
                    # Normaliza o nome do defeito
                    defeito_lower = defeito_nome.lower()
                    parametro_lower = parametro.lower()
                    
                    # Observações para HUMIDADE (C4)
                    if 'humidade' in parametro_lower or 'c4' in parametro_lower:
                        if corr_valor > 0.5:
                            return "⚠️ **Contradição teórica:** A literatura indica que umidade alta é prejudicial, mas seus dados mostram o oposto. Investigue se: (a) a faixa de umidade está abaixo do limiar crítico; (b) a umidade está compensando resfriamento excessivo; (c) há correlação espúria com outra variável (ex: estação do ano)."
                        elif corr_valor < -0.5:
                            if 'quebra no resfriamento' in defeito_lower or 'quebra teste impacto' in defeito_lower:
                                return "✅ **Coerente com teoria:** Umidade mais alta suaviza o choque térmico, reduzindo quebras por resfriamento abrupto. A umidade atua como amortecedor térmico natural."
                            elif 'estourou após furar' in defeito_lower:
                                return "⚠️ **Atenção:** Maior umidade reduzindo este defeito sugere que a peça pode estar chegando muito seca/trincada ao furo. Considere verificar lubrificação da furadeira."
                            else:
                                return "✅ **Efeito benéfico:** Maior umidade reduz este defeito. Possível explicação: a umidade moderada (30-60%) melhora a transferência de calor uniforme durante o resfriamento."
                        else:
                            if 'quebra no resfriamento' in defeito_lower or 'quebra teste impacto' in defeito_lower:
                                return "ℹ️ **Correlação fraca:** Umidade não parece ser o principal fator para este defeito. Priorize análise de temperatura e tempo de resfriamento."
                            else:
                                return "ℹ️ **Correlação fraca:** Umidade tem impacto limitado neste tipo de defeito."
                    
                    # Observações para TEMPERATURA MEIO
                    elif 'temperatura' in parametro_lower or 'meio' in parametro_lower:
                        if corr_valor > 0.5:
                            if 'quebra no resfriamento' in defeito_lower or 'quebra teste impacto' in defeito_lower:
                                return "⚠️ **Temperatura muito alta** pode estar causando tensões residuais excessivas. Considere reduzir temperatura do forno em 5-10°C."
                            elif 'estourou após furar' in defeito_lower:
                                return "⚠️ **Temperatura alta** pode estar fragilizando o vidro na região do furo. Verifique distribuição de temperatura na peça."
                            else:
                                return "⚠️ **Correlação positiva:** Temperaturas mais altas aumentam este defeito. Reduza gradualmente a temperatura e monitore o impacto."
                        elif corr_valor < -0.5:
                            if 'quebra no resfriamento' in defeito_lower or 'quebra teste impacto' in defeito_lower:
                                return "✅ **Temperatura muito baixa** pode estar causando têmpera insuficiente. Aumente temperatura em 5-10°C e verifique resultado."
                            else:
                                return "✅ **Correlação negativa:** Temperaturas mais altas reduzem este defeito. Considere operar no limite superior da faixa especificada."
                        else:
                            return "ℹ️ **Correlação moderada:** Temperatura tem influência, mas não é o fator dominante. Analise também tempo de residência e resfriamento."
                    
                    # Observações para TEMPO C2
                    elif 'tempo' in parametro_lower or 'c2' in parametro_lower:
                        if corr_valor > 0.5:
                            if 'quebra no resfriamento' in defeito_lower or 'quebra teste impacto' in defeito_lower:
                                return "⚠️ **Tempo de residência muito longo** pode estar superaquecendo o vidro. Reduza o tempo gradualmente."
                            elif 'ovalizada' in defeito_lower:
                                return "⚠️ **Tempo excessivo** pode estar causando deformação (ovalização) por amolecimento excessivo do vidro."
                            else:
                                return "⚠️ **Correlação positiva:** Tempos mais longos aumentam este defeito. Reduza o tempo de residência no forno."
                        elif corr_valor < -0.5:
                            if 'quebra teste impacto' in defeito_lower:
                                return "✅ **Tempo insuficiente** para têmpera adequada. Aumente o tempo de residência para garantir aquecimento uniforme."
                            elif 'furada e não fraturou' in defeito_lower:
                                return "⚠️ **Tempo curto** pode resultar em têmpera incompleta, reduzindo a fragmentação. Aumente tempo ou temperatura."
                            else:
                                return "✅ **Correlação negativa:** Tempos mais longos reduzem este defeito. Considere operar com maior tempo de residência."
                        else:
                            return "ℹ️ **Correlação moderada:** Tempo tem influência secundária. Priorize ajuste de temperatura primeiro."
                    
                    return "🔍 Nenhuma correlação forte identificada. Continue monitorando outros parâmetros."
                
                # Exibir correlações em cards
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if 'MEIO' in df_diario.columns:
                        corr_temp = df_diario[col_defeito].corr(df_diario['MEIO'])
                        cor_temp = corr_temp if pd.notna(corr_temp) else 0
                        delta_temp = "📈" if cor_temp > 0 else "📉" if cor_temp < 0 else "➡️"
                        st.metric(f"🌡️ vs Temp. Meio", f"{cor_temp:.2f}", delta_temp)
                        
                        with st.expander("🔍 Análise", expanded=False):
                            obs_temp = get_observacao(defeito_selecionado, "Temperatura Meio", cor_temp)
                            st.markdown(obs_temp)
                
                with col2:
                    if 'C4' in df_diario.columns:
                        corr_hum = df_diario[col_defeito].corr(df_diario['C4'])
                        cor_hum = corr_hum if pd.notna(corr_hum) else 0
                        delta_hum = "📈" if cor_hum > 0 else "📉" if cor_hum < 0 else "➡️"
                        st.metric(f"💧 vs Humidade C4", f"{cor_hum:.2f}", delta_hum)
                        
                        with st.expander("🔍 Análise", expanded=False):
                            obs_hum = get_observacao(defeito_selecionado, "Humidade C4", cor_hum)
                            st.markdown(obs_hum)
                
                with col3:
                    if 'C2' in df_diario.columns:
                        corr_tempo = df_diario[col_defeito].corr(df_diario['C2'])
                        cor_tempo = corr_tempo if pd.notna(corr_tempo) else 0
                        delta_tempo = "📈" if cor_tempo > 0 else "📉" if cor_tempo < 0 else "➡️"
                        st.metric(f"⏱️ vs Tempo C2", f"{cor_tempo:.2f}", delta_tempo)
                        
                        with st.expander("🔍 Análise", expanded=False):
                            obs_tempo = get_observacao(defeito_selecionado, "Tempo C2", cor_tempo)
                            st.markdown(obs_tempo)
                
                # Total do defeito selecionado
                total_defeito = int(df_diario[col_defeito].sum())
                st.info(f"📋 **Total de '{defeito_selecionado}' no período:** {total_defeito} ocorrências")
                
                # ── RESUMO EXECUTIVO COM RECOMENDAÇÕES ──
                st.markdown("#### 🎯 Resumo Executivo e Recomendações")
                
                # Coletar todas as correlações
                correlacoes = {}
                if 'MEIO' in df_diario.columns:
                    correlacoes['Temperatura Meio'] = df_diario[col_defeito].corr(df_diario['MEIO'])
                if 'C4' in df_diario.columns:
                    correlacoes['Humidade C4'] = df_diario[col_defeito].corr(df_diario['C4'])
                if 'C2' in df_diario.columns:
                    correlacoes['Tempo C2'] = df_diario[col_defeito].corr(df_diario['C2'])
                
                # Filtrar correlações válidas
                correlacoes_validas = {k: v for k, v in correlacoes.items() if pd.notna(v)}
                
                if correlacoes_validas:
                    # Encontrar fator mais influente (maior |correlação|)
                    fator_mais_influente = max(correlacoes_validas.items(), key=lambda x: abs(x[1]))
                    
                    st.markdown(f"""
                    <div style="background: {THEME['bg_card2']}; padding: 15px; border-radius: 10px; margin: 10px 0;">
                        <strong>📌 Fator mais influente para "{defeito_selecionado}":</strong> {fator_mais_influente[0]} 
                        (correlação {fator_mais_influente[1]:.2f})
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Recomendações baseadas nas correlações
                    recomendacoes = []
                    
                    # Recomendação para Temperatura
                    if 'Temperatura Meio' in correlacoes_validas:
                        corr_temp = correlacoes_validas['Temperatura Meio']
                        if corr_temp > 0.5:
                            recomendacoes.append("🔧 **Ação:** Reduza a temperatura do forno em 5-10°C e monitore o impacto neste defeito.")
                        elif corr_temp < -0.5:
                            recomendacoes.append("🔧 **Ação:** Aumente a temperatura do forno em 5-10°C e monitore a redução deste defeito.")
                    
                    # Recomendação para Humidade
                    if 'Humidade C4' in correlacoes_validas:
                        corr_hum = correlacoes_validas['Humidade C4']
                        if corr_hum > 0.5:
                            recomendacoes.append("🌧️ **Ação:** Reduza a umidade ambiente (use desumidificadores) abaixo de 50% para diminuir este defeito.")
                        elif corr_hum < -0.5:
                            recomendacoes.append("💧 **Ação:** A umidade mais alta está reduzindo este defeito. Mantenha umidade entre 40-60% e evite ambientes muito secos (<30%).")
                    
                    # Recomendação para Tempo
                    if 'Tempo C2' in correlacoes_validas:
                        corr_tempo = correlacoes_validas['Tempo C2']
                        if corr_tempo > 0.5:
                            recomendacoes.append("⏱️ **Ação:** Reduza o tempo de residência no forno em 10-15% e monitore o efeito.")
                        elif corr_tempo < -0.5:
                            recomendacoes.append("⏱️ **Ação:** Aumente o tempo de residência no forno para garantir têmpera adequada.")
                    
                    if recomendacoes:
                        st.markdown("#### 🔧 Recomendações de Ajuste:")
                        for rec in recomendacoes:
                            st.markdown(rec)
                    else:
                        st.markdown("🔸 Nenhuma correlação forte (>0.5 ou <-0.5) identificada. O processo parece estável para este defeito.")
                else:
                    st.markdown("🔸 Dados insuficientes para análise de correlação.")
                
                # Aviso sobre limitações da análise
                with st.expander("ℹ️ Sobre esta análise", expanded=False):
                    st.markdown("""
                    **Limitações e cuidados:**
                    - Correlação não implica causalidade. Uma correlação forte pode indicar que duas variáveis mudam juntas, mas não necessariamente que uma causa a outra.
                    - Podem existir **variáveis de confundimento** (ex: estação do ano, matéria-prima, operador) que influenciam tanto o defeito quanto o parâmetro.
                    - **Intervalo de confiança:** Correlações baseadas em poucos pontos (menos de 10 dias) têm baixa confiabilidade estatística.
                    - **Faixa de operação:** As correlações são válidas apenas dentro da faixa de dados observada. Extrapolações podem ser perigosas.
                    
                    **Recomendação:** Use estas análises como **guia para investigação**, não como verdade absoluta. Sempre valide ajustes com testes controlados.
                    """)
            else:
                st.info("Nenhum defeito com dados suficientes para análise.")
        else:
            st.info("Dados insuficientes para análise diária (mínimo 2 dias).")
    else:
        st.info("Sem dados disponíveis para análise.")


# ==================================================================================================
# AVISO DE REJEIÇÃO (AR)
# ==================================================================================================
elif aba_selecionada == 'AVISO DE REJEIÇÃO':
    render_page_header("AVISO DE REJEIÇÃO", f"CQ-018 REV004 · Atualizado {get_horario_brasilia()}", THEME['accent_red'])
    
    with st.container():
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, {THEME['bg_card']} 0%, {THEME['bg_card2']} 100%); padding: 15px 20px; border-radius: 10px; border-left: 4px solid {THEME['accent_red']}; margin: 20px 0;">
            <span style="font-size: 20px; margin-right: 10px;">📋</span>
            <span style="font-family: 'Rajdhani', sans-serif; font-size: 18px; font-weight: bold; color: {THEME['accent_red']};">AVISO DE REJEIÇÃO - CQ-018</span>
            <span style="font-family: 'JetBrains Mono', monospace; font-size: 11px; color: {THEME['text_muted']}; margin-left: 15px;">Sistema de Gestão da Qualidade</span>
        </div>
        """, unsafe_allow_html=True)
        
        menu_ar = st.radio("Opções do AR:", ["📝 Novo Registro", "📊 Visualizar Registros", "🔍 Buscar/Editar/Excluir"], horizontal=True, key="menu_ar_principal")
        
        if 'ar_pdf_bytes' not in st.session_state:
            st.session_state.ar_pdf_bytes = None
        if 'ar_pdf_nome' not in st.session_state:
            st.session_state.ar_pdf_nome = None
        if 'ar_mostrar_pdf' not in st.session_state:
            st.session_state.ar_mostrar_pdf = False
        if 'ar_ultimo_registro' not in st.session_state:
            st.session_state.ar_ultimo_registro = None
        if 'ar_etapa_confirmacao' not in st.session_state:
            st.session_state.ar_etapa_confirmacao = 1
        if 'ar_confirmar_email' not in st.session_state:
            st.session_state.ar_confirmar_email = None
        if 'ar_confirmar_imprimir' not in st.session_state:
            st.session_state.ar_confirmar_imprimir = None
        if 'ar_registro_editando' not in st.session_state:
            st.session_state.ar_registro_editando = None
        
        # Função para obter horário de Brasília
        def get_horario_brasilia_ar():
            from datetime import timezone, timedelta
            utc_now = datetime.now(timezone.utc)
            brasilia_offset = timezone(timedelta(hours=-3))
            agora_brasilia = utc_now.astimezone(brasilia_offset)
            return agora_brasilia
        
        # Função para EXCLUIR registro do Google Sheets
        def excluir_registro_ar(numero: int) -> bool:
            """Exclui registro do Google Sheets"""
            try:
                client = get_gspread_client()
                if client is None:
                    st.error("❌ Não foi possível conectar ao Google Sheets")
                    return False
                
                sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_AR)
                
                # Buscar o número na coluna A (primeira coluna)
                celula = sheet.find(str(numero), in_column=1)
                if celula:
                    sheet.delete_rows(celula.row)
                    return True
                else:
                    st.warning(f"⚠️ Registro Nº {numero} não encontrado no Google Sheets")
                    return False
                    
            except Exception as e:
                st.error(f"❌ Erro ao excluir registro: {str(e)}")
                return False
        
        # Função para ALTERAR registro no Google Sheets
        def atualizar_registro_ar(registro: RegistroAR) -> bool:
            """Atualiza registro existente no Google Sheets"""
            try:
                client = get_gspread_client()
                if client is None:
                    st.error("❌ Não foi possível conectar ao Google Sheets")
                    return False
                
                sheet = client.open_by_key(ID_PLANILHA_AR).worksheet(ABA_AR)
                
                # Buscar o número na coluna A
                celula = sheet.find(str(registro.numero), in_column=1)
                if celula:
                    linha = celula.row
                    
                    # Preparar os dados
                    dados = [
                        str(registro.numero),
                        registro.data.strftime("%d/%m/%Y") if registro.data else "",
                        registro.hora,
                        registro.codigo,
                        registro.emissor,
                        registro.referencia,
                        registro.decisao,
                        registro.descricao,
                        registro.status,
                        registro.disposicao,
                        registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else "",
                        registro.turno
                    ]
                    
                    # Atualizar cada célula
                    for col, valor in enumerate(dados, start=1):
                        sheet.update_cell(linha, col, valor)
                    
                    return True
                else:
                    st.warning(f"⚠️ Registro Nº {registro.numero} não encontrado para atualização")
                    return False
                    
            except Exception as e:
                st.error(f"❌ Erro ao atualizar registro: {str(e)}")
                return False
        
        # Função para abrir PDF no navegador e imprimir
        def imprimir_pdf_ar(pdf_bytes: bytes, nome_arquivo: str):
            """Abre o PDF em nova aba para impressão via navegador"""
            try:
                import base64
                import webbrowser
                from datetime import datetime
                
                if not nome_arquivo.lower().endswith('.pdf'):
                    nome_arquivo = nome_arquivo + '.pdf'
                
                # Converter PDF para base64 para abrir no navegador
                pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
                
                # Criar HTML com o PDF embutido e botão de impressão
                html_content = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <title>{nome_arquivo}</title>
                    <meta charset="UTF-8">
                    <style>
                        body {{
                            margin: 0;
                            padding: 0;
                            height: 100vh;
                            display: flex;
                            flex-direction: column;
                            font-family: Arial, sans-serif;
                        }}
                        .toolbar {{
                            background: #2c3e50;
                            padding: 12px 20px;
                            text-align: center;
                            border-bottom: 2px solid #3498db;
                            display: flex;
                            justify-content: center;
                            gap: 15px;
                            flex-wrap: wrap;
                        }}
                        button {{
                            padding: 10px 24px;
                            margin: 0 5px;
                            cursor: pointer;
                            background: #3498db;
                            color: white;
                            border: none;
                            border-radius: 6px;
                            font-size: 14px;
                            font-weight: bold;
                            transition: all 0.3s ease;
                        }}
                        button:hover {{
                            background: #2980b9;
                            transform: scale(1.02);
                        }}
                        .info {{
                            background: #ecf0f1;
                            padding: 8px 16px;
                            border-radius: 6px;
                            color: #2c3e50;
                            font-size: 13px;
                            display: flex;
                            align-items: center;
                            gap: 10px;
                        }}
                        .info span {{
                            font-weight: bold;
                        }}
                        embed {{
                            width: 100%;
                            height: calc(100vh - 70px);
                        }}
                        @media print {{
                            .toolbar {{
                                display: none;
                            }}
                            embed {{
                                height: 100vh;
                            }}
                        }}
                    </style>
                </head>
                <body>
                    <div class="toolbar">
                        <button onclick="window.print()">🖨️ IMPRIMIR (Ctrl+P)</button>
                        <button onclick="window.close()">❌ FECHAR</button>
                        <div class="info">
                            <span>📄 {nome_arquivo}</span>
                            <span>|</span>
                            <span>💡 Pressione Ctrl+P para imprimir</span>
                        </div>
                    </div>
                    <embed src="data:application/pdf;base64,{pdf_base64}" type="application/pdf">
                    <script>
                        // Tentar abrir a impressão automaticamente após 1 segundo
                        setTimeout(function() {{
                            window.print();
                        }}, 1000);
                    </script>
                </body>
                </html>
                """
                
                # Salvar HTML temporário
                temp_html = os.path.join(CAMINHO_PDF_AR, f"temp_print_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
                with open(temp_html, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                
                # Abrir no navegador padrão
                webbrowser.open(f'file://{temp_html}')
                
                # Limpar arquivo temporário após alguns segundos
                import threading
                def limpar_arquivo():
                    import time
                    time.sleep(30)
                    try:
                        if os.path.exists(temp_html):
                            os.remove(temp_html)
                    except:
                        pass
                
                threading.Thread(target=limpar_arquivo, daemon=True).start()
                
                return True, "PDF aberto no navegador. A impressão será iniciada automaticamente ou pressione Ctrl+P."
                
            except Exception as e:
                return False, f"Erro ao abrir PDF: {str(e)}"
        
        if menu_ar == "📝 Novo Registro":
            st.subheader("Novo Aviso de Rejeição")
            st.info("⚠️ Data e hora serão preenchidas automaticamente no salvamento (Horário de Brasília)")
            
            if st.session_state.ar_mostrar_pdf and st.session_state.ar_pdf_bytes:
                st.success(f"✅ Registro salvo com sucesso!")
                if st.session_state.ar_ultimo_registro:
                    reg = st.session_state.ar_ultimo_registro
                    with st.expander("📋 Ver detalhes do registro salvo", expanded=True):
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            st.write(f"**Número:** {reg.numero}")
                            st.write(f"**Data:** {reg.data.strftime('%d/%m/%Y') if reg.data else '-'}")
                            st.write(f"**Hora:** {reg.hora}")
                            st.write(f"**Turno:** {reg.turno}")
                            st.write(f"**Código:** {reg.codigo}")
                        with col_d2:
                            st.write(f"**Emissor:** {reg.emissor}")
                            st.write(f"**Referência:** {reg.referencia[:50]}...")
                            st.write(f"**Decisão:** {reg.decisao}")
                            st.write(f"**Status:** {reg.status}")
                st.markdown("---")
                st.subheader("📋 Confirmações")
                
                if st.session_state.ar_etapa_confirmacao == 1:
                    st.markdown("#### 📧 Etapa 1 de 2")
                    st.write("Deseja enviar este documento por E-MAIL?")
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.button("✅ Sim, enviar e-mail", use_container_width=True):
                            st.session_state.ar_confirmar_email = True
                            st.session_state.ar_etapa_confirmacao = 2
                            st.rerun()
                    with col_btn2:
                        if st.button("❌ Não, pular", use_container_width=True):
                            st.session_state.ar_confirmar_email = False
                            st.session_state.ar_etapa_confirmacao = 2
                            st.rerun()
                elif st.session_state.ar_etapa_confirmacao == 2:
                    if st.session_state.ar_confirmar_email:
                        with st.spinner("Enviando e-mail..."):
                            assunto = f"Aviso de Rejeição - AR {st.session_state.ar_ultimo_registro.numero if st.session_state.ar_ultimo_registro else ''} - {datetime.now().strftime('%d/%m/%Y')}"
                            corpo = f"""Prezados,\n\nSegue em anexo o Aviso de Rejeição.\n\nNúmero: {st.session_state.ar_ultimo_registro.numero if st.session_state.ar_ultimo_registro else ''}\nData: {datetime.now().strftime('%d/%m/%Y')}\nReferência: {st.session_state.ar_ultimo_registro.referencia if st.session_state.ar_ultimo_registro else ''}\n\nAtenciosamente,\nSistema de Gestão da Qualidade - Luvidarte"""
                            if enviar_email_ar(EMAIL_CONFIG_AR["destinatarios"], assunto, corpo, st.session_state.ar_pdf_bytes, st.session_state.ar_pdf_nome):
                                st.success("📧 E-mail enviado com sucesso!")
                            else:
                                st.error("❌ Erro ao enviar e-mail")
                    import time as time_module
                    time_module.sleep(1)
                    st.session_state.ar_etapa_confirmacao = 3
                    st.rerun()
                elif st.session_state.ar_etapa_confirmacao == 3:
                    st.markdown("#### 🖨️ Etapa 2 de 2")
                    st.write("Deseja IMPRIMIR uma cópia deste documento?")
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.button("✅ Sim, imprimir", use_container_width=True):
                            st.session_state.ar_confirmar_imprimir = True
                            st.session_state.ar_etapa_confirmacao = 4
                            st.rerun()
                    with col_btn2:
                        if st.button("❌ Não, pular", use_container_width=True):
                            st.session_state.ar_confirmar_imprimir = False
                            st.session_state.ar_etapa_confirmacao = 4
                            st.rerun()
                elif st.session_state.ar_etapa_confirmacao == 4:
                    if st.session_state.ar_confirmar_imprimir:
                        with st.spinner("Abrindo PDF para impressão..."):
                            success, msg = imprimir_pdf_ar(st.session_state.ar_pdf_bytes, st.session_state.ar_pdf_nome)
                            if success:
                                st.success(f"🖨️ {msg}")
                            else:
                                st.warning(f"⚠️ {msg}")
                    st.markdown("---")
                    st.subheader("✅ Processo concluído!")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(label="📥 Baixar PDF do Registro", data=st.session_state.ar_pdf_bytes, file_name=st.session_state.ar_pdf_nome, mime="application/pdf", use_container_width=True)
                    with col2:
                        if st.button("➕ Novo Registro", use_container_width=True):
                            st.session_state.ar_pdf_bytes = None
                            st.session_state.ar_pdf_nome = None
                            st.session_state.ar_mostrar_pdf = False
                            st.session_state.ar_ultimo_registro = None
                            st.session_state.ar_etapa_confirmacao = 1
                            st.session_state.ar_confirmar_email = None
                            st.session_state.ar_confirmar_imprimir = None
                            st.rerun()
            else:
                with st.form("novo_registro_ar_principal"):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        proximo = obter_proximo_numero_ar()
                        st.info(f"📌 Próximo número disponível: {proximo}")
                        usar_auto = st.checkbox("Usar número automático", value=True)
                        if usar_auto:
                            numero = proximo
                            st.text_input("Número do AR", value=str(numero), disabled=True)
                        else:
                            numero = st.number_input("Número do AR", min_value=1, step=1)
                        turno = st.selectbox("Turno", OPCOES_TURNO_AR)
                    with col2:
                        codigo = st.text_input("Código*")
                        emissor = st.text_input("Emissor*")
                        referencia = st.text_area("Referência*", height=100)
                        decisao = st.selectbox("Decisão*", OPCOES_DECISAO_AR)
                    with col3:
                        status = st.selectbox("Status*", OPCOES_STATUS_AR)
                        descricao = st.text_area("Descrição do Problema*", height=150)
                        disposicao = st.text_area("Disposição", height=100)
                        data_finalizacao = st.date_input("Data de Finalização", datetime.now())
                    submitted = st.form_submit_button("💾 SALVAR REGISTRO", type="primary", use_container_width=True)
                    if submitted:
                        if not codigo or not emissor or not referencia or not descricao:
                            st.error("❌ Preencha todos os campos obrigatórios (*)")
                        else:
                            agora_brasilia = get_horario_brasilia_ar()
                            registro = RegistroAR(
                                numero=numero, 
                                data=agora_brasilia.date(), 
                                hora=agora_brasilia.strftime("%H:%M:%S"), 
                                codigo=codigo, 
                                emissor=emissor, 
                                referencia=referencia, 
                                decisao=decisao, 
                                descricao=descricao, 
                                status=status, 
                                disposicao=disposicao, 
                                data_finalizacao=data_finalizacao, 
                                turno=turno
                            )
                            if salvar_registro_ar(registro, eh_alteracao=False):
                                st.success(f"✅ Registro {numero} salvo com sucesso!")
                                st.info(f"📅 Data: {agora_brasilia.strftime('%d/%m/%Y')} | ⏰ Hora: {agora_brasilia.strftime('%H:%M:%S')} (Horário de Brasília)")
                                
                                pdf_bytes = gerar_pdf_ar(registro)
                                if pdf_bytes:
                                    st.session_state.ar_pdf_bytes = pdf_bytes
                                    st.session_state.ar_pdf_nome = sanitize_filename_ar(f"AR_{numero}_{referencia[:30]}") + ".pdf"
                                    st.session_state.ar_mostrar_pdf = True
                                    st.session_state.ar_ultimo_registro = registro
                                    st.session_state.ar_etapa_confirmacao = 1
                                    st.rerun()
        
        elif menu_ar == "📊 Visualizar Registros":
            st.subheader("Registros de Aviso de Rejeição")
            with st.spinner("Carregando registros..."):
                registros = carregar_registros_ar()
            if registros:
                col_f1, col_f2, col_f3 = st.columns(3)
                with col_f1:
                    filtro_status = st.selectbox("Filtrar por Status", ["Todos"] + OPCOES_STATUS_AR)
                with col_f2:
                    filtro_decisao = st.selectbox("Filtrar por Decisão", ["Todos"] + OPCOES_DECISAO_AR)
                with col_f3:
                    filtro_numero = st.number_input("Filtrar por Nº", min_value=0, step=1, value=0)
                registros_filtrados = registros
                if filtro_status != "Todos":
                    registros_filtrados = [r for r in registros_filtrados if r.status == filtro_status]
                if filtro_decisao != "Todos":
                    registros_filtrados = [r for r in registros_filtrados if r.decisao == filtro_decisao]
                if filtro_numero > 0:
                    registros_filtrados = [r for r in registros_filtrados if r.numero == filtro_numero]
                col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                with col_e1:
                    st.metric("Total Registros", len(registros))
                with col_e2:
                    st.metric("Em Aberto", len([r for r in registros if r.status == "ABERTO"]))
                with col_e3:
                    st.metric("Finalizados", len([r for r in registros if r.status == "FINALIZADO"]))
                with col_e4:
                    st.metric("Aprovados Cond.", len([r for r in registros if r.decisao == "APROVADO CONDICIONAL"]))
                dados = []
                for reg in registros_filtrados[:100]:
                    dados.append({"Nº": reg.numero, "Data": reg.data.strftime("%d/%m/%Y") if reg.data else "-", "Hora": reg.hora, "Código": reg.codigo, "Emissor": reg.emissor, "Referência": reg.referencia[:40] + "..." if len(reg.referencia) > 40 else reg.referencia, "Decisão": reg.decisao, "Status": reg.status, "Turno": reg.turno})
                df = pd.DataFrame(dados)
                st.dataframe(df, use_container_width=True, height=400)
            else:
                st.info("📭 Nenhum registro encontrado na planilha.")
        
        elif menu_ar == "🔍 Buscar/Editar/Excluir":
            st.subheader("Buscar, Editar ou Excluir Registro")
            
            col_b1, col_b2 = st.columns([2, 1])
            with col_b1:
                numero_busca = st.number_input("Digite o número do AR", min_value=1, step=1, key="busca_ar_principal")
            with col_b2:
                st.markdown("<br>", unsafe_allow_html=True)
                buscar_clicked = st.button("🔍 Buscar", key="buscar_ar_principal_btn", use_container_width=True)
            
            if buscar_clicked and numero_busca:
                with st.spinner("Buscando..."):
                    registros = carregar_registros_ar({"numero": numero_busca})
                if registros:
                    st.session_state.ar_registro_editando = registros[0]
                    st.rerun()
            
            if st.session_state.ar_registro_editando:
                reg = st.session_state.ar_registro_editando
                st.success(f"✅ Registro Nº {reg.numero} encontrado!")
                
                with st.expander("📋 Dados completos do registro", expanded=True):
                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        st.write("**📅 Datas e Horários:**")
                        st.write(f"- Data: {reg.data.strftime('%d/%m/%Y') if reg.data else '-'}")
                        st.write(f"- Hora: {reg.hora}")
                        st.write(f"- Turno: {reg.turno}")
                        st.write(f"- Data Finalização: {reg.data_finalizacao.strftime('%d/%m/%Y') if reg.data_finalizacao else '-'}")
                        st.write("**📝 Informações do Produto:**")
                        st.write(f"- Código: {reg.codigo}")
                        st.write(f"- Emissor: {reg.emissor}")
                        st.write(f"- Referência: {reg.referencia}")
                    with col_d2:
                        st.write("**⚖️ Decisão e Status:**")
                        st.write(f"- Decisão: {reg.decisao}")
                        st.write(f"- Status: {reg.status}")
                        st.write("**📋 Descrições:**")
                        st.write(f"- Problema: {reg.descricao[:150]}...")
                        st.write(f"- Disposição: {reg.disposicao[:150]}...")
                
                tab_editar, tab_excluir, tab_acoes = st.tabs(["✏️ Editar Registro", "🗑️ Excluir Registro", "📄 Ações do PDF"])
                
                with tab_editar:
                    st.subheader("Editar Registro")
                    with st.form("editar_ar_principal"):
                        col_e1, col_e2, col_e3 = st.columns(3)
                        with col_e1:
                            data_edt = st.date_input("Data", reg.data if reg.data else datetime.now())
                            hora_edt = st.text_input("Hora", reg.hora)
                            turno_edt = st.selectbox("Turno", OPCOES_TURNO_AR, index=OPCOES_TURNO_AR.index(reg.turno) if reg.turno in OPCOES_TURNO_AR else 0)
                        with col_e2:
                            codigo_edt = st.text_input("Código", reg.codigo)
                            emissor_edt = st.text_input("Emissor", reg.emissor)
                            referencia_edt = st.text_area("Referência", reg.referencia, height=80)
                            decisao_edt = st.selectbox("Decisão", OPCOES_DECISAO_AR, index=OPCOES_DECISAO_AR.index(reg.decisao) if reg.decisao in OPCOES_DECISAO_AR else 0)
                        with col_e3:
                            status_edt = st.selectbox("Status", OPCOES_STATUS_AR, index=OPCOES_STATUS_AR.index(reg.status) if reg.status in OPCOES_STATUS_AR else 0)
                            descricao_edt = st.text_area("Descrição", reg.descricao, height=80)
                            disposicao_edt = st.text_area("Disposição", reg.disposicao, height=80)
                            data_fim_edt = st.date_input("Data Finalização", reg.data_finalizacao if reg.data_finalizacao else datetime.now())
                        
                        submitted_edit = st.form_submit_button("💾 SALVAR ALTERAÇÕES", type="primary", use_container_width=True)
                        
                        if submitted_edit:
                            if not codigo_edt or not emissor_edt or not referencia_edt or not descricao_edt:
                                st.error("❌ Preencha todos os campos obrigatórios")
                            else:
                                registro_atualizado = RegistroAR(
                                    numero=reg.numero, 
                                    data=data_edt, 
                                    hora=hora_edt, 
                                    codigo=codigo_edt, 
                                    emissor=emissor_edt, 
                                    referencia=referencia_edt, 
                                    decisao=decisao_edt, 
                                    descricao=descricao_edt, 
                                    status=status_edt, 
                                    disposicao=disposicao_edt, 
                                    data_finalizacao=data_fim_edt, 
                                    turno=turno_edt
                                )
                                if atualizar_registro_ar(registro_atualizado):
                                    st.success(f"✅ Registro {reg.numero} atualizado com sucesso!")
                                    st.session_state.ar_registro_editando = None
                                    st.rerun()
                                else:
                                    st.error("❌ Erro ao atualizar registro")
                
                with tab_excluir:
                    st.subheader("Excluir Registro")
                    st.error(f"⚠️ **ATENÇÃO!** Exclusão do Registro Nº {reg.numero}")
                    st.warning("Esta ação é **IRREVERSÍVEL** e não pode ser desfeita!")
                    
                    st.write("**Dados do registro que será excluído:**")
                    st.write(f"- Código: {reg.codigo}")
                    st.write(f"- Emissor: {reg.emissor}")
                    st.write(f"- Referência: {reg.referencia[:50]}...")
                    st.write(f"- Data: {reg.data.strftime('%d/%m/%Y') if reg.data else '-'}")
                    
                    confirmar = st.checkbox(f"✅ Confirmo que desejo EXCLUIR permanentemente o Registro Nº {reg.numero}")
                    
                    if confirmar:
                        if st.button("🗑️ CONFIRMAR EXCLUSÃO", type="primary", use_container_width=True):
                            with st.spinner(f"Excluindo registro {reg.numero}..."):
                                if excluir_registro_ar(reg.numero):
                                    st.success(f"✅ Registro {reg.numero} excluído com sucesso!")
                                    st.balloons()
                                    st.session_state.ar_registro_editando = None
                                    st.rerun()
                                else:
                                    st.error(f"❌ Erro ao excluir registro {reg.numero}")
                    else:
                        st.info("Marque a caixa de confirmação para habilitar a exclusão.")
                
                with tab_acoes:
                    st.subheader("Ações para este Registro")
                    
                    if st.button("📄 Gerar PDF", use_container_width=True):
                        pdf_bytes = gerar_pdf_ar(reg)
                        if pdf_bytes:
                            nome_pdf = sanitize_filename_ar(f"AR_{reg.numero}_{reg.referencia[:30]}") + ".pdf"
                            st.session_state.ar_pdf_bytes = pdf_bytes
                            st.session_state.ar_pdf_nome = nome_pdf
                            st.session_state.ar_ultimo_registro = reg
                            st.session_state.ar_etapa_confirmacao = 1
                            st.session_state.ar_mostrar_pdf = True
                            st.rerun()
                    
                    if st.session_state.ar_mostrar_pdf and st.session_state.ar_pdf_bytes:
                        st.markdown("---")
                        st.success("✅ PDF gerado com sucesso!")
                        
                        st.markdown("#### 📧 Enviar por E-mail")
                        if st.button("Enviar este PDF por E-mail", use_container_width=True):
                            with st.spinner("Enviando e-mail..."):
                                assunto = f"Aviso de Rejeição - AR {reg.numero} - {datetime.now().strftime('%d/%m/%Y')}"
                                corpo = f"""Prezados,\n\nSegue em anexo o Aviso de Rejeição.\n\nNúmero: {reg.numero}\nData: {reg.data.strftime('%d/%m/%Y') if reg.data else '-'}\nReferência: {reg.referencia}\n\nAtenciosamente,\nSistema de Gestão da Qualidade - Luvidarte"""
                                if enviar_email_ar(EMAIL_CONFIG_AR["destinatarios"], assunto, corpo, st.session_state.ar_pdf_bytes, st.session_state.ar_pdf_nome):
                                    st.success("📧 E-mail enviado com sucesso!")
                                else:
                                    st.error("❌ Erro ao enviar e-mail")
                        
                        st.markdown("#### 🖨️ Imprimir")
                        if st.button("Imprimir este PDF", use_container_width=True):
                            with st.spinner("Abrindo PDF para impressão..."):
                                success, msg = imprimir_pdf_ar(st.session_state.ar_pdf_bytes, st.session_state.ar_pdf_nome)
                                if success:
                                    st.success(f"🖨️ {msg}")
                                else:
                                    st.warning(f"⚠️ {msg}")
                        
                        st.markdown("#### 📥 Salvar PDF")
                        st.download_button(label="Baixar PDF", data=st.session_state.ar_pdf_bytes, file_name=st.session_state.ar_pdf_nome, mime="application/pdf", use_container_width=True)
                
                # Botão para limpar e voltar
                if st.button("🔍 Nova Busca", use_container_width=True):
                    st.session_state.ar_registro_editando = None
                    st.rerun()
    
    st.markdown(f"""
    <div style="text-align:right;padding:16px 0 8px;
        font-family:'JetBrains Mono',monospace;font-size:10px;
        color:{THEME['text_muted']};letter-spacing:.1em;">
        AVISO DE REJEIÇÃO · {get_horario_brasilia()}
    </div>
    """, unsafe_allow_html=True)

   
# ==================================================================================================
# REQUISIÇÃO DE MANUTENÇÃO (RM) - VERSÃO CORRIGIDA
# ==================================================================================================
elif aba_selecionada == 'REQUISIÇÃO MANUTENÇÃO':
    render_page_header("REQUISIÇÃO DE MANUTENÇÃO", f"MF-001 · Atualizado {get_horario_brasilia()}", THEME['accent_lime'])
    
    # Configurações do RM
    ID_PLANILHA_RM = ID_PLANILHA_AR  # Mesma planilha do AR
    ABA_RM = 'RM'  # Aba específica para RM
    
    # Opções
    OPCOES_CARATER_RM = [
        "1 - Risco Físico/Segurança",
        "2 - Impacto Imediato na Produção", 
        "3 - Impacto a Longo Prazo",
        "4 - Melhoria/Preventiva"
    ]
    OPCOES_SETORES_RM = ["Produção", "Corte", "Vidraria", "Rodaria", "Embalagem", "Expedição", "Qualidade", "Ferramentaria", "Manutenção", "Outros"]
    OPCOES_SETORES2_RM = ["Elétrica", "Mecânica", "Informática", "Ferramentaria", "Manutenção Geral"]
    OPCOES_STATUS_RM = ["ABERTO", "EM ANDAMENTO", "FINALIZADO", "CANCELADO"]
    
    # Criar diretórios RM
    CAMINHO_PDF_RM = r"\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\1-REQUISIÇÃO DE MANUTENÇÃO\1-PDF"
    CAMINHO_PDF_RELATORIO_RM = r"\\srv-luvidarte\dados\DOC\Engenharia_Luvidarte\SGQ - LUVIDARTE - ALTERADAS\1-REQUISIÇÃO DE MANUTENÇÃO\2-PDF"
    os.makedirs(CAMINHO_PDF_RM, exist_ok=True)
    os.makedirs(CAMINHO_PDF_RELATORIO_RM, exist_ok=True)
    
    @dataclass
    class RegistroRM:
        id: Optional[int] = None
        data: Optional[datetime] = None
        hora: str = ""
        emissor: str = ""
        equipamento: str = ""
        setor: str = ""
        caracter: str = ""
        setor2: str = ""
        problema: str = ""
        trabalho: str = ""
        analise: str = ""
        status: str = "ABERTO"
        data_finalizacao: Optional[datetime] = None
        emissor2: str = ""
    
    def obter_proximo_id_rm():
        try:
            client = get_gspread_client()
            if client is None:
                return 1
            sheet = client.open_by_key(ID_PLANILHA_RM).worksheet(ABA_RM)
            todos_dados = sheet.get_all_values()
            if len(todos_dados) < 2:
                return 1
            ids = []
            for row in todos_dados[1:]:
                if row and row[0].strip():
                    try:
                        ids.append(int(row[0]))
                    except:
                        pass
            return max(ids) + 1 if ids else 1
        except:
            return 1
    
    def carregar_registros_rm(filtros: Dict[str, Any] = None) -> List[RegistroRM]:
        registros = []
        try:
            client = get_gspread_client()
            if client is None:
                st.error("❌ Não foi possível conectar ao Google Sheets")
                return registros
            
            # Abrir a aba RM
            spreadsheet = client.open_by_key(ID_PLANILHA_RM)
            
            # Verificar se a aba existe
            try:
                sheet = spreadsheet.worksheet(ABA_RM)
            except Exception as e:
                st.error(f"❌ Aba '{ABA_RM}' não encontrada! Verifique se o nome da aba está correto.")
                return registros
            
            # Obter todos os dados
            todos_dados = sheet.get_all_values()
            
            if len(todos_dados) < 2:
                st.warning(f"⚠️ Nenhum dado encontrado na aba '{ABA_RM}'. Adicione pelo menos uma linha de dados.")
                return registros
            
            # Processar linhas (pular cabeçalho)
            for idx, row in enumerate(todos_dados[1:], start=2):
                if len(row) < 14:
                    continue
                
                try:
                    # Mapear colunas conforme sua planilha
                    # Coluna A (0): ID
                    # Coluna B (1): DATA
                    # Coluna C (2): HORA
                    # Coluna D (3): EMISSOR
                    # Coluna E (4): EQUIPAMENTO
                    # Coluna F (5): SETOR
                    # Coluna G (6): CARÁTER
                    # Coluna H (7): SETOR2
                    # Coluna I (8): PROBLEMA
                    # Coluna J (9): TRABALHO
                    # Coluna K (10): ANÁLISE
                    # Coluna L (11): STATUS
                    # Coluna M (12): DATA FINALIZAÇÃO
                    # Coluna N (13): EMISSOR2
                    
                    registro = RegistroRM()
                    registro.id = int(row[0]) if row[0].strip() else None
                    registro.data = converter_data_br(row[1])
                    registro.hora = row[2] if len(row) > 2 else ""
                    registro.emissor = row[3] if len(row) > 3 else ""
                    registro.equipamento = row[4] if len(row) > 4 else ""
                    registro.setor = row[5] if len(row) > 5 else ""
                    registro.caracter = row[6] if len(row) > 6 else ""
                    registro.setor2 = row[7] if len(row) > 7 else ""
                    registro.problema = row[8] if len(row) > 8 else ""
                    registro.trabalho = row[9] if len(row) > 9 else ""
                    registro.analise = row[10] if len(row) > 10 else ""
                    registro.status = row[11] if len(row) > 11 else "ABERTO"
                    registro.data_finalizacao = converter_data_br(row[12]) if len(row) > 12 else None
                    registro.emissor2 = row[13] if len(row) > 13 else ""
                    
                    # Aplicar filtros
                    if filtros:
                        incluir = True
                        if filtros.get('id') and filtros['id'] != registro.id:
                            incluir = False
                        if filtros.get('equipamento') and filtros['equipamento'].lower() not in registro.equipamento.lower():
                            incluir = False
                        if filtros.get('status') and filtros['status'] != registro.status:
                            incluir = False
                    else:
                        incluir = True
                    
                    if incluir and registro.id is not None:
                        registros.append(registro)
                        
                except Exception as e:
                    st.warning(f"Erro ao processar linha {idx}: {e}")
                    continue
            
            # Ordenar por ID decrescente (mais recente primeiro)
            registros.sort(key=lambda x: x.id if x.id else 0, reverse=True)
            
            if registros:
                st.success(f"✅ {len(registros)} registro(s) carregado(s)")
            
        except Exception as e:
            st.error(f"Erro ao carregar registros RM: {e}")
            traceback.print_exc()
        
        return registros
    
    def salvar_registro_rm(registro: RegistroRM, eh_alteracao: bool = False) -> bool:
        try:
            client = get_gspread_client()
            if client is None:
                st.error("❌ Não foi possível conectar ao Google Sheets")
                return False
            
            sheet = client.open_by_key(ID_PLANILHA_RM).worksheet(ABA_RM)
            
            dados = [
                str(registro.id) if registro.id else "",
                registro.data.strftime("%d/%m/%Y") if registro.data else "",
                registro.hora,
                registro.emissor,
                registro.equipamento,
                registro.setor,
                registro.caracter,
                registro.setor2,
                registro.problema,
                registro.trabalho,
                registro.analise,
                registro.status,
                registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else "",
                registro.emissor2
            ]
            
            if eh_alteracao:
                cell = sheet.find(str(registro.id), in_column=1)
                if cell:
                    row_num = cell.row
                    for col, valor in enumerate(dados, start=1):
                        sheet.update_cell(row_num, col, valor)
                else:
                    sheet.append_row(dados)
            else:
                sheet.insert_row(dados, index=2)
            
            return True
            
        except Exception as e:
            st.error(f"Erro ao salvar registro RM: {e}")
            return False
    
    def excluir_registro_rm(id: int) -> bool:
        try:
            client = get_gspread_client()
            if client is None:
                st.error("❌ Não foi possível conectar ao Google Sheets")
                return False
            
            sheet = client.open_by_key(ID_PLANILHA_RM).worksheet(ABA_RM)
            cell = sheet.find(str(id), in_column=1)
            if cell:
                sheet.delete_rows(cell.row)
                return True
            return False
            
        except Exception as e:
            st.error(f"Erro ao excluir registro RM: {e}")
            return False
    
    def gerar_pdf_rm(registro: RegistroRM) -> Optional[bytes]:
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.lib.units import cm
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            elementos = []
            styles = getSampleStyleSheet()
            styleN = styles["Normal"]
            
            elementos.append(Paragraph("<b>REQUISIÇÃO DE MANUTENÇÃO</b>", ParagraphStyle(name='Titulo', parent=styles["Heading1"], fontSize=16, alignment=1, spaceAfter=12)))
            elementos.append(Paragraph("<b>MF-001 - Luvidarte</b>", ParagraphStyle(name='Subtitulo', parent=styles["Heading2"], fontSize=12, alignment=1, spaceAfter=24)))
            
            data_str = registro.data.strftime("%d/%m/%Y") if registro.data else ""
            data_fim_str = registro.data_finalizacao.strftime("%d/%m/%Y") if registro.data_finalizacao else ""
            
            tabela_dados = Table([
                ["ID:", registro.id, "Data:", data_str, "Hora:", registro.hora],
                ["Emissor:", registro.emissor, "Equipamento:", registro.equipamento, "Setor:", registro.setor],
                ["Caráter:", registro.caracter, "Setor Destino:", registro.setor2, "Status:", registro.status],
                ["Emissor Técnico:", registro.emissor2, "Data Finalização:", data_fim_str, "", ""],
            ], colWidths=[2.5*cm, 4*cm, 2.5*cm, 4*cm, 2*cm, 2.5*cm])
            
            tabela_dados.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.5, colors.black),
                ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("ALIGN", (0,0), (-1,-1), "LEFT"),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("PADDING", (0,0), (-1,-1), 6),
            ]))
            elementos.append(tabela_dados)
            elementos.append(Spacer(1, 24))
            
            elementos.append(Paragraph("<b>DESCRIÇÃO DO PROBLEMA:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
            elementos.append(Paragraph(registro.problema or "-", styleN))
            elementos.append(Spacer(1, 24))
            
            elementos.append(Paragraph("<b>TRABALHO REALIZADO:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
            elementos.append(Paragraph(registro.trabalho or "_________________________", styleN))
            elementos.append(Spacer(1, 24))
            
            elementos.append(Paragraph("<b>ANÁLISE DO SERVIÇO:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
            elementos.append(Paragraph(registro.analise or "_________________________", styleN))
            elementos.append(Spacer(1, 24))
            
            elementos.append(Paragraph("<b>ASSINATURAS:</b>", ParagraphStyle(name='SubtituloSecao', parent=styleN, fontSize=12, spaceAfter=6)))
            tabela_assinatura = Table([
                ["Solicitante", "Responsável Técnico", "Conferência Qualidade"],
                ["_________________________", "_________________________", "_________________________"],
                ["Data: __/__/____", "Data: __/__/____", "Data: __/__/____"]
            ], colWidths=[5.5*cm, 5.5*cm, 5.5*cm])
            tabela_assinatura.setStyle(TableStyle([
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("PADDING", (0,0), (-1,-1), 10),
                ("FONTNAME", (0,0), (0,0), "Helvetica-Bold"),
                ("FONTNAME", (1,0), (1,0), "Helvetica-Bold"),
                ("FONTNAME", (2,0), (2,0), "Helvetica-Bold"),
            ]))
            elementos.append(tabela_assinatura)
            
            elementos.append(Spacer(1, 36))
            elementos.append(Paragraph(f"Documento gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ParagraphStyle(name='Rodape', parent=styleN, fontSize=8, alignment=2)))
            
            doc.build(elementos)
            buffer.seek(0)
            return buffer.getvalue()
        except Exception as e:
            st.error(f"Erro ao gerar PDF RM: {e}")
            return None
    
    with st.container():
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, {THEME['bg_card']} 0%, {THEME['bg_card2']} 100%); padding: 15px 20px; border-radius: 10px; border-left: 4px solid {THEME['accent_lime']}; margin: 20px 0;">
            <span style="font-size: 20px; margin-right: 10px;">🔧</span>
            <span style="font-family: 'Rajdhani', sans-serif; font-size: 18px; font-weight: bold; color: {THEME['accent_lime']};">REQUISIÇÃO DE MANUTENÇÃO - MF-001</span>
            <span style="font-family: 'JetBrains Mono', monospace; font-size: 11px; color: {THEME['text_muted']}; margin-left: 15px;">Sistema de Gestão da Qualidade</span>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("📌 **Tabela de Caráter - Níveis 1 a 4**", expanded=False):
            st.markdown("""
            | Nível | Descrição | Ação Recomendada |
            |-------|-----------|------------------|
            | 🚨 **1** | **Risco Físico/Segurança** | Risco iminente de acidente ou dano físico. Ação imediata! |
            | ⚠️ **2** | **Impacto Imediato na Produção** | Parada total ou parcial da produção. Resolver em até 4h |
            | 📊 **3** | **Impacto a Longo Prazo** | Pode afetar produção futura. Planejar em até 48h |
            | 🔧 **4** | **Melhoria/Preventiva** | Manutenção programada ou melhoria. Agendar conforme disponibilidade |
            """)
        
        menu_rm = st.radio("Opções:", ["📝 Nova Requisição", "📊 Visualizar Requisições", "🔍 Buscar/Editar/Excluir", "📈 Dashboard RM"], horizontal=True, key="menu_rm_principal")
        
        if 'rm_pdf_bytes' not in st.session_state:
            st.session_state.rm_pdf_bytes = None
        if 'rm_pdf_nome' not in st.session_state:
            st.session_state.rm_pdf_nome = None
        if 'rm_mostrar_pdf' not in st.session_state:
            st.session_state.rm_mostrar_pdf = False
        if 'rm_ultimo_registro' not in st.session_state:
            st.session_state.rm_ultimo_registro = None
        if 'rm_registro_editando' not in st.session_state:
            st.session_state.rm_registro_editando = None
        
        if menu_rm == "📝 Nova Requisição":
            st.subheader("Nova Requisição de Manutenção")
            st.info("⚠️ Data e hora serão preenchidas automaticamente no salvamento (Horário de Brasília)")
            
            if st.session_state.rm_mostrar_pdf and st.session_state.rm_pdf_bytes:
                st.success(f"✅ Requisição salva com sucesso!")
                if st.session_state.rm_ultimo_registro:
                    reg = st.session_state.rm_ultimo_registro
                    with st.expander("📋 Ver detalhes da requisição", expanded=True):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**ID:** {reg.id}")
                            st.write(f"**Data:** {reg.data.strftime('%d/%m/%Y') if reg.data else '-'}")
                            st.write(f"**Hora:** {reg.hora}")
                            st.write(f"**Emissor:** {reg.emissor}")
                            st.write(f"**Equipamento:** {reg.equipamento}")
                        with col2:
                            st.write(f"**Setor:** {reg.setor}")
                            st.write(f"**Caráter:** {reg.caracter}")
                            st.write(f"**Status:** {reg.status}")
                            st.write(f"**Setor Destino:** {reg.setor2}")
                st.markdown("---")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.download_button(label="📥 Baixar PDF", data=st.session_state.rm_pdf_bytes, file_name=st.session_state.rm_pdf_nome, mime="application/pdf", use_container_width=True)
                with col2:
                    if st.button("📧 Enviar por E-mail", use_container_width=True):
                        setor_lower = reg.setor2.lower() if reg.setor2 else ""
                        destinatarios = [EMAIL_CONFIG_AR["usuario"]]
                        assunto = f"Requisição de Manutenção #{reg.id} - {reg.equipamento}"
                        corpo = f"""Prezados,\n\nSegue requisição de manutenção #{reg.id}.\n\nEquipamento: {reg.equipamento}\nCaráter: {reg.caracter}\nSetor Destino: {reg.setor2}\n\nAtenciosamente,\nSistema de Gestão - Luvidarte"""
                        if enviar_email_ar(destinatarios, assunto, corpo, st.session_state.rm_pdf_bytes, st.session_state.rm_pdf_nome):
                            st.success("📧 E-mail enviado!")
                        else:
                            st.error("❌ Erro ao enviar e-mail")
                with col3:
                    if st.button("➕ Nova Requisição", use_container_width=True):
                        st.session_state.rm_pdf_bytes = None
                        st.session_state.rm_pdf_nome = None
                        st.session_state.rm_mostrar_pdf = False
                        st.session_state.rm_ultimo_registro = None
                        st.rerun()
            else:
                with st.form("nova_requisicao_rm"):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        proximo = obter_proximo_id_rm()
                        st.info(f"📌 Próximo ID: {proximo}")
                        usar_auto = st.checkbox("Usar ID automático", value=True)
                        if usar_auto:
                            id_reg = proximo
                            st.text_input("ID", value=str(id_reg), disabled=True)
                        else:
                            id_reg = st.number_input("ID", min_value=1, step=1)
                        emissor = st.text_input("Emissor*")
                        equipamento = st.text_input("Equipamento*")
                        setor = st.selectbox("Setor*", OPCOES_SETORES_RM)
                    with col2:
                        caracter = st.selectbox("Caráter*", OPCOES_CARATER_RM)
                        if "1 -" in caracter:
                            st.error("🚨 **RISCO FÍSICO!** Ação imediata necessária!")
                        elif "2 -" in caracter:
                            st.warning("⚠️ **IMPACTO IMEDIATO!** Resolver em até 4h")
                        elif "3 -" in caracter:
                            st.info("📊 **Impacto a longo prazo** - Planejar em até 48h")
                        else:
                            st.success("🔧 **Melhoria/Preventiva** - Agendar programação")
                        setor2 = st.selectbox("Setor Destino*", OPCOES_SETORES2_RM)
                        status = st.selectbox("Status*", OPCOES_STATUS_RM)
                        data_finalizacao = st.date_input("Data Finalização", datetime.now())
                    with col3:
                        problema = st.text_area("Descrição do Problema*", height=120)
                        trabalho = st.text_area("Trabalho Realizado", height=100)
                        analise = st.text_area("Análise do Serviço", height=100)
                        emissor2 = st.text_input("Emissor Técnico")
                    submitted = st.form_submit_button("💾 SALVAR REQUISIÇÃO", type="primary", use_container_width=True)
                    if submitted:
                        if not emissor or not equipamento or not problema:
                            st.error("❌ Preencha todos os campos obrigatórios (*)")
                        else:
                            agora_brasilia = get_horario_brasilia_obj()
                            registro = RegistroRM(
                                id=id_reg, data=agora_brasilia.date(), hora=agora_brasilia.strftime("%H:%M:%S"),
                                emissor=emissor, equipamento=equipamento, setor=setor, caracter=caracter,
                                setor2=setor2, problema=problema, trabalho=trabalho, analise=analise,
                                status=status, data_finalizacao=data_finalizacao, emissor2=emissor2
                            )
                            if salvar_registro_rm(registro, eh_alteracao=False):
                                st.success(f"✅ Requisição {id_reg} salva com sucesso!")
                                pdf_bytes = gerar_pdf_rm(registro)
                                if pdf_bytes:
                                    st.session_state.rm_pdf_bytes = pdf_bytes
                                    st.session_state.rm_pdf_nome = sanitize_filename_ar(f"RM_{id_reg}_{equipamento[:30]}") + ".pdf"
                                    st.session_state.rm_mostrar_pdf = True
                                    st.session_state.rm_ultimo_registro = registro
                                    st.rerun()
        
        elif menu_rm == "📊 Visualizar Requisições":
            st.subheader("Lista de Requisições de Manutenção")
            with st.spinner("Carregando registros..."):
                registros = carregar_registros_rm()
            
            if registros:
                # Filtros
                col_f1, col_f2, col_f3 = st.columns(3)
                with col_f1:
                    filtro_status = st.selectbox("Status", ["Todos"] + OPCOES_STATUS_RM)
                with col_f2:
                    filtro_caracter = st.selectbox("Caráter", ["Todos"] + OPCOES_CARATER_RM)
                with col_f3:
                    filtro_id = st.number_input("ID", min_value=0, step=1, value=0)
                
                dados_filtrados = registros
                if filtro_status != "Todos":
                    dados_filtrados = [r for r in dados_filtrados if r.status == filtro_status]
                if filtro_caracter != "Todos":
                    dados_filtrados = [r for r in dados_filtrados if r.caracter == filtro_caracter]
                if filtro_id > 0:
                    dados_filtrados = [r for r in dados_filtrados if r.id == filtro_id]
                
                # Estatísticas
                col_e1, col_e2, col_e3, col_e4 = st.columns(4)
                with col_e1:
                    st.metric("Total", len(registros))
                with col_e2:
                    abertos = len([r for r in registros if r.status == "ABERTO"])
                    st.metric("Em Aberto", abertos)
                with col_e3:
                    criticos = len([r for r in registros if "1 -" in r.caracter])
                    st.metric("Nível 1 (Crítico)", criticos)
                with col_e4:
                    finalizados = len([r for r in registros if r.status == "FINALIZADO"])
                    st.metric("Finalizados", finalizados)
                
                # Tabela
                dados_tabela = []
                for reg in dados_filtrados[:100]:
                    emoji = "🚨" if "1 -" in reg.caracter else "⚠️" if "2 -" in reg.caracter else "📊" if "3 -" in reg.caracter else "🔧"
                    dados_tabela.append({
                        "ID": reg.id,
                        "Data": reg.data.strftime("%d/%m/%Y") if reg.data else "-",
                        "Equipamento": reg.equipamento[:30] + "..." if len(reg.equipamento) > 30 else reg.equipamento,
                        "Caráter": f"{emoji} {reg.caracter}",
                        "Setor Destino": reg.setor2,
                        "Status": reg.status,
                        "Emissor": reg.emissor
                    })
                df = pd.DataFrame(dados_tabela)
                st.dataframe(df, use_container_width=True, height=400)
            else:
                st.info("📭 Nenhuma requisição encontrada")
        
        elif menu_rm == "🔍 Buscar/Editar/Excluir":
            st.subheader("Buscar, Editar ou Excluir Requisição")
            
            col_b1, col_b2 = st.columns([2, 1])
            with col_b1:
                id_busca = st.number_input("Digite o ID da requisição", min_value=1, step=1, key="busca_rm")
            with col_b2:
                buscar_clicked = st.button("🔍 Buscar", use_container_width=True)
            
            if buscar_clicked and id_busca:
                with st.spinner("Buscando..."):
                    registros = carregar_registros_rm({"id": id_busca})
                if registros:
                    st.session_state.rm_registro_editando = registros[0]
                    st.rerun()
                else:
                    st.error(f"❌ Requisição ID {id_busca} não encontrada!")
            
            if st.session_state.rm_registro_editando:
                reg = st.session_state.rm_registro_editando
                st.success(f"✅ Requisição ID {reg.id} encontrada!")
                
                with st.expander("📋 Dados completos da requisição", expanded=True):
                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        st.write("**📅 Datas e Horários:**")
                        st.write(f"- Data: {reg.data.strftime('%d/%m/%Y') if reg.data else '-'}")
                        st.write(f"- Hora: {reg.hora}")
                        st.write(f"- Data Finalização: {reg.data_finalizacao.strftime('%d/%m/%Y') if reg.data_finalizacao else '-'}")
                        st.write("**📝 Informações:**")
                        st.write(f"- Emissor: {reg.emissor}")
                        st.write(f"- Equipamento: {reg.equipamento}")
                        st.write(f"- Setor: {reg.setor}")
                    with col_d2:
                        st.write("**⚖️ Caráter e Status:**")
                        st.write(f"- Caráter: {reg.caracter}")
                        st.write(f"- Status: {reg.status}")
                        st.write(f"- Setor Destino: {reg.setor2}")
                        st.write(f"- Emissor Técnico: {reg.emissor2}")
                        st.write("**📋 Descrição:**")
                        st.write(f"- Problema: {reg.problema[:150]}...")
                
                tab_editar, tab_excluir = st.tabs(["✏️ Editar", "🗑️ Excluir"])
                
                with tab_editar:
                    with st.form("editar_rm"):
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            data_edt = st.date_input("Data", reg.data if reg.data else datetime.now())
                            hora_edt = st.text_input("Hora", reg.hora)
                            emissor_edt = st.text_input("Emissor", reg.emissor)
                            equipamento_edt = st.text_input("Equipamento", reg.equipamento)
                            setor_edt = st.selectbox("Setor", OPCOES_SETORES_RM, index=OPCOES_SETORES_RM.index(reg.setor) if reg.setor in OPCOES_SETORES_RM else 0)
                        with col2:
                            caracter_edt = st.selectbox("Caráter", OPCOES_CARATER_RM, index=OPCOES_CARATER_RM.index(reg.caracter) if reg.caracter in OPCOES_CARATER_RM else 0)
                            setor2_edt = st.selectbox("Setor Destino", OPCOES_SETORES2_RM, index=OPCOES_SETORES2_RM.index(reg.setor2) if reg.setor2 in OPCOES_SETORES2_RM else 0)
                            status_edt = st.selectbox("Status", OPCOES_STATUS_RM, index=OPCOES_STATUS_RM.index(reg.status) if reg.status in OPCOES_STATUS_RM else 0)
                            data_fim_edt = st.date_input("Data Finalização", reg.data_finalizacao if reg.data_finalizacao else datetime.now())
                        with col3:
                            problema_edt = st.text_area("Descrição do Problema", reg.problema, height=120)
                            trabalho_edt = st.text_area("Trabalho Realizado", reg.trabalho, height=100)
                            analise_edt = st.text_area("Análise do Serviço", reg.analise, height=100)
                            emissor2_edt = st.text_input("Emissor Técnico", reg.emissor2)
                        if st.form_submit_button("💾 SALVAR ALTERAÇÕES", type="primary"):
                            registro_atualizado = RegistroRM(
                                id=reg.id, data=data_edt, hora=hora_edt, emissor=emissor_edt,
                                equipamento=equipamento_edt, setor=setor_edt, caracter=caracter_edt,
                                setor2=setor2_edt, problema=problema_edt, trabalho=trabalho_edt,
                                analise=analise_edt, status=status_edt, data_finalizacao=data_fim_edt,
                                emissor2=emissor2_edt
                            )
                            if salvar_registro_rm(registro_atualizado, eh_alteracao=True):
                                st.success(f"✅ Requisição {reg.id} atualizada!")
                                st.session_state.rm_registro_editando = None
                                st.rerun()
                            else:
                                st.error("❌ Erro ao atualizar requisição")
                
                with tab_excluir:
                    st.error(f"⚠️ **ATENÇÃO!** Exclusão da Requisição ID {reg.id}")
                    st.warning("Esta ação é **IRREVERSÍVEL**!")
                    confirmar = st.checkbox(f"Confirmo exclusão da requisição {reg.id}")
                    if confirmar and st.button("🗑️ EXCLUIR", type="primary"):
                        with st.spinner(f"Excluindo requisição {reg.id}..."):
                            if excluir_registro_rm(reg.id):
                                st.success(f"✅ Requisição {reg.id} excluída!")
                                st.session_state.rm_registro_editando = None
                                st.rerun()
                            else:
                                st.error(f"❌ Erro ao excluir requisição {reg.id}")
        
        elif menu_rm == "📈 Dashboard RM":
            st.subheader("Dashboard de Manutenção")
            with st.spinner("Carregando dados..."):
                registros = carregar_registros_rm()
            if registros:
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Requisições", len(registros))
                with col2:
                    abertos = len([r for r in registros if r.status == "ABERTO"])
                    st.metric("Em Aberto", abertos)
                with col3:
                    criticos = len([r for r in registros if "1 -" in r.caracter])
                    st.metric("Nível 1 (Crítico)", criticos)
                with col4:
                    finalizados = len([r for r in registros if r.status == "FINALIZADO"])
                    st.metric("Finalizados", finalizados)
                
                st.subheader("Distribuição por Caráter")
                caracter_counts = {}
                for c in OPCOES_CARATER_RM:
                    caracter_counts[c] = len([r for r in registros if r.caracter == c])
                fig, ax = plt.subplots(figsize=(10, 4), facecolor=THEME['bg_card'])
                cores = ['#E81123', '#FF8C00', '#FFB900', '#107C10']
                bars = ax.bar(range(len(caracter_counts)), list(caracter_counts.values()), color=cores, alpha=0.8)
                ax.set_xticks(range(len(caracter_counts)))
                ax.set_xticklabels([c.split(' - ')[1] for c in caracter_counts.keys()], rotation=0)
                ax.set_ylabel("Quantidade")
                ax.set_title("Requisições por Caráter")
                for bar, v in zip(bars, caracter_counts.values()):
                    if v > 0:
                        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, str(v), ha='center', va='bottom')
                st.pyplot(fig)
                plt.close(fig)
                
                st.subheader("Status das Requisições")
                status_counts = {}
                for s in OPCOES_STATUS_RM:
                    status_counts[s] = len([r for r in registros if r.status == s])
                fig2, ax2 = plt.subplots(figsize=(8, 4), facecolor=THEME['bg_card'])
                cores_status = ['#E81123' if s == 'ABERTO' else '#FF8C00' if s == 'EM ANDAMENTO' else '#107C10' for s in status_counts.keys()]
                bars2 = ax2.bar(range(len(status_counts)), list(status_counts.values()), color=cores_status, alpha=0.8)
                ax2.set_xticks(range(len(status_counts)))
                ax2.set_xticklabels(status_counts.keys())
                ax2.set_ylabel("Quantidade")
                ax2.set_title("Requisições por Status")
                for bar, v in zip(bars2, status_counts.values()):
                    if v > 0:
                        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, str(v), ha='center', va='bottom')
                st.pyplot(fig2)
                plt.close(fig2)
                
                st.subheader("🚨 Requisições Críticas (Nível 1)")
                criticas = [r for r in registros if "1 -" in r.caracter]
                if criticas:
                    dados_criticos = []
                    for reg in criticas:
                        dados_criticos.append({
                            "ID": reg.id,
                            "Data": reg.data.strftime("%d/%m/%Y") if reg.data else "-",
                            "Equipamento": reg.equipamento,
                            "Setor": reg.setor,
                            "Status": reg.status
                        })
                    st.dataframe(pd.DataFrame(dados_criticos), use_container_width=True)
                else:
                    st.success("✅ Nenhuma requisição crítica no momento!")
            else:
                st.info("Nenhuma requisição encontrada para análise.")
    
    st.markdown(f"""
    <div style="text-align:right;padding:16px 0 8px;
        font-family:'JetBrains Mono',monospace;font-size:10px;
        color:{THEME['text_muted']};letter-spacing:.1em;">
        REQUISIÇÃO DE MANUTENÇÃO · {get_horario_brasilia()}
    </div>
    """, unsafe_allow_html=True)
    
# ==================================================================================================
# FECHAMENTO TURNO - VERSÃO GOOGLE SHEETS
# ==================================================================================================
elif aba_selecionada == 'FECHAMENTO TURNO':
    render_page_header("FECHAMENTO DE TURNO", f"Controle de Produção · Atualizado {get_horario_brasilia()}", THEME['accent_purple'])
    
    # ======================
    # CONFIGURAÇÕES DAS PLANILHAS ONLINE
    # ======================
    ID_PLANILHA_FECHAMENTO = '1_HKkTRCSg24wDJ47v5wSd-UPBkbalLd6plV9IvlTY64'
    ID_PLANILHA_PROGRAMACAO = '1kdsQrTH_vdEa6oqV1AVjqf8JFCAp1OybubnMHFro5lI'
    ID_PLANILHA_FALTAS = '1D4Wqixy60ZW5WPqO026rc1PTHjlVboq9ka0I3VktzDs'
    
    NOME_ABA_PRODUCAO = "PRODUÇÕES"
    NOME_ABA_PROGRAMACAO = "PRODUÇÕES"
    NOME_ABA_CHECKLIST = "CHECK"
    NOME_ABA_FALTAS = "Controle de Faltas"
    
    # ======================
    # FUNÇÕES AUXILIARES DO FECHAMENTO
    # ======================
    def excel_to_date_ft(valor):
        """Converte valor do Excel/Google Sheets para data"""
        if valor is None:
            return None
        if isinstance(valor, (datetime, pd.Timestamp)):
            return valor.date() if hasattr(valor, 'date') else valor
        if isinstance(valor, date):
            return valor
        if isinstance(valor, str):
            try:
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        return datetime.strptime(valor.split()[0], fmt).date()
                    except:
                        continue
            except:
                pass
        return None
    
    def time_to_str_ft(t):
        """Converte time para string HH:MM"""
        if t is None:
            return ""
        if isinstance(t, (datetime, time)):
            return t.strftime("%H:%M")
        return str(t) if t else ""
    
    def str_time_to_minutes_ft(time_str: str) -> int:
        """Converte string HH:MM para minutos totais"""
        try:
            if not time_str or time_str == "00:00":
                return 0
            parts = time_str.split(":")
            if len(parts) >= 2:
                hours = int(parts[0]) if parts[0].isdigit() else 0
                minutes = int(parts[1]) if parts[1].isdigit() else 0
                return hours * 60 + minutes
            return 0
        except:
            return 0
    
    # ======================
    # FUNÇÕES DE CARREGAMENTO (GOOGLE SHEETS)
    # ======================
    @st.cache_data(ttl=300)
    def carregar_producoes_fechamento(data_selecionada: date):
        """Carrega produções do Google Sheets para a data selecionada"""
        producoes = []
        try:
            client = get_gspread_client()
            if client is None:
                st.error("❌ Erro ao conectar ao Google Sheets")
                return producoes
            
            try:
                sheet = client.open_by_key(ID_PLANILHA_FECHAMENTO).worksheet(NOME_ABA_PRODUCAO)
                todos_dados = sheet.get_all_values()
                
                if len(todos_dados) < 2:
                    return producoes
                
                # Pular cabeçalho (linha 0)
                for row in todos_dados[1:]:
                    if len(row) >= 2 and row[1]:  # Coluna B (DATA)
                        data_registro = converter_data_br(row[1])
                        if data_registro and data_registro.date() == data_selecionada:
                            producoes.append({
                                'id': row[0] if len(row) > 0 else None,
                                'data': data_registro,
                                'referencia': row[2] if len(row) > 2 else "",
                                'inicio': row[3] if len(row) > 3 else "",
                                'fim': row[4] if len(row) > 4 else "",
                                'produzido': int(row[5]) if len(row) > 5 and row[5] else 0,
                                'observacoes': row[6] if len(row) > 6 else "",
                                'meta': int(row[7]) if len(row) > 7 and row[7] else 0,
                                'id_prog': row[8] if len(row) > 8 else None,
                                'justificativa': row[9] if len(row) > 9 else "",
                                'setup': row[10] if len(row) > 10 else "",
                                'manut': row[11] if len(row) > 11 else ""
                            })
            except Exception as e:
                st.warning(f"Aba '{NOME_ABA_PRODUCAO}' não encontrada na planilha de fechamento: {e}")
                
        except Exception as e:
            st.error(f"Erro ao carregar produções: {e}")
        
        return producoes
    
    @st.cache_data(ttl=300)
    def carregar_programacao_fechamento(data_selecionada: date):
        """Carrega programação PCP do Google Sheets para a data selecionada"""
        programacao = []
        try:
            client = get_gspread_client()
            if client is None:
                return programacao
            
            try:
                sheet = client.open_by_key(ID_PLANILHA_PROGRAMACAO).worksheet(NOME_ABA_PROGRAMACAO)
                todos_dados = sheet.get_all_values()
                
                if len(todos_dados) < 5:
                    return programacao
                
                # Procurar índices das colunas (cabeçalho)
                cabecalho = todos_dados[4] if len(todos_dados) > 4 else []
                idx_data = -1
                idx_praca = -1
                idx_ref = -1
                idx_meta = -1
                
                for i, col in enumerate(cabecalho):
                    if col and "DATA PRODUÇÃO" in str(col).upper():
                        idx_data = i
                    elif col and "PRAÇA" in str(col).upper():
                        idx_praca = i
                    elif col and "REFERÊNCIA DA PEÇA" in str(col).upper():
                        idx_ref = i
                    elif col and "META" in str(col).upper():
                        idx_meta = i
                
                if idx_data == -1:
                    return programacao
                
                # Processar linhas (começar da linha 6)
                for row in todos_dados[5:]:
                    if len(row) <= max(idx_data, idx_praca, idx_ref, idx_meta):
                        continue
                    if idx_praca >= 0 and not row[idx_praca]:
                        continue
                    if not row[idx_data]:
                        continue
                    
                    data_linha = converter_data_br(row[idx_data])
                    if data_linha and data_linha.date() == data_selecionada:
                        programacao.append({
                            'id_prog': row[0] if len(row) > 0 else None,
                            'praca': str(row[idx_praca]).upper().strip() if idx_praca >= 0 else "",
                            'referencia': str(row[idx_ref] or "").strip() if idx_ref >= 0 else "",
                            'meta': int(row[idx_meta]) if idx_meta >= 0 and row[idx_meta] else 0
                        })
            except Exception as e:
                st.warning(f"Aba '{NOME_ABA_PROGRAMACAO}' não encontrada na planilha de programação: {e}")
                
        except Exception as e:
            st.error(f"Erro ao carregar programação: {e}")
        
        return programacao
    
    @st.cache_data(ttl=300)
    def carregar_checklists_fechamento(data_selecionada: date):
        """Carrega checklists do Google Sheets para a data selecionada"""
        checklists = {"manha": False, "tarde": False, "noite": False}
        detalhes = []
        
        try:
            client = get_gspread_client()
            if client is None:
                return checklists, detalhes
            
            try:
                sheet = client.open_by_key(ID_PLANILHA_FECHAMENTO).worksheet(NOME_ABA_CHECKLIST)
                todos_dados = sheet.get_all_values()
                
                if len(todos_dados) < 2:
                    return checklists, detalhes
                
                for row in todos_dados[1:]:
                    if len(row) >= 2 and row[0]:
                        data_registro = converter_data_br(row[0])
                        if data_registro and data_registro.date() == data_selecionada:
                            turno = str(row[1]).lower().strip() if len(row) > 1 else ""
                            if turno in checklists:
                                checklists[turno] = True
                            detalhes.append({
                                'turno': row[1] if len(row) > 1 else "",
                                'faltas': row[2] if len(row) > 2 else "",
                                'temp_forno': row[4] if len(row) > 4 else "",
                                'temp_obs': row[5] if len(row) > 5 else "",
                                'aspecto_vidro': row[6] if len(row) > 6 else "",
                                'aspecto_obs': row[7] if len(row) > 7 else ""
                            })
            except Exception as e:
                st.warning(f"Aba '{NOME_ABA_CHECKLIST}' não encontrada na planilha de fechamento: {e}")
                
        except Exception as e:
            st.error(f"Erro ao carregar checklists: {e}")
        
        return checklists, detalhes
    
    @st.cache_data(ttl=300)
    def carregar_faltas_fechamento(data_selecionada: date):
        """Carrega faltas do Google Sheets para a data selecionada"""
        faltas = []
        try:
            client = get_gspread_client()
            if client is None:
                return faltas
            
            try:
                sheet = client.open_by_key(ID_PLANILHA_FALTAS).worksheet(NOME_ABA_FALTAS)
                todos_dados = sheet.get_all_values()
                
                if len(todos_dados) < 2:
                    return faltas
                
                for row in todos_dados[1:]:
                    if len(row) >= 6 and row[5]:  # Coluna 5 = Data
                        data_falta = converter_data_br(row[5])
                        if data_falta and data_falta.date() == data_selecionada:
                            faltas.append({
                                'id': row[0] if len(row) > 0 else "",
                                'chapa': row[1] if len(row) > 1 else "",
                                'nome': row[2] if len(row) > 2 else "",
                                'motivo': row[3] if len(row) > 3 else "",
                                'justificativa': row[6] if len(row) > 6 else ""
                            })
            except Exception as e:
                st.warning(f"Aba '{NOME_ABA_FALTAS}' não encontrada na planilha de faltas: {e}")
                
        except Exception as e:
            st.error(f"Erro ao carregar faltas: {e}")
        
        return faltas
    
    # ======================
    # INTERFACE DO FECHAMENTO TURNO
    # ======================
    
    # Seleção de data
    col_data1, col_data2 = st.columns([1, 3])
    with col_data1:
        st.markdown("#### 📅 Selecione a Data")
        data_fechamento = st.date_input(
            "Data do Fechamento",
            value=datetime.now().date(),
            key="fechamento_data"
        )
    
    with col_data2:
        st.markdown("""
        <div style="background: #e8f4fd; padding: 10px; border-radius: 8px; border-left: 4px solid #0078D4;">
            <small>ℹ️ Os dados são carregados diretamente do Google Sheets.</small>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # Carregar dados
    with st.spinner("Carregando dados do Google Sheets..."):
        producoes = carregar_producoes_fechamento(data_fechamento)
        programacao = carregar_programacao_fechamento(data_fechamento)
        checklists, checklists_detalhes = carregar_checklists_fechamento(data_fechamento)
        faltas = carregar_faltas_fechamento(data_fechamento)
    
    # KPIs do dia
    st.markdown("### 📊 Resumo do Dia")
    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    
    total_produzido = sum(p.get('produzido', 0) or 0 for p in producoes)
    total_meta = sum(p.get('meta', 0) or 0 for p in producoes)
    eficiencia = (total_produzido / total_meta * 100) if total_meta > 0 else 0
    
    total_setup_min = sum(str_time_to_minutes_ft(p.get('setup', '')) for p in producoes)
    total_manut_min = sum(str_time_to_minutes_ft(p.get('manut', '')) for p in producoes)
    
    with col_k1:
        st.metric("📦 Total Produzido", f"{total_produzido:,}".replace(",", "."))
    with col_k2:
        st.metric("🎯 Meta Total", f"{total_meta:,}".replace(",", "."))
    with col_k3:
        cor_eficiencia = "🟢" if eficiencia >= 85 else "🟡" if eficiencia >= 70 else "🔴"
        st.metric(f"{cor_eficiencia} Eficiência", f"{eficiencia:.1f}%")
    with col_k4:
        st.metric("🔧 Setup Total", minutos_para_horas_str(total_setup_min))
    with col_k5:
        st.metric("⚙️ Manutenção Total", minutos_para_horas_str(total_manut_min))
    
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # Tabs para diferentes visualizações
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Produções do Dia", "📅 Programação PCP", "✅ Checklists Turno", "🟥 Faltas"])
    
    with tab1:
        st.subheader("Registros de Produção")
        
        if producoes:
            # Preparar dados para tabela
            df_producoes = pd.DataFrame(producoes)
            
            # Verificar se as colunas existem
            colunas_necessarias = ['referencia', 'inicio', 'fim', 'produzido', 'meta', 'setup', 'manut', 'observacoes', 'justificativa']
            colunas_existentes = [col for col in colunas_necessarias if col in df_producoes.columns]
            
            df_display = df_producoes[colunas_existentes].copy()
            df_display.columns = ['Referência', 'Início', 'Fim', 'Produzido', 'Meta', 'Setup', 'Manut.', 'Observações', 'Justificativa'][:len(colunas_existentes)]
            
            # Calcular percentual
            if 'Produzido' in df_display.columns and 'Meta' in df_display.columns:
                df_display['% Meta'] = df_display.apply(
                    lambda row: (row['Produzido'] / row['Meta'] * 100) if row['Meta'] > 0 else 0, axis=1
                )
                
                # Função para colorir
                def color_eficiencia(val):
                    if isinstance(val, (int, float)):
                        if val >= 85:
                            return 'background-color: #d4f5d4'
                        elif val >= 70:
                            return 'background-color: #fff3cd'
                        else:
                            return 'background-color: #f8d7da'
                    return ''
                
                styled_df = df_display.style.applymap(
                    color_eficiencia, subset=['% Meta']
                ).format({
                    'Produzido': '{:,.0f}'.format,
                    'Meta': '{:,.0f}'.format,
                    '% Meta': '{:.1f}%'
                })
                
                st.dataframe(styled_df, use_container_width=True, height=400)
            else:
                st.dataframe(df_display, use_container_width=True, height=400)
        else:
            st.info("📭 Nenhuma produção registrada para esta data.")
    
    with tab2:
        st.subheader("Programação PCP")
        
        if programacao:
            df_programacao = pd.DataFrame(programacao)
            df_programacao = df_programacao[['praca', 'referencia', 'meta']]
            df_programacao.columns = ['Praça', 'Referência', 'Meta']
            
            # Verificar quais itens já foram produzidos
            referencias_produzidas = set(p.get('referencia', '') for p in producoes)
            
            def status_producao(row):
                if row['Referência'] in referencias_produzidas:
                    return '✅ Produzido'
                return '⏳ Pendente'
            
            df_programacao['Status'] = df_programacao.apply(status_producao, axis=1)
            
            st.dataframe(df_programacao, use_container_width=True, height=400)
        else:
            st.info("📭 Nenhuma programação encontrada para esta data.")
    
    with tab3:
        st.subheader("Checklists de Início de Turno")
        
        col_c1, col_c2, col_c3 = st.columns(3)
        
        with col_c1:
            if checklists.get("manha", False):
                st.success("✅ Turno da MANHÃ - Realizado")
            else:
                st.warning("⏳ Turno da MANHÃ - Pendente")
        
        with col_c2:
            if checklists.get("tarde", False):
                st.success("✅ Turno da TARDE - Realizado")
            else:
                st.warning("⏳ Turno da TARDE - Pendente")
        
        with col_c3:
            if checklists.get("noite", False):
                st.success("✅ Turno da NOITE - Realizado")
            else:
                st.warning("⏳ Turno da NOITE - Pendente")
        
        if checklists_detalhes:
            st.markdown("---")
            st.subheader("📋 Detalhes dos Checklists")
            for checklist in checklists_detalhes:
                with st.expander(f"📌 Checklist - Turno {checklist['turno']}"):
                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        st.write(f"**Faltas:** {checklist.get('faltas', '-')}")
                        st.write(f"**Temperatura Forno:** {checklist.get('temp_forno', '-')}")
                        if checklist.get('temp_obs'):
                            st.write(f"**Obs Temperatura:** {checklist['temp_obs']}")
                    with col_d2:
                        st.write(f"**Aspecto Vidro:** {checklist.get('aspecto_vidro', '-')}")
                        if checklist.get('aspecto_obs'):
                            st.write(f"**Obs Aspecto:** {checklist['aspecto_obs']}")
    
    with tab4:
        st.subheader("Registro de Faltas")
        
        if faltas:
            df_faltas = pd.DataFrame(faltas)
            colunas_faltas = ['chapa', 'nome', 'motivo', 'justificativa']
            colunas_existentes = [col for col in colunas_faltas if col in df_faltas.columns]
            if colunas_existentes:
                df_faltas = df_faltas[colunas_existentes]
                df_faltas.columns = ['Chapa', 'Nome', 'Motivo', 'Justificativa'][:len(colunas_existentes)]
                st.dataframe(df_faltas, use_container_width=True, height=300)
                st.markdown(f"**Total de faltas no dia:** {len(faltas)}")
        else:
            st.success("✅ Nenhuma falta registrada para esta data.")
    
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # ======================
    # GRÁFICOS E ANÁLISES
    # ======================
    st.subheader("📈 Análise do Dia")
    
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        if producoes:
            df_grafico = pd.DataFrame(producoes)
            df_grafico = df_grafico[['referencia', 'produzido', 'meta']].head(15)
            
            fig, ax = plt.subplots(figsize=(10, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "Produção vs Meta por Referência", ylabel="Quantidade")
            
            x = range(len(df_grafico))
            width = 0.35
            
            bars1 = ax.bar([i - width/2 for i in x], df_grafico['produzido'], width, label='Produzido', color=THEME['accent_cyan'], alpha=0.8)
            bars2 = ax.bar([i + width/2 for i in x], df_grafico['meta'], width, label='Meta', color=THEME['accent_orange'], alpha=0.8)
            
            ax.set_xticks(x)
            ax.set_xticklabels(df_grafico['referencia'], rotation=45, ha='right', fontsize=8)
            ax.legend(loc='upper right')
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    with col_g2:
        if producoes:
            df_eficiencia = pd.DataFrame(producoes)
            df_eficiencia['eficiencia'] = df_eficiencia.apply(
                lambda row: (row['produzido'] / row['meta'] * 100) if row['meta'] > 0 else 0, axis=1
            )
            df_eficiencia = df_eficiencia[['referencia', 'eficiencia']].head(15)
            
            fig, ax = plt.subplots(figsize=(10, 5), facecolor=THEME['bg_card'])
            apply_chart_style(ax, fig, "Eficiência por Referência", ylabel="Eficiência (%)")
            
            cores = [THEME['accent_lime'] if v >= 85 else THEME['accent_orange'] if v >= 70 else THEME['accent_red'] for v in df_eficiencia['eficiencia']]
            
            bars = ax.barh(range(len(df_eficiencia)), df_eficiencia['eficiencia'], color=cores, alpha=0.8)
            ax.axvline(x=85, color=THEME['accent_lime'], linestyle='--', alpha=0.7, label='Meta 85%')
            ax.axvline(x=70, color=THEME['accent_orange'], linestyle='--', alpha=0.7, label='Alerta 70%')
            
            ax.set_yticks(range(len(df_eficiencia)))
            ax.set_yticklabels(df_eficiencia['referencia'], fontsize=8)
            ax.set_xlim(0, 110)
            ax.legend(loc='lower right')
            
            for bar, v in zip(bars, df_eficiencia['eficiencia']):
                ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, f'{v:.1f}%', va='center', fontsize=8)
            
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
    
    st.markdown(f"""
    <div style="text-align:right;padding:16px 0 8px;
        font-family:'JetBrains Mono',monospace;font-size:10px;
        color:{THEME['text_muted']};letter-spacing:.1em;">
        FECHAMENTO DE TURNO · {get_horario_brasilia()}
    </div>
    """, unsafe_allow_html=True)
