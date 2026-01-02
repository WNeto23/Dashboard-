# dashboard_unimed_streamlit_cloud_completo.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
from datetime import datetime, timedelta
import warnings
import io
import xlsxwriter
import hashlib
import base64

warnings.filterwarnings('ignore')

# ============================================
# CONFIGURAÇÃO DA PÁGINA (OBRIGATÓRIO NO TOPO)
# ============================================
st.set_page_config(
    page_title="Dashboard Unimed - Diárias Hospitalares",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://www.unimed.com.br',
        'Report a bug': None,
        'About': """
        Dashboard de Diárias Hospitalares - Unimed
        Desenvolvido por Waltuiro Neto
        """
    }
)

# ============================================
# CONFIGURAÇÕES GLOBAIS
# ============================================
# Configuração inicial do tema
if 'tema' not in st.session_state:
    st.session_state.tema = "dark"

# Definir template Plotly baseado no tema
if st.session_state.tema == "dark":
    pio.templates.default = "plotly_dark"
else:
    pio.templates.default = "plotly_white"

# Informações do Sistema
DESENVOLVEDOR = "Waltuiro Neto - Analista de Relacionamento com a Rede"
PERIODO_BASE = "Janeiro a Novembro"
VERSAO = "9.0.0 (Streamlit Cloud Completo)"
LINK_ACESSO = "https://dashboard-unimed.streamlit.app"

# ============================================
# CSS PERSONALIZADO DINÂMICO
# ============================================
def aplicar_css():
    """Aplica CSS baseado no tema selecionado"""
    
    if st.session_state.tema == "dark":
        css = """
        <style>
            /* Modo escuro profissional */
            .main {
                background-color: #0f172a !important;
                color: #e2e8f0 !important;
                font-family: 'Segoe UI', sans-serif !important;
            }
            
            /* Cabeçalhos */
            h1, h2, h3, h4 {
                color: #ffffff !important;
                font-weight: 600 !important;
            }
            
            /* Métricas */
            .stMetric {
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%) !important;
                border-radius: 10px !important;
                padding: 15px !important;
                border-left: 4px solid !important;
                margin-bottom: 10px !important;
                border: 1px solid #334155 !important;
            }
            
            .metric-local { border-left-color: #22c55e !important; }
            .metric-intercambio { border-left-color: #3b82f6 !important; }
            .metric-total { border-left-color: #f59e0b !important; }
            .metric-pacientes { border-left-color: #8b5cf6 !important; }
            .metric-media { border-left-color: #ec4899 !important; }
            
            div[data-testid="stMetricValue"] {
                font-size: 24px !important;
                font-weight: 700 !important;
                color: white !important;
                line-height: 1.2 !important;
            }
            
            div[data-testid="stMetricLabel"] {
                font-size: 14px !important;
                color: #94a3b8 !important;
                font-weight: 500 !important;
            }
            
            .metric-subtext {
                font-size: 12px !important;
                color: #64748b !important;
                margin-top: 5px !important;
            }
            
            /* Abas */
            .stTabs [data-baseweb="tab-list"] {
                gap: 8px !important;
                background-color: #1e293b !important;
                padding: 8px !important;
                border-radius: 8px !important;
                margin-bottom: 20px !important;
                border: 1px solid #334155 !important;
            }
            
            .stTabs [data-baseweb="tab"] {
                border-radius: 6px !important;
                padding: 10px 20px !important;
                background-color: #334155 !important;
                color: #94a3b8 !important;
                font-weight: 500 !important;
                transition: all 0.3s ease !important;
                border: 1px solid #475569 !important;
            }
            
            .stTabs [aria-selected="true"] {
                background-color: #3b82f6 !important;
                color: white !important;
                border-color: #3b82f6 !important;
            }
            
            /* Sidebar */
            section[data-testid="stSidebar"] {
                background-color: #1e293b !important;
                border-right: 1px solid #334155 !important;
            }
            
            /* Cards com sombra e hover */
            .info-card {
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%) !important;
                border-radius: 10px !important;
                padding: 20px !important;
                margin-bottom: 15px !important;
                border-left: 4px solid !important;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3) !important;
                border: 1px solid #334155 !important;
                transition: all 0.3s ease !important;
            }
            
            .info-card:hover {
                box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3) !important;
                transform: translateY(-2px) !important;
            }
            
            .card-local { border-left-color: #22c55e !important; }
            .card-intercambio { border-left-color: #3b82f6 !important; }
            .card-info { border-left-color: #f59e0b !important; }
            .card-success { border-left-color: #10b981 !important; }
            
            /* Botões */
            .stButton > button {
                border-radius: 8px !important;
                font-weight: 600 !important;
                transition: all 0.3s ease !important;
                border: 1px solid #475569 !important;
            }
            
            .stButton > button:hover {
                transform: translateY(-2px) !important;
                box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
            }
            
            /* Tabelas */
            .dataframe {
                background-color: #1e293b !important;
                border-radius: 8px !important;
                overflow: hidden !important;
                border: 1px solid #334155 !important;
            }
            
            .dataframe th {
                background-color: #334155 !important;
                color: white !important;
                font-weight: 600 !important;
                border-bottom: 1px solid #475569 !important;
            }
            
            .dataframe td {
                color: #e2e8f0 !important;
                border-bottom: 1px solid #475569 !important;
            }
            
            /* Selectboxes e inputs */
            .stSelectbox, .stNumberInput, .stTextInput {
                background-color: #1e293b !important;
            }
            
            /* Status info */
            .status-info {
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%) !important;
                border-radius: 10px !important;
                padding: 15px !important;
                border-left: 4px solid #3b82f6 !important;
                margin: 10px 0 !important;
                border: 1px solid #334155 !important;
            }
            
            /* Scrollbar */
            ::-webkit-scrollbar {
                width: 8px;
                height: 8px;
            }
            
            ::-webkit-scrollbar-track {
                background: #1e293b;
            }
            
            ::-webkit-scrollbar-thumb {
                background: #475569;
                border-radius: 4px;
            }
            
            ::-webkit-scrollbar-thumb:hover {
                background: #64748b;
            }
        </style>
        """
    else:
        css = """
        <style>
            /* Modo claro profissional */
            .main {
                background-color: #f8fafc !important;
                color: #334155 !important;
                font-family: 'Segoe UI', sans-serif !important;
            }
            
            /* Cabeçalhos */
            h1, h2, h3, h4 {
                color: #1e293b !important;
                font-weight: 600 !important;
            }
            
            /* Métricas */
            .stMetric {
                background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%) !important;
                border-radius: 10px !important;
                padding: 15px !important;
                border-left: 4px solid !important;
                margin-bottom: 10px !important;
                border: 1px solid #e2e8f0 !important;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
            }
            
            .metric-local { border-left-color: #16a34a !important; }
            .metric-intercambio { border-left-color: #2563eb !important; }
            .metric-total { border-left-color: #d97706 !important; }
            .metric-pacientes { border-left-color: #7c3aed !important; }
            .metric-media { border-left-color: #db2777 !important; }
            
            div[data-testid="stMetricValue"] {
                font-size: 24px !important;
                font-weight: 700 !important;
                color: #1e293b !important;
                line-height: 1.2 !important;
            }
            
            div[data-testid="stMetricLabel"] {
                font-size: 14px !important;
                color: #64748b !important;
                font-weight: 500 !important;
            }
            
            .metric-subtext {
                font-size: 12px !important;
                color: #94a3b8 !important;
                margin-top: 5px !important;
            }
            
            /* Abas */
            .stTabs [data-baseweb="tab-list"] {
                gap: 8px !important;
                background-color: #f1f5f9 !important;
                padding: 8px !important;
                border-radius: 8px !important;
                margin-bottom: 20px !important;
                border: 1px solid #e2e8f0 !important;
            }
            
            .stTabs [data-baseweb="tab"] {
                border-radius: 6px !important;
                padding: 10px 20px !important;
                background-color: #ffffff !important;
                color: #64748b !important;
                font-weight: 500 !important;
                transition: all 0.3s ease !important;
                border: 1px solid #e2e8f0 !important;
            }
            
            .stTabs [aria-selected="true"] {
                background-color: #3b82f6 !important;
                color: white !important;
                border-color: #3b82f6 !important;
            }
            
            /* Sidebar */
            section[data-testid="stSidebar"] {
                background-color: #f1f5f9 !important;
                border-right: 1px solid #e2e8f0 !important;
            }
            
            /* Cards com sombra e hover */
            .info-card {
                background: white !important;
                border-radius: 10px !important;
                padding: 20px !important;
                margin-bottom: 15px !important;
                border-left: 4px solid !important;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
                border: 1px solid #e2e8f0 !important;
                transition: all 0.3s ease !important;
            }
            
            .info-card:hover {
                box-shadow: 0 8px 12px -2px rgba(0, 0, 0, 0.1) !important;
                transform: translateY(-2px) !important;
            }
            
            .card-local { border-left-color: #16a34a !important; }
            .card-intercambio { border-left-color: #2563eb !important; }
            .card-info { border-left-color: #d97706 !important; }
            .card-success { border-left-color: #059669 !important; }
            
            /* Botões */
            .stButton > button {
                border-radius: 8px !important;
                font-weight: 600 !important;
                transition: all 0.3s ease !important;
                border: 1px solid #e2e8f0 !important;
            }
            
            .stButton > button:hover {
                transform: translateY(-2px) !important;
                box-shadow: 0 4px 12px rgba(59, 130, 246, 0.1) !important;
            }
            
            /* Tabelas */
            .dataframe {
                background-color: white !important;
                border-radius: 8px !important;
                overflow: hidden !important;
                border: 1px solid #e2e8f0 !important;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
            }
            
            .dataframe th {
                background-color: #f8fafc !important;
                color: #1e293b !important;
                font-weight: 600 !important;
                border-bottom: 1px solid #e2e8f0 !important;
            }
            
            .dataframe td {
                color: #334155 !important;
                border-bottom: 1px solid #f1f5f9 !important;
            }
            
            /* Selectboxes e inputs */
            .stSelectbox, .stNumberInput, .stTextInput {
                background-color: white !important;
            }
            
            /* Status info */
            .status-info {
                background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%) !important;
                border-radius: 10px !important;
                padding: 15px !important;
                border-left: 4px solid #3b82f6 !important;
                margin: 10px 0 !important;
                border: 1px solid #e2e8f0 !important;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important;
            }
            
            /* Scrollbar */
            ::-webkit-scrollbar {
                width: 8px;
                height: 8px;
            }
            
            ::-webkit-scrollbar-track {
                background: #f1f5f9;
            }
            
            ::-webkit-scrollbar-thumb {
                background: #cbd5e1;
                border-radius: 4px;
            }
            
            ::-webkit-scrollbar-thumb:hover {
                background: #94a3b8;
            }
        </style>
        """
    
    st.markdown(css, unsafe_allow_html=True)

# Aplicar CSS
aplicar_css()

# ============================================
# SISTEMA DE LOGIN CORPORATIVO AVANÇADO
# ============================================
def check_password():
    """Sistema de login corporativo"""
    def password_entered():
        """Verifica se a senha está correta."""
        try:
            # Tentar usar Secrets do Streamlit Cloud
            SENHA_CORPORATIVA = st.secrets["SENHA_CORPORATIVA"]
        except:
            # Fallback para desenvolvimento local
            SENHA_CORPORATIVA = "Unimed@2024!Dashboard"
        
        # Verificar senha
        if st.session_state["password"] == SENHA_CORPORATIVA:
            st.session_state["password_correct"] = True
            st.session_state["login_time"] = datetime.now()
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False
            
            # Registrar tentativa falha
            if 'failed_attempts' not in st.session_state:
                st.session_state.failed_attempts = 0
            st.session_state.failed_attempts += 1
    
    # Se já está logado, verificar tempo de sessão (8 horas)
    if "password_correct" in st.session_state and st.session_state["password_correct"]:
        if "login_time" in st.session_state:
            session_duration = datetime.now() - st.session_state["login_time"]
            if session_duration.total_seconds() > 28800:  # 8 horas
                st.warning("⚠️ Sessão expirada. Faça login novamente.")
                del st.session_state["password_correct"]
                del st.session_state["login_time"]
                st.rerun()
        return True
    
    # Tela de login
    if "password_correct" not in st.session_state:
        col_left, col_center, col_right = st.columns([1, 2, 1])
        
        with col_center:
            st.markdown(f"""
            <div style="text-align: center; padding: 40px; background: {'linear-gradient(135deg, #1e293b 0%, #0f172a 100%)' if st.session_state.tema == 'dark' else 'linear-gradient(135deg, #f1f5f9 0%, #ffffff 100%)'}; 
                        border-radius: 15px; border: 1px solid {'#334155' if st.session_state.tema == 'dark' else '#e2e8f0'}; 
                        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2); margin: 50px 0;">
                <h1 style="color: #3b82f6; margin-bottom: 10px;">🏥 UNIMED</h1>
                <h3 style="color: {'#e2e8f0' if st.session_state.tema == 'dark' else '#334155'}; margin-bottom: 5px;">
                    Dashboard de Diárias Hospitalares
                </h3>
                <p style="color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; font-size: 14px; margin-bottom: 30px;">
                    Acesso restrito à corporação
                </p>
                
                <div style="background: {'#0f172a' if st.session_state.tema == 'dark' else '#f8fafc'}; 
                            padding: 20px; border-radius: 10px; 
                            border: 1px solid {'#334155' if st.session_state.tema == 'dark' else '#e2e8f0'};">
                    <h4 style="color: {'#e2e8f0' if st.session_state.tema == 'dark' else '#334155'}; margin-bottom: 20px;">
                        🔒 Autenticação Requerida
                    </h4>
            """, unsafe_allow_html=True)
            
            # Campo de senha
            st.text_input(
                "**Senha Corporativa:**",
                type="password",
                on_change=password_entered,
                key="password",
                label_visibility="collapsed",
                placeholder="Digite a senha de acesso..."
            )
            
            st.markdown("""
                </div>
            """, unsafe_allow_html=True)
            
            # Mensagem de erro
            if "password_correct" in st.session_state and not st.session_state["password_correct"]:
                attempts = st.session_state.get('failed_attempts', 0)
                st.error(f"❌ **Senha incorreta.** Tentativa {attempts} de 5.")
        
        return False
    
    return st.session_state["password_correct"]

# ============================================
# FUNÇÕES AUXILIARES GERAIS
# ============================================
def formatar_moeda_br(valor):
    """Formata valor monetário no padrão brasileiro: R$ 9.000,00"""
    try:
        if pd.isna(valor) or valor is None:
            return "R$ 0,00"
        
        valor_float = float(valor)
        if valor_float == 0:
            return "R$ 0,00"
        
        # Formatar com separadores de milhar e duas casas decimais
        valor_formatado = f"{valor_float:,.2f}"
        valor_formatado = valor_formatado.replace(",", "X").replace(".", ",").replace("X", ".")
        
        return f"R$ {valor_formatado}"
    except:
        return "R$ 0,00"

def formatar_inteiro_br(valor):
    """Formata inteiro no padrão brasileiro: 9.000"""
    try:
        if pd.isna(valor) or valor is None:
            return "0"
        
        valor_int = int(float(valor))
        if valor_int == 0:
            return "0"
        
        # Formatar com separadores de milhar
        valor_str = f"{abs(valor_int):,}"
        valor_str = valor_str.replace(",", ".")
        
        if valor_int < 0:
            return f"-{valor_str}"
        return valor_str
    except:
        return "0"

def obter_nome_mes(numero_mes):
    """Retorna o nome do mês a partir do número"""
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return meses.get(numero_mes, f"Mês {numero_mes}")

def obter_numero_mes(nome_mes):
    """Retorna o número do mês a partir do nome"""
    meses = {
        "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
        "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
        "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
    }
    return meses.get(nome_mes, 13)

def ordenar_meses(lista_meses):
    """Ordena lista de meses no formato 'Mês/Ano' cronologicamente"""
    def chave_ordenacao(mes_ano):
        try:
            if mes_ano == "TODAS AS COMPETÊNCIAS":
                return (0, 0)
            if '/' in mes_ano:
                mes_nome, ano = mes_ano.split('/')
                return (int(ano), obter_numero_mes(mes_nome))
        except:
            pass
        return (9999, 13)
    
    return sorted(lista_meses, key=chave_ordenacao)

def criar_categoria_ordenada(df, coluna_mes_ano):
    """
    Cria uma coluna categórica ordenada para garantir a ordenação nos gráficos
    """
    if coluna_mes_ano not in df.columns:
        return df
    
    # Criar colunas temporárias para ordenação
    df['ANO_TEMP'] = df[coluna_mes_ano].apply(lambda x: int(x.split('/')[1]) if '/' in str(x) else 0)
    df['MES_NUM_TEMP'] = df[coluna_mes_ano].apply(
        lambda x: obter_numero_mes(x.split('/')[0]) if '/' in str(x) else 13
    )
    
    # Ordenar o dataframe
    df = df.sort_values(['ANO_TEMP', 'MES_NUM_TEMP'])
    
    # Criar coluna categórica ordenada
    categorias_ordenadas = df[coluna_mes_ano].unique().tolist()
    df[coluna_mes_ano] = pd.Categorical(
        df[coluna_mes_ano], 
        categories=categorias_ordenadas, 
        ordered=True
    )
    
    # Remover colunas temporárias
    df = df.drop(columns=['ANO_TEMP', 'MES_NUM_TEMP'])
    
    return df

def alternar_tema():
    """Alterna entre tema dark e light"""
    if st.session_state.tema == "dark":
        st.session_state.tema = "light"
    else:
        st.session_state.tema = "dark"
    st.rerun()

# ============================================
# FUNÇÃO DE CARREGAMENTO DE DADOS DO CSV
# ============================================
@st.cache_data(ttl=3600, show_spinner="📊 Carregando dados do CSV...")
def carregar_dados_csv():
    """
    Carrega dados do arquivo CSV 'dados_reais.csv'
    """
    try:
        # Tentar carregar dados do CSV
        df = pd.read_csv('dados_reais.csv', encoding='utf-8')
        
        st.success(f"✅ Dados carregados com sucesso! {len(df)} registros encontrados.")
        
        # Limpeza básica
        df = df.dropna(subset=['VL_LIBERADO', 'QT_ITEM'])
        df = df[df['VL_LIBERADO'] > 0]
        df = df[df['QT_ITEM'] > 0]
        
        # Verificar colunas necessárias
        colunas_necessarias = ['NM_PRESTADOR_EXEC', 'DS_PROCEDIMENTO', 'CD_BENEFICIARIO', 'QT_ITEM', 'VL_LIBERADO']
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        
        if colunas_faltantes:
            st.warning(f"⚠️ Colunas faltantes no CSV: {', '.join(colunas_faltantes)}")
            return None
        
        # Processar competência se existir
        if 'COMPETENCIA' in df.columns:
            # Converter para string e limpar
            df['COMPETENCIA'] = df['COMPETENCIA'].astype(str).str.strip()
            
            # Extrair ano e mês do formato AAAAMM
            def extrair_ano_mes(competencia):
                try:
                    competencia_str = str(competencia).strip()
                    if len(competencia_str) == 6 and competencia_str.isdigit():
                        ano = int(competencia_str[:4])
                        mes = int(competencia_str[4:6])
                        if 1 <= mes <= 12:
                            return ano, mes
                except:
                    pass
                return 2024, np.random.randint(1, 13)  # Fallback
            
            df[['ANO_COMP', 'MES_COMP']] = df.apply(
                lambda x: pd.Series(extrair_ano_mes(x['COMPETENCIA'])), 
                axis=1
            )
            
            df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
            df['MES_ANO_FORMATADO'] = df.apply(
                lambda x: f"{obter_nome_mes(x['MES_COMP'])}/{x['ANO_COMP']}", 
                axis=1
            )
            
            df['DATA_COMPETENCIA'] = df.apply(
                lambda x: datetime(int(x['ANO_COMP']), int(x['MES_COMP']), 1), 
                axis=1
            )
        else:
            # Criar competências fictícias se não existir
            df['ANO_COMP'] = 2024
            df['MES_COMP'] = np.random.randint(1, 13, len(df))
            df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
            df['MES_ANO_FORMATADO'] = df.apply(lambda x: f"{obter_nome_mes(x['MES_COMP'])}/2024", axis=1)
            df['DATA_COMPETENCIA'] = df.apply(lambda x: datetime(2024, int(x['MES_COMP']), 1), axis=1)
        
        # Classificar LOCAL/INTERCÂMBIO
        palavras_intercambio = [
            'SÍRIO', 'SIRIO', 'LIBANÊS', 'ALBERT EINSTEIN', 'EINSTEIN',
            'MOINHOS DE VENTO', 'MOINHOS', 'HOSPITAL SÃO PAULO',
            'SÃO PAULO', 'RIO DE JANEIRO', 'BRASÍLIA', 'HCFMUSP',
            'HOSPITAL DAS CLÍNICAS', 'CLÍNICAS'
        ]
        
        def classificar_por_nome(nome):
            if pd.isna(nome):
                return 'LOCAL'
            
            nome_str = str(nome).upper()
            
            if 'UNIMED' in nome_str:
                return 'LOCAL'
            
            for palavra in palavras_intercambio:
                if palavra in nome_str:
                    return 'INTERCÂMBIO'
            
            return 'LOCAL'
        
        df['TP_PRESTADOR_CLASSIFICADO'] = df['NM_PRESTADOR_EXEC'].apply(classificar_por_nome)
        
        # Adicionar valor por diária
        df['VL_POR_DIARIA'] = df['VL_LIBERADO'] / df['QT_ITEM']
        
        # Ordenar por competência
        df = df.sort_values('DATA_COMPETENCIA')
        
        return df
        
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados do CSV: {str(e)}")
        return None

# ============================================
# SIDEBAR
# ============================================
def render_sidebar():
    """Renderiza a sidebar"""
    with st.sidebar:
        # Cabeçalho
        st.markdown(f"""
        <div style="text-align: center; margin-bottom: 20px; padding-bottom: 20px; border-bottom: 1px solid {'#334155' if st.session_state.tema == 'dark' else '#e2e8f0'};">
            <h2 style="color: #3b82f6; margin-bottom: 5px;">🏥 UNIMED</h2>
            <p style="color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; font-size: 14px; margin-bottom: 5px;">
                Dashboard de Diárias Hospitalares
            </p>
            <p style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 12px;">
                v{VERSAO} • Streamlit Cloud
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Controle de tema
        col_theme1, col_theme2 = st.columns([3, 1])
        with col_theme1:
            tema_icon = "🌙" if st.session_state.tema == "dark" else "☀️"
            tema_texto = "Modo Escuro" if st.session_state.tema == "dark" else "Modo Claro"
            st.write(f"**{tema_icon} {tema_texto}**")
        with col_theme2:
            if st.button("🔄", help="Alternar tema"):
                alternar_tema()
        
        st.markdown("---")
        
        # Status do sistema
        st.markdown("### 📊 Status do Sistema")
        
        # Carregar dados para status
        df_status = carregar_dados_csv()
        if df_status is not None:
            col_stat1, col_stat2 = st.columns(2)
            with col_stat1:
                st.metric("📈 Registros", formatar_inteiro_br(len(df_status)))
            with col_stat2:
                valor_total_status = df_status['VL_LIBERADO'].sum()
                st.metric("💰 Total", formatar_moeda_br(valor_total_status))
            
            with st.expander("ℹ️ Informações detalhadas"):
                st.write(f"**Período:** {PERIODO_BASE}")
                st.write(f"**Prestadores únicos:** {df_status['NM_PRESTADOR_EXEC'].nunique()}")
                st.write(f"**Procedimentos:** {df_status['DS_PROCEDIMENTO'].nunique()}")
                
                if 'MUNICIPIO_PRESTADOR' in df_status.columns:
                    st.write(f"**Municípios:** {df_status['MUNICIPIO_PRESTADOR'].nunique()}")
                
                # Distribuição
                local_count = len(df_status[df_status['TP_PRESTADOR_CLASSIFICADO'] == 'LOCAL'])
                intercambio_count = len(df_status[df_status['TP_PRESTADOR_CLASSIFICADO'] == 'INTERCÂMBIO'])
                total = len(df_status)
                if total > 0:
                    st.write(f"**Local:** {local_count} ({local_count/total*100:.1f}%)")
                    st.write(f"**Intercâmbio:** {intercambio_count} ({intercambio_count/total*100:.1f}%)")
        
        st.markdown("---")
        
        # Botão de logout
        if st.button("🚪 Sair do Sistema", use_container_width=True, type="secondary"):
            keys_to_clear = ['password_correct', 'login_time', 'failed_attempts']
            for key in keys_to_clear:
                if key in st.session_state:
                    del st.session_state[key]
            st.success("✅ Logout realizado com sucesso!")
            st.rerun()
        
        st.markdown("---")
        
        # Informações do desenvolvedor
        st.markdown(f"""
        <div style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 11px; text-align: center;">
            <p><strong>Desenvolvido por:</strong><br>{DESENVOLVEDOR}</p>
            <p>📅 {PERIODO_BASE} • 🕐 {datetime.now().strftime('%H:%M')}</p>
            <p>🌐 <code>{LINK_ACESSO}</code></p>
        </div>
        """, unsafe_allow_html=True)

# ============================================
# DASHBOARD PRINCIPAL COM TODOS OS GRÁFICOS
# ============================================
def dashboard_principal():
    """Dashboard principal com TODOS os gráficos"""
    
    # Cabeçalho principal
    col_header1, col_header2, col_header3 = st.columns([4, 1, 1])
    
    with col_header1:
        st.title("🏥 Dashboard de Diárias Hospitalares")
        st.markdown(f"""
        <div style="color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; font-size: 14px;">
        Análise completa com classificação LOCAL/INTERCÂMBIO • {PERIODO_BASE}
        </div>
        """, unsafe_allow_html=True)
    
    with col_header2:
        st.markdown(f"""
        <div class="status-info">
            <div style="font-size: 12px; color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'};">
                🌐 <strong>Streamlit Cloud</strong><br>
                📅 {datetime.now().strftime('%d/%m/%Y')}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_header3:
        if st.button("🔄 Atualizar", use_container_width=True, help="Recarregar dados"):
            st.cache_data.clear()
            st.rerun()
    
    # Carregar dados do CSV
    with st.spinner("📊 Carregando dados do CSV..."):
        df = carregar_dados_csv()
    
    if df is None:
        st.error("""
        ❌ **Não foi possível carregar os dados do CSV.**
        
        **Solução:**
        1. Verifique se o arquivo `dados_reais.csv` está na mesma pasta do dashboard
        2. Verifique se o CSV tem as colunas obrigatórias:
           - `NM_PRESTADOR_EXEC`
           - `DS_PROCEDIMENTO`
           - `CD_BENEFICIARIO`
           - `QT_ITEM`
           - `VL_LIBERADO`
        3. Recarregue a página
        """)
        return
    
    # ============================================
    # SEÇÃO DE FILTROS
    # ============================================
    st.markdown("## 🔧 Filtros Avançados")
    st.markdown(f"*Período base: {PERIODO_BASE}*")
    
    with st.container():
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            # FILTRO DE COMPETÊNCIA
            if 'MES_ANO_FORMATADO' in df.columns:
                competencias_lista = df['MES_ANO_FORMATADO'].unique().tolist()
                competencias_ordenadas = ordenar_meses(competencias_lista)
                competencias_opcoes = ['TODAS AS COMPETÊNCIAS'] + competencias_ordenadas
                
                competencia_selecionada = st.selectbox(
                    "📅 Competência:",
                    competencias_opcoes,
                    index=0,
                    key="filtro_competencia"
                )
            else:
                competencia_selecionada = 'TODAS AS COMPETÊNCIAS'
        
        with col2:
            # FILTRO DE TIPO
            tipo_opcoes = ['TODOS', 'LOCAL', 'INTERCÂMBIO']
            tipo_selecionado = st.selectbox(
                "🏢 Tipo:",
                tipo_opcoes,
                index=0,
                key="filtro_tipo"
            )
        
        with col3:
            # FILTRO DE MUNICÍPIO (se existir)
            if 'MUNICIPIO_PRESTADOR' in df.columns:
                df_filtro_municipio = df.copy()
                
                if competencia_selecionada != 'TODAS AS COMPETÊNCIAS':
                    df_filtro_municipio = df_filtro_municipio[df_filtro_municipio['MES_ANO_FORMATADO'] == competencia_selecionada]
                
                if tipo_selecionado != 'TODOS':
                    df_filtro_municipio = df_filtro_municipio[df_filtro_municipio['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
                
                municipios_disponiveis = sorted(df_filtro_municipio['MUNICIPIO_PRESTADOR'].unique().tolist())
                municipios_opcoes = ['TODOS OS MUNICÍPIOS'] + municipios_disponiveis
                
                municipio_selecionado = st.selectbox(
                    "📍 Município:",
                    municipios_opcoes,
                    index=0,
                    key="filtro_municipio"
                )
            else:
                municipio_selecionado = 'TODOS OS MUNICÍPIOS'
                st.info("ℹ️ Coluna de município não encontrada")
        
        with col4:
            # FILTRO DE PRESTADOR
            df_filtro_prestador = df.copy()
            
            if competencia_selecionada != 'TODAS AS COMPETÊNCIAS':
                df_filtro_prestador = df_filtro_prestador[df_filtro_prestador['MES_ANO_FORMATADO'] == competencia_selecionada]
            
            if tipo_selecionado != 'TODOS':
                df_filtro_prestador = df_filtro_prestador[df_filtro_prestador['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
            
            if municipio_selecionado != 'TODOS OS MUNICÍPIOS' and 'MUNICIPIO_PRESTADOR' in df_filtro_prestador.columns:
                df_filtro_prestador = df_filtro_prestador[df_filtro_prestador['MUNICIPIO_PRESTADOR'] == municipio_selecionado]
            
            prestadores_disponiveis = sorted(df_filtro_prestador['NM_PRESTADOR_EXEC'].unique().tolist())
            prestadores_opcoes = ['TODOS OS PRESTADORES'] + prestadores_disponiveis[:50]  # Limitar a 50
            
            prestador_selecionado = st.selectbox(
                "👨‍⚕️ Prestador:",
                prestadores_opcoes,
                index=0,
                key="filtro_prestador"
            )
    
    # Segunda linha de filtros
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        # FILTRO DE PROCEDIMENTO
        df_filtro_procedimento = df.copy()
        
        if competencia_selecionada != 'TODAS AS COMPETÊNCIAS':
            df_filtro_procedimento = df_filtro_procedimento[df_filtro_procedimento['MES_ANO_FORMATADO'] == competencia_selecionada]
        
        if tipo_selecionado != 'TODOS':
            df_filtro_procedimento = df_filtro_procedimento[df_filtro_procedimento['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
        
        if prestador_selecionado != 'TODOS OS PRESTADORES':
            df_filtro_procedimento = df_filtro_procedimento[df_filtro_procedimento['NM_PRESTADOR_EXEC'] == prestador_selecionado]
        
        procedimentos_disponiveis = sorted(df_filtro_procedimento['DS_PROCEDIMENTO'].unique().tolist())
        procedimentos_opcoes = ['TODOS OS PROCEDIMENTOS'] + procedimentos_disponiveis
        
        procedimento_selecionado = st.selectbox(
            "🩺 Procedimento:",
            procedimentos_opcoes,
            index=0,
            key="filtro_procedimento"
        )
    
    with col6:
        # FILTRO DE VALOR MÍNIMO
        valor_min = float(df['VL_LIBERADO'].min())
        valor_max = float(df['VL_LIBERADO'].max())
        
        valor_min_selecionado = st.number_input(
            "💰 Valor Mínimo (R$):",
            min_value=valor_min,
            max_value=valor_max,
            value=valor_min,
            step=10.0,
            key="filtro_valor_min"
        )
    
    with col7:
        # FILTRO DE VALOR MÁXIMO
        valor_max_selecionado = st.number_input(
            "💰 Valor Máximo (R$):",
            min_value=valor_min,
            max_value=valor_max,
            value=valor_max,
            step=10.0,
            key="filtro_valor_max"
        )
    
    with col8:
        # FILTRO DE QUANTIDADE
        qt_min = int(df['QT_ITEM'].min())
        qt_max = int(df['QT_ITEM'].max())
        
        qt_max_selecionada = st.number_input(
            "📊 Máx. Diárias:",
            min_value=qt_min,
            max_value=qt_max,
            value=qt_max,
            step=1,
            key="filtro_quantidade_max"
        )
    
    # Botões de controle
    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
    with col_btn1:
        aplicar_filtros = st.button("✅ Aplicar Filtros", use_container_width=True, type="primary")
    with col_btn2:
        if st.button("🔄 Atualizar Dados", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with col_btn3:
        if st.button("🗑️ Limpar Filtros", use_container_width=True):
            keys_to_remove = [k for k in st.session_state.keys() if k.startswith('filtro_')]
            for key in keys_to_remove:
                del st.session_state[key]
            st.rerun()
    with col_btn4:
        exportar_clicked = st.button("📤 Exportar Excel", use_container_width=True, type="secondary")
    
    # ============================================
    # APLICAR FILTROS
    # ============================================
    df_filtrado = df.copy()
    
    filtros_info = []
    
    # Aplicar filtros sequencialmente
    if competencia_selecionada != 'TODAS AS COMPETÊNCIAS' and 'MES_ANO_FORMATADO' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['MES_ANO_FORMATADO'] == competencia_selecionada]
        filtros_info.append(f"📅 {competencia_selecionada}")
    
    if tipo_selecionado != 'TODOS':
        df_filtrado = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
        filtros_info.append(f"🏢 {tipo_selecionado}")
    
    if municipio_selecionado != 'TODOS OS MUNICÍPIOS' and 'MUNICIPIO_PRESTADOR' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['MUNICIPIO_PRESTADOR'] == municipio_selecionado]
        filtros_info.append(f"📍 {municipio_selecionado}")
    
    if prestador_selecionado != 'TODOS OS PRESTADORES':
        df_filtrado = df_filtrado[df_filtrado['NM_PRESTADOR_EXEC'] == prestador_selecionado]
        filtros_info.append(f"👨‍⚕️ {prestador_selecionado[:20]}{'...' if len(prestador_selecionado) > 20 else ''}")
    
    if procedimento_selecionado != 'TODOS OS PROCEDIMENTOS':
        df_filtrado = df_filtrado[df_filtrado['DS_PROCEDIMENTO'] == procedimento_selecionado]
        filtros_info.append(f"🩺 {procedimento_selecionado[:20]}{'...' if len(procedimento_selecionado) > 20 else ''}")
    
    # Filtros numéricos
    df_filtrado = df_filtrado[
        (df_filtrado['VL_LIBERADO'] >= valor_min_selecionado) & 
        (df_filtrado['VL_LIBERADO'] <= valor_max_selecionado)
    ]
    
    df_filtrado = df_filtrado[
        (df_filtrado['QT_ITEM'] <= qt_max_selecionada)
    ]
    
    # Mostrar filtros ativos
    if filtros_info:
        st.markdown(f"""
        <div class="info-card card-info">
            <div style="display: flex; align-items: center; justify-content: space-between;">
                <div>
                    <strong>🔧 Filtros Ativos:</strong> {' • '.join(filtros_info)}
                </div>
                <div>
                    <strong>📊 Registros:</strong> {formatar_inteiro_br(len(df_filtrado))}
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # ============================================
    # KPIs PRINCIPAIS
    # ============================================
    st.markdown("## 📈 Métricas Principais")
    
    # Calcular métricas
    pacientes_unicos = df_filtrado['CD_BENEFICIARIO'].nunique()
    total_diarias = df_filtrado['QT_ITEM'].sum()
    valor_total = df_filtrado['VL_LIBERADO'].sum()
    valor_medio = df_filtrado['VL_LIBERADO'].mean() if len(df_filtrado) > 0 else 0
    
    local_df = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == 'LOCAL']
    intercambio_df = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == 'INTERCÂMBIO']
    
    valor_local = local_df['VL_LIBERADO'].sum() if len(local_df) > 0 else 0
    valor_intercambio = intercambio_df['VL_LIBERADO'].sum() if len(intercambio_df) > 0 else 0
    perc_local = (valor_local / valor_total * 100) if valor_total > 0 else 0
    perc_intercambio = (valor_intercambio / valor_total * 100) if valor_total > 0 else 0
    
    # Layout de métricas
    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    
    with col_k1:
        st.markdown(f"""
        <div class="stMetric metric-pacientes">
            <div data-testid="stMetricLabel">👥 Pacientes Únicos</div>
            <div data-testid="stMetricValue">{formatar_inteiro_br(pacientes_unicos)}</div>
            <div class="metric-subtext">
                {formatar_inteiro_br(total_diarias)} diárias totais
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k2:
        st.markdown(f"""
        <div class="stMetric metric-local">
            <div data-testid="stMetricLabel">🏥 Local</div>
            <div data-testid="stMetricValue">{formatar_moeda_br(valor_local)}</div>
            <div class="metric-subtext">
                {formatar_inteiro_br(len(local_df))} reg • {perc_local:.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k3:
        st.markdown(f"""
        <div class="stMetric metric-intercambio">
            <div data-testid="stMetricLabel">🌐 Intercâmbio</div>
            <div data-testid="stMetricValue">{formatar_moeda_br(valor_intercambio)}</div>
            <div class="metric-subtext">
                {formatar_inteiro_br(len(intercambio_df))} reg • {perc_intercambio:.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k4:
        st.markdown(f"""
        <div class="stMetric metric-total">
            <div data-testid="stMetricLabel">💰 Valor Total</div>
            <div data-testid="stMetricValue">{formatar_moeda_br(valor_total)}</div>
            <div class="metric-subtext">
                Média: {formatar_moeda_br(valor_medio)}/diária
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k5:
        media_por_paciente = valor_total / pacientes_unicos if pacientes_unicos > 0 else 0
        st.markdown(f"""
        <div class="stMetric metric-media">
            <div data-testid="stMetricLabel">📊 Média/Paciente</div>
            <div data-testid="stMetricValue">{formatar_moeda_br(media_por_paciente)}</div>
            <div class="metric-subtext">
                Diárias/paciente: {(total_diarias/pacientes_unicos):.1f}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # ============================================
    # GRÁFICO 1: EVOLUÇÃO TEMPORAL
    # ============================================
    st.markdown("---")
    st.markdown("## 📈 Evolução Temporal")
    
    if 'MES_ANO_FORMATADO' in df_filtrado.columns and df_filtrado['MES_ANO_FORMATADO'].nunique() > 1:
        analise_temporal = df_filtrado.groupby(['ANO_COMP', 'MES_COMP', 'MES_ANO_FORMATADO']).agg({
            'VL_LIBERADO': 'sum',
            'CD_BENEFICIARIO': 'nunique',
            'QT_ITEM': 'sum'
        }).reset_index()
        
        # Ordenar cronologicamente
        analise_temporal = analise_temporal.sort_values(['ANO_COMP', 'MES_COMP'])
        analise_temporal = criar_categoria_ordenada(analise_temporal, 'MES_ANO_FORMATADO')
        
        col_t1, col_t2 = st.columns(2)
        
        with col_t1:
            # Gráfico de linha - Evolução do valor
            fig_evolucao = px.line(
                analise_temporal,
                x='MES_ANO_FORMATADO',
                y='VL_LIBERADO',
                title='📈 Evolução do Valor Total por Competência',
                labels={'VL_LIBERADO': 'Valor Total (R$)', 'MES_ANO_FORMATADO': 'Competência'},
                markers=True,
                line_shape='spline',
                height=400
            )
            
            fig_evolucao.update_traces(
                line=dict(color='#3b82f6', width=3),
                marker=dict(size=8, color='#22c55e')
            )
            
            fig_evolucao.update_layout(
                xaxis_tickangle=-45
            )
            
            fig_evolucao.update_yaxes(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
            
            st.plotly_chart(fig_evolucao, use_container_width=True)
        
        with col_t2:
            # Gráfico de barras - Distribuição por tipo
            analise_tipo_temporal = df_filtrado.groupby(['MES_ANO_FORMATADO', 'TP_PRESTADOR_CLASSIFICADO']).agg({
                'VL_LIBERADO': 'sum'
            }).reset_index()
            
            analise_tipo_temporal = criar_categoria_ordenada(analise_tipo_temporal, 'MES_ANO_FORMATADO')
            
            fig_tipo_temporal = px.bar(
                analise_tipo_temporal,
                x='MES_ANO_FORMATADO',
                y='VL_LIBERADO',
                color='TP_PRESTADOR_CLASSIFICADO',
                title='🏢 Distribuição Local vs Intercâmbio',
                labels={'VL_LIBERADO': 'Valor (R$)', 'MES_ANO_FORMATADO': 'Competência'},
                color_discrete_map={'LOCAL': '#22c55e', 'INTERCÂMBIO': '#3b82f6'},
                barmode='group',
                height=400
            )
            
            fig_tipo_temporal.update_layout(
                xaxis_tickangle=-45
            )
            
            fig_tipo_temporal.update_yaxes(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
            
            st.plotly_chart(fig_tipo_temporal, use_container_width=True)
    else:
        st.info("ℹ️ Dados insuficientes para análise temporal")
    
    # ============================================
    # GRÁFICO 2: ANÁLISE POR PROCEDIMENTO
    # ============================================
    st.markdown("## 🩺 Análise por Tipo de Diária")
    
    analise_procedimento = df_filtrado.groupby('DS_PROCEDIMENTO').agg({
        'VL_LIBERADO': 'sum',
        'CD_BENEFICIARIO': 'nunique',
        'QT_ITEM': 'sum'
    }).reset_index().sort_values('VL_LIBERADO', ascending=False)
    
    col_g3, col_g4 = st.columns(2)
    
    with col_g3:
        # Treemap de procedimentos
        fig_treemap = px.treemap(
            analise_procedimento,
            path=['DS_PROCEDIMENTO'],
            values='VL_LIBERADO',
            title='🌳 Distribuição de Valor por Procedimento',
            color='VL_LIBERADO',
            color_continuous_scale='Greens',
            hover_data=['CD_BENEFICIARIO', 'QT_ITEM'],
            height=500
        )
        
        fig_treemap.update_traces(
            hovertemplate="<b>%{label}</b><br>Valor: R$ %{value:,.2f}<br>Pacientes: %{customdata[0]}<br>Diárias: %{customdata[1]}"
        )
        
        st.plotly_chart(fig_treemap, use_container_width=True)
    
    with col_g4:
        # Gráfico de barras horizontal (Top 10)
        analise_procedimento_top10 = analise_procedimento.head(10)
        
        fig_barras_proc = px.bar(
            analise_procedimento_top10,
            x='VL_LIBERADO',
            y='DS_PROCEDIMENTO',
            orientation='h',
            title='📊 Top 10 Procedimentos por Valor',
            color='VL_LIBERADO',
            color_continuous_scale='Viridis',
            text='VL_LIBERADO',
            height=500,
            hover_data=['CD_BENEFICIARIO', 'QT_ITEM']
        )
        
        fig_barras_proc.update_traces(
            texttemplate='R$ %{text:,.2f}',
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Valor: R$ %{x:,.2f}<br>Pacientes: %{customdata[0]}<br>Diárias: %{customdata[1]}'
        )
        
        fig_barras_proc.update_layout(
            yaxis={'categoryorder': 'total ascending'},
            coloraxis_showscale=False,
            xaxis=dict(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
        )
        
        st.plotly_chart(fig_barras_proc, use_container_width=True)
    
    # ============================================
    # GRÁFICO 3: RANKING DE PRESTADORES
    # ============================================
    st.markdown("## 🏆 Ranking de Prestadores")
    
    ranking = df_filtrado.groupby(['NM_PRESTADOR_EXEC', 'TP_PRESTADOR_CLASSIFICADO']).agg({
        'CD_BENEFICIARIO': 'nunique',
        'VL_LIBERADO': 'sum',
        'QT_ITEM': 'sum'
    }).reset_index()
    
    ranking['Valor Médio'] = ranking['VL_LIBERADO'] / ranking['QT_ITEM']
    ranking = ranking.sort_values('VL_LIBERADO', ascending=False)
    ranking['Pos'] = range(1, len(ranking) + 1)
    
    # Top 10 prestadores
    top_10 = ranking.head(10)
    
    col_g5, col_g6 = st.columns(2)
    
    with col_g5:
        # Gráfico de barras horizontais
        fig_barras = px.bar(
            top_10,
            x='VL_LIBERADO',
            y='NM_PRESTADOR_EXEC',
            orientation='h',
            title='🏅 Top 10 Prestadores por Valor',
            color='TP_PRESTADOR_CLASSIFICADO',
            color_discrete_map={'LOCAL': '#22c55e', 'INTERCÂMBIO': '#3b82f6'},
            text='VL_LIBERADO',
            height=500,
            hover_data=['CD_BENEFICIARIO']
        )
        
        fig_barras.update_traces(
            texttemplate='R$ %{text:,.2f}',
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Valor: R$ %{x:,.2f}<br>Pacientes: %{customdata[0]}<br>Tipo: %{marker.color}'
        )
        
        fig_barras.update_layout(
            yaxis={'categoryorder': 'total ascending'},
            showlegend=True,
            legend_title_text='Tipo',
            xaxis=dict(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
        )
        
        st.plotly_chart(fig_barras, use_container_width=True)
    
    with col_g6:
        # Scatter plot: Valor vs Pacientes
        fig_scatter = px.scatter(
            top_10,
            x='CD_BENEFICIARIO',
            y='VL_LIBERADO',
            size='QT_ITEM',
            color='TP_PRESTADOR_CLASSIFICADO',
            title='📈 Relação: Pacientes vs Valor Total',
            labels={
                'CD_BENEFICIARIO': 'Pacientes Únicos',
                'VL_LIBERADO': 'Valor Total (R$)',
                'TP_PRESTADOR_CLASSIFICADO': 'Tipo',
                'QT_ITEM': 'Diárias'
            },
            hover_name='NM_PRESTADOR_EXEC',
            size_max=60,
            height=500,
            color_discrete_map={'LOCAL': '#22c55e', 'INTERCÂMBIO': '#3b82f6'}
        )
        
        fig_scatter.update_traces(
            marker=dict(line=dict(width=1, color='DarkSlateGrey')),
            hovertemplate='<b>%{hovertext}</b><br>Pacientes: %{x}<br>Valor: R$ %{y:,.2f}<br>Diárias: %{marker.size}<br>Tipo: %{marker.color}'
        )
        
        fig_scatter.update_layout(
            showlegend=True,
            legend_title_text='Tipo',
            xaxis_title='Número de Pacientes Únicos',
            yaxis=dict(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
        )
        
        st.plotly_chart(fig_scatter, use_container_width=True)
    
    # ============================================
    # TABELAS DETALHADAS
    # ============================================
    st.markdown("---")
    st.markdown("## 📋 Tabelas Detalhadas")
    
    tab1, tab2, tab3 = st.tabs(["🏥 Ranking Completo", "🩺 Detalhes por Procedimento", "📊 Resumo por Competência"])
    
    with tab1:
        # Tabela ranking prestadores formatada
        ranking_formatado = ranking.copy()
        
        # Aplicar formatação brasileira
        ranking_formatado['VL_LIBERADO'] = ranking_formatado['VL_LIBERADO'].apply(formatar_moeda_br)
        ranking_formatado['CD_BENEFICIARIO'] = ranking_formatado['CD_BENEFICIARIO'].apply(formatar_inteiro_br)
        ranking_formatado['QT_ITEM'] = ranking_formatado['QT_ITEM'].apply(formatar_inteiro_br)
        ranking_formatado['Valor Médio'] = ranking_formatado['Valor Médio'].apply(formatar_moeda_br)
        
        st.dataframe(
            ranking_formatado[['Pos', 'NM_PRESTADOR_EXEC', 'TP_PRESTADOR_CLASSIFICADO', 
                              'VL_LIBERADO', 'CD_BENEFICIARIO', 'Valor Médio']],
            use_container_width=True,
            height=400,
            column_config={
                'Pos': 'Posição',
                'NM_PRESTADOR_EXEC': 'Prestador',
                'TP_PRESTADOR_CLASSIFICADO': 'Tipo',
                'VL_LIBERADO': 'Valor Total',
                'CD_BENEFICIARIO': 'Pacientes',
                'Valor Médio': 'Média/Diária'
            }
        )
        
        # Estatísticas da tabela
        col_stats1, col_stats2, col_stats3, col_stats4 = st.columns(4)
        with col_stats1:
            st.metric("Total Prestadores", formatar_inteiro_br(len(ranking)))
        with col_stats2:
            st.metric("Média Valor", formatar_moeda_br(ranking['VL_LIBERADO'].mean()))
        with col_stats3:
            st.metric("Mediana Valor", formatar_moeda_br(ranking['VL_LIBERADO'].median()))
        with col_stats4:
            st.metric("Desvio Padrão", formatar_moeda_br(ranking['VL_LIBERADO'].std()))
    
    with tab2:
        # Tabela procedimentos formatada
        tabela_procedimentos = df_filtrado.groupby('DS_PROCEDIMENTO').agg({
            'VL_LIBERADO': ['sum', 'mean', 'count'],
            'CD_BENEFICIARIO': 'nunique',
            'QT_ITEM': 'sum'
        }).reset_index()
        
        tabela_procedimentos.columns = [
            'Procedimento', 
            'Valor Total', 
            'Valor Médio', 
            'Qtde Registros',
            'Pacientes Únicos',
            'Total Diárias'
        ]
        
        tabela_procedimentos = tabela_procedimentos.sort_values('Valor Total', ascending=False)
        
        tabela_formatada = tabela_procedimentos.copy()
        tabela_formatada['Valor Total'] = tabela_formatada['Valor Total'].apply(formatar_moeda_br)
        tabela_formatada['Valor Médio'] = tabela_formatada['Valor Médio'].apply(formatar_moeda_br)
        tabela_formatada['Qtde Registros'] = tabela_formatada['Qtde Registros'].apply(formatar_inteiro_br)
        tabela_formatada['Pacientes Únicos'] = tabela_formatada['Pacientes Únicos'].apply(formatar_inteiro_br)
        tabela_formatada['Total Diárias'] = tabela_formatada['Total Diárias'].apply(formatar_inteiro_br)
        
        st.dataframe(
            tabela_formatada,
            use_container_width=True,
            height=400
        )
    
    with tab3:
        # Tabela competências (se disponível)
        if 'MES_ANO_FORMATADO' in df_filtrado.columns:
            tabela_competencia = df_filtrado.groupby('MES_ANO_FORMATADO').agg({
                'VL_LIBERADO': 'sum',
                'CD_BENEFICIARIO': 'nunique',
                'QT_ITEM': 'sum',
                'NM_PRESTADOR_EXEC': 'nunique',
                'DS_PROCEDIMENTO': 'nunique'
            }).reset_index()
            
            # Ordenar por competência (cronologicamente)
            tabela_competencia['ORDEM'] = tabela_competencia['MES_ANO_FORMATADO'].apply(
                lambda x: (int(x.split('/')[1]), obter_numero_mes(x.split('/')[0])) if '/' in x else (9999, 13)
            )
            tabela_competencia = tabela_competencia.sort_values('ORDEM')
            tabela_competencia = tabela_competencia.drop(columns=['ORDEM'])
            
            tabela_competencia.columns = ['Competência', 'Valor Total', 'Pacientes Únicos', 
                                          'Total Diárias', 'Prestadores Únicos', 'Procedimentos Únicos']
            
            tabela_formatada = tabela_competencia.copy()
            tabela_formatada['Valor Total'] = tabela_formatada['Valor Total'].apply(formatar_moeda_br)
            tabela_formatada['Pacientes Únicos'] = tabela_formatada['Pacientes Únicos'].apply(formatar_inteiro_br)
            tabela_formatada['Total Diárias'] = tabela_formatada['Total Diárias'].apply(formatar_inteiro_br)
            tabela_formatada['Prestadores Únicos'] = tabela_formatada['Prestadores Únicos'].apply(formatar_inteiro_br)
            tabela_formatada['Procedimentos Únicos'] = tabela_formatada['Procedimentos Únicos'].apply(formatar_inteiro_br)
            
            st.dataframe(
                tabela_formatada,
                use_container_width=True,
                height=400
            )
        else:
            st.info("ℹ️ Dados de competência não disponíveis")
    
    # ============================================
    # EXPORTAÇÃO EM EXCEL
    # ============================================
    if exportar_clicked:
        st.markdown("---")
        st.markdown("## 📤 Exportar Dados em Excel")
        
        col_e1, col_e2 = st.columns(2)
        
        with col_e1:
            # Exportar Excel com múltiplas abas
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Dados filtrados
                df_filtrado.to_excel(writer, sheet_name='Dados Filtrados', index=False)
                
                # Ranking
                ranking.to_excel(writer, sheet_name='Ranking Prestadores', index=False)
                
                # Análise por procedimento
                tabela_procedimentos.to_excel(writer, sheet_name='Análise Procedimentos', index=False)
                
                # Análise temporal (se disponível)
                if 'MES_ANO_FORMATADO' in df_filtrado.columns:
                    tabela_competencia_raw = df_filtrado.groupby('MES_ANO_FORMATADO').agg({
                        'VL_LIBERADO': 'sum',
                        'CD_BENEFICIARIO': 'nunique',
                        'QT_ITEM': 'sum'
                    }).reset_index()
                    tabela_competencia_raw.to_excel(writer, sheet_name='Análise Temporal', index=False)
                
                # Resumo executivo
                resumo_df = pd.DataFrame({
                    'Métrica': ['Registros Filtrados', 'Pacientes Únicos', 'Valor Total', 
                               'Total Diárias', 'Valor Médio por Diária', 'Valor Local',
                               'Valor Intercâmbio', 'Percentual Local', 'Percentual Intercâmbio'],
                    'Valor': [
                        len(df_filtrado),
                        pacientes_unicos,
                        valor_total,
                        total_diarias,
                        valor_medio,
                        valor_local,
                        valor_intercambio,
                        f"{perc_local:.1f}%",
                        f"{perc_intercambio:.1f}%"
                    ]
                })
                resumo_df.to_excel(writer, sheet_name='Resumo Executivo', index=False)
                
                # Formatar números no Excel
                workbook = writer.book
                money_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
                int_format = workbook.add_format({'num_format': '#,##0'})
                percent_format = workbook.add_format({'num_format': '0.0%"'})
                
                # Aplicar formatação
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    if sheet_name == 'Dados Filtrados':
                        # Encontrar coluna VL_LIBERADO
                        for i, col in enumerate(df_filtrado.columns):
                            if 'VL_' in col or 'VALOR' in col.upper():
                                worksheet.set_column(i, i, 15, money_format)
                            elif col in ['QT_ITEM', 'CD_BENEFICIARIO']:
                                worksheet.set_column(i, i, 12, int_format)
                    
                    elif sheet_name == 'Ranking Prestadores':
                        worksheet.set_column(2, 2, 20, money_format)  # VL_LIBERADO
                        worksheet.set_column(5, 5, 15, money_format)  # Valor Médio
                        worksheet.set_column(3, 3, 12, int_format)   # CD_BENEFICIARIO
                        worksheet.set_column(4, 4, 12, int_format)   # QT_ITEM
                    
                    elif sheet_name == 'Resumo Executivo':
                        worksheet.set_column(1, 1, 25)  # Coluna Valor
            
            output.seek(0)
            
            st.download_button(
                label="📊 Baixar Excel Completo",
                data=output,
                file_name=f"dashboard_unimed_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_e2:
            # Relatório executivo em texto
            relatorio = f"""
            RELATÓRIO EXECUTIVO - DASHBOARD UNIMED
            ==========================================
            
            📅 Período: {PERIODO_BASE}
            🕐 Data geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}
            👨‍💻 Gerado por: {DESENVOLVEDOR}
            
            📊 MÉTRICAS PRINCIPAIS:
            • Registros Filtrados: {formatar_inteiro_br(len(df_filtrado))}
            • Pacientes Únicos: {formatar_inteiro_br(pacientes_unicos)}
            • Total de Diárias: {formatar_inteiro_br(total_diarias)}
            • Valor Total: {formatar_moeda_br(valor_total)}
            • Valor Médio por Diária: {formatar_moeda_br(valor_medio)}
            
            🏥 DISTRIBUIÇÃO LOCAL/INTERCÂMBIO:
            • Local: {formatar_moeda_br(valor_local)} ({perc_local:.1f}%)
            • Intercâmbio: {formatar_moeda_br(valor_intercambio)} ({perc_intercambio:.1f}%)
            
            🔧 FILTROS APLICADOS:
            {', '.join(filtros_info) if filtros_info else 'Nenhum filtro aplicado'}
            
            📈 TOP 5 PRESTADORES:
            """
            
            for i, row in ranking.head(5).iterrows():
                relatorio += f"\n{i+1}. {row['NM_PRESTADOR_EXEC']}: {formatar_moeda_br(row['VL_LIBERADO'])}"
            
            relatorio += f"""
            
            📋 RESUMO:
            • Prestadores únicos: {formatar_inteiro_br(df_filtrado['NM_PRESTADOR_EXEC'].nunique())}
            • Procedimentos únicos: {formatar_inteiro_br(df_filtrado['DS_PROCEDIMENTO'].nunique())}
            """
            
            st.download_button(
                label="📝 Baixar Relatório (TXT)",
                data=relatorio,
                file_name=f"relatorio_unimed_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain",
                use_container_width=True
            )
    
    # ============================================
    # RODAPÉ
    # ============================================
    st.markdown("---")
    st.markdown(f"""
    <div style="text-align: center; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 12px; padding: 20px;">
        <p>🏥 <strong>Dashboard Unimed - Diárias Hospitalares</strong> | v{VERSAO}</p>
        <p>📅 {PERIODO_BASE} | 👨‍💻 {DESENVOLVEDOR}</p>
        <p>📊 {formatar_inteiro_br(len(df_filtrado))} registros | 💰 {formatar_moeda_br(valor_total)}</p>
        <p>🏥 Local: {formatar_moeda_br(valor_local)} ({perc_local:.1f}%) | 🌐 Intercâmbio: {formatar_moeda_br(valor_intercambio)} ({perc_intercambio:.1f}%)</p>
        <p>🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | 🌓 Tema: {st.session_state.tema.title()}</p>
    </div>
    """, unsafe_allow_html=True)

# ============================================
# EXECUÇÃO PRINCIPAL
# ============================================
if __name__ == "__main__":
    # Aplicar CSS
    aplicar_css()
    
    # Renderizar sidebar
    render_sidebar()
    
    # Verificar autenticação
    if check_password():
        # Inicializar dashboard
        dashboard_principal()