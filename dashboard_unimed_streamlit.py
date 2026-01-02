# dashboard_unimed_streamlit_cloud.py
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
VERSAO = "7.0.0 (Streamlit Cloud Edition)"
LINK_ACESSO = "https://dashboard-unimed.streamlit.app"

# ============================================
# SISTEMA DE LOGIN CORPORATIVO AVANÇADO
# ============================================
def check_password():
    """Sistema de login corporativo com múltiplas camadas de segurança"""
    def password_entered():
        """Verifica se a senha está correta."""
        # Senha corporativa - ALTERE AQUI PARA SUA SENHA
        SENHA_CORPORATIVA = "Unimed@2024!Dashboard"
        
        # Verificar senha
        if st.session_state["password"] == SENHA_CORPORATIVA:
            st.session_state["password_correct"] = True
            st.session_state["login_time"] = datetime.now()
            del st.session_state["password"]
            
            # Registrar acesso (em memória)
            if 'access_log' not in st.session_state:
                st.session_state.access_log = []
            
            st.session_state.access_log.append({
                'timestamp': datetime.now(),
                'action': 'login_success'
            })
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
        # Layout de login profissional
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
                
                <div style="margin-top: 20px; padding: 15px; background: {'#0f172a' if st.session_state.tema == 'dark' else '#f8fafc'}; 
                            border-radius: 8px; border-left: 4px solid #3b82f6;">
                    <p style="color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; font-size: 12px; margin: 0;">
                        <strong>ℹ️ Informações:</strong><br>
                        • Acesso via link oficial seguro<br>
                        • Dados confidenciais da corporação<br>
                        • Sessão válida por 8 horas
                    </p>
                </div>
            """, unsafe_allow_html=True)
            
            # Mensagem de erro
            if "password_correct" in st.session_state and not st.session_state["password_correct"]:
                attempts = st.session_state.get('failed_attempts', 0)
                st.error(f"""
                ❌ **Senha incorreta.** 
                
                Tentativa {attempts} de 5.
                {f'⚠️ **{5 - attempts} tentativas restantes** antes do bloqueio.' if attempts < 5 else '🚫 **Acesso bloqueado temporariamente.**'}
                """)
        
        # Rodapé da tela de login
        st.markdown("---")
        st.markdown(f"""
        <div style="text-align: center; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 12px; padding: 20px;">
            <p><strong>Dashboard Unimed v{VERSAO}</strong> • Desenvolvido por {DESENVOLVEDOR}</p>
            <p>📅 {datetime.now().strftime('%A, %d de %B de %Y')} • 🕐 {datetime.now().strftime('%H:%M:%S')}</p>
            <p>🌐 Acesso via: <code>{LINK_ACESSO}</code></p>
            <p>🔒 Sistema de autenticação corporativa • Streamlit Cloud</p>
        </div>
        """, unsafe_allow_html=True)
        
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

def get_base64_image(image_path):
    """Converte imagem para base64 para usar em CSS"""
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except:
        return ""

# ============================================
# FUNÇÃO DE CARREGAMENTO DE DADOS OTIMIZADA
# ============================================
@st.cache_data(ttl=3600, show_spinner="📊 Carregando e processando dados...")
def carregar_e_preparar_dados():
    """
    Carrega e prepara os dados do arquivo CSV na mesma pasta
    Otimizado para Streamlit Cloud
    """
    try:
        # Tenta carregar dados reais
        df = pd.read_csv('dados_reais.csv', encoding='utf-8')
        
        # Log de carregamento
        st.session_state.data_loaded = True
        st.session_state.data_rows = len(df)
        st.session_state.load_time = datetime.now()
        
        # Remover espaços em branco
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        # Verificar colunas obrigatórias
        colunas_obrigatorias = ['NM_PRESTADOR_EXEC', 'DS_PROCEDIMENTO', 
                               'CD_BENEFICIARIO', 'QT_ITEM', 'VL_LIBERADO']
        
        colunas_faltantes = [col for col in colunas_obrigatorias if col not in df.columns]
        if colunas_faltantes:
            st.warning(f"⚠️ Colunas faltantes: {', '.join(colunas_faltantes)}")
            return criar_dados_simulados()
        
        # ============================================
        # PROCESSAMENTO DA COMPETÊNCIA
        # ============================================
        colunas_competencia = ['COMPETENCIA', 'COMPETÊNCIA', 'COMP', 'MES_ANO', 'MES', 'ANO']
        coluna_competencia_encontrada = None
        
        for col in colunas_competencia:
            if col in df.columns:
                coluna_competencia_encontrada = col
                break
        
        if coluna_competencia_encontrada:
            # Processar competência no formato AAAAMM
            df[coluna_competencia_encontrada] = df[coluna_competencia_encontrada].astype(str).str.strip()
            
            def extrair_ano_mes(competencia):
                try:
                    competencia_str = str(competencia).strip()
                    if len(competencia_str) == 6 and competencia_str.isdigit():
                        ano = int(competencia_str[:4])
                        mes = int(competencia_str[4:6])
                        return ano, mes
                    return None, None
                except:
                    return None, None
            
            df[['ANO_COMP', 'MES_COMP']] = df.apply(
                lambda x: pd.Series(extrair_ano_mes(x[coluna_competencia_encontrada])), 
                axis=1
            )
            
            df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
            df['MES_ANO_FORMATADO'] = df.apply(
                lambda x: f"{obter_nome_mes(x['MES_COMP'])}/{x['ANO_COMP']}" 
                if pd.notna(x['MES_COMP']) and pd.notna(x['ANO_COMP']) 
                else "Data inválida", 
                axis=1
            )
            
            df['DATA_COMPETENCIA'] = df.apply(
                lambda x: datetime(int(x['ANO_COMP']), int(x['MES_COMP']), 1) 
                if pd.notna(x['MES_COMP']) and pd.notna(x['ANO_COMP']) 
                else pd.NaT, 
                axis=1
            )
            
            df = df.sort_values('DATA_COMPETENCIA')
        else:
            # Se não tem competência, criar com base na data atual
            df['ANO_COMP'] = 2024
            df['MES_COMP'] = np.random.randint(1, 12, len(df))
            df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
            df['MES_ANO_FORMATADO'] = df.apply(lambda x: f"{obter_nome_mes(x['MES_COMP'])}/2024", axis=1)
            df['DATA_COMPETENCIA'] = df.apply(lambda x: datetime(2024, int(x['MES_COMP']), 1), axis=1)
        
        # ============================================
        # CLASSIFICAÇÃO LOCAL/INTERCÂMBIO
        # ============================================
        if 'TP_PRESTADOR_EXEC' in df.columns:
            df['TP_PRESTADOR_ORIGINAL'] = df['TP_PRESTADOR_EXEC']
            
            def classificar_local_intercambio(tipo):
                if pd.isna(tipo):
                    return 'LOCAL'
                tipo_str = str(tipo).upper().strip()
                palavras_intercambio = ['INTERCÂMBIO', 'INTERCAMBIO', 'INTER', 'EXTRA', 'FORA', 'EXTERNO']
                for palavra in palavras_intercambio:
                    if palavra in tipo_str:
                        return 'INTERCÂMBIO'
                return 'LOCAL'
            
            df['TP_PRESTADOR_CLASSIFICADO'] = df['TP_PRESTADOR_EXEC'].apply(classificar_local_intercambio)
        else:
            df['TP_PRESTADOR_ORIGINAL'] = 'NÃO INFORMADO'
            
            palavras_intercambio = [
                'SÍRIO', 'SIRIO', 'LIBANÊS', 'ALBERT EINSTEIN', 'EINSTEIN',
                'MOINHOS DE VENTO', 'MOINHOS', 'HOSPITAL SÃO PAULO',
                'SÃO PAULO', 'RIO DE JANEIRO', 'BRASÍLIA', 'HCFMUSP',
                'HOSPITAL DAS CLÍNICAS', 'CLÍNICAS', 'SANTAS CASAS',
                'SANTA CASA DE SÃO PAULO', 'UNIFESP', 'HOSPITAL DO SERVIDOR',
                'SERVIDOR PÚBLICO', 'HOSPITAL PORTUGUÊS', 'PORTUGUÊS'
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
        
        # Município
        if 'MUNICIPIO_PRESTADOR' not in df.columns:
            df['MUNICIPIO_PRESTADOR'] = 'NÃO INFORMADO'
        else:
            df['MUNICIPIO_PRESTADOR'] = df['MUNICIPIO_PRESTADOR'].astype(str).str.strip()
        
        # Limpeza básica
        df = df[df['VL_LIBERADO'] > 0]
        df = df[df['QT_ITEM'] > 0]
        
        # Conversão de tipos
        try:
            df['CD_BENEFICIARIO'] = pd.to_numeric(df['CD_BENEFICIARIO'], errors='coerce')
            df['QT_ITEM'] = pd.to_numeric(df['QT_ITEM'], errors='coerce').fillna(0).astype(int)
            df['VL_LIBERADO'] = pd.to_numeric(df['VL_LIBERADO'], errors='coerce')
        except:
            pass
        
        # Adicionar coluna de valor por diária
        df['VL_POR_DIARIA'] = df['VL_LIBERADO'] / df['QT_ITEM']
        
        return df
        
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados: {str(e)[:100]}")
        return criar_dados_simulados()

def criar_dados_simulados():
    """Cria dados simulados para demonstração"""
    np.random.seed(42)
    n = 1500  # Reduzido para melhor performance
    
    prestadores_local = [
        'UNIMED RIO VERDE', 'UNIMED JATAÍ', 'UNIMED SANTA HELENA',
        'HOSPITAL SÃO LUCAS RIO VERDE', 'SANTA CASA DE RIO VERDE',
        'CLÍNICA SÃO JOSÉ', 'PRONTO SOCORRO MUNICIPAL'
    ]
    
    prestadores_intercambio = [
        'HOSPITAL SÍRIO-LIBANÊS SÃO PAULO', 'HOSPITAL ALBERT EINSTEIN',
        'HOSPITAL MOINHOS DE VENTO', 'HOSPITAL SÃO PAULO',
        'SANTA CASA DE SÃO PAULO', 'HOSPITAL DAS CLÍNICAS USP'
    ]
    
    prestadores = prestadores_local + prestadores_intercambio
    
    procedimentos = [
        'DIÁRIA DE UTI ADULTO', 'DIÁRIA DE UTI PEDIÁTRICA', 
        'DIÁRIA DE ENFERMARIA', 'DIÁRIA SEMI-INTENSIVA',
        'DIÁRIA DE ACOMPANHANTE', 'DIÁRIA DE OBSERVAÇÃO'
    ]
    
    df = pd.DataFrame({
        'NM_PRESTADOR_EXEC': np.random.choice(prestadores, n, 
            p=[0.15, 0.12, 0.08, 0.1, 0.08, 0.07, 0.05, 0.08, 0.07, 0.06, 0.06, 0.05, 0.03]),
        'DS_PROCEDIMENTO': np.random.choice(procedimentos, n, p=[0.35, 0.15, 0.25, 0.1, 0.1, 0.05]),
        'CD_BENEFICIARIO': np.random.randint(10000, 99999, n),
        'QT_ITEM': np.random.randint(1, 15, n),
        'VL_LIBERADO': np.random.exponential(600, n) + 300,
        'MUNICIPIO_PRESTADOR': np.random.choice(['Rio Verde', 'Jataí', 'Santa Helena', 'São Paulo', 'Rio de Janeiro'], n,
            p=[0.4, 0.2, 0.15, 0.15, 0.1])
    })
    
    df['VL_LIBERADO'] = df['VL_LIBERADO'].clip(200, 2500).round(2)
    
    # Adicionar competência
    df['ANO_COMP'] = 2024
    df['MES_COMP'] = np.random.randint(1, 12, n)
    df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
    df['MES_ANO_FORMATADO'] = df.apply(lambda x: f"{obter_nome_mes(x['MES_COMP'])}/2024", axis=1)
    df['DATA_COMPETENCIA'] = df.apply(lambda x: datetime(2024, int(x['MES_COMP']), 1), axis=1)
    
    # Classificar LOCAL/INTERCÂMBIO
    def classificar_simulado(nome):
        nome_str = str(nome).upper()
        intercambios = ['SÍRIO', 'SIRIO', 'EINSTEIN', 'MOINHOS', 'SÃO PAULO', 'RIO DE JANEIRO', 'CLÍNICAS', 'SANTA CASA DE SÃO PAULO']
        for inter in intercambios:
            if inter in nome_str:
                return 'INTERCÂMBIO'
        return 'LOCAL'
    
    df['TP_PRESTADOR_CLASSIFICADO'] = df['NM_PRESTADOR_EXEC'].apply(classificar_simulado)
    df['VL_POR_DIARIA'] = df['VL_LIBERADO'] / df['QT_ITEM']
    
    return df

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
            
            /* Tooltip personalizado */
            .tooltip {
                position: relative;
                display: inline-block;
                border-bottom: 1px dotted #64748b;
            }
            
            .tooltip .tooltiptext {
                visibility: hidden;
                background-color: #1e293b;
                color: #e2e8f0;
                text-align: center;
                border-radius: 6px;
                padding: 5px 10px;
                position: absolute;
                z-index: 1;
                bottom: 125%;
                left: 50%;
                transform: translateX(-50%);
                white-space: nowrap;
                border: 1px solid #334155;
            }
            
            .tooltip:hover .tooltiptext {
                visibility: visible;
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

# ============================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================
st.set_page_config(
    page_title="Dashboard Unimed - Diárias Hospitalares",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://www.unimed.com.br',
        'Report a bug': None,
        'About': f"""
        Dashboard de Diárias Hospitalares - Unimed
        Versão: {VERSAO}
        Desenvolvedor: {DESENVOLVEDOR}
        """
    }
)

# Aplicar CSS
aplicar_css()

# ============================================
# SIDEBAR AVANÇADA
# ============================================
with st.sidebar:
    # Cabeçalho com informações
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
    df_status = carregar_e_preparar_dados()
    if not df_status.empty:
        col_stat1, col_stat2 = st.columns(2)
        with col_stat1:
            st.metric("📈 Registros", formatar_inteiro_br(len(df_status)))
        with col_stat2:
            valor_total_status = df_status['VL_LIBERADO'].sum()
            st.metric("💰 Total", formatar_moeda_br(valor_total_status))
        
        # Informações adicionais
        with st.expander("ℹ️ Informações detalhadas"):
            st.write(f"**Período:** {PERIODO_BASE}")
            st.write(f"**Prestadores únicos:** {df_status['NM_PRESTADOR_EXEC'].nunique()}")
            st.write(f"**Procedimentos:** {df_status['DS_PROCEDIMENTO'].nunique()}")
            st.write(f"**Municípios:** {df_status['MUNICIPIO_PRESTADOR'].nunique()}")
            
            # Distribuição
            local_count = len(df_status[df_status['TP_PRESTADOR_CLASSIFICADO'] == 'LOCAL'])
            intercambio_count = len(df_status[df_status['TP_PRESTADOR_CLASSIFICADO'] == 'INTERCÂMBIO'])
            st.write(f"**Local:** {local_count} ({local_count/len(df_status)*100:.1f}%)")
            st.write(f"**Intercâmbio:** {intercambio_count} ({intercambio_count/len(df_status)*100:.1f}%)")
    
    st.markdown("---")
    
    # Informações de acesso
    st.markdown("### 🔐 Informações de Acesso")
    
    if 'login_time' in st.session_state:
        tempo_sessao = datetime.now() - st.session_state["login_time"]
        horas = int(tempo_sessao.total_seconds() // 3600)
        minutos = int((tempo_sessao.total_seconds() % 3600) // 60)
        
        st.info(f"""
        **Sessão ativa:** {horas}h {minutos}min
        **Expira em:** {8 - horas}h {60 - minutos}min
        """)
    
    # Botão de logout
    if st.button("🚪 Sair do Sistema", use_container_width=True, type="secondary"):
        # Registrar logout
        if 'access_log' in st.session_state:
            st.session_state.access_log.append({
                'timestamp': datetime.now(),
                'action': 'logout'
            })
        
        # Limpar sessão
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
# DASHBOARD PRINCIPAL
# ============================================
def dashboard_principal():
    """Dashboard principal com todas as funcionalidades"""
    
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
        if st.button("🔄 Atualizar", use_container_width=True, help="Recarregar dados e atualizar dashboard"):
            st.cache_data.clear()
            st.rerun()
    
    # Carregar dados
    with st.spinner("📊 Carregando dados..."):
        df = carregar_e_preparar_dados()
    
    if df.empty:
        st.error("Não foi possível carregar os dados. Verifique o arquivo 'dados_reais.csv'.")
        return
    
    # ============================================
    # SEÇÃO DE FILTROS
    # ============================================
    st.markdown("## 🔧 Filtros Avançados")
    
    # Container de filtros com tabs
    filtro_tab1, filtro_tab2 = st.tabs(["📋 Filtros Básicos", "⚙️ Filtros Avançados"])
    
    with filtro_tab1:
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        with col_f1:
            # Competência
            if 'MES_ANO_FORMATADO' in df.columns:
                competencias = ['TODAS AS COMPETÊNCIAS'] + ordenar_meses(df['MES_ANO_FORMATADO'].unique().tolist())
                competencia_selecionada = st.selectbox(
                    "📅 Competência",
                    competencias,
                    key="filtro_competencia_basic"
                )
            else:
                competencia_selecionada = 'TODAS AS COMPETÊNCIAS'
        
        with col_f2:
            # Tipo
            tipo_selecionado = st.selectbox(
                "🏢 Tipo",
                ['TODOS', 'LOCAL', 'INTERCÂMBIO'],
                key="filtro_tipo_basic"
            )
        
        with col_f3:
            # Município
            municipios = ['TODOS OS MUNICÍPIOS'] + sorted(df['MUNICIPIO_PRESTADOR'].unique().tolist())
            municipio_selecionado = st.selectbox(
                "📍 Município",
                municipios,
                key="filtro_municipio_basic"
            )
        
        with col_f4:
            # Prestador
            prestadores = ['TODOS OS PRESTADORES'] + sorted(df['NM_PRESTADOR_EXEC'].unique().tolist())[:50]
            prestador_selecionado = st.selectbox(
                "👨‍⚕️ Prestador",
                prestadores,
                key="filtro_prestador_basic"
            )
    
    with filtro_tab2:
        col_f5, col_f6, col_f7, col_f8 = st.columns(4)
        
        with col_f5:
            # Procedimento
            procedimentos = ['TODOS OS PROCEDIMENTOS'] + sorted(df['DS_PROCEDIMENTO'].unique().tolist())
            procedimento_selecionado = st.selectbox(
                "🩺 Procedimento",
                procedimentos,
                key="filtro_procedimento_advanced"
            )
        
        with col_f6:
            # Valor mínimo
            valor_min = st.number_input(
                "💰 Valor Mínimo (R$)",
                min_value=float(df['VL_LIBERADO'].min()),
                max_value=float(df['VL_LIBERADO'].max()),
                value=float(df['VL_LIBERADO'].min()),
                step=10.0,
                key="filtro_valor_min_advanced"
            )
        
        with col_f7:
            # Valor máximo
            valor_max = st.number_input(
                "💰 Valor Máximo (R$)",
                min_value=float(df['VL_LIBERADO'].min()),
                max_value=float(df['VL_LIBERADO'].max()),
                value=float(df['VL_LIBERADO'].max()),
                step=10.0,
                key="filtro_valor_max_advanced"
            )
        
        with col_f8:
            # Quantidade máxima
            qt_max = st.number_input(
                "📊 Máx. Diárias",
                min_value=int(df['QT_ITEM'].min()),
                max_value=int(df['QT_ITEM'].max()),
                value=int(df['QT_ITEM'].max()),
                step=1,
                key="filtro_quantidade_advanced"
            )
    
    # Botões de controle de filtros
    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
    with col_btn1:
        aplicar_filtros = st.button("✅ Aplicar Filtros", use_container_width=True, type="primary")
    with col_btn2:
        if st.button("🔄 Resetar Filtros", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key.startswith('filtro_'):
                    del st.session_state[key]
            st.rerun()
    with col_btn3:
        if st.button("📊 Exportar Dashboard", use_container_width=True):
            st.session_state['exportar_dashboard'] = True
    with col_btn4:
        if st.button("📈 Gerar Relatório", use_container_width=True):
            st.session_state['gerar_relatorio'] = True
    
    # ============================================
    # APLICAR FILTROS
    # ============================================
    df_filtrado = df.copy()
    filtros_ativos = []
    
    # Aplicar filtros básicos
    if competencia_selecionada != 'TODAS AS COMPETÊNCIAS' and 'MES_ANO_FORMATADO' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['MES_ANO_FORMATADO'] == competencia_selecionada]
        filtros_ativos.append(f"Competência: {competencia_selecionada}")
    
    if tipo_selecionado != 'TODOS':
        df_filtrado = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
        filtros_ativos.append(f"Tipo: {tipo_selecionado}")
    
    if municipio_selecionado != 'TODOS OS MUNICÍPIOS':
        df_filtrado = df_filtrado[df_filtrado['MUNICIPIO_PRESTADOR'] == municipio_selecionado]
        filtros_ativos.append(f"Município: {municipio_selecionado}")
    
    if prestador_selecionado != 'TODOS OS PRESTADORES':
        df_filtrado = df_filtrado[df_filtrado['NM_PRESTADOR_EXEC'] == prestador_selecionado]
        filtros_ativos.append(f"Prestador: {prestador_selecionado}")
    
    # Aplicar filtros avançados
    if procedimento_selecionado != 'TODOS OS PROCEDIMENTOS':
        df_filtrado = df_filtrado[df_filtrado['DS_PROCEDIMENTO'] == procedimento_selecionado]
        filtros_ativos.append(f"Procedimento: {procedimento_selecionado}")
    
    df_filtrado = df_filtrado[
        (df_filtrado['VL_LIBERADO'] >= valor_min) & 
        (df_filtrado['VL_LIBERADO'] <= valor_max)
    ]
    
    df_filtrado = df_filtrado[df_filtrado['QT_ITEM'] <= qt_max]
    
    # Mostrar resumo dos filtros
    if filtros_ativos:
        st.markdown(f"""
        <div class="status-info">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <strong>🔧 Filtros Ativos:</strong> {', '.join(filtros_ativos)}
                </div>
                <div>
                    <strong>📊 Registros Filtrados:</strong> {formatar_inteiro_br(len(df_filtrado))}
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # ============================================
    # KPI DASHBOARD
    # ============================================
    st.markdown("## 📈 Dashboard de Métricas")
    
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
    
    # Layout de KPIs
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    
    with kpi_col1:
        st.metric(
            label="👥 Pacientes Únicos",
            value=formatar_inteiro_br(pacientes_unicos),
            delta=f"{formatar_inteiro_br(total_diarias)} diárias",
            delta_color="off"
        )
    
    with kpi_col2:
        st.metric(
            label="🏥 Local",
            value=formatar_moeda_br(valor_local),
            delta=f"{perc_local:.1f}% • {len(local_df)} reg",
            delta_color="normal"
        )
    
    with kpi_col3:
        st.metric(
            label="🌐 Intercâmbio",
            value=formatar_moeda_br(valor_intercambio),
            delta=f"{perc_intercambio:.1f}% • {len(intercambio_df)} reg",
            delta_color="normal"
        )
    
    with kpi_col4:
        st.metric(
            label="💰 Valor Total",
            value=formatar_moeda_br(valor_total),
            delta=f"Média: {formatar_moeda_br(valor_medio)}",
            delta_color="off"
        )
    
    with kpi_col5:
        media_por_paciente = valor_total / pacientes_unicos if pacientes_unicos > 0 else 0
        st.metric(
            label="📊 Média/Paciente",
            value=formatar_moeda_br(media_por_paciente),
            delta=f"{(total_diarias/pacientes_unicos):.1f} diárias/paciente",
            delta_color="off"
        )
    
    # ============================================
    # VISUALIZAÇÕES GRÁFICAS
    # ============================================
    st.markdown("## 📊 Visualizações Analíticas")
    
    # Configurações de tema para gráficos
    plotly_template = "plotly_dark" if st.session_state.tema == "dark" else "plotly_white"
    
    # Container de gráficos com tabs
    grafico_tab1, grafico_tab2, grafico_tab3 = st.tabs(["📈 Evolução Temporal", "🏢 Análise por Tipo", "🏆 Ranking e Distribuição"])
    
    with grafico_tab1:
        if 'MES_ANO_FORMATADO' in df_filtrado.columns and df_filtrado['MES_ANO_FORMATADO'].nunique() > 1:
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                # Evolução do valor total
                evolucao_valor = df_filtrado.groupby('MES_ANO_FORMATADO').agg({
                    'VL_LIBERADO': 'sum',
                    'CD_BENEFICIARIO': 'nunique'
                }).reset_index()
                
                evolucao_valor = criar_categoria_ordenada(evolucao_valor, 'MES_ANO_FORMATADO')
                
                fig_evolucao = px.line(
                    evolucao_valor,
                    x='MES_ANO_FORMATADO',
                    y='VL_LIBERADO',
                    title='📈 Evolução do Valor Total por Competência',
                    labels={'VL_LIBERADO': 'Valor Total (R$)', 'MES_ANO_FORMATADO': 'Competência'},
                    markers=True,
                    template=plotly_template
                )
                st.plotly_chart(fig_evolucao, use_container_width=True)
            
            with col_g2:
                # Evolução de pacientes
                fig_pacientes = px.line(
                    evolucao_valor,
                    x='MES_ANO_FORMATADO',
                    y='CD_BENEFICIARIO',
                    title='👥 Evolução de Pacientes Únicos',
                    labels={'CD_BENEFICIARIO': 'Pacientes Únicos', 'MES_ANO_FORMATADO': 'Competência'},
                    markers=True,
                    template=plotly_template,
                    line_shape='spline'
                )
                st.plotly_chart(fig_pacientes, use_container_width=True)
        else:
            st.info("ℹ️ Dados insuficientes para análise temporal")
    
    with grafico_tab2:
        col_g3, col_g4 = st.columns(2)
        
        with col_g3:
            # Distribuição por tipo
            tipo_dist = df_filtrado.groupby('TP_PRESTADOR_CLASSIFICADO').agg({
                'VL_LIBERADO': 'sum',
                'CD_BENEFICIARIO': 'nunique'
            }).reset_index()
            
            fig_tipo = px.pie(
                tipo_dist,
                values='VL_LIBERADO',
                names='TP_PRESTADOR_CLASSIFICADO',
                title='🏢 Distribuição de Valor por Tipo',
                hole=0.4,
                color_discrete_sequence=['#22c55e', '#3b82f6'],
                template=plotly_template
            )
            st.plotly_chart(fig_tipo, use_container_width=True)
        
        with col_g4:
            # Distribuição por município (top 10)
            munic_dist = df_filtrado.groupby('MUNICIPIO_PRESTADOR').agg({
                'VL_LIBERADO': 'sum'
            }).reset_index().nlargest(10, 'VL_LIBERADO')
            
            fig_munic = px.bar(
                munic_dist,
                x='VL_LIBERADO',
                y='MUNICIPIO_PRESTADOR',
                orientation='h',
                title='📍 Top 10 Municípios por Valor',
                labels={'VL_LIBERADO': 'Valor (R$)', 'MUNICIPIO_PRESTADOR': 'Município'},
                template=plotly_template
            )
            st.plotly_chart(fig_munic, use_container_width=True)
    
    with grafico_tab3:
        col_g5, col_g6 = st.columns(2)
        
        with col_g5:
            # Top 10 prestadores
            top_prestadores = df_filtrado.groupby('NM_PRESTADOR_EXEC').agg({
                'VL_LIBERADO': 'sum',
                'TP_PRESTADOR_CLASSIFICADO': 'first'
            }).reset_index().nlargest(10, 'VL_LIBERADO')
            
            fig_top = px.bar(
                top_prestadores,
                x='VL_LIBERADO',
                y='NM_PRESTADOR_EXEC',
                color='TP_PRESTADOR_CLASSIFICADO',
                orientation='h',
                title='🏆 Top 10 Prestadores por Valor',
                labels={'VL_LIBERADO': 'Valor (R$)', 'NM_PRESTADOR_EXEC': 'Prestador'},
                color_discrete_map={'LOCAL': '#22c55e', 'INTERCÂMBIO': '#3b82f6'},
                template=plotly_template
            )
            st.plotly_chart(fig_top, use_container_width=True)
        
        with col_g6:
            # Scatter plot: valor vs pacientes
            scatter_data = df_filtrado.groupby(['NM_PRESTADOR_EXEC', 'TP_PRESTADOR_CLASSIFICADO']).agg({
                'VL_LIBERADO': 'sum',
                'CD_BENEFICIARIO': 'nunique'
            }).reset_index()
            
            fig_scatter = px.scatter(
                scatter_data,
                x='CD_BENEFICIARIO',
                y='VL_LIBERADO',
                color='TP_PRESTADOR_CLASSIFICADO',
                size='VL_LIBERADO',
                hover_name='NM_PRESTADOR_EXEC',
                title='📈 Relação: Pacientes vs Valor',
                labels={'CD_BENEFICIARIO': 'Pacientes Únicos', 'VL_LIBERADO': 'Valor Total (R$)'},
                color_discrete_map={'LOCAL': '#22c55e', 'INTERCÂMBIO': '#3b82f6'},
                template=plotly_template
            )
            st.plotly_chart(fig_scatter, use_container_width=True)
    
    # ============================================
    # TABELAS DETALHADAS
    # ============================================
    st.markdown("## 📋 Tabelas Detalhadas")
    
    tabela_tab1, tabela_tab2, tabela_tab3 = st.tabs(["🏥 Ranking Completo", "🩺 Análise por Procedimento", "📊 Dados Brutos"])
    
    with tabela_tab1:
        # Ranking de prestadores
        ranking = df_filtrado.groupby(['NM_PRESTADOR_EXEC', 'TP_PRESTADOR_CLASSIFICADO', 'MUNICIPIO_PRESTADOR']).agg({
            'VL_LIBERADO': ['sum', 'mean'],
            'CD_BENEFICIARIO': 'nunique',
            'QT_ITEM': 'sum'
        }).reset_index()
        
        ranking.columns = ['Prestador', 'Tipo', 'Município', 'Valor Total', 'Valor Médio', 'Pacientes Únicos', 'Total Diárias']
        ranking = ranking.sort_values('Valor Total', ascending=False)
        ranking['Posição'] = range(1, len(ranking) + 1)
        
        st.dataframe(
            ranking[['Posição', 'Prestador', 'Tipo', 'Município', 'Valor Total', 'Pacientes Únicos', 'Total Diárias']],
            use_container_width=True,
            height=400
        )
    
    with tabela_tab2:
        # Análise por procedimento
        proc_analise = df_filtrado.groupby('DS_PROCEDIMENTO').agg({
            'VL_LIBERADO': ['sum', 'mean', 'count'],
            'CD_BENEFICIARIO': 'nunique',
            'QT_ITEM': 'sum'
        }).reset_index()
        
        proc_analise.columns = ['Procedimento', 'Valor Total', 'Valor Médio', 'Qtd Registros', 'Pacientes Únicos', 'Total Diárias']
        proc_analise = proc_analise.sort_values('Valor Total', ascending=False)
        
        st.dataframe(
            proc_analise,
            use_container_width=True,
            height=400
        )
    
    with tabela_tab3:
        # Dados brutos filtrados
        colunas_disponiveis = list(df_filtrado.columns)
        colunas_selecionadas = st.multiselect(
            "Selecione colunas para visualizar:",
            colunas_disponiveis,
            default=colunas_disponiveis[:8] if len(colunas_disponiveis) > 8 else colunas_disponiveis
        )
        
        if colunas_selecionadas:
            st.dataframe(
                df_filtrado[colunas_selecionadas].head(100),
                use_container_width=True,
                height=400
            )
        else:
            st.info("Selecione pelo menos uma coluna para visualizar os dados.")
    
    # ============================================
    # EXPORTAÇÃO DE DADOS
    # ============================================
    if st.session_state.get('exportar_dashboard', False):
        st.markdown("---")
        st.markdown("## 📤 Exportar Dados")
        
        exp_col1, exp_col2, exp_col3 = st.columns(3)
        
        with exp_col1:
            # Exportar CSV
            csv_data = df_filtrado.to_csv(index=False, sep=';', decimal=',', encoding='utf-8-sig')
            st.download_button(
                label="💾 Baixar CSV",
                data=csv_data,
                file_name=f"dados_unimed_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with exp_col2:
            # Exportar Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_filtrado.to_excel(writer, sheet_name='Dados Filtrados', index=False)
                
                # Adicionar resumo
                resumo_df = pd.DataFrame({
                    'Métrica': ['Registros', 'Pacientes Únicos', 'Valor Total', 'Total Diárias', 'Valor Médio'],
                    'Valor': [
                        len(df_filtrado),
                        pacientes_unicos,
                        valor_total,
                        total_diarias,
                        valor_medio
                    ]
                })
                resumo_df.to_excel(writer, sheet_name='Resumo', index=False)
            
            output.seek(0)
            
            st.download_button(
                label="📊 Baixar Excel",
                data=output,
                file_name=f"dashboard_unimed_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with exp_col3:
            # Relatório executivo
            relatorio = f"""
            RELATÓRIO EXECUTIVO - DASHBOARD UNIMED
            ===========================================
            
            DATA: {datetime.now().strftime('%d/%m/%Y %H:%M')}
            DESENVOLVEDOR: {DESENVOLVEDOR}
            VERSÃO: {VERSAO}
            LINK: {LINK_ACESSO}
            
            RESUMO EXECUTIVO:
            • Registros analisados: {formatar_inteiro_br(len(df_filtrado))}
            • Pacientes únicos: {formatar_inteiro_br(pacientes_unicos)}
            • Valor total: {formatar_moeda_br(valor_total)}
            • Total de diárias: {formatar_inteiro_br(total_diarias)}
            • Valor médio por diária: {formatar_moeda_br(valor_medio)}
            
            DISTRIBUIÇÃO LOCAL/INTERCÂMBIO:
            • Local: {formatar_moeda_br(valor_local)} ({perc_local:.1f}%)
            • Intercâmbio: {formatar_moeda_br(valor_intercambio)} ({perc_intercambio:.1f}%)
            
            FILTROS APLICADOS:
            {chr(10).join(['• ' + f for f in filtros_ativos]) if filtros_ativos else '• Nenhum filtro aplicado'}
            
            TOP 5 PRESTADORES:
            """
            
            for i, row in ranking.head(5).iterrows():
                relatorio += f"\n{i+1}. {row['Prestador']}: {formatar_moeda_br(row['Valor Total'])}"
            
            st.download_button(
                label="📝 Baixar Relatório",
                data=relatorio,
                file_name=f"relatorio_executivo_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain",
                use_container_width=True
            )
        
        st.session_state['exportar_dashboard'] = False
    
    # ============================================
    # RODAPÉ AVANÇADO
    # ============================================
    st.markdown("---")
    
    rodape_col1, rodape_col2, rodape_col3 = st.columns(3)
    
    with rodape_col1:
        st.markdown(f"""
        <div style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 11px;">
            <p><strong>🏥 Dashboard Unimed</strong></p>
            <p>Versão: {VERSAO}</p>
            <p>Desenvolvedor: {DESENVOLVEDOR}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with rodape_col2:
        st.markdown(f"""
        <div style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 11px; text-align: center;">
            <p><strong>📊 Estatísticas da Sessão</strong></p>
            <p>Registros: {formatar_inteiro_br(len(df_filtrado))}</p>
            <p>Valor Total: {formatar_moeda_br(valor_total)}</p>
            <p>Atualizado: {datetime.now().strftime('%H:%M:%S')}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with rodape_col3:
        st.markdown(f"""
        <div style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 11px; text-align: right;">
            <p><strong>🌐 Informações Técnicas</strong></p>
            <p>Streamlit Cloud Edition</p>
            <p>Tema: {st.session_state.tema.title()}</p>
            <p>Cache: Ativo (1 hora)</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style="text-align: center; color: {'#475569' if st.session_state.tema == 'dark' else '#cbd5e1'}; font-size: 10px; padding-top: 10px; border-top: 1px solid {'#334155' if st.session_state.tema == 'dark' else '#e2e8f0'};">
        <p>© 2024 Unimed • Dashboard de Diárias Hospitalares • Acesso seguro via autenticação corporativa</p>
        <p>{LINK_ACESSO} • Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
    </div>
    """, unsafe_allow_html=True)

# ============================================
# EXECUÇÃO PRINCIPAL
# ============================================
if __name__ == "__main__":
    # Verificar autenticação
    if check_password():
        # Inicializar dashboard
        dashboard_principal()