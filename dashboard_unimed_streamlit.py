# dashboard_unimed_streamlit_cloud_final_v12.py
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
import os
from pathlib import Path

warnings.filterwarnings('ignore')

# ============================================
# 1. CONFIGURAÇÃO DA PÁGINA (NO TOPO!)
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
# 2. CONFIGURAÇÕES GLOBAIS E SESSÃO
# ============================================
# Configuração inicial do tema
if 'tema' not in st.session_state:
    st.session_state.tema = "dark"

if 'auth_sucesso' not in st.session_state:
    st.session_state.auth_sucesso = False

if 'login_time' not in st.session_state:
    st.session_state.login_time = None

# Definir template Plotly baseado no tema
if st.session_state.tema == "dark":
    pio.templates.default = "plotly_dark"
else:
    pio.templates.default = "plotly_white"

# Informações do Sistema
DESENVOLVEDOR = "Waltuiro Neto - Analista de Relacionamento com a Rede"
PERIODO_BASE = "Janeiro a Novembro"
VERSAO = "12.0.0 (Streamlit Cloud Otimizado - Completo)"
LINK_ACESSO = "https://dashboard-unimed.streamlit.app"

# ============================================
# 3. FUNÇÕES DE SUPORTE - CAMINHOS E DADOS
# ============================================
def obter_caminho_arquivo(nome_arquivo):
    """
    Localiza o arquivo no ambiente local ou GitHub/Streamlit Cloud.
    Tenta múltiplos caminhos possíveis.
    """
    # Lista de caminhos possíveis para tentar
    caminhos_tentativas = [
        nome_arquivo,  # Raiz do projeto
        os.path.join(os.path.dirname(__file__), nome_arquivo),  # Pasta do script
        os.path.join(os.path.dirname(__file__), "data", nome_arquivo),  # Subpasta data
        os.path.join(os.path.dirname(__file__), "dados", nome_arquivo),  # Subpasta dados
        os.path.join(Path(__file__).parent.parent, nome_arquivo),  # Pasta pai
    ]
    
    for caminho in caminhos_tentativas:
        if os.path.exists(caminho):
            return caminho
    
    return None

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

def alternar_tema():
    """Alterna entre tema dark e light"""
    if st.session_state.tema == "dark":
        st.session_state.tema = "light"
    else:
        st.session_state.tema = "dark"
    st.rerun()

# ============================================
# 4. CLASSIFICAÇÃO DE PRESTADORES - LÓGICA DO CÓDIGO 2
# ============================================
def classificar_prestador_local_intercambio(df):
    """
    Classifica prestadores como LOCAL ou INTERCÂMBIO com base no nome do prestador
    Retorna DataFrame com coluna 'TP_PRESTADOR_CLASSIFICADO'
    """
    # Lista corrigida de palavras-chave para intercâmbio
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
        
        # Primeiro, verificar se é claramente local (Unimed)
        if 'UNIMED' in nome_str:
            return 'LOCAL'
        
        # Verificar se é intercâmbio
        for palavra in palavras_intercambio:
            if palavra in nome_str:
                return 'INTERCÂMBIO'
        
        # Se não for nenhum dos casos acima, considerar LOCAL
        return 'LOCAL'
    
    # Criar coluna de classificação
    df['TP_PRESTADOR_CLASSIFICADO'] = df['NM_PRESTADOR_EXEC'].apply(classificar_por_nome)
    
    return df

# ============================================
# 5. CARREGAMENTO DE DADOS COM CACHE
# ============================================
@st.cache_data(ttl=3600, show_spinner="📊 Carregando e processando dados...")
def carregar_e_preparar_dados():
    """
    Carrega e prepara os dados com fallback para dados simulados
    se o CSV não for encontrado. Inclui classificação LOCAL/INTERCÂMBIO.
    """
    # Primeiro, tentar carregar dados reais
    caminho_csv = obter_caminho_arquivo('dados_reais.csv')
    
    if caminho_csv:
        try:
            st.info(f"✅ Carregando dados de: {os.path.basename(caminho_csv)}")
            df = pd.read_csv(caminho_csv, encoding='utf-8')
            
            # Limpeza básica
            df.columns = [c.strip().upper() for c in df.columns]
            
            # Verificar colunas obrigatórias
            colunas_obrigatorias = ['NM_PRESTADOR_EXEC', 'DS_PROCEDIMENTO', 
                                   'CD_BENEFICIARIO', 'QT_ITEM', 'VL_LIBERADO']
            
            colunas_faltantes = [col for col in colunas_obrigatorias if col not in df.columns]
            if colunas_faltantes:
                st.warning(f"⚠️ Colunas faltantes: {', '.join(colunas_faltantes)}")
                return criar_dados_simulados()
            
            # Processar competência se existir
            if 'COMPETENCIA' in df.columns:
                df['COMPETENCIA'] = df['COMPETENCIA'].astype(str).str.strip()
                
                def extrair_ano_mes(competencia):
                    try:
                        competencia_str = str(competencia).strip()
                        if len(competencia_str) == 6 and competencia_str.isdigit():
                            ano = int(competencia_str[:4])
                            mes = int(competencia_str[4:6])
                            return ano, mes
                        return 2024, np.random.randint(1, 13)
                    except:
                        return 2024, np.random.randint(1, 13)
                
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
                # Criar competências fictícias
                df['ANO_COMP'] = 2024
                df['MES_COMP'] = np.random.randint(1, 13, len(df))
                df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
                df['MES_ANO_FORMATADO'] = df.apply(
                    lambda x: f"{obter_nome_mes(x['MES_COMP'])}/2024", 
                    axis=1
                )
                df['DATA_COMPETENCIA'] = df.apply(lambda x: datetime(2024, int(x['MES_COMP']), 1), axis=1)
            
            # Aplicar classificação LOCAL/INTERCÂMBIO
            df = classificar_prestador_local_intercambio(df)
            
            # Adicionar valor por diária
            df['VL_POR_DIARIA'] = df['VL_LIBERADO'] / df['QT_ITEM']
            
            return df
            
        except Exception as e:
            st.error(f"❌ Erro ao processar CSV: {str(e)[:200]}")
            return criar_dados_simulados()
    else:
        # Se não encontrar CSV, usar dados simulados
        st.info("📊 Usando dados simulados para demonstração.")
        st.info("💡 Para usar dados reais, adicione o arquivo 'dados_reais.csv' ao seu repositório.")
        return criar_dados_simulados()

def criar_dados_simulados():
    """Cria dados simulados para demonstração com classificação LOCAL/INTERCÂMBIO"""
    np.random.seed(42)
    n = 2000
    
    # Prestadores realistas (incluindo intercâmbios)
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
        'QT_ITEM': np.random.randint(1, 20, n),
        'VL_LIBERADO': np.random.exponential(800, n) + 200,
        'MUNICIPIO_PRESTADOR': np.random.choice(['Rio Verde', 'Jataí', 'Santa Helena', 'São Paulo', 'Rio de Janeiro'], n,
            p=[0.4, 0.2, 0.15, 0.15, 0.1])
    })
    
    df['VL_LIBERADO'] = df['VL_LIBERADO'].clip(150, 3000).round(2)
    
    # Adicionar competência
    df['ANO_COMP'] = 2024
    df['MES_COMP'] = np.random.randint(1, 13, n)
    df['MES_NOME'] = df['MES_COMP'].apply(obter_nome_mes)
    df['MES_ANO_FORMATADO'] = df.apply(lambda x: f"{obter_nome_mes(x['MES_COMP'])}/2024", axis=1)
    df['DATA_COMPETENCIA'] = df.apply(lambda x: datetime(2024, int(x['MES_COMP']), 1), axis=1)
    
    # Aplicar classificação LOCAL/INTERCÂMBIO
    df = classificar_prestador_local_intercambio(df)
    
    # Adicionar valor por diária
    df['VL_POR_DIARIA'] = df['VL_LIBERADO'] / df['QT_ITEM']
    
    return df

# ============================================
# 6. CSS PERSONALIZADO DINÂMICO
# ============================================
def aplicar_estilo():
    """Aplica CSS baseado no tema selecionado"""
    tema = st.session_state.tema
    
    if tema == "dark":
        bg_color = "#0f172a"
        card_bg = "linear-gradient(135deg, #1e293b 0%, #0f172a 100%)"
        border_color = "#334155"
        text_color = "#ffffff"
        text_secondary = "#94a3b8"
        success_color = "#22c55e"
        intercambio_color = "#3b82f6"
        total_color = "#f59e0b"
    else:
        bg_color = "#f8fafc"
        card_bg = "#ffffff"
        border_color = "#e2e8f0"
        text_color = "#1e293b"
        text_secondary = "#64748b"
        success_color = "#16a34a"
        intercambio_color = "#2563eb"
        total_color = "#d97706"
    
    css = f"""
    <style>
        /* Fundo principal */
        .stApp {{
            background-color: {bg_color} !important;
            color: {text_color} !important;
            font-family: 'Segoe UI', sans-serif !important;
        }}
        
        /* Cabeçalhos */
        h1, h2, h3, h4 {{
            color: {text_color} !important;
            font-weight: 600 !important;
        }}
        
        /* Container de autenticação */
        .auth-container {{
            background: {card_bg};
            padding: 40px;
            border-radius: 15px;
            border: 1px solid {border_color};
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
            text-align: center;
            max-width: 500px;
            margin: 50px auto;
        }}
        
        /* Métricas - estilo profissional */
        .metric-card {{
            background: {card_bg};
            border-radius: 10px;
            padding: 20px;
            border-left: 5px solid #3b82f6;
            border: 1px solid {border_color};
            margin-bottom: 15px;
            height: 120px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        
        .metric-local {{ border-left-color: {success_color} !important; }}
        .metric-intercambio {{ border-left-color: {intercambio_color} !important; }}
        .metric-total {{ border-left-color: {total_color} !important; }}
        .metric-pacientes {{ border-left-color: #8b5cf6 !important; }}
        .metric-media {{ border-left-color: #ec4899 !important; }}
        
        /* Cards info */
        .info-card {{
            background: {card_bg};
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 15px;
            border-left: 4px solid;
            border: 1px solid {border_color};
            transition: all 0.3s ease;
        }}
        
        .info-card:hover {{
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.2);
            transform: translateY(-2px);
        }}
        
        .card-local {{ border-left-color: {success_color} !important; }}
        .card-intercambio {{ border-left-color: {intercambio_color} !important; }}
        .card-info {{ border-left-color: {total_color} !important; }}
        .card-success {{ border-left-color: #10b981 !important; }}
        
        /* Tabelas */
        .dataframe {{
            background-color: {card_bg} !important;
            border-radius: 8px !important;
            border: 1px solid {border_color} !important;
        }}
        
        .dataframe th {{
            background-color: {border_color} !important;
            color: {text_color} !important;
            font-weight: 600 !important;
        }}
        
        .dataframe td {{
            color: {text_color} !important;
            border-bottom: 1px solid {border_color} !important;
        }}
        
        /* Botões */
        .stButton > button {{
            border-radius: 8px !important;
            font-weight: 600 !important;
            transition: all 0.3s ease !important;
            border: 1px solid {border_color} !important;
        }}
        
        .stButton > button:hover {{
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
        }}
        
        /* Sidebar */
        section[data-testid="stSidebar"] {{
            background-color: {card_bg} !important;
            border-right: 1px solid {border_color} !important;
        }}
        
        /* Status info */
        .status-info {{
            background: {card_bg};
            border-radius: 10px;
            padding: 15px;
            border-left: 4px solid #3b82f6;
            margin: 10px 0;
            border: 1px solid {border_color};
        }}
        
        /* Textos secundários */
        .text-secondary {{
            color: {text_secondary} !important;
        }}
        
        /* Abas */
        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px !important;
            background-color: {border_color} !important;
            padding: 8px !important;
            border-radius: 8px !important;
            margin-bottom: 20px !important;
        }}
        
        .stTabs [data-baseweb="tab"] {{
            border-radius: 6px !important;
            padding: 10px 20px !important;
            background-color: {card_bg} !important;
            color: {text_secondary} !important;
            font-weight: 500 !important;
            transition: all 0.3s ease !important;
            border: 1px solid {border_color} !important;
        }}
        
        .stTabs [aria-selected="true"] {{
            background-color: #3b82f6 !important;
            color: white !important;
            border-color: #3b82f6 !important;
        }}
        
        /* Scrollbar */
        ::-webkit-scrollbar {{
            width: 8px;
            height: 8px;
        }}
        
        ::-webkit-scrollbar-track {{
            background: {bg_color};
        }}
        
        ::-webkit-scrollbar-thumb {{
            background: {border_color};
            border-radius: 4px;
        }}
        
        ::-webkit-scrollbar-thumb:hover {{
            background: {text_secondary};
        }}
    </style>
    """
    
    st.markdown(css, unsafe_allow_html=True)

# ============================================
# 7. SISTEMA DE LOGIN
# ============================================
def tela_login():
    """Exibe a tela de login"""
    aplicar_estilo()
    
    # Layout de login centralizado
    col_left, col_central, col_right = st.columns([1, 2, 1])
    
    with col_central:
        # Container de autenticação
        st.markdown(f"""
        <div class="auth-container">
            <h1 style="color: #3b82f6; margin-bottom: 10px;">🏥 UNIMED</h1>
            <h3 style="margin-bottom: 5px;">Dashboard de Diárias Hospitalares</h3>
            <p style="color: #94a3b8; font-size: 14px; margin-bottom: 30px;">
                Acesso restrito à corporação • Versão {VERSAO}
            </p>
            <h4 style="margin-bottom: 20px;">🔒 Autenticação Requerida</h4>
        </div>
        """, unsafe_allow_html=True)
        
        # Campo de senha (FORA do HTML para funcionar corretamente)
        senha_digitada = st.text_input(
            "Senha Corporativa",
            type="password",
            label_visibility="collapsed",
            placeholder="Digite a senha de acesso...",
            key="senha_login"
        )
        
        # Botão de login
        if st.button("✅ Entrar no Sistema", use_container_width=True, type="primary"):
            # Usar Secrets do Streamlit Cloud ou fallback
            try:
                senha_correta = st.secrets["SENHA_CORPORATIVA"]
            except:
                senha_correta = "Diarias@202!Dashboard"
                st.warning("⚠️ Usando senha padrão. Configure os Secrets no Streamlit Cloud.")
            
            if senha_digitada == senha_correta:
                st.session_state.auth_sucesso = True
                st.session_state.login_time = datetime.now()
                st.success("✅ Login realizado com sucesso!")
                st.rerun()
            else:
                st.error("❌ Senha incorreta. Tente novamente.")
        
        # Informações adicionais
        st.markdown("""
        <div style="text-align: center; color: #64748b; font-size: 12px; margin-top: 20px;">
            <p>👨‍💻 Desenvolvido por: Waltuiro Neto</p>
            <p>📅 Período base: Janeiro a Novembro</p>
            <p>🌐 Acesso via: dashboard-unimed.streamlit.app</p>
        </div>
        """, unsafe_allow_html=True)

# ============================================
# 8. DASHBOARD PRINCIPAL COMPLETO
# ============================================
def mostrar_dashboard():
    """Exibe o dashboard principal completo"""
    aplicar_estilo()
    
    # ============================================
    # SIDEBAR
    # ============================================
    with st.sidebar:
        # Logo e informações
        st.markdown(f"""
        <div style="text-align: center; margin-bottom: 20px;">
            <h2 style="color: #3b82f6; margin-bottom: 5px;">🏥 UNIMED</h2>
            <p style="color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; font-size: 14px;">
                Dashboard de Diárias
            </p>
            <p style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 12px;">
                v{VERSAO} • Streamlit Cloud
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Controle de tema
        st.markdown("---")
        col_theme1, col_theme2 = st.columns([3, 1])
        with col_theme1:
            tema_icon = "🌙" if st.session_state.tema == "dark" else "☀️"
            tema_texto = "Modo Escuro" if st.session_state.tema == "dark" else "Modo Claro"
            st.write(f"**{tema_icon} {tema_texto}**")
        with col_theme2:
            if st.button("🔄", help="Alternar tema", use_container_width=True):
                alternar_tema()
        
        # Status do sistema
        st.markdown("---")
        st.markdown("### 📊 Status do Sistema")
        
        # Carregar dados para status
        df_status = carregar_e_preparar_dados()
        if df_status is not None and not df_status.empty:
            col_stat1, col_stat2 = st.columns(2)
            with col_stat1:
                st.metric("📈 Registros", formatar_inteiro_br(len(df_status)))
            with col_stat2:
                valor_total_status = df_status['VL_LIBERADO'].sum() if 'VL_LIBERADO' in df_status.columns else 0
                st.metric("💰 Total", formatar_moeda_br(valor_total_status))
            
            with st.expander("ℹ️ Informações detalhadas"):
                st.write(f"**Período:** {PERIODO_BASE}")
                st.write(f"**Prestadores únicos:** {formatar_inteiro_br(df_status['NM_PRESTADOR_EXEC'].nunique())}")
                st.write(f"**Procedimentos:** {formatar_inteiro_br(df_status['DS_PROCEDIMENTO'].nunique())}")
                
                if 'MUNICIPIO_PRESTADOR' in df_status.columns:
                    st.write(f"**Municípios:** {formatar_inteiro_br(df_status['MUNICIPIO_PRESTADOR'].nunique())}")
                
                # Distribuição
                local_count = len(df_status[df_status['TP_PRESTADOR_CLASSIFICADO'] == 'LOCAL'])
                intercambio_count = len(df_status[df_status['TP_PRESTADOR_CLASSIFICADO'] == 'INTERCÂMBIO'])
                total = len(df_status)
                if total > 0:
                    st.write(f"**Local:** {formatar_inteiro_br(local_count)} ({local_count/total*100:.1f}%)")
                    st.write(f"**Intercâmbio:** {formatar_inteiro_br(intercambio_count)} ({intercambio_count/total*100:.1f}%)")
        
        # Informações de sessão
        st.markdown("---")
        st.markdown("### 🔐 Sessão")
        
        if st.session_state.login_time:
            tempo_sessao = datetime.now() - st.session_state.login_time
            horas = int(tempo_sessao.total_seconds() // 3600)
            minutos = int((tempo_sessao.total_seconds() % 3600) // 60)
            
            st.info(f"""
            **Ativa:** {horas}h {minutos}min
            **Expira em:** {8 - horas}h {60 - minutos}min
            """)
        
        # Botão de logout
        if st.button("🚪 Sair do Sistema", use_container_width=True, type="secondary"):
            st.session_state.auth_sucesso = False
            st.session_state.login_time = None
            st.success("✅ Logout realizado com sucesso!")
            st.rerun()
        
        # Rodapé
        st.markdown("---")
        st.markdown(f"""
        <div style="color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; font-size: 11px; text-align: center;">
            <p><strong>Desenvolvido por:</strong><br>{DESENVOLVEDOR}</p>
            <p>📅 {PERIODO_BASE}</p>
            <p>🕐 {datetime.now().strftime('%H:%M')}</p>
        </div>
        """, unsafe_allow_html=True)
    
    # ============================================
    # CONTEÚDO PRINCIPAL
    # ============================================
    # Cabeçalho
    col_header1, col_header2, col_header3 = st.columns([4, 1, 1])
    
    with col_header1:
        st.title("🏥 Dashboard de Diárias Hospitalares")
        st.markdown(f"""
        <div class="text-secondary">
        Análise completa com classificação LOCAL/INTERCÂMBIO • {PERIODO_BASE}
        </div>
        """, unsafe_allow_html=True)
    
    with col_header2:
        st.markdown(f"""
        <div class="status-info">
            <div style="font-size: 12px;">
                🌐 <strong>Streamlit Cloud</strong><br>
                📅 {datetime.now().strftime('%d/%m/%Y')}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_header3:
        if st.button("🔄 Atualizar", use_container_width=True, help="Recarregar dados"):
            st.cache_data.clear()
            st.rerun()
    
    # Carregar dados
    with st.spinner("📊 Carregando dados..."):
        df = carregar_e_preparar_dados()
    
    if df is None or df.empty:
        st.error("❌ Não foi possível carregar os dados.")
        return
    
    # ============================================
    # FILTROS AVANÇADOS
    # ============================================
    st.markdown("## 🔧 Filtros Avançados")
    st.markdown(f"*Período base: {PERIODO_BASE}*")
    
    with st.container():
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            # Competência
            if 'MES_ANO_FORMATADO' in df.columns:
                competencias = ordenar_meses(df['MES_ANO_FORMATADO'].unique().tolist())
                competencia_selecionada = st.selectbox(
                    "📅 Competência",
                    ['TODAS AS COMPETÊNCIAS'] + competencias,
                    key="filtro_competencia"
                )
            else:
                competencia_selecionada = 'TODAS AS COMPETÊNCIAS'
        
        with col2:
            # Tipo (LOCAL/INTERCÂMBIO)
            tipo_selecionado = st.selectbox(
                "🏢 Tipo",
                ['TODOS', 'LOCAL', 'INTERCÂMBIO'],
                key="filtro_tipo"
            )
        
        with col3:
            # Município
            if 'MUNICIPIO_PRESTADOR' in df.columns:
                municipios = ['TODOS OS MUNICÍPIOS'] + sorted(df['MUNICIPIO_PRESTADOR'].unique().tolist())
                municipio_selecionado = st.selectbox(
                    "📍 Município",
                    municipios,
                    key="filtro_municipio"
                )
            else:
                municipio_selecionado = 'TODOS OS MUNICÍPIOS'
        
        with col4:
            # Prestador
            df_filtro_prestador = df.copy()
            
            if competencia_selecionada != 'TODAS AS COMPETÊNCIAS':
                df_filtro_prestador = df_filtro_prestador[df_filtro_prestador['MES_ANO_FORMATADO'] == competencia_selecionada]
            
            if tipo_selecionado != 'TODOS':
                df_filtro_prestador = df_filtro_prestador[df_filtro_prestador['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
            
            if municipio_selecionado != 'TODOS OS MUNICÍPIOS':
                df_filtro_prestador = df_filtro_prestador[df_filtro_prestador['MUNICIPIO_PRESTADOR'] == municipio_selecionado]
            
            prestadores = ['TODOS OS PRESTADORES'] + sorted(df_filtro_prestador['NM_PRESTADOR_EXEC'].unique().tolist())[:50]
            prestador_selecionado = st.selectbox(
                "👨‍⚕️ Prestador",
                prestadores,
                key="filtro_prestador"
            )
    
    # Segunda linha de filtros
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        # Procedimento
        df_filtro_procedimento = df_filtro_prestador.copy()
        
        if prestador_selecionado != 'TODOS OS PRESTADORES':
            df_filtro_procedimento = df_filtro_procedimento[df_filtro_procedimento['NM_PRESTADOR_EXEC'] == prestador_selecionado]
        
        procedimentos = ['TODOS OS PROCEDIMENTOS'] + sorted(df_filtro_procedimento['DS_PROCEDIMENTO'].unique().tolist())
        procedimento_selecionado = st.selectbox(
            "🩺 Procedimento",
            procedimentos,
            key="filtro_procedimento"
        )
    
    with col6:
        # Valor mínimo
        if 'VL_LIBERADO' in df.columns:
            valor_min = st.number_input(
                "💰 Valor Mínimo (R$)",
                min_value=float(df['VL_LIBERADO'].min()),
                max_value=float(df['VL_LIBERADO'].max()),
                value=float(df['VL_LIBERADO'].min()),
                step=10.0,
                key="filtro_valor_min"
            )
    
    with col7:
        # Valor máximo
        if 'VL_LIBERADO' in df.columns:
            valor_max = st.number_input(
                "💰 Valor Máximo (R$)",
                min_value=float(df['VL_LIBERADO'].min()),
                max_value=float(df['VL_LIBERADO'].max()),
                value=float(df['VL_LIBERADO'].max()),
                step=10.0,
                key="filtro_valor_max"
            )
    
    with col8:
        # Quantidade máxima
        if 'QT_ITEM' in df.columns:
            qt_max = st.number_input(
                "📊 Máx. Diárias",
                min_value=int(df['QT_ITEM'].min()),
                max_value=int(df['QT_ITEM'].max()),
                value=int(df['QT_ITEM'].max()),
                step=1,
                key="filtro_quantidade"
            )
    
    # Botões de controle - CORREÇÃO AQUI
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
    filtros_ativos = []
    
    # Aplicar filtros sequencialmente
    if competencia_selecionada != 'TODAS AS COMPETÊNCIAS' and 'MES_ANO_FORMATADO' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['MES_ANO_FORMATADO'] == competencia_selecionada]
        filtros_ativos.append(f"Competência: {competencia_selecionada}")
    
    if tipo_selecionado != 'TODOS' and 'TP_PRESTADOR_CLASSIFICADO' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == tipo_selecionado]
        filtros_ativos.append(f"Tipo: {tipo_selecionado}")
    
    if municipio_selecionado != 'TODOS OS MUNICÍPIOS' and 'MUNICIPIO_PRESTADOR' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['MUNICIPIO_PRESTADOR'] == municipio_selecionado]
        filtros_ativos.append(f"Município: {municipio_selecionado}")
    
    if prestador_selecionado != 'TODOS OS PRESTADORES':
        df_filtrado = df_filtrado[df_filtrado['NM_PRESTADOR_EXEC'] == prestador_selecionado]
        filtros_ativos.append(f"Prestador: {prestador_selecionado[:20]}{'...' if len(prestador_selecionado) > 20 else ''}")
    
    if procedimento_selecionado != 'TODOS OS PROCEDIMENTOS':
        df_filtrado = df_filtrado[df_filtrado['DS_PROCEDIMENTO'] == procedimento_selecionado]
        filtros_ativos.append(f"Procedimento: {procedimento_selecionado[:20]}{'...' if len(procedimento_selecionado) > 20 else ''}")
    
    # Filtros numéricos
    if 'VL_LIBERADO' in df_filtrado.columns:
        df_filtrado = df_filtrado[
            (df_filtrado['VL_LIBERADO'] >= valor_min) & 
            (df_filtrado['VL_LIBERADO'] <= valor_max)
        ]
    
    if 'QT_ITEM' in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado['QT_ITEM'] <= qt_max]
    
    # Mostrar filtros ativos
    if filtros_ativos:
        st.markdown(f"""
        <div class="info-card card-info">
            <div style="display: flex; align-items: center; justify-content: space-between;">
                <div>
                    <strong>🔧 Filtros Ativos:</strong> {' • '.join(filtros_ativos)}
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
    pacientes_unicos = df_filtrado['CD_BENEFICIARIO'].nunique() if 'CD_BENEFICIARIO' in df_filtrado.columns else 0
    total_diarias = df_filtrado['QT_ITEM'].sum() if 'QT_ITEM' in df_filtrado.columns else 0
    valor_total = df_filtrado['VL_LIBERADO'].sum() if 'VL_LIBERADO' in df_filtrado.columns else 0
    valor_medio = valor_total / len(df_filtrado) if len(df_filtrado) > 0 else 0
    
    # Local/Intercâmbio específico
    if 'TP_PRESTADOR_CLASSIFICADO' in df_filtrado.columns:
        local_df = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == 'LOCAL']
        intercambio_df = df_filtrado[df_filtrado['TP_PRESTADOR_CLASSIFICADO'] == 'INTERCÂMBIO']
        
        valor_local = local_df['VL_LIBERADO'].sum() if len(local_df) > 0 else 0
        valor_intercambio = intercambio_df['VL_LIBERADO'].sum() if len(intercambio_df) > 0 else 0
        perc_local = (valor_local / valor_total * 100) if valor_total > 0 else 0
        perc_intercambio = (valor_intercambio / valor_total * 100) if valor_total > 0 else 0
    else:
        valor_local = 0
        valor_intercambio = 0
        perc_local = 0
        perc_intercambio = 0
    
    # Layout de métricas
    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    
    with col_k1:
        st.markdown(f"""
        <div class="metric-card metric-pacientes">
            <div style="font-size: 14px; color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; margin-bottom: 10px;">
                👥 Pacientes Únicos
            </div>
            <div style="font-size: 24px; font-weight: 700; color: {'#ffffff' if st.session_state.tema == 'dark' else '#1e293b'};">
                {formatar_inteiro_br(pacientes_unicos)}
            </div>
            <div style="font-size: 12px; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; margin-top: 5px;">
                {formatar_inteiro_br(total_diarias)} diárias
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k2:
        st.markdown(f"""
        <div class="metric-card metric-local">
            <div style="font-size: 14px; color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; margin-bottom: 10px;">
                🏥 Local
            </div>
            <div style="font-size: 24px; font-weight: 700; color: {'#ffffff' if st.session_state.tema == 'dark' else '#1e293b'};">
                {formatar_moeda_br(valor_local)}
            </div>
            <div style="font-size: 12px; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; margin-top: 5px;">
                {formatar_inteiro_br(len(local_df) if 'local_df' in locals() else 0)} reg • {perc_local:.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k3:
        st.markdown(f"""
        <div class="metric-card metric-intercambio">
            <div style="font-size: 14px; color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; margin-bottom: 10px;">
                🌐 Intercâmbio
            </div>
            <div style="font-size: 24px; font-weight: 700; color: {'#ffffff' if st.session_state.tema == 'dark' else '#1e293b'};">
                {formatar_moeda_br(valor_intercambio)}
            </div>
            <div style="font-size: 12px; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; margin-top: 5px;">
                {formatar_inteiro_br(len(intercambio_df) if 'intercambio_df' in locals() else 0)} reg • {perc_intercambio:.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k4:
        st.markdown(f"""
        <div class="metric-card metric-total">
            <div style="font-size: 14px; color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; margin-bottom: 10px;">
                💰 Valor Total
            </div>
            <div style="font-size: 24px; font-weight: 700; color: {'#ffffff' if st.session_state.tema == 'dark' else '#1e293b'};">
                {formatar_moeda_br(valor_total)}
            </div>
            <div style="font-size: 12px; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; margin-top: 5px;">
                Média: {formatar_moeda_br(valor_medio)}/diária
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_k5:
        media_por_paciente = valor_total / pacientes_unicos if pacientes_unicos > 0 else 0
        diarias_por_paciente = total_diarias / pacientes_unicos if pacientes_unicos > 0 else 0

        st.markdown(
            f"""
            <div class="metric-card metric-media">
                <div style="font-size: 14px; color: {'#94a3b8' if st.session_state.tema == 'dark' else '#64748b'}; margin-bottom: 10px;">
                    📊 Média/Paciente
                </div>
                <div style="font-size: 24px; font-weight: 700; color: {'#ffffff' if st.session_state.tema == 'dark' else '#1e293b'};">
                    {formatar_moeda_br(media_por_paciente)}
                </div>
                <div style="font-size: 12px; color: {'#64748b' if st.session_state.tema == 'dark' else '#94a3b8'}; margin-top: 5px;">
                    Diárias/paciente: {diarias_por_paciente:.1f}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    # ============================================
    # GRÁFICOS COMPLETOS (ESTILO CÓDIGO 2)
    # ============================================
    st.markdown("---")
    st.markdown("## 📊 Visualizações")
    
    # Configurar template Plotly baseado no tema
    plotly_template = "plotly_dark" if st.session_state.tema == "dark" else "plotly_white"
    font_color = "white" if st.session_state.tema == "dark" else "#1e293b"
    bg_color = "rgba(0,0,0,0)" if st.session_state.tema == "dark" else "rgba(255,255,255,0)"
    paper_bg_color = "rgba(0,0,0,0)" if st.session_state.tema == "dark" else "rgba(255,255,255,0)"
    
    # GRÁFICO 1: Evolução Temporal (se houver dados temporais)
    if 'MES_ANO_FORMATADO' in df_filtrado.columns and df_filtrado['MES_ANO_FORMATADO'].nunique() > 1:
        st.markdown("### 📈 Evolução Temporal")
        
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
                height=400,
                template=plotly_template
            )
            
            fig_evolucao.update_traces(
                line=dict(color='#3b82f6', width=3),
                marker=dict(size=8, color='#22c55e')
            )
            
            fig_evolucao.update_layout(
                plot_bgcolor=bg_color,
                paper_bgcolor=paper_bg_color,
                font_color=font_color,
                xaxis_tickangle=-45
            )
            
            fig_evolucao.update_yaxes(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
            
            st.plotly_chart(fig_evolucao, use_container_width=True)
        
        with col_t2:
            # Gráfico de distribuição por tipo
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
                height=400,
                template=plotly_template
            )
            
            fig_tipo_temporal.update_layout(
                plot_bgcolor=bg_color,
                paper_bgcolor=paper_bg_color,
                font_color=font_color,
                xaxis_tickangle=-45,
                legend_title_text='Tipo de Prestador'
            )
            
            fig_tipo_temporal.update_yaxes(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
            
            st.plotly_chart(fig_tipo_temporal, use_container_width=True)
    
    # GRÁFICO 2: Análise por Procedimento
    st.markdown("### 🩺 Análise por Tipo de Diária")
    
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
            height=500,
            template=plotly_template
        )
        
        fig_treemap.update_layout(
            plot_bgcolor=bg_color,
            paper_bgcolor=paper_bg_color,
            font_color=font_color
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
            hover_data=['CD_BENEFICIARIO', 'QT_ITEM'],
            template=plotly_template
        )
        
        fig_barras_proc.update_traces(
            texttemplate='R$ %{text:,.2f}',
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Valor: R$ %{x:,.2f}<br>Pacientes: %{customdata[0]}<br>Diárias: %{customdata[1]}'
        )
        
        fig_barras_proc.update_layout(
            plot_bgcolor=bg_color,
            paper_bgcolor=paper_bg_color,
            font_color=font_color,
            yaxis={'categoryorder': 'total ascending'},
            coloraxis_showscale=False,
            xaxis=dict(
                tickprefix="R$ ",
                tickformat=",.2f"
            )
        )
        
        st.plotly_chart(fig_barras_proc, use_container_width=True)
    
    # GRÁFICO 3: Ranking de Prestadores
    st.markdown("### 🏆 Ranking de Prestadores")
    
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
            hover_data=['CD_BENEFICIARIO'],
            template=plotly_template
        )
        
        fig_barras.update_traces(
            texttemplate='R$ %{text:,.2f}',
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Valor: R$ %{x:,.2f}<br>Pacientes: %{customdata[0]}<br>Tipo: %{marker.color}'
        )
        
        fig_barras.update_layout(
            plot_bgcolor=bg_color,
            paper_bgcolor=paper_bg_color,
            font_color=font_color,
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
            color_discrete_map={'LOCAL': '#22c55e', 'INTERCÂMBIO': '#3b82f6'},
            template=plotly_template
        )
        
        fig_scatter.update_traces(
            marker=dict(line=dict(width=1, color='white' if st.session_state.tema == "dark" else "#1e293b")),
            hovertemplate='<b>%{hovertext}</b><br>Pacientes: %{x}<br>Valor: R$ %{y:,.2f}<br>Diárias: %{marker.size}<br>Tipo: %{marker.color}'
        )
        
        fig_scatter.update_layout(
            plot_bgcolor=bg_color,
            paper_bgcolor=paper_bg_color,
            font_color=font_color,
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
    st.markdown("## 📋 Dados Detalhados")
    
    tab1, tab2, tab3 = st.tabs(["🏥 Ranking Completo", "🩺 Detalhes por Procedimento", "📊 Resumo por Competência"])
    
    with tab1:
        # Tabela ranking prestadores formatada
        ranking_formatado = ranking.copy()
        
        # Aplicar formatação brasileira
        ranking_formatado['VL_LIBERADO'] = ranking_formatado['VL_LIBERADO'].apply(formatar_moeda_br)
        ranking_formatado['CD_BENEFICIARIO'] = ranking_formatado['CD_BENEFICIARio'].apply(formatar_inteiro_br)
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
    # EXPORTAÇÃO COMPLETA
    # ============================================
    if exportar_clicked:
        st.markdown("---")
        st.markdown("## 📤 Exportar Dados")
        
        col_export1, col_export2 = st.columns(2)
        
        with col_export1:
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
            
            output.seek(0)
            
            st.download_button(
                label="📊 Baixar Excel Completo",
                data=output,
                file_name=f"dashboard_unimed_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_export2:
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
            {', '.join(filtros_ativos) if filtros_ativos else 'Nenhum filtro aplicado'}
            
            📈 TOP 5 PRESTADORES:
            """
            
            for i, row in ranking.head().iterrows():
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
    # RODAPÉ COMPLETO
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
# 9. EXECUÇÃO PRINCIPAL
# ============================================
def main():
    """Função principal do aplicativo"""
    # Verificar sessão
    if st.session_state.login_time:
        tempo_sessao = datetime.now() - st.session_state.login_time
        if tempo_sessao.total_seconds() > 28800:  # 8 horas
            st.session_state.auth_sucesso = False
            st.session_state.login_time = None
    
    # Exibir tela apropriada
    if not st.session_state.auth_sucesso:
        tela_login()
    else:
        mostrar_dashboard()

# ============================================
# 10. INICIALIZAÇÃO
# ============================================
if __name__ == "__main__":
    main()