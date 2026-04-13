import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import hashlib
import os

# ========================================================================
# CONFIGURACAO
# ========================================================================
st.set_page_config(
    page_title="Dashboard de Pedidos - Autron",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "dados"))

# ========================================================================
# AUTENTICACAO
# ========================================================================
# Usuarios configurados via variavel de ambiente AUTH_USERS
# Formato: "email1:senha1,email2:senha2"
# Exemplo: "dani@autron.com:Autron2026,admin@autron.com:Admin123"
# Se AUTH_USERS nao estiver definida, login fica desabilitado (acesso livre)

AUTH_USERS_RAW = os.environ.get("AUTH_USERS", "")


def get_usuarios():
    """Carrega usuarios da variavel de ambiente."""
    if not AUTH_USERS_RAW.strip():
        return {}
    usuarios = {}
    for par in AUTH_USERS_RAW.split(","):
        par = par.strip()
        if ":" in par:
            email, senha = par.split(":", 1)
            usuarios[email.strip().lower()] = senha.strip()
    return usuarios


def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()


def verificar_login(email, senha):
    usuarios = get_usuarios()
    email = email.strip().lower()
    if email in usuarios:
        return usuarios[email] == senha
    return False


def tela_login():
    """Exibe tela de login estilizada."""
    st.markdown("""
    <style>
        .login-container {
            max-width: 420px;
            margin: 80px auto;
            padding: 40px;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            border: 1px solid #2F5496;
            border-radius: 16px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        }
        .login-title {
            text-align: center;
            color: #4FC3F7;
            font-size: 1.8rem;
            font-weight: 700;
            margin-bottom: 5px;
        }
        .login-subtitle {
            text-align: center;
            color: #B0BEC5;
            font-size: 0.95rem;
            margin-bottom: 30px;
        }
        .login-logo {
            text-align: center;
            font-size: 3rem;
            margin-bottom: 10px;
        }
        [data-testid="stForm"] {
            border: none !important;
        }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        logo_login = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
        if os.path.exists(logo_login):
            import base64
            with open(logo_login, "rb") as f:
                logo_b64 = base64.b64encode(f.read()).decode()
            st.markdown(f'<div style="text-align:center;"><img src="data:image/png;base64,{logo_b64}" width="220"></div>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle">Dashboard de Pedidos</div>', unsafe_allow_html=True)

        with st.form("login_form"):
            email = st.text_input("📧 Email", placeholder="seu@email.com")
            senha = st.text_input("🔒 Senha", type="password", placeholder="Digite sua senha")
            submit = st.form_submit_button("Entrar", use_container_width=True, type="primary")

            if submit:
                if email and senha:
                    if verificar_login(email, senha):
                        st.session_state["autenticado"] = True
                        st.session_state["usuario"] = email.strip().lower()
                        st.rerun()
                    else:
                        st.error("Email ou senha incorretos.")
                else:
                    st.warning("Preencha email e senha.")

    st.stop()


# Verificar se login esta habilitado
LOGIN_HABILITADO = bool(AUTH_USERS_RAW.strip())

if LOGIN_HABILITADO:
    if not st.session_state.get("autenticado", False):
        tela_login()

CORES = {
    "azul": "#2F5496",
    "azul_claro": "#4472C4",
    "verde": "#2ECC71",
    "amarelo": "#F39C12",
    "vermelho": "#E74C3C",
    "vermelho_escuro": "#C0392B",
    "cinza": "#95A5A6",
    "branco": "#FFFFFF",
    "fundo": "#0E1117",
}


# ========================================================================
# CSS CUSTOMIZADO
# ========================================================================
st.markdown("""
<style>
    .main .block-container { padding-top: 1rem; max-width: 100%; }

    .kpi-card {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 1px solid #2F5496;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin: 5px 0;
    }
    .kpi-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #4FC3F7;
        margin: 0;
    }
    .kpi-label {
        font-size: 0.85rem;
        color: #B0BEC5;
        margin-top: 5px;
    }
    .kpi-card-alert {
        background: linear-gradient(135deg, #4a1a1a 0%, #3e1616 100%);
        border: 1px solid #E74C3C;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin: 5px 0;
    }
    .kpi-value-alert {
        font-size: 2.2rem;
        font-weight: 700;
        color: #E74C3C;
        margin: 0;
    }
    .kpi-card-ok {
        background: linear-gradient(135deg, #1a3a1a 0%, #163e16 100%);
        border: 1px solid #2ECC71;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin: 5px 0;
    }
    .kpi-value-ok {
        font-size: 2.2rem;
        font-weight: 700;
        color: #2ECC71;
        margin: 0;
    }
    .kpi-card-warn {
        background: linear-gradient(135deg, #3a3a1a 0%, #3e3616 100%);
        border: 1px solid #F39C12;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin: 5px 0;
    }
    .kpi-value-warn {
        font-size: 2.2rem;
        font-weight: 700;
        color: #F39C12;
        margin: 0;
    }

    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 1px solid #2F5496;
        border-radius: 10px;
        padding: 15px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #1a1a2e;
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        border: 1px solid #2F5496;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2F5496;
    }

    h1, h2, h3 { color: #E0E0E0; }

    .status-finalizado { color: #2ECC71; font-weight: bold; }
    .status-aberto { color: #E74C3C; font-weight: bold; }

    .upload-section {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 2px dashed #2F5496;
        border-radius: 12px;
        padding: 20px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)


def kpi_card(valor, label, tipo="normal"):
    css_class = {
        "normal": ("kpi-card", "kpi-value"),
        "alert": ("kpi-card-alert", "kpi-value-alert"),
        "ok": ("kpi-card-ok", "kpi-value-ok"),
        "warn": ("kpi-card-warn", "kpi-value-warn"),
    }.get(tipo, ("kpi-card", "kpi-value"))

    return f"""
    <div class="{css_class[0]}">
        <p class="{css_class[1]}">{valor}</p>
        <p class="kpi-label">{label}</p>
    </div>
    """


def filtro_data(data_df, col_data, label, tab_key):
    """Exibe filtro de data (inicio/fim) e retorna o dataframe filtrado.
    Filtra graficos e tabelas da aba correspondente."""
    datas_validas = data_df[col_data].dropna()
    if len(datas_validas) == 0:
        return data_df

    data_min = datas_validas.min().date()
    data_max = datas_validas.max().date()

    col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
    with col_f1:
        dt_ini = st.date_input(
            f"De ({label})", value=data_min, min_value=data_min,
            max_value=data_max, key=f"dt_ini_{tab_key}",
            format="DD/MM/YYYY"
        )
    with col_f2:
        dt_fim = st.date_input(
            f"Ate ({label})", value=data_max, min_value=data_min,
            max_value=data_max, key=f"dt_fim_{tab_key}",
            format="DD/MM/YYYY"
        )
    with col_f3:
        st.markdown("")

    dt_ini = pd.Timestamp(dt_ini)
    dt_fim = pd.Timestamp(dt_fim) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    mask = (data_df[col_data] >= dt_ini) & (data_df[col_data] <= dt_fim)
    # Manter linhas sem data (nao excluir por falta de informacao)
    mask = mask | data_df[col_data].isna()
    return data_df[mask]


# ========================================================================
# CARGA E PROCESSAMENTO DE DADOS
# ========================================================================
@st.cache_data(ttl=300)
def carregar_e_processar():
    """Carrega os 5 arquivos e processa toda a logica de negocio."""

    REPO_DIR = os.path.dirname(os.path.abspath(__file__))

    def resolver_arquivo(nome):
        """Busca arquivo em DATA_DIR primeiro, depois na raiz do repo.
        Se existir nos dois, usa o mais recente."""
        arq_dados = os.path.join(DATA_DIR, nome)
        arq_repo = os.path.join(REPO_DIR, nome)
        existe_dados = os.path.exists(arq_dados)
        existe_repo = os.path.exists(arq_repo)
        if existe_dados and existe_repo:
            # Usar o mais recente
            if os.path.getmtime(arq_repo) > os.path.getmtime(arq_dados):
                return arq_repo
            return arq_dados
        if existe_dados:
            return arq_dados
        if existe_repo:
            return arq_repo
        return arq_dados  # retorna path padrao (vai falhar no check)

    arq_entrada = resolver_arquivo("entrada_pedido.xlsx")
    arq_followup = resolver_arquivo("followup.xlsx")
    arq_estoque = resolver_arquivo("mata010.xlsx")
    arq_faturamento = resolver_arquivo("faturamento.xlsx")
    arq_budget = resolver_arquivo("Metas_Budget_Grupo_Autron_2026.xlsx")
    # Aceitar ambos os nomes do CSV
    arq_sciozmq = resolver_arquivo("sciozvs0.csv")
    if not os.path.exists(arq_sciozmq):
        arq_sciozmq = resolver_arquivo("sciozmq0.csv")

    for arq in [arq_entrada, arq_followup, arq_estoque, arq_sciozmq]:
        if not os.path.exists(arq):
            return None, None, None, f"Arquivo nao encontrado: {os.path.basename(arq)}"

    # Carregar
    ep = pd.read_excel(arq_entrada, header=1)
    fu = pd.read_excel(arq_followup, header=1)
    # Faturamento (opcional - pode nao existir)
    fat = None
    if os.path.exists(arq_faturamento):
        fat = pd.read_excel(arq_faturamento, header=1)
        fat['Emissao'] = pd.to_datetime(fat['Emissao'], errors='coerce')
        col_vlr = [c for c in fat.columns if 'vlr' in c.lower() and 'total' in c.lower()]
        if col_vlr:
            fat.rename(columns={col_vlr[0]: 'Vlr.Total.Fat'}, inplace=True)
        else:
            fat['Vlr.Total.Fat'] = 0
        fat['Vlr.Total.Fat'] = pd.to_numeric(fat['Vlr.Total.Fat'], errors='coerce')
        fat['Faturamento Bruto'] = pd.to_numeric(fat.get('Faturamento Bruto', 0), errors='coerce')
        fat['Faturamento Liquido'] = pd.to_numeric(fat.get('Faturamento Liquido', 0), errors='coerce')
        fat['No do Pedido'] = fat['No do Pedido'].astype(str).str.replace('.', '', regex=False).str.strip()
        fat['Mes_Faturamento'] = fat['Emissao'].dt.to_period('M').astype(str)
    # Tentar diferentes linhas de header para mata010
    mt = pd.read_excel(arq_estoque)
    mt.columns = mt.columns.astype(str).str.strip()
    # Se nao encontrar 'Saldo Atual' ou similar, tentar header=1, header=2
    _has_saldo = any('saldo' in c.lower() for c in mt.columns)
    _has_codigo = any(c.lower() in ('codigo', 'código') or 'cod' in c.lower() for c in mt.columns)
    if not _has_saldo and not _has_codigo:
        mt = pd.read_excel(arq_estoque, header=1)
        mt.columns = mt.columns.astype(str).str.strip()
        _has_saldo = any('saldo' in c.lower() for c in mt.columns)
        if not _has_saldo:
            mt = pd.read_excel(arq_estoque, header=2)
            mt.columns = mt.columns.astype(str).str.strip()
    sc_csv = pd.read_csv(arq_sciozmq, encoding='latin-1', sep=';', header=2, low_memory=False)

    # Limpar entrada_pedido
    ep['Num. Pedido'] = ep['Num. Pedido'].astype(str).str.replace('.', '', regex=False).str.strip()
    ep['Numero SC'] = pd.to_numeric(ep['Numero SC'], errors='coerce')
    ep['Quantidade'] = pd.to_numeric(ep['Quantidade'], errors='coerce')
    ep['DT Emissao'] = pd.to_datetime(ep['DT Emissao'], errors='coerce')
    ep['DT. Ofertada'] = pd.to_datetime(ep['DT. Ofertada'], errors='coerce')
    ep['DT. Fat. Cli'] = pd.to_datetime(ep['DT. Fat. Cli'], errors='coerce')
    ep['Entrega'] = pd.to_datetime(ep['Entrega'], errors='coerce')
    ep['Nota Fiscal'] = pd.to_numeric(ep['Nota Fiscal'], errors='coerce')
    ep['Item'] = pd.to_numeric(ep['Item'], errors='coerce')
    ep['Vlr.Total'] = pd.to_numeric(ep['Vlr.Total'], errors='coerce')
    ep['Margem'] = pd.to_numeric(ep['Margem'], errors='coerce')
    if 'Numero OP' not in ep.columns:
        ep['Numero OP'] = np.nan

    # Limpar follow-up
    fu['No. da S.C.'] = pd.to_numeric(fu['No. da S.C.'], errors='coerce')
    fu['Dt. Confirma'] = pd.to_datetime(fu['Dt. Confirma'], errors='coerce')
    fu['Dt. Pre entr'] = pd.to_datetime(fu['Dt. Pre entr'], errors='coerce')
    fu['Dt Chegada Autron'] = pd.to_datetime(fu['Dt Chegada Autron'], errors='coerce')
    fu['PV informado na SC'] = pd.to_numeric(fu['PV informado na SC'], errors='coerce')
    fu['Numero do PV'] = fu['Numero do PV'].astype(str).str.replace('.', '', regex=False).str.strip()
    if 'OP na SC' not in fu.columns:
        fu['OP na SC'] = np.nan

    # Limpar estoque - normalizar nomes de colunas
    mt.columns = mt.columns.astype(str).str.strip()
    col_saldo = [c for c in mt.columns if 'saldo' in c.lower() and 'atual' in c.lower()]
    if col_saldo:
        mt.rename(columns={col_saldo[0]: 'Saldo Atual'}, inplace=True)
    col_codigo = [c for c in mt.columns if c.lower().strip() in ('codigo', 'código', 'cod', 'cod.')]
    if col_codigo:
        mt.rename(columns={col_codigo[0]: 'Codigo'}, inplace=True)
    if 'Codigo' not in mt.columns:
        col_codigo_alt = [c for c in mt.columns if 'codig' in c.lower() or 'cod' in c.lower() or 'produto' in c.lower() or 'item' in c.lower() or 'b1_cod' in c.lower()]
        if col_codigo_alt:
            mt.rename(columns={col_codigo_alt[0]: 'Codigo'}, inplace=True)
        else:
            # Usar primeira coluna como codigo (fallback)
            mt.rename(columns={mt.columns[0]: 'Codigo'}, inplace=True)
    if 'Saldo Atual' not in mt.columns:
        # Tentar coluna que contenha 'saldo'
        col_saldo_alt = [c for c in mt.columns if 'saldo' in c.lower()]
        if col_saldo_alt:
            mt.rename(columns={col_saldo_alt[0]: 'Saldo Atual'}, inplace=True)
        else:
            mt['Saldo Atual'] = 0
    mt['Saldo Atual'] = pd.to_numeric(mt['Saldo Atual'], errors='coerce').fillna(0)

    # Classificacao Comprando/Produzindo
    sc_csv['Produto'] = sc_csv['Produto'].astype(str).str.strip()
    sc_csv['Prod.Solict'] = sc_csv['Prod.Solict'].astype(str).str.strip()
    tipo_prod = sc_csv.groupby('Produto')['Prod.Solict'].agg(
        lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'Indefinido'
    ).reset_index()
    tipo_prod.rename(columns={'Prod.Solict': 'Tipo_Produto'}, inplace=True)

    # --- Follow-up por SC ---
    fu_por_sc = fu.groupby('No. da S.C.').agg({
        'Dt. Confirma': 'first', 'Dt. Pre entr': 'first', 'Pasta': 'first',
        'No. do P.O.': 'first', 'Dt Chegada Autron': 'first', 'OP na SC': 'first'
    }).reset_index()
    fu_por_sc.rename(columns={'No. da S.C.': 'Numero SC'}, inplace=True)

    # --- Follow-up por PV+Produto ---
    fu_pv = fu.copy()
    fu_pv['Numero do PV'] = fu_pv['Numero do PV'].str.strip()
    fu_por_pv = fu_pv.groupby(['Numero do PV', 'Codigo Item']).agg({
        'Dt. Confirma': 'first', 'Dt. Pre entr': 'first', 'Pasta': 'first',
        'No. do P.O.': 'first', 'Dt Chegada Autron': 'first', 'OP na SC': 'first'
    }).reset_index()
    fu_por_pv.rename(columns={
        'Numero do PV': 'Num. Pedido', 'Codigo Item': 'Produto',
        'Dt. Confirma': 'FU_Dt_Confirma_PV', 'Dt. Pre entr': 'FU_Dt_Pre_entr_PV',
        'Pasta': 'FU_Pasta_PV', 'No. do P.O.': 'FU_PO_PV',
        'Dt Chegada Autron': 'FU_Dt_Chegada_PV', 'OP na SC': 'FU_OP_na_SC_PV'
    }, inplace=True)

    # Merges
    merged = ep.merge(fu_por_sc, on='Numero SC', how='left', suffixes=('', '_fu'))
    merged = merged.merge(fu_por_pv, on=['Num. Pedido', 'Produto'], how='left')

    # Consolidar
    merged['FU_Dt_Confirma'] = merged['Dt. Confirma'].combine_first(merged['FU_Dt_Confirma_PV'])
    merged['FU_Dt_Pre_Entr'] = merged['Dt. Pre entr'].combine_first(merged['FU_Dt_Pre_entr_PV'])
    merged['FU_Pasta'] = merged['Pasta'].combine_first(merged['FU_Pasta_PV'])
    merged['FU_PO'] = merged['No. do P.O.'].combine_first(merged['FU_PO_PV'])
    merged['FU_Dt_Chegada_Autron'] = merged['Dt Chegada Autron'].combine_first(merged['FU_Dt_Chegada_PV'])
    merged['FU_OP_na_SC'] = merged['OP na SC'].combine_first(merged['FU_OP_na_SC_PV'])

    merged['Prazo_Real_Entrega'] = merged['FU_Dt_Confirma'].combine_first(merged['FU_Dt_Pre_Entr'])
    merged['Semana_Entrega'] = np.where(merged['FU_Dt_Confirma'].notna(), merged['FU_Pasta'], None)

    # Tipo produto
    merged = merged.merge(tipo_prod, on='Produto', how='left')
    merged['Tipo_Produto'] = merged['Tipo_Produto'].fillna('Nao classificado')

    # Status
    merged['Status_Pedido'] = np.where(merged['Nota Fiscal'].notna(), 'FINALIZADO', 'EM ABERTO')

    # Estoque
    stock = mt.groupby('Codigo')['Saldo Atual'].sum().reset_index()
    stock.rename(columns={'Codigo': 'Produto', 'Saldo Atual': 'Estoque_Disponivel'}, inplace=True)
    merged = merged.merge(stock, on='Produto', how='left')
    merged['Estoque_Disponivel'] = merged['Estoque_Disponivel'].fillna(0)

    # Alocacao
    open_orders = merged[merged['Status_Pedido'] == 'EM ABERTO'].copy()
    open_orders = open_orders.sort_values(['Produto', 'DT Emissao', 'Num. Pedido', 'Item'])

    allocation = {}
    for produto in open_orders['Produto'].unique():
        prod_orders = open_orders[open_orders['Produto'] == produto]
        available = prod_orders['Estoque_Disponivel'].iloc[0] if len(prod_orders) > 0 else 0
        for idx, row in prod_orders.iterrows():
            qty = row['Quantidade'] if pd.notna(row['Quantidade']) else 0
            alloc = min(available, qty)
            allocation[idx] = {
                'Qtd_Alocada': alloc,
                'Disponivel_Estoque': 'SIM' if alloc >= qty else ('PARCIAL' if alloc > 0 else 'NAO')
            }
            available = max(0, available - qty)

    alloc_df = pd.DataFrame.from_dict(allocation, orient='index')
    merged = merged.join(alloc_df, how='left')
    merged.loc[merged['Status_Pedido'] == 'FINALIZADO', 'Disponivel_Estoque'] = 'N/A'
    merged.loc[merged['Status_Pedido'] == 'FINALIZADO', 'Qtd_Alocada'] = None

    # Regras SC/OP
    merged['Tem_SC'] = merged['Numero SC'].notna()
    merged['Tem_OP'] = (
        (merged['Numero OP'].notna() & merged['Numero OP'].astype(str).str.strip().ne('') & merged['Numero OP'].astype(str).str.strip().ne('nan')) |
        (merged['FU_OP_na_SC'].notna() & merged['FU_OP_na_SC'].astype(str).str.strip().ne('') & merged['FU_OP_na_SC'].astype(str).str.strip().ne('nan'))
    )

    def gerar_acao(row):
        if row['Status_Pedido'] == 'FINALIZADO':
            return 'Finalizado'
        if row['Disponivel_Estoque'] == 'SIM':
            return 'Estoque OK'
        tipo = str(row.get('Tipo_Produto', '')).strip()
        if tipo == 'Comprando':
            if row['Tem_OP']: return 'ERRO no CADASTRO'
            elif row['Tem_SC']: return 'SC gerada - Aguardando'
            else: return 'Necessario gerar SC'
        elif tipo == 'Produzindo':
            if row['Tem_OP']: return 'OP gerada - Aguardando'
            else: return 'Necessario gerar OP'
        else:
            return 'Verificar classificacao'

    merged['Acao_Necessaria'] = merged.apply(gerar_acao, axis=1)

    # Indicadores
    today = pd.Timestamp.now().normalize()
    merged['Dias_Atraso_Cliente'] = np.where(
        (merged['Status_Pedido'] == 'EM ABERTO') & (merged['DT. Fat. Cli'].notna()),
        (today - merged['DT. Fat. Cli']).dt.days, None
    )
    merged['Dias_Atraso_Cliente'] = pd.to_numeric(merged['Dias_Atraso_Cliente'], errors='coerce')

    merged['Dias_Atraso_Ofertada'] = np.where(
        (merged['Status_Pedido'] == 'EM ABERTO') & (merged['DT. Ofertada'].notna()),
        (today - merged['DT. Ofertada']).dt.days, None
    )
    merged['Dias_Atraso_Ofertada'] = pd.to_numeric(merged['Dias_Atraso_Ofertada'], errors='coerce')

    merged['Mes_Emissao'] = merged['DT Emissao'].dt.to_period('M').astype(str)
    merged['Ano_Emissao'] = merged['DT Emissao'].dt.year

    merged['Pronto_para_Fazer'] = np.where(
        merged['Status_Pedido'] == 'FINALIZADO', 'FINALIZADO',
        np.where(
            (merged['Disponivel_Estoque'] == 'SIM') & (merged['FU_Dt_Confirma'].notna()), 'SIM',
            np.where(
                merged['Disponivel_Estoque'] == 'SIM', 'PARCIAL - Sem Follow-up',
                np.where(
                    merged['FU_Dt_Confirma'].notna(), 'PARCIAL - Sem Estoque', 'NAO'
                )
            )
        )
    )

    # Mes previsto de faturamento (Previsao tab): baseado em Dt Chegada Autron
    merged['Mes_Prev_Faturamento'] = merged['FU_Dt_Chegada_Autron'].dt.to_period('M').astype(str)
    merged.loc[merged['FU_Dt_Chegada_Autron'].isna(), 'Mes_Prev_Faturamento'] = 'Sem previsao'

    # Mes de entrega (Faturamento quadro): baseado no campo Entrega de entrada_pedido
    merged['Mes_Entrega'] = merged['Entrega'].dt.to_period('M').astype(str)
    merged.loc[merged['Entrega'].isna(), 'Mes_Entrega'] = 'Sem data'

    # Budget / Metas
    budget = None
    if os.path.exists(arq_budget):
        raw = pd.read_excel(arq_budget, header=1)
        meses_col = [c for c in raw.columns[2:14]]  # 12 meses (Jan-Dez)
        labels = raw.iloc[:, 1].astype(str).str.strip()

        def extrair_meta(secao_inicio, linha_label):
            """Extrai valores mensais de uma linha dentro de uma secao."""
            for i in range(secao_inicio, min(secao_inicio + 12, len(raw))):
                lbl = str(labels.iloc[i]).strip()
                if lbl == linha_label:
                    vals = raw.iloc[i, 2:14].values
                    return [float(v) if pd.notna(v) else 0 for v in vals]
            return [0] * 12

        # Secao AUTRON comeca na linha 0, GRUPO mais abaixo
        autron_idx = 0
        grupo_idx = None
        for i, l in enumerate(labels):
            if str(l).upper() == 'GRUPO':
                grupo_idx = i
                break

        meta_faturamento = extrair_meta(autron_idx, 'Receita de vendas e Serviços')
        meta_entrada_autron = extrair_meta(autron_idx, 'Entrada de pedido')

        # ERGOMEC section
        ergomec_idx = None
        for i, l in enumerate(labels):
            if str(l).upper() == 'ERGOMEC':
                ergomec_idx = i
                break
        meta_entrada_ergomec = extrair_meta(ergomec_idx, 'Entrada de pedido') if ergomec_idx else [0]*12

        # Meta Receita Liquida = GRUPO
        if grupo_idx is not None:
            meta_receita_liquida = extrair_meta(grupo_idx, 'Receita líquida')
            meta_entrada_grupo = extrair_meta(grupo_idx, 'Entrada de pedido')
        else:
            meta_receita_liquida = extrair_meta(autron_idx, 'Receita líquida')
            meta_entrada_grupo = meta_entrada_autron

        meses_str = []
        for c in meses_col:
            if hasattr(c, 'strftime'):
                meses_str.append(pd.Timestamp(c).to_period('M').strftime('%Y-%m'))
            else:
                meses_str.append(str(c))

        budget = pd.DataFrame({
            'Mes': meses_str,
            'Meta_Faturamento': meta_faturamento,
            'Meta_Receita_Liquida': meta_receita_liquida,
            'Meta_Entrada': meta_entrada_grupo,
            'Meta_Entrada_Autron': meta_entrada_autron,
            'Meta_Entrada_Ergomec': meta_entrada_ergomec,
        })

    return merged, fat, budget, None


# ========================================================================
# UPLOAD E GESTAO DE ARQUIVOS
# ========================================================================
os.makedirs(DATA_DIR, exist_ok=True)
ARQUIVOS_NECESSARIOS = ['entrada_pedido.xlsx', 'followup.xlsx', 'mata010.xlsx', 'sciozvs0.csv']
ARQUIVOS_OPCIONAIS = ['faturamento.xlsx', 'Metas_Budget_Grupo_Autron_2026.xlsx']


def verificar_arquivos():
    """Retorna lista de arquivos presentes e ausentes."""
    presentes = []
    ausentes = []
    for f in ARQUIVOS_NECESSARIOS:
        caminho = os.path.join(DATA_DIR, f)
        if os.path.exists(caminho) and os.path.getsize(caminho) > 0:
            mod_time = datetime.fromtimestamp(os.path.getmtime(caminho))
            presentes.append((f, mod_time))
        else:
            ausentes.append(f)
    return presentes, ausentes


def tela_upload():
    """Exibe tela de upload dos arquivos."""
    st.markdown("""
    <div style="text-align: center; padding: 40px 0 20px 0;">
        <h1 style="color: #4FC3F7;">Dashboard de Pedidos</h1>
        <p style="color: #B0BEC5; font-size: 1.1rem;">
            Envie os 4 relatorios para gerar o dashboard
        </p>
    </div>
    """, unsafe_allow_html=True)

    presentes, ausentes = verificar_arquivos()

    if presentes:
        st.markdown("#### ✅ Arquivos ja carregados:")
        for nome, dt in presentes:
            st.markdown(f"- **{nome}** — atualizado em {dt.strftime('%d/%m/%Y %H:%M')}")
        st.markdown("")

    if ausentes:
        st.markdown(f"#### 📤 Arquivos pendentes ({len(ausentes)}):")
        for nome in ausentes:
            st.markdown(f"- ❌ {nome}")
        st.markdown("")

    st.markdown("---")
    st.markdown("### Envie os arquivos:")

    col1, col2, col3 = st.columns(3)
    with col1:
        up_entrada = st.file_uploader(
            "📋 entrada_pedido.xlsx",
            type=['xlsx'], key='up_entrada',
            help="Relatorio de entrada de pedidos do ERP"
        )
        up_followup = st.file_uploader(
            "📅 followup.xlsx",
            type=['xlsx'], key='up_followup',
            help="Relatorio de follow-up de compras"
        )
    with col2:
        up_mata = st.file_uploader(
            "📦 mata010.xlsx",
            type=['xlsx'], key='up_mata',
            help="Relatorio de estoque (MATA010)"
        )
        up_scio = st.file_uploader(
            "🏭 sciozvs0.csv",
            type=['csv'], key='up_scio',
            help="Relatorio SC com classificacao Comprando/Produzindo"
        )
    with col3:
        up_fat = st.file_uploader(
            "💰 faturamento.xlsx",
            type=['xlsx'], key='up_fat',
            help="Relatorio de faturamento (NFs emitidas)"
        )
        up_budget = st.file_uploader(
            "📊 Metas_Budget.xlsx",
            type=['xlsx'], key='up_budget',
            help="Metas e Budget do Grupo Autron 2026"
        )

    uploads = {
        'entrada_pedido.xlsx': up_entrada,
        'followup.xlsx': up_followup,
        'mata010.xlsx': up_mata,
        'sciozvs0.csv': up_scio,
        'faturamento.xlsx': up_fat,
        'Metas_Budget_Grupo_Autron_2026.xlsx': up_budget,
    }

    # Salvar arquivos enviados
    algum_novo = False
    for nome, uploaded in uploads.items():
        if uploaded is not None:
            with open(os.path.join(DATA_DIR, nome), 'wb') as f:
                f.write(uploaded.getbuffer())
            algum_novo = True

    # Verificar novamente apos upload
    presentes, ausentes = verificar_arquivos()

    if len(ausentes) == 0:
        st.success("✅ Todos os arquivos prontos! Processando dashboard...")
        st.cache_data.clear()
        st.rerun()
    elif algum_novo:
        st.info(f"Faltam {len(ausentes)} arquivo(s): {', '.join(ausentes)}")
        st.rerun()
    else:
        st.info(f"Envie os {len(ausentes)} arquivo(s) pendente(s) para continuar.")
        st.stop()


# Verificar se precisa mostrar tela de upload
presentes, ausentes = verificar_arquivos()
if len(ausentes) > 0:
    tela_upload()


# ========================================================================
# CARREGAR DADOS
# ========================================================================
df, df_fat, df_budget, erro = carregar_e_processar()

if erro:
    st.error(f"Erro ao carregar dados: {erro}")
    # Limpar arquivos corrompidos e forcar novo upload
    st.warning("Pode haver um problema com os arquivos. Tente enviar novamente.")
    if st.button("🔄 Reenviar arquivos"):
        for f in ARQUIVOS_NECESSARIOS:
            caminho = os.path.join(DATA_DIR, f)
            if os.path.exists(caminho):
                os.remove(caminho)
        st.cache_data.clear()
        st.rerun()
    st.stop()


# ========================================================================
# SIDEBAR - FILTROS
# ========================================================================
with st.sidebar:
    logo_path = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=180)
    st.markdown("## 🔍 Filtros")
    st.markdown("---")

    # Status
    status_opts = sorted(df['Status_Pedido'].unique())
    status_sel = st.multiselect("Status", status_opts, default=['EM ABERTO'])

    # Ano
    anos = sorted(df['Ano_Emissao'].dropna().unique())
    ano_sel = st.multiselect("Ano Emissao", anos, default=anos[-3:] if len(anos) >= 3 else anos)

    # Mes
    meses_disp = sorted(df[df['Ano_Emissao'].isin(ano_sel)]['Mes_Emissao'].dropna().unique()) if ano_sel else []
    mes_sel = st.multiselect("Mes Emissao", meses_disp, default=[])

    # Tipo Produto
    tipo_opts = sorted(df['Tipo_Produto'].unique())
    tipo_sel = st.multiselect("Tipo Produto", tipo_opts, default=[])

    # Vendedor
    vendedores = sorted(df['Nome do Vendedor'].dropna().unique())
    vend_sel = st.multiselect("Vendedor", vendedores, default=[])

    # Pronto para Fazer
    pronto_opts = sorted(df['Pronto_para_Fazer'].unique())
    pronto_sel = st.multiselect("Pronto p/ Fazer?", pronto_opts, default=[])

    st.markdown("---")

    # Mostrar data dos arquivos carregados
    presentes, _ = verificar_arquivos()
    if presentes:
        st.markdown("##### 📁 Arquivos carregados")
        for nome, dt in presentes:
            st.caption(f"{nome}: {dt.strftime('%d/%m/%Y %H:%M')}")

    # Botao para atualizar arquivos
    if st.button("📤 Atualizar Arquivos", use_container_width=True):
        for f in ARQUIVOS_NECESSARIOS:
            caminho = os.path.join(DATA_DIR, f)
            if os.path.exists(caminho):
                os.remove(caminho)
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    ultima_att = datetime.now().strftime('%d/%m/%Y %H:%M')
    st.caption(f"Ultima atualizacao: {ultima_att}")

    if st.button("🔄 Atualizar Dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # --- Download da planilha consolidada ---
    st.markdown("---")
    try:
        import io
        buffer = io.BytesIO()
        today_ts = pd.Timestamp.now().normalize()
        mes_atual_period = str(today_ts.to_period('M'))

        # --- 1) Base Consolidada ---
        cols_base = [
            'Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
            'Ped Cliente', 'Cliente', 'Razao Social', 'Nome do Vendedor',
            'DT Emissao', 'Mes_Emissao', 'Ano_Emissao',
            'DT. Fat. Cli', 'DT. Ofertada',
            'Quantidade', 'Prc Unitario', 'Vlr.Total', 'Margem',
            'Nota Fiscal', 'Status_Pedido',
            'Tipo_Produto',
            'Estoque_Disponivel', 'Qtd_Alocada', 'Disponivel_Estoque',
            'Numero SC', 'Numero OP', 'FU_OP_na_SC',
            'Acao_Necessaria',
            'FU_Dt_Confirma', 'FU_Dt_Pre_Entr', 'Prazo_Real_Entrega',
            'Semana_Entrega', 'FU_Dt_Chegada_Autron',
            'Mes_Prev_Faturamento',
            'Dias_Atraso_Cliente', 'Dias_Atraso_Ofertada',
            'Pronto_para_Fazer',
        ]
        cols_base = [c for c in cols_base if c in df.columns]
        base_cons = df[cols_base].copy()

        # --- 2) Entrada Pedidos Mes ---
        entrada_mes = df.groupby('Mes_Emissao').agg(
            Qtd_Linhas=('Vlr.Total', 'count'),
            Qtd_PVs=('Num. Pedido', 'nunique'),
            Valor_Total=('Vlr.Total', 'sum'),
        ).reset_index().sort_values('Mes_Emissao')
        entrada_mes.rename(columns={'Mes_Emissao': 'Mes'}, inplace=True)

        # --- 3) Faturamento Mes ---
        fat_mes_export = pd.DataFrame()
        if df_fat is not None:
            fat_mes_export = df_fat.groupby('Mes_Faturamento').agg(
                Qtd_NFs=('Num. Docto.', 'nunique'),
                Qtd_Linhas=('Vlr.Total.Fat', 'count'),
                Valor_Total=('Vlr.Total.Fat', 'sum'),
                Faturamento_Bruto=('Faturamento Bruto', 'sum'),
                Faturamento_Liquido=('Faturamento Liquido', 'sum'),
            ).reset_index().sort_values('Mes_Faturamento')
            fat_mes_export.rename(columns={'Mes_Faturamento': 'Mes'}, inplace=True)

        # --- 4) Previsao Faturamento ---
        abertos_exp = df[df['Status_Pedido'] == 'EM ABERTO'].copy()
        prev_fat = abertos_exp.groupby('Mes_Prev_Faturamento').agg(
            Qtd_Linhas=('Vlr.Total', 'count'),
            Qtd_PVs=('Num. Pedido', 'nunique'),
            Valor_Previsto=('Vlr.Total', 'sum'),
        ).reset_index().sort_values('Mes_Prev_Faturamento')
        prev_fat.rename(columns={'Mes_Prev_Faturamento': 'Mes_Previsto'}, inplace=True)

        # --- 5) Faturado vs A Faturar (mes atual) ---
        valor_faturado = 0
        valor_fat_bruto = 0
        valor_fat_liquido = 0
        qtd_nfs = 0
        if df_fat is not None:
            fat_m = df_fat[df_fat['Mes_Faturamento'] == mes_atual_period]
            valor_faturado = fat_m['Vlr.Total.Fat'].sum()
            valor_fat_bruto = fat_m['Faturamento Bruto'].sum()
            valor_fat_liquido = fat_m['Faturamento Liquido'].sum()
            qtd_nfs = fat_m['Num. Docto.'].nunique() if 'Num. Docto.' in fat_m.columns else 0

        a_faturar_m = abertos_exp[abertos_exp['Mes_Prev_Faturamento'] == mes_atual_period]
        valor_a_faturar = a_faturar_m['Vlr.Total'].sum()
        qtd_pvs_a_fat = a_faturar_m['Num. Pedido'].nunique()

        resumo_mes = pd.DataFrame([
            {'Indicador': 'Mes Referencia', 'Valor': mes_atual_period},
            {'Indicador': 'Valor Faturado (Total)', 'Valor': f"R$ {valor_faturado:,.2f}"},
            {'Indicador': 'Faturamento Bruto', 'Valor': f"R$ {valor_fat_bruto:,.2f}"},
            {'Indicador': 'Faturamento Liquido', 'Valor': f"R$ {valor_fat_liquido:,.2f}"},
            {'Indicador': 'Qtd NFs Emitidas', 'Valor': str(qtd_nfs)},
            {'Indicador': 'Valor a Faturar (Previsao)', 'Valor': f"R$ {valor_a_faturar:,.2f}"},
            {'Indicador': 'Qtd PVs a Faturar', 'Valor': str(qtd_pvs_a_fat)},
            {'Indicador': 'Total Previsto no Mes', 'Valor': f"R$ {valor_faturado + valor_a_faturar:,.2f}"},
        ])

        # --- 6) Detalhe A Faturar Mes ---
        cols_afat = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
                    'Ped Cliente', 'Vlr.Total', 'FU_Dt_Chegada_Autron',
                    'Pronto_para_Fazer', 'Acao_Necessaria']
        cols_afat = [c for c in cols_afat if c in a_faturar_m.columns]
        detalhe_a_faturar = a_faturar_m[cols_afat].copy() if len(a_faturar_m) > 0 else pd.DataFrame(columns=cols_afat)

        # --- 7) Alertas Erro Cadastro ---
        erros = df[df['Acao_Necessaria'] == 'ERRO no CADASTRO']
        cols_erro = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
                    'Tipo_Produto', 'Numero SC', 'Numero OP', 'FU_OP_na_SC',
                    'Acao_Necessaria']
        cols_erro = [c for c in cols_erro if c in erros.columns]
        erros_export = erros[cols_erro].copy() if len(erros) > 0 else pd.DataFrame([{'Info': 'Nenhum erro de cadastro encontrado'}])

        # Escrever arquivo (formato original dashboard_base.xlsx)
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            base_cons.to_excel(writer, sheet_name='Base Consolidada', index=False)
            entrada_mes.to_excel(writer, sheet_name='Entrada Pedidos Mes', index=False)
            if len(fat_mes_export) > 0:
                fat_mes_export.to_excel(writer, sheet_name='Faturamento Mes', index=False)
            prev_fat.to_excel(writer, sheet_name='Previsao Faturamento', index=False)
            resumo_mes.to_excel(writer, sheet_name='Faturado vs A Faturar', index=False)
            detalhe_a_faturar.to_excel(writer, sheet_name='Detalhe A Faturar Mes', index=False)
            erros_export.to_excel(writer, sheet_name='Alertas Erro Cadastro', index=False)
        buffer.seek(0)

        st.download_button(
            label="⬇️ Baixar Planilha",
            data=buffer,
            file_name=f"dashboard_pedidos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.caption(f"Erro ao gerar planilha: {e}")

    # Usuario logado e logout
    if LOGIN_HABILITADO:
        st.markdown("---")
        usuario = st.session_state.get("usuario", "")
        st.caption(f"👤 {usuario}")
        if st.button("🚪 Sair", use_container_width=True):
            st.session_state["autenticado"] = False
            st.session_state["usuario"] = ""
            st.rerun()

# Aplicar filtros
mask = pd.Series(True, index=df.index)
if status_sel:
    mask &= df['Status_Pedido'].isin(status_sel)
if ano_sel:
    mask &= df['Ano_Emissao'].isin(ano_sel)
if mes_sel:
    mask &= df['Mes_Emissao'].isin(mes_sel)
if tipo_sel:
    mask &= df['Tipo_Produto'].isin(tipo_sel)
if vend_sel:
    mask &= df['Nome do Vendedor'].isin(vend_sel)
if pronto_sel:
    mask &= df['Pronto_para_Fazer'].isin(pronto_sel)

filtered = df[mask]


# ========================================================================
# HEADER
# ========================================================================
logo_header = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
if os.path.exists(logo_header):
    st.image(logo_header, width=180)
st.markdown("# Dashboard de Pedidos")
st.markdown(f"<p style='color: #888;'>Dados filtrados: {len(filtered):,} linhas de {len(df):,} | {filtered['Num. Pedido'].nunique()} PVs</p>", unsafe_allow_html=True)


# ========================================================================
# ABAS PRINCIPAIS
# ========================================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📊 Visao Geral", "📋 Entrada de Pedidos", "✅ Prontidao", "📅 Previsao Entrega",
    "📦 Estoque & SC/OP", "💰 Faturamento", "📈 Previsao Faturamento"
])


# ===== ABA 1: VISAO GERAL =====
with tab1:
    filtered_t1 = filtro_data(filtered, 'DT Emissao', 'DT Emissao', 'tab1')

    # KPIs
    total_pvs = filtered_t1['Num. Pedido'].nunique()
    total_linhas = len(filtered_t1)
    em_aberto = len(filtered_t1[filtered_t1['Status_Pedido'] == 'EM ABERTO'])
    finalizados = len(filtered_t1[filtered_t1['Status_Pedido'] == 'FINALIZADO'])
    valor_total = filtered_t1['Vlr.Total'].sum()
    valor_aberto = filtered_t1.loc[filtered_t1['Status_Pedido'] == 'EM ABERTO', 'Vlr.Total'].sum()
    pct = (finalizados / total_linhas * 100) if total_linhas > 0 else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.markdown(kpi_card(f"{total_pvs}", "Total PVs"), unsafe_allow_html=True)
    c2.markdown(kpi_card(f"{total_linhas:,}", "Total Linhas"), unsafe_allow_html=True)
    c3.markdown(kpi_card(f"{em_aberto}", "Em Aberto", "alert"), unsafe_allow_html=True)
    c4.markdown(kpi_card(f"{finalizados}", "Finalizados", "ok"), unsafe_allow_html=True)
    c5.markdown(kpi_card(f"R$ {valor_aberto:,.0f}", "Valor Em Aberto", "warn"), unsafe_allow_html=True)
    c6.markdown(kpi_card(f"{pct:.1f}%", "% Conclusao", "ok" if pct > 80 else "warn"), unsafe_allow_html=True)

    st.markdown("")

    col1, col2 = st.columns(2)

    with col1:
        # Entrada de pedidos por mes
        if len(filtered_t1) > 0:
            pedidos_mes = filtered_t1.groupby('Mes_Emissao').size().reset_index(name='Qtd')
            pedidos_mes = pedidos_mes.sort_values('Mes_Emissao')
            fig = px.bar(pedidos_mes, x='Mes_Emissao', y='Qtd',
                        title='Entrada de Pedidos por Mes',
                        color_discrete_sequence=[CORES['azul_claro']])
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', xaxis_title='', yaxis_title='Quantidade',
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Status
        if len(filtered_t1) > 0:
            status_count = filtered_t1['Status_Pedido'].value_counts().reset_index()
            status_count.columns = ['Status', 'Qtd']
            cores_status = {'FINALIZADO': CORES['verde'], 'EM ABERTO': CORES['vermelho']}
            fig2 = px.pie(status_count, values='Qtd', names='Status',
                         title='Distribuicao por Status',
                         color='Status', color_discrete_map=cores_status,
                         hole=0.4)
            fig2.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=400
            )
            st.plotly_chart(fig2, use_container_width=True)

    # Pedidos por vendedor
    if len(filtered_t1) > 0 and 'Nome do Vendedor' in filtered_t1.columns:
        vend_data = filtered_t1.groupby(['Nome do Vendedor', 'Status_Pedido']).size().reset_index(name='Qtd')
        fig3 = px.bar(vend_data, x='Nome do Vendedor', y='Qtd', color='Status_Pedido',
                     title='Pedidos por Vendedor',
                     color_discrete_map={'FINALIZADO': CORES['verde'], 'EM ABERTO': CORES['vermelho']},
                     barmode='stack')
        fig3.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#E0E0E0', xaxis_title='', yaxis_title='Quantidade',
            height=400
        )
        st.plotly_chart(fig3, use_container_width=True)


# ===== ABA 2: ENTRADA DE PEDIDOS =====
with tab2:
    filtered_t2 = filtro_data(filtered, 'DT Emissao', 'DT Emissao', 'tab2_entrada')

    # KPIs
    total_ep = len(filtered_t2)
    total_pvs_ep = filtered_t2['Num. Pedido'].nunique()
    valor_total_ep = filtered_t2['Vlr.Total'].sum()

    # Entrada 2026
    ep_2026 = filtered_t2[filtered_t2['Ano_Emissao'] == 2026]
    valor_2026 = ep_2026['Vlr.Total'].sum()

    # Separar ERGOMEC (IND21) e AUTRON (demais)
    ep_2026_ergomec = ep_2026[ep_2026['Unidade Negocio'] == 'IND21'] if 'Unidade Negocio' in ep_2026.columns else pd.DataFrame()
    ep_2026_autron = ep_2026[ep_2026['Unidade Negocio'] != 'IND21'] if 'Unidade Negocio' in ep_2026.columns else ep_2026

    # Meta entrada GRUPO (do budget)
    meta_entrada_grupo = 0
    if df_budget is not None and 'Meta_Entrada' in df_budget.columns:
        meta_entrada_grupo = df_budget['Meta_Entrada'].sum()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.markdown(kpi_card(f"{total_pvs_ep}", "Total PVs"), unsafe_allow_html=True)
    c2.markdown(kpi_card(f"{total_ep:,}", "Total Linhas"), unsafe_allow_html=True)
    c3.markdown(kpi_card(f"R$ {valor_total_ep:,.0f}", "Valor Total"), unsafe_allow_html=True)
    c4.markdown(kpi_card(f"R$ {valor_2026:,.0f}", "Valor 2026"), unsafe_allow_html=True)
    c5.markdown(kpi_card(f"R$ {meta_entrada_grupo:,.0f}", "Meta 2026 Grupo"), unsafe_allow_html=True)

    st.markdown("")

    # Entrada de pedidos mes a mes
    entrada_mes = filtered_t2.groupby('Mes_Emissao').agg(
        Qtd_Linhas=('Vlr.Total', 'count'),
        Qtd_PVs=('Num. Pedido', 'nunique'),
        Valor_Total=('Vlr.Total', 'sum'),
    ).reset_index().sort_values('Mes_Emissao')
    entrada_mes.columns = ['Mes', 'Qtd Linhas', 'Qtd PVs', 'Valor Total']

    col1, col2 = st.columns([2, 1])
    with col1:
        fig_ent = px.bar(entrada_mes, x='Mes', y='Valor Total',
                        title='Entrada de Pedidos - Valor por Mes (DT Emissao)',
                        color_discrete_sequence=[CORES['azul_claro']],
                        text='Qtd PVs')
        fig_ent.update_traces(texttemplate='%{text} PVs', textposition='outside')
        fig_ent.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#E0E0E0', xaxis_title='', yaxis_title='R$', height=450
        )
        st.plotly_chart(fig_ent, use_container_width=True)
    with col2:
        st.dataframe(entrada_mes, use_container_width=True, height=450)

    # === Meta x Realizado GRUPO ===
    if df_budget is not None and 'Meta_Entrada' in df_budget.columns:
        st.markdown("### 📊 Meta x Realizado - Grupo (Entrada de Pedidos 2026)")

        ep_mes_2026 = ep_2026.groupby('Mes_Emissao').agg(
            Realizado=('Vlr.Total', 'sum'),
        ).reset_index().rename(columns={'Mes_Emissao': 'Mes'})

        comp_ent = df_budget[['Mes', 'Meta_Entrada']].copy()
        comp_ent.rename(columns={'Meta_Entrada': 'Meta'}, inplace=True)
        comp_ent = comp_ent.merge(ep_mes_2026, on='Mes', how='outer').fillna(0).sort_values('Mes')
        comp_ent = comp_ent[comp_ent['Mes'].str.startswith('2026')]
        comp_ent['Ating. %'] = np.where(
            comp_ent['Meta'] > 0,
            (comp_ent['Realizado'] / comp_ent['Meta'] * 100), 0
        )

        col1, col2 = st.columns([2, 1])
        with col1:
            fig_comp_ent = go.Figure()
            fig_comp_ent.add_trace(go.Bar(
                x=comp_ent['Mes'], y=comp_ent['Meta'], name='Meta',
                marker_color=CORES['azul_claro'], opacity=0.7
            ))
            fig_comp_ent.add_trace(go.Bar(
                x=comp_ent['Mes'], y=comp_ent['Realizado'], name='Realizado',
                marker_color=CORES['verde']
            ))
            fig_comp_ent.update_layout(
                barmode='group',
                title='Meta Grupo x Realizado - 2026',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', xaxis_title='', yaxis_title='R$', height=400,
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
            )
            st.plotly_chart(fig_comp_ent, use_container_width=True)
        with col2:
            comp_ent_display = comp_ent.copy()
            comp_ent_display['Meta'] = comp_ent_display['Meta'].apply(lambda x: f"R$ {x:,.0f}")
            comp_ent_display['Realizado'] = comp_ent_display['Realizado'].apply(lambda x: f"R$ {x:,.0f}")
            comp_ent_display['Ating. %'] = comp_ent_display['Ating. %'].apply(lambda x: f"{x:.1f}%")
            st.dataframe(comp_ent_display, use_container_width=True, height=400)

        # === Meta x Realizado AUTRON (sem IND21) ===
        st.markdown("### 📊 Meta x Realizado - Autron (sem Ergomec)")

        ep_autron_mes = ep_2026_autron.groupby('Mes_Emissao').agg(
            Realizado=('Vlr.Total', 'sum'),
        ).reset_index().rename(columns={'Mes_Emissao': 'Mes'}) if len(ep_2026_autron) > 0 else pd.DataFrame(columns=['Mes', 'Realizado'])

        comp_autron = df_budget[['Mes', 'Meta_Entrada_Autron']].copy()
        comp_autron.rename(columns={'Meta_Entrada_Autron': 'Meta'}, inplace=True)
        comp_autron = comp_autron.merge(ep_autron_mes, on='Mes', how='outer').fillna(0).sort_values('Mes')
        comp_autron = comp_autron[comp_autron['Mes'].str.startswith('2026')]
        comp_autron['Ating. %'] = np.where(comp_autron['Meta'] > 0,
            (comp_autron['Realizado'] / comp_autron['Meta'] * 100), 0)

        col1, col2 = st.columns([2, 1])
        with col1:
            fig_a = go.Figure()
            fig_a.add_trace(go.Bar(x=comp_autron['Mes'], y=comp_autron['Meta'], name='Meta', marker_color='#4472C4', opacity=0.7))
            fig_a.add_trace(go.Bar(x=comp_autron['Mes'], y=comp_autron['Realizado'], name='Realizado', marker_color='#ED7D31'))
            fig_a.update_layout(barmode='group', title='Meta Autron x Realizado - 2026',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', xaxis_title='', yaxis_title='R$', height=350,
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5))
            st.plotly_chart(fig_a, use_container_width=True)
        with col2:
            ca_d = comp_autron.copy()
            ca_d['Meta'] = ca_d['Meta'].apply(lambda x: f"R$ {x:,.0f}")
            ca_d['Realizado'] = ca_d['Realizado'].apply(lambda x: f"R$ {x:,.0f}")
            ca_d['Ating. %'] = ca_d['Ating. %'].apply(lambda x: f"{x:.1f}%")
            st.dataframe(ca_d, use_container_width=True, height=350)

        # === Meta x Realizado ERGOMEC (IND21) ===
        st.markdown("### 📊 Meta x Realizado - Ergomec (IND21)")

        ep_ergomec_mes = ep_2026_ergomec.groupby('Mes_Emissao').agg(
            Realizado=('Vlr.Total', 'sum'),
        ).reset_index().rename(columns={'Mes_Emissao': 'Mes'}) if len(ep_2026_ergomec) > 0 else pd.DataFrame(columns=['Mes', 'Realizado'])

        comp_ergomec = df_budget[['Mes', 'Meta_Entrada_Ergomec']].copy()
        comp_ergomec.rename(columns={'Meta_Entrada_Ergomec': 'Meta'}, inplace=True)
        comp_ergomec = comp_ergomec.merge(ep_ergomec_mes, on='Mes', how='outer').fillna(0).sort_values('Mes')
        comp_ergomec = comp_ergomec[comp_ergomec['Mes'].str.startswith('2026')]
        comp_ergomec['Ating. %'] = np.where(comp_ergomec['Meta'] > 0,
            (comp_ergomec['Realizado'] / comp_ergomec['Meta'] * 100), 0)

        col1, col2 = st.columns([2, 1])
        with col1:
            fig_e = go.Figure()
            fig_e.add_trace(go.Bar(x=comp_ergomec['Mes'], y=comp_ergomec['Meta'], name='Meta', marker_color='#4472C4', opacity=0.7))
            fig_e.add_trace(go.Bar(x=comp_ergomec['Mes'], y=comp_ergomec['Realizado'], name='Realizado', marker_color='#70AD47'))
            fig_e.update_layout(barmode='group', title='Meta Ergomec x Realizado - 2026',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', xaxis_title='', yaxis_title='R$', height=350,
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5))
            st.plotly_chart(fig_e, use_container_width=True)
        with col2:
            ce_d = comp_ergomec.copy()
            ce_d['Meta'] = ce_d['Meta'].apply(lambda x: f"R$ {x:,.0f}")
            ce_d['Realizado'] = ce_d['Realizado'].apply(lambda x: f"R$ {x:,.0f}")
            ce_d['Ating. %'] = ce_d['Ating. %'].apply(lambda x: f"{x:.1f}%")
            st.dataframe(ce_d, use_container_width=True, height=350)


# ===== ABA 3: PRONTIDAO =====
with tab3:
    filtered_t3 = filtro_data(filtered, 'DT Emissao', 'DT Emissao', 'tab3_prontidao')
    abertos = filtered_t3[filtered_t3['Status_Pedido'] == 'EM ABERTO']

    prontos = len(abertos[abertos['Pronto_para_Fazer'] == 'SIM'])
    parciais = len(abertos[abertos['Pronto_para_Fazer'].str.contains('PARCIAL', na=False)])
    nao_prontos = len(abertos[abertos['Pronto_para_Fazer'] == 'NAO'])
    erros = len(abertos[abertos['Acao_Necessaria'] == 'ERRO no CADASTRO'])

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(kpi_card(f"{prontos}", "Prontos p/ Fazer", "ok"), unsafe_allow_html=True)
    c2.markdown(kpi_card(f"{parciais}", "Parcialmente Prontos", "warn"), unsafe_allow_html=True)
    c3.markdown(kpi_card(f"{nao_prontos}", "Nao Prontos", "alert"), unsafe_allow_html=True)
    c4.markdown(kpi_card(f"{erros}", "ERROS Cadastro", "alert" if erros > 0 else "ok"), unsafe_allow_html=True)

    st.markdown("")

    col1, col2 = st.columns(2)

    with col1:
        if len(abertos) > 0:
            pronto_count = abertos['Pronto_para_Fazer'].value_counts().reset_index()
            pronto_count.columns = ['Status', 'Qtd']
            cores_pronto = {
                'SIM': CORES['verde'], 'NAO': CORES['vermelho'],
                'PARCIAL - Sem Follow-up': CORES['amarelo'],
                'PARCIAL - Sem Estoque': '#E67E22'
            }
            fig = px.pie(pronto_count, values='Qtd', names='Status',
                        title='Distribuicao - Pronto p/ Fazer?',
                        color='Status', color_discrete_map=cores_pronto, hole=0.4)
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=400
            )
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if len(abertos) > 0:
            tipo_count = abertos['Tipo_Produto'].value_counts().reset_index()
            tipo_count.columns = ['Tipo', 'Qtd']
            fig2 = px.bar(tipo_count, x='Tipo', y='Qtd',
                         title='Comprando vs Produzindo (Em Aberto)',
                         color='Tipo',
                         color_discrete_map={'Comprando': CORES['azul_claro'], 'Produzindo': CORES['amarelo'],
                                           'Nao classificado': CORES['cinza']})
            fig2.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=400, showlegend=False
            )
            st.plotly_chart(fig2, use_container_width=True)

    # Tabela detalhada
    st.markdown("### 📋 Detalhes - Pedidos Em Aberto")

    if len(abertos) > 0:
        tab_cols = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto', 'Tipo_Produto',
                   'Quantidade', 'Pronto_para_Fazer', 'Disponivel_Estoque', 'Acao_Necessaria',
                   'Prazo_Real_Entrega', 'Semana_Entrega', 'Ped Cliente']
        tab_cols = [c for c in tab_cols if c in abertos.columns]

        display_df = abertos[tab_cols].copy()
        display_df.columns = ['PV', 'Item', 'Produto', 'Descricao', 'Tipo',
                             'Qtd', 'Pronto?', 'Estoque', 'Acao',
                             'Prazo Entrega', 'Semana', 'Ped Cliente'][:len(tab_cols)]

        def color_pronto(val):
            if val == 'SIM': return 'background-color: #1a3a1a; color: #2ECC71'
            elif 'PARCIAL' in str(val): return 'background-color: #3a3a1a; color: #F39C12'
            elif val == 'NAO': return 'background-color: #3a1a1a; color: #E74C3C'
            return ''

        def color_acao(val):
            if val == 'ERRO no CADASTRO': return 'background-color: #C0392B; color: white; font-weight: bold'
            elif val == 'Estoque OK': return 'background-color: #1a3a1a; color: #2ECC71'
            elif 'Necessario' in str(val): return 'background-color: #3a1a1a; color: #E74C3C'
            elif 'gerada' in str(val): return 'background-color: #3a3a1a; color: #F39C12'
            return ''

        styled = display_df.style.applymap(color_pronto, subset=['Pronto?'] if 'Pronto?' in display_df.columns else [])
        if 'Acao' in display_df.columns:
            styled = styled.applymap(color_acao, subset=['Acao'])

        st.dataframe(styled, use_container_width=True, height=500)
    else:
        st.info("Nenhum pedido em aberto com os filtros selecionados.")


# ===== ABA 4: PREVISAO ENTREGA =====
with tab4:
    filtered_t4 = filtro_data(filtered, 'DT Emissao', 'DT Emissao', 'tab4_previsao_entrega')
    abertos = filtered_t4[filtered_t4['Status_Pedido'] == 'EM ABERTO']

    atrasados = len(abertos[abertos['Dias_Atraso_Cliente'] > 0])
    media_atraso = abertos.loc[abertos['Dias_Atraso_Cliente'] > 0, 'Dias_Atraso_Cliente'].mean()
    max_atraso = abertos['Dias_Atraso_Cliente'].max()

    c1, c2, c3 = st.columns(3)
    c1.markdown(kpi_card(f"{atrasados}", "Pedidos Atrasados", "alert" if atrasados > 0 else "ok"), unsafe_allow_html=True)
    c2.markdown(kpi_card(f"{media_atraso:.0f} dias" if pd.notna(media_atraso) else "0", "Media Atraso", "warn"), unsafe_allow_html=True)
    c3.markdown(kpi_card(f"{max_atraso:.0f} dias" if pd.notna(max_atraso) else "0", "Maior Atraso", "alert"), unsafe_allow_html=True)

    st.markdown("")

    col1, col2 = st.columns(2)

    with col1:
        # Pedidos por semana de entrega
        semana_data = abertos[abertos['Semana_Entrega'].notna()]
        if len(semana_data) > 0:
            sem_count = semana_data['Semana_Entrega'].value_counts().reset_index()
            sem_count.columns = ['Semana', 'Qtd']
            sem_count = sem_count.sort_values('Semana')
            fig = px.bar(sem_count, x='Semana', y='Qtd',
                        title='Pedidos por Semana de Entrega (Pasta)',
                        color_discrete_sequence=[CORES['azul_claro']])
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=400
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhum pedido com semana de entrega definida.")

    with col2:
        # Top 10 mais atrasados
        top_atrasados = abertos.nlargest(10, 'Dias_Atraso_Cliente')[
            ['Num. Pedido', 'Produto', 'Dias_Atraso_Cliente']
        ].dropna(subset=['Dias_Atraso_Cliente'])

        if len(top_atrasados) > 0:
            top_atrasados['Label'] = top_atrasados['Num. Pedido'] + ' - ' + top_atrasados['Produto']
            fig2 = px.bar(top_atrasados, x='Dias_Atraso_Cliente', y='Label',
                         title='Top 10 - Mais Atrasados (vs Cliente)',
                         orientation='h', color_discrete_sequence=[CORES['vermelho']])
            fig2.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', yaxis_title='', xaxis_title='Dias de Atraso',
                height=400
            )
            st.plotly_chart(fig2, use_container_width=True)

    # Tabela previsao
    st.markdown("### 📋 Previsao de Entrega")

    if len(abertos) > 0:
        ent_cols = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto', 'Quantidade',
                   'Ped Cliente', 'DT. Fat. Cli', 'DT. Ofertada', 'Prazo_Real_Entrega',
                   'Semana_Entrega', 'FU_Dt_Chegada_Autron', 'Dias_Atraso_Cliente', 'Pronto_para_Fazer']
        ent_cols = [c for c in ent_cols if c in abertos.columns]

        ent_df = abertos[ent_cols].copy().sort_values('Dias_Atraso_Cliente', ascending=False)
        rename_map = {
            'Num. Pedido': 'PV', 'Descricao do Produto': 'Descricao',
            'Quantidade': 'Qtd', 'Ped Cliente': 'Ped.Cliente',
            'DT. Fat. Cli': 'Dt.Solicitada', 'DT. Ofertada': 'Dt.Ofertada',
            'Prazo_Real_Entrega': 'Prazo Real', 'Semana_Entrega': 'Semana',
            'FU_Dt_Chegada_Autron': 'Chegada Autron',
            'Dias_Atraso_Cliente': 'Dias Atraso', 'Pronto_para_Fazer': 'Pronto?'
        }
        ent_df = ent_df.rename(columns=rename_map)

        def color_atraso(val):
            if pd.isna(val): return ''
            if val > 30: return 'background-color: #3a1a1a; color: #E74C3C'
            elif val > 0: return 'background-color: #3a3a1a; color: #F39C12'
            else: return 'background-color: #1a3a1a; color: #2ECC71'

        styled = ent_df.style
        if 'Dias Atraso' in ent_df.columns:
            styled = styled.applymap(color_atraso, subset=['Dias Atraso'])

        st.dataframe(styled, use_container_width=True, height=500)


# ===== ABA 5: ESTOQUE & SC/OP =====
with tab5:
    filtered_t5 = filtro_data(filtered, 'DT Emissao', 'DT Emissao', 'tab5_estoque')
    abertos = filtered_t5[filtered_t5['Status_Pedido'] == 'EM ABERTO']

    com_estoque = len(abertos[abertos['Disponivel_Estoque'] == 'SIM'])
    sem_estoque = len(abertos[abertos['Disponivel_Estoque'] == 'NAO'])
    parcial_est = len(abertos[abertos['Disponivel_Estoque'] == 'PARCIAL'])
    necessita_sc = len(abertos[abertos['Acao_Necessaria'] == 'Necessario gerar SC'])
    necessita_op = len(abertos[abertos['Acao_Necessaria'] == 'Necessario gerar OP'])
    erros = len(abertos[abertos['Acao_Necessaria'] == 'ERRO no CADASTRO'])

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.markdown(kpi_card(f"{com_estoque}", "Com Estoque", "ok"), unsafe_allow_html=True)
    c2.markdown(kpi_card(f"{parcial_est}", "Estoque Parcial", "warn"), unsafe_allow_html=True)
    c3.markdown(kpi_card(f"{sem_estoque}", "Sem Estoque", "alert"), unsafe_allow_html=True)
    c4.markdown(kpi_card(f"{necessita_sc}", "Necessitam SC", "alert" if necessita_sc > 0 else "ok"), unsafe_allow_html=True)
    c5.markdown(kpi_card(f"{necessita_op}", "Necessitam OP", "alert" if necessita_op > 0 else "ok"), unsafe_allow_html=True)
    c6.markdown(kpi_card(f"{erros}", "ERROS Cadastro", "alert" if erros > 0 else "ok"), unsafe_allow_html=True)

    st.markdown("")

    col1, col2 = st.columns(2)

    with col1:
        if len(abertos) > 0:
            est_count = abertos['Disponivel_Estoque'].value_counts().reset_index()
            est_count.columns = ['Disponibilidade', 'Qtd']
            cores_est = {'SIM': CORES['verde'], 'NAO': CORES['vermelho'], 'PARCIAL': CORES['amarelo']}
            fig = px.pie(est_count, values='Qtd', names='Disponibilidade',
                        title='Disponibilidade de Estoque',
                        color='Disponibilidade', color_discrete_map=cores_est, hole=0.4)
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=380
            )
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if len(abertos) > 0:
            acao_count = abertos[abertos['Acao_Necessaria'] != 'Finalizado']['Acao_Necessaria'].value_counts().reset_index()
            acao_count.columns = ['Acao', 'Qtd']
            cores_acao = {
                'Estoque OK': CORES['verde'],
                'Necessario gerar SC': CORES['vermelho'],
                'Necessario gerar OP': '#E67E22',
                'SC gerada - Aguardando': CORES['amarelo'],
                'OP gerada - Aguardando': '#F1C40F',
                'ERRO no CADASTRO': CORES['vermelho_escuro'],
                'Verificar classificacao': CORES['cinza']
            }
            fig2 = px.bar(acao_count, x='Acao', y='Qtd',
                         title='Acoes Necessarias',
                         color='Acao', color_discrete_map=cores_acao)
            fig2.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', height=380, showlegend=False,
                xaxis_title='', yaxis_title='Quantidade'
            )
            st.plotly_chart(fig2, use_container_width=True)

    # Tabela estoque
    st.markdown("### 📋 Detalhe Estoque & SC/OP")

    if len(abertos) > 0:
        est_cols = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto', 'Tipo_Produto',
                   'Quantidade', 'Estoque_Disponivel', 'Qtd_Alocada', 'Disponivel_Estoque',
                   'Acao_Necessaria', 'Numero SC', 'Numero OP', 'FU_OP_na_SC', 'DT Emissao']
        est_cols = [c for c in est_cols if c in abertos.columns]

        est_df = abertos[est_cols].copy().sort_values(['Produto', 'DT Emissao'])
        rename_est = {
            'Num. Pedido': 'PV', 'Descricao do Produto': 'Descricao',
            'Tipo_Produto': 'Tipo', 'Quantidade': 'Qtd Pedida',
            'Estoque_Disponivel': 'Estoque', 'Qtd_Alocada': 'Alocado',
            'Disponivel_Estoque': 'Disponivel', 'Acao_Necessaria': 'Acao',
            'Numero SC': 'SC', 'Numero OP': 'OP', 'FU_OP_na_SC': 'OP na SC',
            'DT Emissao': 'Dt.Emissao'
        }
        est_df = est_df.rename(columns=rename_est)

        def color_acao_tab(val):
            if val == 'ERRO no CADASTRO': return 'background-color: #C0392B; color: white; font-weight: bold'
            elif val == 'Estoque OK': return 'background-color: #1a3a1a; color: #2ECC71'
            elif 'Necessario' in str(val): return 'background-color: #3a1a1a; color: #E74C3C'
            elif 'gerada' in str(val): return 'background-color: #3a3a1a; color: #F39C12'
            return ''

        styled = est_df.style
        if 'Acao' in est_df.columns:
            styled = styled.applymap(color_acao_tab, subset=['Acao'])

        st.dataframe(styled, use_container_width=True, height=500)

    # Alertas de erro
    erros_df = abertos[abertos['Acao_Necessaria'] == 'ERRO no CADASTRO']
    if len(erros_df) > 0:
        st.markdown("### ⚠️ ERROS DE CADASTRO - Item Comprando com OP")
        st.error(f"Foram encontrados {len(erros_df)} itens classificados como 'Comprando' que possuem OP gerada. Verificar cadastro!")

        erro_display = erros_df[['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto', 'Numero OP', 'FU_OP_na_SC']].copy()
        erro_display.columns = ['PV', 'Item', 'Produto', 'Descricao', 'OP (Entrada)', 'OP na SC (Follow-up)']
        st.dataframe(erro_display, use_container_width=True)


# ===== ABA 6: FATURAMENTO =====
with tab6:
    today_period = pd.Timestamp.now().normalize().to_period('M')
    mes_atual_str = str(today_period)

    if df_fat is not None and len(df_fat) > 0:
        df_fat_f = filtro_data(df_fat, 'Emissao', 'Emissao NF', 'tab6_fat')

        # --- Dados base ---
        meses_com_fat = sorted(df_fat['Mes_Faturamento'].dropna().unique())
        ultimo_mes_fechado = meses_com_fat[-1] if meses_com_fat else mes_atual_str

        # Mes anterior ao atual (regra: primeiro bloco sempre mostra mes anterior)
        mes_anterior = str((today_period - 1))

        # Faturamento LIQUIDO por mes (base do realizado)
        fat_liq_mes = df_fat.groupby('Mes_Faturamento').agg(
            Fat_Bruto=('Faturamento Bruto', 'sum'),
            Fat_Liquido=('Faturamento Liquido', 'sum'),
        ).reset_index().rename(columns={'Mes_Faturamento': 'Mes'})

        # Total a faturar por mes (pedidos em aberto, campo Entrega de entrada_pedido)
        abertos_all = df[df['Status_Pedido'] == 'EM ABERTO'].copy()
        a_faturar_mes = abertos_all[abertos_all['Mes_Entrega'] != 'Sem data'].groupby(
            'Mes_Entrega'
        ).agg(Total_a_Faturar=('Vlr.Total', 'sum')).reset_index().rename(
            columns={'Mes_Entrega': 'Mes'}
        )

        # Montar quadro consolidado 2026
        if df_budget is not None and 'Meta_Receita_Liquida' in df_budget.columns:
            quadro = df_budget[['Mes', 'Meta_Receita_Liquida']].copy()
            quadro.rename(columns={'Meta_Receita_Liquida': 'Receita_Liquida_Meta'}, inplace=True)
        else:
            quadro = pd.DataFrame({'Mes': [f'2026-{m:02d}' for m in range(1, 13)], 'Receita_Liquida_Meta': 0})

        quadro = quadro.merge(fat_liq_mes, on='Mes', how='left')
        quadro = quadro.merge(a_faturar_mes, on='Mes', how='left')
        quadro = quadro.fillna(0).sort_values('Mes')
        quadro = quadro[quadro['Mes'].str.startswith('2026')].copy()

        # Status: fechado, atual, futuro
        quadro['Status_Mes'] = quadro['Mes'].apply(
            lambda m: 'Fechado' if m < mes_atual_str else ('Atual' if m == mes_atual_str else 'Futuro')
        )

        # Para meses fechados, Total a faturar = 0 (ja realizado)
        quadro.loc[quadro['Status_Mes'] == 'Fechado', 'Total_a_Faturar'] = 0

        # Total líquido a faturar = Total a faturar - 17%
        quadro['Total_Liq_a_Faturar'] = quadro['Total_a_Faturar'] * 0.83

        # Total líquido = Fat. Líquido (realizado) + Total líq a faturar
        quadro['Total_Liquido'] = quadro['Fat_Liquido'] + quadro['Total_Liq_a_Faturar']

        # Realizado vs Meta (so para meses com dados)
        tem_dados = (quadro['Fat_Liquido'] > 0) | (quadro['Total_a_Faturar'] > 0)
        quadro['Diferenca_Meta'] = np.where(tem_dados,
            quadro['Total_Liquido'] - quadro['Receita_Liquida_Meta'], np.nan)
        quadro['Ating_Pct'] = np.where(
            tem_dados & (quadro['Receita_Liquida_Meta'] > 0),
            (quadro['Total_Liquido'] / quadro['Receita_Liquida_Meta'] * 100), np.nan
        )

        # --- KPIs do mes anterior (sempre o mes anterior ao atual) ---
        if mes_anterior not in quadro['Mes'].values:
            mes_anterior = ultimo_mes_fechado
        ult = quadro[quadro['Mes'] == mes_anterior]
        if len(ult) > 0:
            ult = ult.iloc[0]
            st.markdown(f"#### Resultado {mes_anterior}")
            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(kpi_card(f"R$ {ult['Receita_Liquida_Meta']:,.0f}", "Meta (Rec. Líquida)"), unsafe_allow_html=True)
            c2.markdown(kpi_card(f"R$ {ult['Fat_Liquido']:,.0f}", "Fat. Líquido",
                        "ok" if ult['Fat_Liquido'] >= ult['Receita_Liquida_Meta'] else "warn"), unsafe_allow_html=True)
            c3.markdown(kpi_card(f"{ult['Ating_Pct']:.0f}%", "Atingimento",
                        "ok" if ult['Ating_Pct'] >= 100 else ("warn" if ult['Ating_Pct'] >= 80 else "alert")), unsafe_allow_html=True)
            c4.markdown(kpi_card(f"R$ {ult['Diferenca_Meta']:+,.0f}", "Diferenca",
                        "ok" if ult['Diferenca_Meta'] >= 0 else "alert"), unsafe_allow_html=True)

        # --- KPIs do mes atual com Total Liquido Projetado ---
        atual = quadro[quadro['Mes'] == mes_atual_str]
        if len(atual) > 0:
            atual = atual.iloc[0]
            total_liq_proj = atual['Fat_Liquido'] + atual['Total_Liq_a_Faturar']
            st.markdown(f"#### Resultado {mes_atual_str}")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.markdown(kpi_card(f"R$ {atual['Receita_Liquida_Meta']:,.0f}", "Meta (Rec. Líquida)"), unsafe_allow_html=True)
            c2.markdown(kpi_card(f"R$ {atual['Fat_Liquido']:,.0f}", "Fat. Líquido Atual"), unsafe_allow_html=True)
            c3.markdown(kpi_card(f"R$ {atual['Total_Liq_a_Faturar']:,.0f}", "Líq. a Faturar (-17%)"), unsafe_allow_html=True)
            c4.markdown(kpi_card(f"R$ {total_liq_proj:,.0f}", "Total Líquido Projetado", "ok" if total_liq_proj >= atual['Receita_Liquida_Meta'] else "warn"), unsafe_allow_html=True)
            ating_atual = (total_liq_proj / atual['Receita_Liquida_Meta'] * 100) if atual['Receita_Liquida_Meta'] > 0 else 0
            c5.markdown(kpi_card(f"{ating_atual:.0f}%", "Atingimento Projetado",
                        "ok" if ating_atual >= 100 else ("warn" if ating_atual >= 80 else "alert")), unsafe_allow_html=True)

        # --- Consolidado Q1 (fechado) ---
        quadro['Trimestre'] = quadro['Mes'].apply(
            lambda m: f"Q{(int(m.split('-')[1]) - 1) // 3 + 1}"
        )

        q1_data = quadro[quadro['Trimestre'] == 'Q1']
        meta_q1 = q1_data['Receita_Liquida_Meta'].sum()
        fat_liq_q1 = q1_data['Fat_Liquido'].sum()
        total_liq_q1 = q1_data['Total_Liquido'].sum()
        diff_q1 = fat_liq_q1 - meta_q1
        ating_q1 = (fat_liq_q1 / meta_q1 * 100) if meta_q1 > 0 else 0

        st.markdown("#### Consolidado Q1")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(kpi_card(f"R$ {meta_q1:,.0f}", "Meta Q1"), unsafe_allow_html=True)
        c2.markdown(kpi_card(f"R$ {fat_liq_q1:,.0f}", "Fat. Líquido Q1"), unsafe_allow_html=True)
        c3.markdown(kpi_card(f"R$ {total_liq_q1:,.0f}", "Total Líquido Q1"), unsafe_allow_html=True)
        c4.markdown(kpi_card(f"R$ {diff_q1:+,.0f}", "Diferenca",
                    "ok" if diff_q1 >= 0 else "alert"), unsafe_allow_html=True)
        c5.markdown(kpi_card(f"{ating_q1:.0f}%", "Atingimento",
                    "ok" if ating_q1 >= 100 else ("warn" if ating_q1 >= 80 else "alert")), unsafe_allow_html=True)

        # --- Consolidado trimestre atual (se diferente de Q1) ---
        tri_atual = f"Q{(today_period.month - 1) // 3 + 1}"
        if tri_atual != 'Q1':
            tri_data = quadro[quadro['Trimestre'] == tri_atual]
            meta_tri = tri_data['Receita_Liquida_Meta'].sum()
            fat_liq_tri = tri_data['Fat_Liquido'].sum()
            total_liq_tri = tri_data['Total_Liquido'].sum()
            diff_tri = total_liq_tri - meta_tri
            ating_tri = (total_liq_tri / meta_tri * 100) if meta_tri > 0 else 0

            st.markdown(f"#### Consolidado {tri_atual}")
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.markdown(kpi_card(f"R$ {meta_tri:,.0f}", f"Meta {tri_atual}"), unsafe_allow_html=True)
            c2.markdown(kpi_card(f"R$ {fat_liq_tri:,.0f}", "Fat. Líquido"), unsafe_allow_html=True)
            c3.markdown(kpi_card(f"R$ {total_liq_tri:,.0f}", f"Total Líquido {tri_atual}"), unsafe_allow_html=True)
            c4.markdown(kpi_card(f"R$ {diff_tri:+,.0f}", "Diferenca",
                        "ok" if diff_tri >= 0 else "alert"), unsafe_allow_html=True)
            c5.markdown(kpi_card(f"{ating_tri:.0f}%", "Atingimento",
                        "ok" if ating_tri >= 100 else ("warn" if ating_tri >= 80 else "alert")), unsafe_allow_html=True)

        # --- Total de Carteira (todos pedidos em aberto) ---
        carteira = df[df['Status_Pedido'] == 'EM ABERTO']
        carteira_bruto = carteira['Vlr.Total'].sum()
        carteira_liquido = carteira_bruto * 0.83
        carteira_pvs = carteira['Num. Pedido'].nunique()
        carteira_linhas = len(carteira)

        st.markdown("#### Total de Carteira")
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(kpi_card(f"R$ {carteira_bruto:,.0f}", "Valor Bruto"), unsafe_allow_html=True)
        c2.markdown(kpi_card(f"R$ {carteira_liquido:,.0f}", "Valor Líquido (-17%)"), unsafe_allow_html=True)
        c3.markdown(kpi_card(f"{carteira_pvs}", "PVs em Aberto"), unsafe_allow_html=True)
        c4.markdown(kpi_card(f"{carteira_linhas}", "Linhas em Aberto"), unsafe_allow_html=True)

        st.markdown("")

        # --- COMPARATIVO 2025 x 2026 (Fat. Liquido mes a mes) ---
        st.markdown("### 📊 Faturamento Líquido - 2025 x 2026")

        fat_anos = df_fat.copy()
        fat_anos['Ano'] = fat_anos['Emissao'].dt.year
        fat_anos['Mes_Num'] = fat_anos['Emissao'].dt.month

        fat_25 = fat_anos[fat_anos['Ano'] == 2025].groupby('Mes_Num')['Faturamento Liquido'].sum()
        fat_26 = fat_anos[fat_anos['Ano'] == 2026].groupby('Mes_Num')['Faturamento Liquido'].sum()

        meses_nomes_curto = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
        comp_ano = pd.DataFrame({
            'Mes': meses_nomes_curto,
            '2025': [fat_25.get(m+1, 0) for m in range(12)],
            '2026': [fat_26.get(m+1, 0) for m in range(12)],
        })
        comp_ano['Diferenca_Pct'] = np.where(
            comp_ano['2025'] > 0,
            ((comp_ano['2026'] - comp_ano['2025']) / comp_ano['2025'] * 100),
            np.nan
        )

        fig_anos = go.Figure()
        fig_anos.add_trace(go.Bar(
            x=comp_ano['Mes'], y=comp_ano['2025'], name='2025',
            marker_color='#4472C4',
            text=comp_ano['2025'].apply(lambda x: f"R$ {x:,.0f}".replace(',', '.') if x > 0 else ''),
            textposition='outside', textfont=dict(size=10, color='#4472C4'),
        ))
        fig_anos.add_trace(go.Bar(
            x=comp_ano['Mes'], y=comp_ano['2026'], name='2026',
            marker_color='#70AD47',
            text=comp_ano.apply(
                lambda r: f"R$ {r['2026']:,.0f}".replace(',', '.') +
                          (f"<br>{r['Diferenca_Pct']:+.1f}%" if pd.notna(r['Diferenca_Pct']) and r['2026'] > 0 else '')
                if r['2026'] > 0 else '', axis=1
            ),
            textposition='outside', textfont=dict(size=10, color='#70AD47'),
        ))
        fig_anos.update_layout(
            barmode='group',
            title='Faturamento Líquido - 2025 x 2026',
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#E0E0E0', xaxis_title='', yaxis_title='R$',
            height=500,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
            yaxis=dict(showticklabels=False, showgrid=False),
        )
        st.plotly_chart(fig_anos, use_container_width=True)

        # Totais anuais
        total_25 = comp_ano['2025'].sum()
        total_26 = comp_ano['2026'].sum()
        diff_total = ((total_26 - total_25) / total_25 * 100) if total_25 > 0 else 0
        c1, c2, c3 = st.columns(3)
        c1.markdown(kpi_card(f"R$ {total_25:,.0f}", "Total 2025"), unsafe_allow_html=True)
        c2.markdown(kpi_card(f"R$ {total_26:,.0f}", "Total 2026"), unsafe_allow_html=True)
        c3.markdown(kpi_card(f"{diff_total:+.1f}%", "Variacao", "ok" if diff_total >= 0 else "alert"), unsafe_allow_html=True)

        st.markdown("")

        # --- QUADRO PRINCIPAL (estilo planilha da Daniele) ---
        st.markdown("### 📊 Quadro Faturamento - Meta x Realizado")

        # Tabela transposta (meses como colunas)
        meses_2026 = quadro['Mes'].tolist()
        meses_labels = [m.replace('2026-', '') for m in meses_2026]

        linhas = {
            'Receita Líquida (Meta)': quadro['Receita_Liquida_Meta'].tolist(),
            'Total a Faturar': quadro['Total_a_Faturar'].tolist(),
            'Faturamento Líquido': quadro['Fat_Liquido'].tolist(),
            'Total Líq. a Faturar (-17%)': quadro['Total_Liq_a_Faturar'].tolist(),
            'Total Líquido': quadro['Total_Liquido'].tolist(),
            'Realizado vs Meta': quadro['Diferenca_Meta'].tolist(),
            'Atingimento %': quadro['Ating_Pct'].tolist(),
        }

        quadro_display = pd.DataFrame(linhas, index=meses_labels).T
        quadro_display.index.name = 'Indicador'

        # Formatar valores
        quadro_fmt = quadro_display.copy()
        for col in quadro_fmt.columns:
            for idx in quadro_fmt.index:
                val = quadro_fmt.at[idx, col]
                if pd.isna(val):
                    quadro_fmt.at[idx, col] = "-"
                elif idx == 'Atingimento %':
                    quadro_fmt.at[idx, col] = f"{val:.0f}%"
                elif val == 0:
                    quadro_fmt.at[idx, col] = "-"
                else:
                    quadro_fmt.at[idx, col] = f"R$ {val:,.0f}"

        st.dataframe(quadro_fmt, use_container_width=True, height=320)

        st.markdown("")

        # --- GRAFICO FATURAMENTO: Meta x Realizado x Previsao ---
        # Realizado (Fat. Bruto) por mes - apenas 2026
        fat_real_2026 = df_fat.groupby('Mes_Faturamento').agg(
            Realizado=('Faturamento Bruto', 'sum'),
        ).reset_index().rename(columns={'Mes_Faturamento': 'Mes'})
        fat_real_2026 = fat_real_2026[fat_real_2026['Mes'].str.startswith('2026')]

        # Montar grafico com os 12 meses usando dados do quadro
        graf = quadro[['Mes', 'Receita_Liquida_Meta', 'Fat_Liquido', 'Total_Liquido', 'Status_Mes']].copy()
        graf.rename(columns={'Receita_Liquida_Meta': 'Meta'}, inplace=True)
        graf = graf.merge(fat_real_2026, on='Mes', how='left')
        graf['Realizado'] = graf['Realizado'].fillna(0)
        # Previsao = Total Liquido para meses atuais/futuros com a faturar
        graf['Previsao'] = np.where(
            graf['Status_Mes'] != 'Fechado', graf['Total_Liquido'], 0
        )

        # Labels no formato "jan-26"
        meses_nomes = {
            '01': 'jan', '02': 'fev', '03': 'mar', '04': 'abr',
            '05': 'mai', '06': 'jun', '07': 'jul', '08': 'ago',
            '09': 'set', '10': 'out', '11': 'nov', '12': 'dez'
        }
        graf['Label'] = graf['Mes'].apply(
            lambda m: f"{meses_nomes.get(m.split('-')[1], m.split('-')[1])}-{m.split('-')[0][2:]}"
        )

        fig_fat = go.Figure()
        # Meta (azul)
        fig_fat.add_trace(go.Bar(
            x=graf['Label'], y=graf['Meta'], name='Meta',
            marker_color='#4472C4',
            text=graf['Meta'].apply(lambda x: f"{x:,.0f}".replace(',', '.') if x > 0 else ''),
            textposition='outside', textfont=dict(size=10)
        ))
        # Realizado (laranja)
        fig_fat.add_trace(go.Bar(
            x=graf['Label'], y=graf['Realizado'], name='Realizado',
            marker_color='#ED7D31',
            text=graf['Realizado'].apply(lambda x: f"{x:,.0f}".replace(',', '.') if x > 0 else ''),
            textposition='outside', textfont=dict(size=10, color='#ED7D31')
        ))
        # Previsao (verde) = Total Liquido para meses nao fechados
        fig_fat.add_trace(go.Bar(
            x=graf['Label'], y=graf['Previsao'], name='Previsao',
            marker_color='#70AD47',
            text=graf['Previsao'].apply(lambda x: f"{x:,.0f}".replace(',', '.') if x > 0 else ''),
            textposition='outside', textfont=dict(size=10, color='#70AD47')
        ))

        fig_fat.update_layout(
            barmode='group',
            title=dict(text='Faturamento', font=dict(size=20), x=0.5),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font_color='#E0E0E0', xaxis_title='', yaxis_title='',
            height=500,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
            yaxis=dict(showticklabels=False, showgrid=False),
            margin=dict(t=80)
        )
        st.plotly_chart(fig_fat, use_container_width=True)

        # Detalhe a faturar (campo Entrega)
        abertos_mes_atual = df[
            (df['Status_Pedido'] == 'EM ABERTO') &
            (df['Mes_Entrega'] == mes_atual_str)
        ]
        if len(abertos_mes_atual) > 0:
            st.markdown(f"### 📋 Detalhe - A Faturar em {mes_atual_str} (campo Entrega)")
            cols_afat = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
                        'Ped Cliente', 'Vlr.Total', 'Entrega', 'FU_Dt_Chegada_Autron',
                        'Pronto_para_Fazer', 'Acao_Necessaria']
            cols_afat = [c for c in cols_afat if c in abertos_mes_atual.columns]
            st.dataframe(abertos_mes_atual[cols_afat], use_container_width=True, height=300)
    else:
        st.warning("Arquivo faturamento.xlsx nao carregado. Envie o arquivo para ver esta aba.")


# ===== ABA 7: PREVISAO FATURAMENTO =====
with tab7:
    filtered_t7 = filtro_data(filtered, 'DT Emissao', 'DT Emissao', 'tab7_previsao')
    abertos = filtered_t7[filtered_t7['Status_Pedido'] == 'EM ABERTO']

    if len(abertos) > 0:
        mes_atual_prev = str(pd.Timestamp.now().normalize().to_period('M'))

        # Filtrar apenas do mes atual em diante (sem datas passadas)
        prev_fat = abertos[
            (abertos['Mes_Prev_Faturamento'] != 'Sem previsao') &
            (abertos['Mes_Prev_Faturamento'] >= mes_atual_prev)
        ].groupby('Mes_Prev_Faturamento').agg(
            Qtd_Linhas=('Vlr.Total', 'count'),
            Qtd_PVs=('Num. Pedido', 'nunique'),
            Valor_Previsto=('Vlr.Total', 'sum'),
        ).reset_index().sort_values('Mes_Prev_Faturamento')

        # KPIs
        com_previsao = abertos[
            (abertos['Mes_Prev_Faturamento'] != 'Sem previsao') &
            (abertos['Mes_Prev_Faturamento'] >= mes_atual_prev)
        ]
        sem_previsao = abertos[
            (abertos['Mes_Prev_Faturamento'] == 'Sem previsao') |
            (abertos['Mes_Prev_Faturamento'] < mes_atual_prev)
        ]
        valor_com = com_previsao['Vlr.Total'].sum()
        valor_sem = sem_previsao['Vlr.Total'].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(kpi_card(f"{len(com_previsao)}", "Com Previsao", "ok"), unsafe_allow_html=True)
        c2.markdown(kpi_card(f"R$ {valor_com:,.0f}", "Valor c/ Previsao", "ok"), unsafe_allow_html=True)
        c3.markdown(kpi_card(f"{len(sem_previsao)}", "Sem Previsao", "alert"), unsafe_allow_html=True)
        c4.markdown(kpi_card(f"R$ {valor_sem:,.0f}", "Valor s/ Previsao", "warn"), unsafe_allow_html=True)

        st.markdown("")
        st.caption("Datas usadas: Dt Chegada Autron > Prazo Real Entrega > DT. Ofertada")

        # Grafico previsao - com VALOR nas barras
        if len(prev_fat) > 0:
            fig_prev = px.bar(prev_fat, x='Mes_Prev_Faturamento', y='Valor_Previsto',
                             title='Previsao de Faturamento Mes a Mes',
                             color_discrete_sequence=[CORES['azul_claro']],
                             text='Valor_Previsto')
            fig_prev.update_traces(texttemplate='R$ %{text:,.0f}', textposition='outside')
            fig_prev.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                font_color='#E0E0E0', xaxis_title='', yaxis_title='R$', height=450
            )
            st.plotly_chart(fig_prev, use_container_width=True)

        # Tabela previsao detalhada
        st.markdown("### 📋 Previsao Detalhada")
        prev_display = prev_fat.copy()
        prev_display.columns = ['Mes Previsto', 'Qtd Linhas', 'Qtd PVs', 'Valor Previsto']
        st.dataframe(prev_display, use_container_width=True, height=300)
    else:
        st.info("Nenhum pedido em aberto com os filtros selecionados.")
