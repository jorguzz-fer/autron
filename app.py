import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import io
import json
import tempfile

# ========================================================================
# CONFIGURACAO
# ========================================================================
st.set_page_config(
    page_title="Dashboard de Pedidos",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "dados"))

# Google Drive - IDs das pastas/arquivos (configurar via variavel de ambiente ou secrets)
# Para usar: coloque os arquivos numa pasta do Google Drive e compartilhe com o service account
GDRIVE_ENABLED = os.environ.get("GDRIVE_ENABLED", "false").lower() == "true"
GDRIVE_FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID", "")
GDRIVE_CREDENTIALS = os.environ.get("GDRIVE_CREDENTIALS", "")


# ========================================================================
# GOOGLE DRIVE - DOWNLOAD DOS ARQUIVOS
# ========================================================================
def baixar_do_gdrive():
    """Baixa os 4 arquivos do Google Drive para a pasta local."""
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        # Carregar credenciais
        if os.path.exists(GDRIVE_CREDENTIALS):
            creds = service_account.Credentials.from_service_account_file(
                GDRIVE_CREDENTIALS,
                scopes=['https://www.googleapis.com/auth/drive.readonly']
            )
        else:
            creds_json = json.loads(GDRIVE_CREDENTIALS)
            creds = service_account.Credentials.from_service_account_info(
                creds_json,
                scopes=['https://www.googleapis.com/auth/drive.readonly']
            )

        service = build('drive', 'v3', credentials=creds)

        # Listar arquivos na pasta
        arquivos_esperados = {
            'entrada_pedido.xlsx': None,
            'followup.xlsx': None,
            'mata010.xlsx': None,
            'sciozmq0.csv': None
        }

        results = service.files().list(
            q=f"'{GDRIVE_FOLDER_ID}' in parents and trashed=false",
            fields="files(id, name, modifiedTime)"
        ).execute()

        for f in results.get('files', []):
            if f['name'] in arquivos_esperados:
                arquivos_esperados[f['name']] = f['id']

        # Verificar se todos foram encontrados
        faltando = [n for n, fid in arquivos_esperados.items() if fid is None]
        if faltando:
            return False, f"Arquivos nao encontrados no Google Drive: {', '.join(faltando)}"

        # Baixar cada arquivo
        os.makedirs(DATA_DIR, exist_ok=True)
        for nome, file_id in arquivos_esperados.items():
            request = service.files().get_media(fileId=file_id)
            destino = os.path.join(DATA_DIR, nome)
            with open(destino, 'wb') as fout:
                downloader = MediaIoBaseDownload(fout, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

        return True, "Arquivos baixados com sucesso do Google Drive!"

    except ImportError:
        return False, "Bibliotecas do Google Drive nao instaladas. Instale: pip install google-api-python-client google-auth"
    except Exception as e:
        return False, f"Erro ao acessar Google Drive: {str(e)}"

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


# ========================================================================
# CARGA E PROCESSAMENTO DE DADOS
# ========================================================================
@st.cache_data(ttl=300)
def carregar_e_processar():
    """Carrega os 4 arquivos e processa toda a logica de negocio."""

    arq_entrada = os.path.join(DATA_DIR, "entrada_pedido.xlsx")
    arq_followup = os.path.join(DATA_DIR, "followup.xlsx")
    arq_estoque = os.path.join(DATA_DIR, "mata010.xlsx")
    arq_sciozmq = os.path.join(DATA_DIR, "sciozmq0.csv")

    for arq in [arq_entrada, arq_followup, arq_estoque, arq_sciozmq]:
        if not os.path.exists(arq):
            return None, f"Arquivo nao encontrado: {os.path.basename(arq)}"

    # Carregar
    ep = pd.read_excel(arq_entrada, header=1)
    fu = pd.read_excel(arq_followup, header=1)
    mt = pd.read_excel(arq_estoque)
    sc_csv = pd.read_csv(arq_sciozmq, encoding='latin-1', sep=';', header=2, low_memory=False)

    # Limpar entrada_pedido
    ep['Num. Pedido'] = ep['Num. Pedido'].astype(str).str.replace('.', '', regex=False).str.strip()
    ep['Numero SC'] = pd.to_numeric(ep['Numero SC'], errors='coerce')
    ep['Quantidade'] = pd.to_numeric(ep['Quantidade'], errors='coerce')
    ep['DT Emissao'] = pd.to_datetime(ep['DT Emissao'], errors='coerce')
    ep['DT. Ofertada'] = pd.to_datetime(ep['DT. Ofertada'], errors='coerce')
    ep['DT. Fat. Cli'] = pd.to_datetime(ep['DT. Fat. Cli'], errors='coerce')
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

    # Limpar estoque
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

    return merged, None


# ========================================================================
# GOOGLE DRIVE - SINCRONIZAR SE HABILITADO
# ========================================================================
if GDRIVE_ENABLED and GDRIVE_FOLDER_ID:
    with st.spinner("Sincronizando com Google Drive..."):
        ok, msg = baixar_do_gdrive()
        if not ok:
            st.warning(f"Google Drive: {msg}")


# ========================================================================
# UPLOAD MANUAL (fallback se nao tem Google Drive nem arquivos locais)
# ========================================================================
os.makedirs(DATA_DIR, exist_ok=True)
arquivos_necessarios = ['entrada_pedido.xlsx', 'followup.xlsx', 'mata010.xlsx', 'sciozmq0.csv']
arquivos_existentes = [f for f in arquivos_necessarios if os.path.exists(os.path.join(DATA_DIR, f))]

if len(arquivos_existentes) < 4:
    st.markdown("# 📦 Dashboard de Pedidos")
    st.markdown("### 📤 Upload dos Arquivos")
    st.markdown("Envie os 4 arquivos para iniciar o dashboard:")

    col1, col2 = st.columns(2)
    with col1:
        up_entrada = st.file_uploader("entrada_pedido.xlsx", type=['xlsx'], key='up1')
        up_followup = st.file_uploader("followup.xlsx", type=['xlsx'], key='up2')
    with col2:
        up_mata = st.file_uploader("mata010.xlsx", type=['xlsx'], key='up3')
        up_scio = st.file_uploader("sciozmq0.csv", type=['csv'], key='up4')

    uploads = {'entrada_pedido.xlsx': up_entrada, 'followup.xlsx': up_followup,
               'mata010.xlsx': up_mata, 'sciozmq0.csv': up_scio}

    if all(u is not None for u in uploads.values()):
        for nome, uploaded in uploads.items():
            with open(os.path.join(DATA_DIR, nome), 'wb') as f:
                f.write(uploaded.getbuffer())
        st.success("Arquivos carregados! Reprocessando...")
        st.cache_data.clear()
        st.rerun()
    else:
        enviados = sum(1 for u in uploads.values() if u is not None)
        st.info(f"Enviados {enviados}/4 arquivos. Envie todos para continuar.")
        st.stop()


# ========================================================================
# CARREGAR DADOS
# ========================================================================
df, erro = carregar_e_processar()

if erro:
    st.error(f"Erro ao carregar dados: {erro}")
    st.info("Coloque os 4 arquivos na pasta 'dados/' ou use o upload acima.")
    st.stop()


# ========================================================================
# SIDEBAR - FILTROS
# ========================================================================
with st.sidebar:
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
    ultima_att = datetime.now().strftime('%d/%m/%Y %H:%M')
    st.caption(f"Ultima atualizacao: {ultima_att}")

    if st.button("🔄 Atualizar Dados", use_container_width=True):
        st.cache_data.clear()
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
st.markdown("# 📦 Dashboard de Pedidos")
st.markdown(f"<p style='color: #888;'>Dados filtrados: {len(filtered):,} linhas de {len(df):,} | {filtered['Num. Pedido'].nunique()} PVs</p>", unsafe_allow_html=True)


# ========================================================================
# ABAS PRINCIPAIS
# ========================================================================
tab1, tab2, tab3, tab4 = st.tabs(["📊 Visao Geral", "✅ Prontidao", "📅 Previsao Entrega", "📦 Estoque & SC/OP"])


# ===== ABA 1: VISAO GERAL =====
with tab1:
    # KPIs
    total_pvs = filtered['Num. Pedido'].nunique()
    total_linhas = len(filtered)
    em_aberto = len(filtered[filtered['Status_Pedido'] == 'EM ABERTO'])
    finalizados = len(filtered[filtered['Status_Pedido'] == 'FINALIZADO'])
    valor_total = filtered['Vlr.Total'].sum()
    valor_aberto = filtered.loc[filtered['Status_Pedido'] == 'EM ABERTO', 'Vlr.Total'].sum()
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
        if len(filtered) > 0:
            pedidos_mes = filtered.groupby('Mes_Emissao').size().reset_index(name='Qtd')
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
        if len(filtered) > 0:
            status_count = filtered['Status_Pedido'].value_counts().reset_index()
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
    if len(filtered) > 0 and 'Nome do Vendedor' in filtered.columns:
        vend_data = filtered.groupby(['Nome do Vendedor', 'Status_Pedido']).size().reset_index(name='Qtd')
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


# ===== ABA 2: PRONTIDAO =====
with tab2:
    abertos = filtered[filtered['Status_Pedido'] == 'EM ABERTO']

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


# ===== ABA 3: PREVISAO ENTREGA =====
with tab3:
    abertos = filtered[filtered['Status_Pedido'] == 'EM ABERTO']

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


# ===== ABA 4: ESTOQUE & SC/OP =====
with tab4:
    abertos = filtered[filtered['Status_Pedido'] == 'EM ABERTO']

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
