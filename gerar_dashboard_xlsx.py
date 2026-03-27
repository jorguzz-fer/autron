#!/usr/bin/env python3
"""
Gera planilha consolidada dashboard_pedidos_powerbi.xlsx
com todas as visoes solicitadas para Power BI.
"""
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ========================================================================
# CARGA DE DADOS
# ========================================================================
print("Carregando arquivos...")

ep = pd.read_excel('entrada_pedido.xlsx', header=1)
fu = pd.read_excel('followup.xlsx', header=1)
mt = pd.read_excel('mata010.xlsx', header=1)
fat = pd.read_excel('faturamento.xlsx', header=1)
sc_csv = pd.read_csv('sciozvs0.csv', encoding='latin-1', sep=';', header=2, low_memory=False)

print(f"  entrada_pedido: {ep.shape}")
print(f"  followup:       {fu.shape}")
print(f"  mata010:        {mt.shape}")
print(f"  faturamento:    {fat.shape}")
print(f"  sciozvs0:       {sc_csv.shape}")

# ========================================================================
# LIMPEZA ENTRADA DE PEDIDOS
# ========================================================================
print("Processando entrada de pedidos...")

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

# ========================================================================
# LIMPEZA FOLLOW-UP
# ========================================================================
print("Processando follow-up...")

fu['No. da S.C.'] = pd.to_numeric(fu['No. da S.C.'], errors='coerce')
fu['Dt. Confirma'] = pd.to_datetime(fu['Dt. Confirma'], errors='coerce')
fu['Dt. Pre entr'] = pd.to_datetime(fu['Dt. Pre entr'], errors='coerce')
fu['Dt Chegada Autron'] = pd.to_datetime(fu['Dt Chegada Autron'], errors='coerce')
fu['PV informado na SC'] = pd.to_numeric(fu['PV informado na SC'], errors='coerce')
fu['Numero do PV'] = fu['Numero do PV'].astype(str).str.replace('.', '', regex=False).str.strip()
if 'OP na SC' not in fu.columns:
    fu['OP na SC'] = np.nan

# ========================================================================
# LIMPEZA ESTOQUE
# ========================================================================
print("Processando estoque...")

mt.columns = mt.columns.astype(str).str.strip()
mt['Saldo Atual'] = pd.to_numeric(mt['Saldo Atual'], errors='coerce').fillna(0)

# ========================================================================
# LIMPEZA FATURAMENTO
# ========================================================================
print("Processando faturamento...")

fat.columns = fat.columns.astype(str).str.strip()
fat['Emissao'] = pd.to_datetime(fat['Emissao'], errors='coerce')
vlr_col = [c for c in fat.columns if 'vlr' in c.lower() and 'total' in c.lower()]
if vlr_col:
    fat.rename(columns={vlr_col[0]: 'Vlr_Total_Fat'}, inplace=True)
else:
    fat['Vlr_Total_Fat'] = 0
fat['Vlr_Total_Fat'] = pd.to_numeric(fat['Vlr_Total_Fat'], errors='coerce').fillna(0)
if 'Faturamento Bruto' in fat.columns:
    fat['Faturamento Bruto'] = pd.to_numeric(fat['Faturamento Bruto'], errors='coerce').fillna(0)
if 'Faturamento Liquido' in fat.columns:
    fat['Faturamento Liquido'] = pd.to_numeric(fat['Faturamento Liquido'], errors='coerce').fillna(0)
if 'No do Pedido' in fat.columns:
    fat['No do Pedido'] = fat['No do Pedido'].astype(str).str.replace('.', '', regex=False).str.strip()
fat['Mes_Faturamento'] = fat['Emissao'].dt.to_period('M').astype(str)

# ========================================================================
# CLASSIFICACAO COMPRANDO/PRODUZINDO (sciozvs0)
# ========================================================================
print("Processando classificacao Comprando/Produzindo...")

sc_csv['Produto'] = sc_csv['Produto'].astype(str).str.strip()
sc_csv['Prod.Solict'] = sc_csv['Prod.Solict'].astype(str).str.strip()
tipo_prod = sc_csv.groupby('Produto')['Prod.Solict'].agg(
    lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'Indefinido'
).reset_index()
tipo_prod.rename(columns={'Prod.Solict': 'Tipo_Produto'}, inplace=True)

# ========================================================================
# FOLLOW-UP: AGRUPAMENTO POR SC
# ========================================================================
print("Cruzando follow-up por SC...")

fu_por_sc = fu.groupby('No. da S.C.').agg({
    'Dt. Confirma': 'first', 'Dt. Pre entr': 'first', 'Pasta': 'first',
    'No. do P.O.': 'first', 'Dt Chegada Autron': 'first', 'OP na SC': 'first'
}).reset_index()
fu_por_sc.rename(columns={'No. da S.C.': 'Numero SC'}, inplace=True)

# ========================================================================
# FOLLOW-UP: AGRUPAMENTO POR PV + PRODUTO
# ========================================================================
print("Cruzando follow-up por PV+Produto...")

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

# ========================================================================
# MERGES
# ========================================================================
print("Mesclando dados...")

merged = ep.merge(fu_por_sc, on='Numero SC', how='left', suffixes=('', '_fu'))
merged = merged.merge(fu_por_pv, on=['Num. Pedido', 'Produto'], how='left')

# Consolidar colunas do follow-up
merged['FU_Dt_Confirma'] = merged['Dt. Confirma'].combine_first(merged['FU_Dt_Confirma_PV'])
merged['FU_Dt_Pre_Entr'] = merged['Dt. Pre entr'].combine_first(merged['FU_Dt_Pre_entr_PV'])
merged['FU_Pasta'] = merged['Pasta'].combine_first(merged['FU_Pasta_PV'])
merged['FU_PO'] = merged['No. do P.O.'].combine_first(merged['FU_PO_PV'])
merged['FU_Dt_Chegada_Autron'] = merged['Dt Chegada Autron'].combine_first(merged['FU_Dt_Chegada_PV'])
merged['FU_OP_na_SC'] = merged['OP na SC'].combine_first(merged['FU_OP_na_SC_PV'])

# Prazo real = Dt Confirma, se vazio usa Dt Pre entr
merged['Prazo_Real_Entrega'] = merged['FU_Dt_Confirma'].combine_first(merged['FU_Dt_Pre_Entr'])
# Semana de entrega (do campo Pasta) so quando Dt Confirma preenchido
merged['Semana_Entrega'] = np.where(merged['FU_Dt_Confirma'].notna(), merged['FU_Pasta'], None)

# Tipo produto
merged = merged.merge(tipo_prod, on='Produto', how='left')
merged['Tipo_Produto'] = merged['Tipo_Produto'].fillna('Nao classificado')

# Status (Nota Fiscal preenchida = FINALIZADO)
merged['Status_Pedido'] = np.where(merged['Nota Fiscal'].notna(), 'FINALIZADO', 'EM ABERTO')

# ========================================================================
# ESTOQUE COM ALOCACAO POR PRIORIDADE (data emissao mais antiga primeiro)
# ========================================================================
print("Calculando estoque e alocacao...")

stock = mt.groupby('Codigo')['Saldo Atual'].sum().reset_index()
stock.rename(columns={'Codigo': 'Produto', 'Saldo Atual': 'Estoque_Disponivel'}, inplace=True)
merged = merged.merge(stock, on='Produto', how='left')
merged['Estoque_Disponivel'] = merged['Estoque_Disponivel'].fillna(0)

# Alocacao somente para pedidos em aberto
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

# ========================================================================
# REGRAS SC / OP
# ========================================================================
print("Aplicando regras SC/OP...")

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

# ========================================================================
# INDICADORES DE ATRASO
# ========================================================================
print("Calculando indicadores...")

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

# Pronto para fazer
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

# Previsao mes (baseado em Dt Chegada Autron)
merged['Mes_Previsao_Faturamento'] = merged['FU_Dt_Chegada_Autron'].dt.to_period('M').astype(str)

# ========================================================================
# MONTAR ABAS DA PLANILHA
# ========================================================================
print("Gerando planilha Excel...")

# --- ABA 1: Pedidos Consolidado ---
cols_consolidado = [
    'Num. Pedido', 'Item', 'Produto', 'Descricao do Produto', 'Tipo_Produto',
    'Quantidade', 'Prc Unitario', 'Vlr.Total', 'Margem',
    'DT Emissao', 'DT. Ofertada', 'DT. Fat. Cli', 'Ped Cliente',
    'Cliente', 'Razao Social', 'Nome Fantasia', 'Nome do Vendedor',
    'Nota Fiscal', 'Status_Pedido',
    'Prazo_Real_Entrega', 'Semana_Entrega', 'FU_Dt_Chegada_Autron',
    'Estoque_Disponivel', 'Qtd_Alocada', 'Disponivel_Estoque',
    'Numero SC', 'Numero OP', 'FU_OP_na_SC', 'FU_PO',
    'Acao_Necessaria', 'Pronto_para_Fazer',
    'Dias_Atraso_Cliente', 'Dias_Atraso_Ofertada',
    'Mes_Emissao', 'Mes_Previsao_Faturamento',
    'TP Venda (PV)', 'Tipo Negocio (PV)', 'Unidade Negocio',
    'Segmento 1', 'Nome do Segmento 1', 'Estado', 'Regional (PV)'
]
cols_consolidado = [c for c in cols_consolidado if c in merged.columns]
aba_consolidado = merged[cols_consolidado].copy()

# --- ABA 2: Pedidos Em Aberto (detalhado) ---
abertos = merged[merged['Status_Pedido'] == 'EM ABERTO'].copy()
cols_abertos = [
    'Num. Pedido', 'Item', 'Produto', 'Descricao do Produto', 'Tipo_Produto',
    'Quantidade', 'Vlr.Total', 'DT Emissao', 'DT. Ofertada', 'DT. Fat. Cli',
    'Ped Cliente', 'Nome do Vendedor', 'Razao Social',
    'Prazo_Real_Entrega', 'Semana_Entrega', 'FU_Dt_Chegada_Autron',
    'Estoque_Disponivel', 'Qtd_Alocada', 'Disponivel_Estoque',
    'Numero SC', 'Numero OP', 'FU_OP_na_SC', 'FU_PO',
    'Acao_Necessaria', 'Pronto_para_Fazer',
    'Dias_Atraso_Cliente', 'Dias_Atraso_Ofertada'
]
cols_abertos = [c for c in cols_abertos if c in abertos.columns]
aba_abertos = abertos[cols_abertos].sort_values(['Num. Pedido', 'Item'])

# --- ABA 3: Entrada de Pedidos Mes a Mes ---
entrada_mes = merged.groupby('Mes_Emissao').agg(
    Qtd_Linhas=('Mes_Emissao', 'size'),
    Qtd_PVs=('Num. Pedido', 'nunique'),
    Valor_Total=('Vlr.Total', 'sum')
).reset_index().sort_values('Mes_Emissao')
entrada_mes.rename(columns={'Mes_Emissao': 'Mes'}, inplace=True)

# --- ABA 4: Faturamento Realizado Mes a Mes ---
fat_mes = fat.groupby('Mes_Faturamento').agg(
    Qtd_NFs=('Num. Docto.', 'nunique'),
    Valor_Faturado=('Vlr_Total_Fat', 'sum'),
).reset_index().sort_values('Mes_Faturamento')
fat_mes.rename(columns={'Mes_Faturamento': 'Mes'}, inplace=True)

if 'Faturamento Bruto' in fat.columns:
    fat_bruto = fat.groupby(fat['Emissao'].dt.to_period('M').astype(str))['Faturamento Bruto'].sum().reset_index()
    fat_bruto.columns = ['Mes', 'Faturamento_Bruto']
    fat_mes = fat_mes.merge(fat_bruto, on='Mes', how='left')

if 'Faturamento Liquido' in fat.columns:
    fat_liq = fat.groupby(fat['Emissao'].dt.to_period('M').astype(str))['Faturamento Liquido'].sum().reset_index()
    fat_liq.columns = ['Mes', 'Faturamento_Liquido']
    fat_mes = fat_mes.merge(fat_liq, on='Mes', how='left')

# --- ABA 5: Previsao Faturamento Mes a Mes (Dt Chegada Autron) ---
abertos_prev = abertos[abertos['FU_Dt_Chegada_Autron'].notna()].copy()
if len(abertos_prev) > 0:
    prev_mes = abertos_prev.groupby('Mes_Previsao_Faturamento').agg(
        Qtd_Linhas=('Mes_Previsao_Faturamento', 'size'),
        Qtd_PVs=('Num. Pedido', 'nunique'),
        Valor_Previsto=('Vlr.Total', 'sum')
    ).reset_index().sort_values('Mes_Previsao_Faturamento')
    prev_mes.rename(columns={'Mes_Previsao_Faturamento': 'Mes'}, inplace=True)
else:
    prev_mes = pd.DataFrame(columns=['Mes', 'Qtd_Linhas', 'Qtd_PVs', 'Valor_Previsto'])

# --- ABA 6: Mes Atual - Faturado vs A Faturar ---
mes_atual_str = today.to_period('M').strftime('%Y-%m')

fat_mes_atual = fat[fat['Mes_Faturamento'] == mes_atual_str]
valor_faturado_mes = fat_mes_atual['Vlr_Total_Fat'].sum()

abertos_chegada = abertos[abertos['FU_Dt_Chegada_Autron'].notna()].copy()
abertos_chegada['Mes_Chegada'] = abertos_chegada['FU_Dt_Chegada_Autron'].dt.to_period('M').astype(str)
a_faturar_mes = abertos_chegada[abertos_chegada['Mes_Chegada'] == mes_atual_str]
valor_a_faturar_mes = a_faturar_mes['Vlr.Total'].sum()

resumo_mes = pd.DataFrame([{
    'Mes_Referencia': mes_atual_str,
    'Valor_Faturado': valor_faturado_mes,
    'Valor_A_Faturar': valor_a_faturar_mes,
    'Total_Previsto': valor_faturado_mes + valor_a_faturar_mes,
    'Pct_Faturado': (valor_faturado_mes / (valor_faturado_mes + valor_a_faturar_mes) * 100)
        if (valor_faturado_mes + valor_a_faturar_mes) > 0 else 0
}])

# Detalhe dos itens a faturar no mes atual
if len(a_faturar_mes) > 0:
    cols_det = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
                'Quantidade', 'Vlr.Total', 'FU_Dt_Chegada_Autron',
                'Pronto_para_Fazer', 'Ped Cliente', 'Nome do Vendedor', 'Razao Social']
    cols_det = [c for c in cols_det if c in a_faturar_mes.columns]
    detalhe_a_faturar = a_faturar_mes[cols_det].sort_values('FU_Dt_Chegada_Autron')
else:
    detalhe_a_faturar = pd.DataFrame()

# --- ABA 7: Faturamento Detalhado (do relatorio faturamento.xlsx) ---
cols_fat_det = [
    'Emissao', 'Num. Docto.', 'Produto', 'Descricao Produto', 'Quantidade',
    'Vlr_Total_Fat', 'No do Pedido', 'Item Pv', 'Razao Social', 'Nome Fantasia',
    'UF', 'Nome do Vendedor', 'Tipo Negocio', 'Mes_Faturamento'
]
if 'Faturamento Bruto' in fat.columns:
    cols_fat_det.append('Faturamento Bruto')
if 'Faturamento Liquido' in fat.columns:
    cols_fat_det.append('Faturamento Liquido')
cols_fat_det = [c for c in cols_fat_det if c in fat.columns]
aba_fat_det = fat[cols_fat_det].sort_values('Emissao', ascending=False)

# ========================================================================
# EXPORTAR XLSX
# ========================================================================
output_file = 'dashboard_pedidos_powerbi_REV2.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    aba_consolidado.to_excel(writer, sheet_name='Pedidos_Consolidado', index=False)
    aba_abertos.to_excel(writer, sheet_name='Pedidos_Em_Aberto', index=False)
    entrada_mes.to_excel(writer, sheet_name='Entrada_Pedidos_Mes', index=False)
    fat_mes.to_excel(writer, sheet_name='Faturamento_Mes', index=False)
    prev_mes.to_excel(writer, sheet_name='Previsao_Faturamento_Mes', index=False)
    resumo_mes.to_excel(writer, sheet_name='Mes_Atual_Fat_vs_AFat', index=False)
    if len(detalhe_a_faturar) > 0:
        detalhe_a_faturar.to_excel(writer, sheet_name='Detalhe_A_Faturar_Mes', index=False)
    aba_fat_det.to_excel(writer, sheet_name='Faturamento_Detalhado', index=False)

print(f"\n✅ Arquivo gerado: {output_file}")
print(f"\nAbas:")
print(f"  1. Pedidos_Consolidado       - {len(aba_consolidado)} linhas (todos os pedidos com todas as informacoes)")
print(f"  2. Pedidos_Em_Aberto         - {len(aba_abertos)} linhas (somente pedidos sem NF)")
print(f"  3. Entrada_Pedidos_Mes       - {len(entrada_mes)} meses (entrada de pedidos por DT Emissao)")
print(f"  4. Faturamento_Mes           - {len(fat_mes)} meses (faturamento realizado por Emissao NF)")
print(f"  5. Previsao_Faturamento_Mes  - {len(prev_mes)} meses (previsao baseada em Dt Chegada Autron)")
print(f"  6. Mes_Atual_Fat_vs_AFat     - Resumo mes atual: faturado vs a faturar")
if len(detalhe_a_faturar) > 0:
    print(f"  7. Detalhe_A_Faturar_Mes     - {len(detalhe_a_faturar)} itens a faturar no mes atual")
print(f"  8. Faturamento_Detalhado     - {len(aba_fat_det)} linhas (NFs detalhadas)")
