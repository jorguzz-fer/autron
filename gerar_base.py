"""
Gera planilha consolidada para Power BI com todas as regras de negocio.
Saida: dashboard_base.xlsx com multiplas abas.
"""
import pandas as pd
import numpy as np
from datetime import datetime
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TODAY = pd.Timestamp.now().normalize()
MES_ATUAL = TODAY.to_period('M')

print("=" * 60)
print("GERACAO DA BASE CONSOLIDADA PARA POWER BI")
print("=" * 60)


# ========================================================================
# 1. CARGA DOS ARQUIVOS
# ========================================================================
print("\n[1/7] Carregando arquivos...")

ep = pd.read_excel(os.path.join(BASE_DIR, "entrada_pedido.xlsx"), header=1)
fu = pd.read_excel(os.path.join(BASE_DIR, "followup.xlsx"), header=1)
mt = pd.read_excel(os.path.join(BASE_DIR, "mata010.xlsx"), header=1)
mt.columns = mt.columns.astype(str).str.strip()
fat = pd.read_excel(os.path.join(BASE_DIR, "faturamento.xlsx"), header=1)
sc_csv = pd.read_csv(
    os.path.join(BASE_DIR, "sciozvs0.csv"),
    encoding='latin-1', sep=';', header=2, low_memory=False
)

print(f"  entrada_pedido: {len(ep)} linhas")
print(f"  followup:       {len(fu)} linhas")
print(f"  mata010:        {len(mt)} linhas")
print(f"  faturamento:    {len(fat)} linhas")
print(f"  sciozvs0:       {len(sc_csv)} linhas")


# ========================================================================
# 2. LIMPEZA E PREPARACAO
# ========================================================================
print("\n[2/7] Limpando dados...")

# --- Entrada de Pedidos ---
ep['Num. Pedido'] = ep['Num. Pedido'].astype(str).str.replace('.', '', regex=False).str.strip()
ep['Numero SC'] = pd.to_numeric(ep['Numero SC'], errors='coerce')
ep['Quantidade'] = pd.to_numeric(ep['Quantidade'], errors='coerce')
ep['Vlr.Total'] = pd.to_numeric(ep['Vlr.Total'], errors='coerce')
ep['Margem'] = pd.to_numeric(ep['Margem'], errors='coerce')
ep['DT Emissao'] = pd.to_datetime(ep['DT Emissao'], errors='coerce')
ep['DT. Ofertada'] = pd.to_datetime(ep['DT. Ofertada'], errors='coerce')
ep['DT. Fat. Cli'] = pd.to_datetime(ep['DT. Fat. Cli'], errors='coerce')
ep['Nota Fiscal'] = pd.to_numeric(ep['Nota Fiscal'], errors='coerce')
ep['Item'] = pd.to_numeric(ep['Item'], errors='coerce')
if 'Numero OP' not in ep.columns:
    ep['Numero OP'] = np.nan

# --- Follow-up ---
fu['No. da S.C.'] = pd.to_numeric(fu['No. da S.C.'], errors='coerce')
fu['Dt. Confirma'] = pd.to_datetime(fu['Dt. Confirma'], errors='coerce')
fu['Dt. Pre entr'] = pd.to_datetime(fu['Dt. Pre entr'], errors='coerce')
fu['Dt Chegada Autron'] = pd.to_datetime(fu['Dt Chegada Autron'], errors='coerce')
if 'PV informado na SC' in fu.columns:
    fu['PV informado na SC'] = pd.to_numeric(fu['PV informado na SC'], errors='coerce')
fu['Numero do PV'] = fu['Numero do PV'].astype(str).str.replace('.', '', regex=False).str.strip()
if 'OP na SC' not in fu.columns:
    fu['OP na SC'] = np.nan

# --- Estoque (mata010) ---
col_saldo = [c for c in mt.columns if 'saldo' in c.lower() and 'atual' in c.lower()]
if col_saldo:
    mt.rename(columns={col_saldo[0]: 'Saldo Atual'}, inplace=True)
col_codigo = [c for c in mt.columns if c.lower().strip() in ('codigo', 'código', 'cod', 'cod.')]
if col_codigo:
    mt.rename(columns={col_codigo[0]: 'Codigo'}, inplace=True)
if 'Codigo' not in mt.columns:
    alt = [c for c in mt.columns if 'codig' in c.lower() or 'cod' in c.lower()]
    if alt:
        mt.rename(columns={alt[0]: 'Codigo'}, inplace=True)
    else:
        mt.rename(columns={mt.columns[0]: 'Codigo'}, inplace=True)
if 'Saldo Atual' not in mt.columns:
    alt = [c for c in mt.columns if 'saldo' in c.lower()]
    if alt:
        mt.rename(columns={alt[0]: 'Saldo Atual'}, inplace=True)
    else:
        mt['Saldo Atual'] = 0
mt['Saldo Atual'] = pd.to_numeric(mt['Saldo Atual'], errors='coerce').fillna(0)

# --- Faturamento ---
fat['Emissao'] = pd.to_datetime(fat['Emissao'], errors='coerce')
col_vlr_total_fat = [c for c in fat.columns if 'vlr' in c.lower() and 'total' in c.lower()]
if col_vlr_total_fat:
    fat.rename(columns={col_vlr_total_fat[0]: 'Vlr.Total.Fat'}, inplace=True)
else:
    fat['Vlr.Total.Fat'] = 0
fat['Vlr.Total.Fat'] = pd.to_numeric(fat['Vlr.Total.Fat'], errors='coerce')
fat['Faturamento Bruto'] = pd.to_numeric(fat.get('Faturamento Bruto', 0), errors='coerce')
fat['Faturamento Liquido'] = pd.to_numeric(fat.get('Faturamento Liquido', 0), errors='coerce')
fat['No do Pedido'] = fat['No do Pedido'].astype(str).str.replace('.', '', regex=False).str.strip()

# --- sciozvs0 (classificacao Comprando/Produzindo) ---
sc_csv['Produto'] = sc_csv['Produto'].astype(str).str.strip()
sc_csv['Prod.Solict'] = sc_csv['Prod.Solict'].astype(str).str.strip()
tipo_prod = sc_csv.groupby('Produto')['Prod.Solict'].agg(
    lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'Indefinido'
).reset_index()
tipo_prod.rename(columns={'Prod.Solict': 'Tipo_Produto'}, inplace=True)


# ========================================================================
# 3. FOLLOW-UP: MERGE POR SC E POR PV+PRODUTO
# ========================================================================
print("\n[3/7] Cruzando follow-up...")

# Por SC
fu_por_sc = fu.groupby('No. da S.C.').agg({
    'Dt. Confirma': 'first',
    'Dt. Pre entr': 'first',
    'Pasta': 'first',
    'No. do P.O.': 'first',
    'Dt Chegada Autron': 'first',
    'OP na SC': 'first',
}).reset_index()
fu_por_sc.rename(columns={'No. da S.C.': 'Numero SC'}, inplace=True)

# Por PV + Produto
fu_pv = fu.copy()
fu_pv['Numero do PV'] = fu_pv['Numero do PV'].str.strip()
fu_por_pv = fu_pv.groupby(['Numero do PV', 'Codigo Item']).agg({
    'Dt. Confirma': 'first',
    'Dt. Pre entr': 'first',
    'Pasta': 'first',
    'No. do P.O.': 'first',
    'Dt Chegada Autron': 'first',
    'OP na SC': 'first',
}).reset_index()
fu_por_pv.rename(columns={
    'Numero do PV': 'Num. Pedido',
    'Codigo Item': 'Produto',
    'Dt. Confirma': 'FU_Dt_Confirma_PV',
    'Dt. Pre entr': 'FU_Dt_Pre_entr_PV',
    'Pasta': 'FU_Pasta_PV',
    'No. do P.O.': 'FU_PO_PV',
    'Dt Chegada Autron': 'FU_Dt_Chegada_PV',
    'OP na SC': 'FU_OP_na_SC_PV',
}, inplace=True)

# Merges
merged = ep.merge(fu_por_sc, on='Numero SC', how='left', suffixes=('', '_fu'))
merged = merged.merge(fu_por_pv, on=['Num. Pedido', 'Produto'], how='left')

# Consolidar campos do follow-up (SC tem prioridade, depois PV)
merged['FU_Dt_Confirma'] = merged['Dt. Confirma'].combine_first(merged.get('FU_Dt_Confirma_PV'))
merged['FU_Dt_Pre_Entr'] = merged['Dt. Pre entr'].combine_first(merged.get('FU_Dt_Pre_entr_PV'))
merged['FU_Pasta'] = merged['Pasta'].combine_first(merged.get('FU_Pasta_PV'))
merged['FU_PO'] = merged['No. do P.O.'].combine_first(merged.get('FU_PO_PV'))
merged['FU_Dt_Chegada_Autron'] = merged['Dt Chegada Autron'].combine_first(merged.get('FU_Dt_Chegada_PV'))
merged['FU_OP_na_SC'] = merged['OP na SC'].combine_first(merged.get('FU_OP_na_SC_PV'))

# Prazo real: Dt.Confirma; se vazio, Dt.Pre entr
merged['Prazo_Real_Entrega'] = merged['FU_Dt_Confirma'].combine_first(merged['FU_Dt_Pre_Entr'])

# Semana de entrega: so quando Dt.Confirma esta preenchida
merged['Semana_Entrega'] = np.where(
    merged['FU_Dt_Confirma'].notna(), merged['FU_Pasta'], None
)

# Tipo produto (Comprando / Produzindo)
merged = merged.merge(tipo_prod, on='Produto', how='left')
merged['Tipo_Produto'] = merged['Tipo_Produto'].fillna('Nao classificado')


# ========================================================================
# 4. STATUS E ESTOQUE
# ========================================================================
print("\n[4/7] Calculando status e estoque...")

# Status: Nota Fiscal preenchida = FINALIZADO
merged['Status_Pedido'] = np.where(merged['Nota Fiscal'].notna(), 'FINALIZADO', 'EM ABERTO')

# Estoque
stock = mt.groupby('Codigo')['Saldo Atual'].sum().reset_index()
stock.rename(columns={'Codigo': 'Produto', 'Saldo Atual': 'Estoque_Disponivel'}, inplace=True)
merged = merged.merge(stock, on='Produto', how='left')
merged['Estoque_Disponivel'] = merged['Estoque_Disponivel'].fillna(0)

# Alocacao de estoque por prioridade (data emissao mais antiga primeiro)
# Apenas para pedidos EM ABERTO
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
            'Disponivel_Estoque': 'SIM' if alloc >= qty else ('PARCIAL' if alloc > 0 else 'NAO'),
        }
        available = max(0, available - qty)

alloc_df = pd.DataFrame.from_dict(allocation, orient='index')
merged = merged.join(alloc_df, how='left')
merged.loc[merged['Status_Pedido'] == 'FINALIZADO', 'Disponivel_Estoque'] = 'N/A'
merged.loc[merged['Status_Pedido'] == 'FINALIZADO', 'Qtd_Alocada'] = None


# ========================================================================
# 5. REGRAS SC/OP E ALERTAS
# ========================================================================
print("\n[5/7] Aplicando regras SC/OP...")

merged['Tem_SC'] = merged['Numero SC'].notna()
merged['Tem_OP'] = (
    (merged['Numero OP'].notna()
     & merged['Numero OP'].astype(str).str.strip().ne('')
     & merged['Numero OP'].astype(str).str.strip().ne('nan'))
    | (merged['FU_OP_na_SC'].notna()
       & merged['FU_OP_na_SC'].astype(str).str.strip().ne('')
       & merged['FU_OP_na_SC'].astype(str).str.strip().ne('nan'))
)


def gerar_acao(row):
    if row['Status_Pedido'] == 'FINALIZADO':
        return 'Finalizado'
    if row['Disponivel_Estoque'] == 'SIM':
        return 'Estoque OK'
    tipo = str(row.get('Tipo_Produto', '')).strip()
    if tipo == 'Comprando':
        if row['Tem_OP']:
            return 'ERRO no CADASTRO'
        elif row['Tem_SC']:
            return 'SC gerada - Aguardando'
        else:
            return 'Necessario gerar SC'
    elif tipo == 'Produzindo':
        if row['Tem_OP']:
            return 'OP gerada - Aguardando'
        else:
            return 'Necessario gerar OP'
    else:
        return 'Verificar classificacao'


merged['Acao_Necessaria'] = merged.apply(gerar_acao, axis=1)


# ========================================================================
# 6. INDICADORES CALCULADOS
# ========================================================================
print("\n[6/7] Calculando indicadores...")

# Dias de atraso vs cliente
merged['Dias_Atraso_Cliente'] = np.where(
    (merged['Status_Pedido'] == 'EM ABERTO') & (merged['DT. Fat. Cli'].notna()),
    (TODAY - merged['DT. Fat. Cli']).dt.days, None
)
merged['Dias_Atraso_Cliente'] = pd.to_numeric(merged['Dias_Atraso_Cliente'], errors='coerce')

# Dias de atraso vs ofertada
merged['Dias_Atraso_Ofertada'] = np.where(
    (merged['Status_Pedido'] == 'EM ABERTO') & (merged['DT. Ofertada'].notna()),
    (TODAY - merged['DT. Ofertada']).dt.days, None
)
merged['Dias_Atraso_Ofertada'] = pd.to_numeric(merged['Dias_Atraso_Ofertada'], errors='coerce')

# Periodos
merged['Mes_Emissao'] = merged['DT Emissao'].dt.to_period('M').astype(str)
merged['Ano_Emissao'] = merged['DT Emissao'].dt.year

# Pronto para fazer?
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

# Mes previsto de faturamento (baseado em Dt Chegada Autron)
merged['Mes_Prev_Faturamento'] = merged['FU_Dt_Chegada_Autron'].dt.to_period('M').astype(str)
merged.loc[merged['FU_Dt_Chegada_Autron'].isna(), 'Mes_Prev_Faturamento'] = 'Sem previsao'


# ========================================================================
# 7. GERAR PLANILHA
# ========================================================================
print("\n[7/7] Gerando planilha dashboard_base.xlsx...")

# --- Aba 1: Base Consolidada (principal) ---
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
cols_base = [c for c in cols_base if c in merged.columns]
df_base = merged[cols_base].copy()

# --- Aba 2: Entrada de Pedidos mes a mes ---
entrada_mes = merged.groupby('Mes_Emissao').agg(
    Qtd_Linhas=('Vlr.Total', 'count'),
    Qtd_PVs=('Num. Pedido', 'nunique'),
    Valor_Total=('Vlr.Total', 'sum'),
).reset_index().sort_values('Mes_Emissao')
entrada_mes.rename(columns={'Mes_Emissao': 'Mes'}, inplace=True)

# --- Aba 3: Faturamento mes a mes (NFs emitidas) ---
fat['Mes_Faturamento'] = fat['Emissao'].dt.to_period('M').astype(str)
fat_mes = fat.groupby('Mes_Faturamento').agg(
    Qtd_NFs=('Num. Docto.', 'nunique'),
    Qtd_Linhas=('Vlr.Total.Fat', 'count'),
    Valor_Total=('Vlr.Total.Fat', 'sum'),
    Faturamento_Bruto=('Faturamento Bruto', 'sum'),
    Faturamento_Liquido=('Faturamento Liquido', 'sum'),
).reset_index().sort_values('Mes_Faturamento')
fat_mes.rename(columns={'Mes_Faturamento': 'Mes'}, inplace=True)

# --- Aba 4: Previsao de Faturamento mes a mes (Dt Chegada Autron) ---
abertos = merged[merged['Status_Pedido'] == 'EM ABERTO'].copy()
prev_fat = abertos.groupby('Mes_Prev_Faturamento').agg(
    Qtd_Linhas=('Vlr.Total', 'count'),
    Qtd_PVs=('Num. Pedido', 'nunique'),
    Valor_Previsto=('Vlr.Total', 'sum'),
).reset_index().sort_values('Mes_Prev_Faturamento')
prev_fat.rename(columns={'Mes_Prev_Faturamento': 'Mes_Previsto'}, inplace=True)

# --- Aba 5: Faturado vs A Faturar (mes atual) ---
mes_atual_str = str(MES_ATUAL)

# Valor faturado no mes atual (do relatorio faturamento)
fat_mes_atual = fat[fat['Mes_Faturamento'] == mes_atual_str]
valor_faturado = fat_mes_atual['Vlr.Total.Fat'].sum()
valor_faturado_bruto = fat_mes_atual['Faturamento Bruto'].sum()
valor_faturado_liquido = fat_mes_atual['Faturamento Liquido'].sum()

# Valor a faturar no mes atual (pedidos em aberto com Dt Chegada Autron no mes atual)
a_faturar = abertos[abertos['Mes_Prev_Faturamento'] == mes_atual_str]
valor_a_faturar = a_faturar['Vlr.Total'].sum()
qtd_a_faturar = a_faturar['Num. Pedido'].nunique()

resumo_mes = pd.DataFrame([
    {'Indicador': 'Mes Referencia', 'Valor': mes_atual_str},
    {'Indicador': 'Valor Faturado (Total)', 'Valor': f"R$ {valor_faturado:,.2f}"},
    {'Indicador': 'Faturamento Bruto', 'Valor': f"R$ {valor_faturado_bruto:,.2f}"},
    {'Indicador': 'Faturamento Liquido', 'Valor': f"R$ {valor_faturado_liquido:,.2f}"},
    {'Indicador': 'Qtd NFs Emitidas', 'Valor': str(fat_mes_atual['Num. Docto.'].nunique() if 'Num. Docto.' in fat_mes_atual.columns else 0)},
    {'Indicador': 'Valor a Faturar (Previsao)', 'Valor': f"R$ {valor_a_faturar:,.2f}"},
    {'Indicador': 'Qtd PVs a Faturar', 'Valor': str(qtd_a_faturar)},
    {'Indicador': 'Total Previsto no Mes', 'Valor': f"R$ {valor_faturado + valor_a_faturar:,.2f}"},
])

# Detalhe do a faturar no mes atual
if len(a_faturar) > 0:
    cols_afat = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
                 'Ped Cliente', 'Vlr.Total', 'FU_Dt_Chegada_Autron',
                 'Pronto_para_Fazer', 'Acao_Necessaria']
    cols_afat = [c for c in cols_afat if c in a_faturar.columns]
    detalhe_a_faturar = a_faturar[cols_afat].copy()
else:
    detalhe_a_faturar = pd.DataFrame(columns=['Sem dados para o mes atual'])

# --- Aba 6: Alertas (ERRO no CADASTRO) ---
erros = merged[merged['Acao_Necessaria'] == 'ERRO no CADASTRO']
if len(erros) > 0:
    cols_erro = ['Num. Pedido', 'Item', 'Produto', 'Descricao do Produto',
                 'Tipo_Produto', 'Numero SC', 'Numero OP', 'FU_OP_na_SC',
                 'Acao_Necessaria']
    cols_erro = [c for c in cols_erro if c in erros.columns]
    df_erros = erros[cols_erro].copy()
else:
    df_erros = pd.DataFrame([{'Info': 'Nenhum erro de cadastro encontrado'}])


# --- Escrever Excel ---
output_path = os.path.join(BASE_DIR, "dashboard_base.xlsx")

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_base.to_excel(writer, sheet_name='Base Consolidada', index=False)
    entrada_mes.to_excel(writer, sheet_name='Entrada Pedidos Mes', index=False)
    fat_mes.to_excel(writer, sheet_name='Faturamento Mes', index=False)
    prev_fat.to_excel(writer, sheet_name='Previsao Faturamento', index=False)
    resumo_mes.to_excel(writer, sheet_name='Faturado vs A Faturar', index=False)
    detalhe_a_faturar.to_excel(writer, sheet_name='Detalhe A Faturar Mes', index=False)
    df_erros.to_excel(writer, sheet_name='Alertas Erro Cadastro', index=False)

print(f"\n{'=' * 60}")
print(f"PLANILHA GERADA: {output_path}")
print(f"{'=' * 60}")
print(f"\nAbas geradas:")
print(f"  1. Base Consolidada       - {len(df_base)} linhas (dados completos por linha de pedido)")
print(f"  2. Entrada Pedidos Mes    - {len(entrada_mes)} meses")
print(f"  3. Faturamento Mes        - {len(fat_mes)} meses (NFs emitidas)")
print(f"  4. Previsao Faturamento   - {len(prev_fat)} periodos (baseado Dt Chegada Autron)")
print(f"  5. Faturado vs A Faturar  - Resumo mes atual ({mes_atual_str})")
print(f"  6. Detalhe A Faturar Mes  - {len(detalhe_a_faturar)} linhas previstas no mes")
print(f"  7. Alertas Erro Cadastro  - {len(df_erros)} itens com alerta")
print(f"\nResumo mes atual ({mes_atual_str}):")
print(f"  Faturado:    R$ {valor_faturado:>14,.2f}")
print(f"  A Faturar:   R$ {valor_a_faturar:>14,.2f}")
print(f"  Total:       R$ {valor_faturado + valor_a_faturar:>14,.2f}")
