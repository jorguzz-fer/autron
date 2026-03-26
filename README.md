# Dashboard de Pedidos - Autron

Dashboard web para gestao de pedidos, estoque e follow-up.

## Funcionalidades
- Visao geral de pedidos (KPIs, graficos, filtros)
- Analise de prontidao (estoque + follow-up)
- Previsao de entrega com dias de atraso
- Analise de estoque com alocacao por prioridade
- Regras SC/OP com alertas de erro de cadastro
- Upload de dados pelo navegador

## Deploy no Coolify
1. New Resource > Public Repository > `https://github.com/jorguzz-fer/autron.git`
2. Build Pack: **Dockerfile**
3. Porta: **8501**
4. Volume: `/data/dashboard-pedidos` -> `/app/dados`
5. Deploy!

Ver `DEPLOY_COOLIFY.md` para instrucoes detalhadas.

## Arquivos de Dados
Envie pelo navegador ao acessar o dashboard:
- `entrada_pedido.xlsx` - Pedidos de venda
- `followup.xlsx` - Follow-up de compras
- `mata010.xlsx` - Estoque
- `sciozmq0.csv` - Classificacao Comprando/Produzindo
