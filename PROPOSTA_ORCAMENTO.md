# PROPOSTA COMERCIAL

---

**De:** Tudo Mudou!
**Para:** Autron
**Data:** 27/04/2026
**Ref:** Dashboard de Gestao de Pedidos e Faturamento

---

## 1. OBJETIVO

Desenvolvimento de plataforma web de Business Intelligence para consolidacao e analise de dados operacionais da Autron, integrando informacoes de multiplos relatorios do ERP Protheus, CRM Plumes e planilhas de Budget/Metas em um dashboard interativo e centralizado.

---

## 2. ESCOPO DO PROJETO

### 2.1 Plataforma Web

| Item | Descricao |
|------|-----------|
| Autenticacao | Sistema de login com controle de acesso por usuario |
| Upload de dados | Interface para envio e atualizacao dos relatorios diretamente pelo navegador |
| Carga inteligente | Leitura automatica de 7 fontes de dados com deteccao de estrutura (multi-aba, multi-header) |
| Atualizacao em tempo real | Possibilidade de atualizar dados via upload ou via repositorio Git |
| Download consolidado | Exportacao de planilha consolidada com 7 abas prontas para Power BI |
| Identidade visual | Interface customizada com logo oficial da Autron |
| Formatacao regional | Valores monetarios, datas e numeros no padrao brasileiro |

### 2.2 Modulos / Visoes do Dashboard

#### Modulo 1 - Visao Geral
- KPIs: Total PVs, Linhas, Em Aberto, Finalizados, Valor, % Conclusao
- Grafico de entrada de pedidos por mes
- Distribuicao por status (Finalizado / Em Aberto)
- Pedidos por vendedor
- Filtro de data interativo

#### Modulo 2 - Entrada de Pedidos
- Visao mensal de entrada de pedidos (DT Emissao)
- Comparativo Meta x Realizado - Grupo Autron
- Comparativo Meta x Realizado - Autron (sem Ergomec)
- Comparativo Meta x Realizado - Ergomec (IND21)
- Graficos e tabelas interativas

#### Modulo 3 - Disponibilidade (Prontidao)
- Analise de disponibilidade de material por pedido
- Quadro de datas por PV: Necessidade do Cliente, Entrega Adaptada, Prazo Prometido
- Visao de atendimento de prazo do cliente (Dentro / Fora do prazo com percentuais)
- Identificacao de SC e OP por linha de pedido (match direto + match por produto)
- Filtros interativos: Disponibilidade, Tipo Produto, Atendimento Prazo, Mes de Faturamento
- Tabela detalhada com ordenacao funcional por valor e quantidade

#### Modulo 4 - Previsao de Entrega
- KPIs de atraso (quantidade, media, maximo)
- Pedidos por semana de entrega
- Top 10 pedidos mais atrasados
- Tabela de previsao com datas e status

#### Modulo 5 - Estoque e SC/OP
- Disponibilidade de estoque com alocacao por prioridade (data de emissao)
- Regras de negocio: Comprando precisa SC, Produzindo precisa OP
- Alerta automatico "ERRO no CADASTRO" (item Comprando com OP)
- Graficos de distribuicao e acoes necessarias

#### Modulo 6 - Faturamento
- Resultado do mes anterior (Meta, Fat. Liquido, Atingimento, Diferenca)
- Resultado do mes atual com Total Liquido Projetado
- Consolidado trimestral (Q1, Q2...)
- Total de Carteira (valor bruto e liquido)
- Quadro Meta x Realizado no formato de planilha (12 meses)
- Grafico Faturamento: Meta x Realizado x Previsao (estilo Excel)
- Comparativo Faturamento Liquido 2025 x 2026 (YoY)
- Detalhe de pedidos a faturar no mes

#### Modulo 7 - Previsao de Faturamento
- Previsao baseada em Dt Chegada Autron (apenas meses futuros)
- Valores em R$ nos graficos
- Tabela detalhada de previsao

#### Modulo 8 - Comparativo Plumes x Protheus
- Integracao com dados do CRM (Ganhas.xlsx)
- Match automatico por Pedido de Compra do Cliente
- Identificacao de divergencias: valor diferente, mes diferente, pedido ausente
- KPIs, graficos de distribuicao e tabela detalhada com filtros

### 2.3 Integracao de Dados

| Fonte | Arquivo | Dados |
|-------|---------|-------|
| Protheus | entrada_pedido.xlsx | Pedidos de venda, produtos, datas, valores |
| Protheus | followup.xlsx | Acompanhamento de compras, datas de entrega |
| Protheus | mata010.xlsx | Cadastro de produtos e saldo de estoque |
| Protheus | faturamento.xlsx | Notas fiscais de venda e devolucao (multi-aba) |
| Protheus | sciozvs0.csv | Classificacao Comprando/Produzindo |
| Budget | Metas_Budget_Grupo_Autron_2026.xlsx | Metas Autron, Ergomec e Grupo |
| CRM Plumes | Ganhas.xlsx | Pedidos ganhos no CRM |

### 2.4 Regras de Negocio Implementadas

- Pedido finalizado quando campo Nota Fiscal preenchido
- Prioridade de estoque por data de emissao do PV (mais antigo primeiro)
- Prazo real = Dt. Confirma; fallback Dt. Pre Entr
- Total a Faturar baseado no campo Entrega
- Total Liquido a Faturar = Total a Faturar - 17% (deducoes)
- Meta = Receita Liquida do Grupo (nao Autron isolado)
- Ergomec = tudo que e IND21 nos relatorios
- Faturamento unificado: vendas + devolucoes com deduplicacao

### 2.5 Infraestrutura e Deploy

- Containerizacao com Docker
- Deploy automatizado via Git (push = deploy)
- Configuracao de servidor (Coolify)
- Integracao com repositorio privado GitHub
- Script de exportacao independente (gerar_base.py)

---

## 3. ENTREGAVEIS

1. Plataforma web funcional com 8 modulos de analise
2. Codigo-fonte completo no repositorio GitHub
3. Dockerfile para deploy em qualquer servidor
4. Script de exportacao de planilha consolidada
5. Documentacao de deploy (DEPLOY_COOLIFY.md)

---

## 4. INVESTIMENTO

| Item | Descricao | Valor |
|------|-----------|-------|
| Desenvolvimento da plataforma | 8 modulos de dashboard com regras de negocio | R$ ________ |
| Integracao de dados | 7 fontes de dados com carga inteligente | R$ ________ |
| Deploy e infraestrutura | Docker, Coolify, integracao GitHub | R$ ________ |
| Ajustes e validacao com usuario | Iteracoes de refinamento em producao | R$ ________ |
| **TOTAL** | | **R$ ________** |

*Valores a serem preenchidos conforme negociacao.*

---

## 5. PRAZO

Projeto entregue e em producao.

---

## 6. CONDICOES

- Eventuais novas funcionalidades ou modulos serao cotados separadamente
- Suporte tecnico e manutencao evolutiva podem ser contratados em separado
- Os dados sao de propriedade exclusiva da Autron
- O codigo-fonte e entregue junto com o projeto

---

## 7. VALIDADE

Esta proposta tem validade de 30 dias a partir da data de emissao.

---

**Tudo Mudou!**
Desenvolvimento de Solucoes em Tecnologia

---
