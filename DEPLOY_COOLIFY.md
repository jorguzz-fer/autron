# Deploy no Coolify - Dashboard de Pedidos

## Passo a passo

### 1. No Coolify, criar novo recurso
- Clique em **New Resource**
- Selecione **Public Repository** (ou conecte seu GitHub)
- URL: `https://github.com/jorguzz-fer/autron.git`
- Branch: `main`
- Build Pack: **Dockerfile**

### 2. Configurar porta
- Container Port: **8501**
- A porta publica sera gerada automaticamente pelo Coolify

### 3. Configurar volume persistente (IMPORTANTE)
Para que os dados nao se percam ao reiniciar o container:
- Va em **Storages** (ou Volumes)
- Adicione um volume:
  - **Source**: `/data/dashboard-pedidos` (ou caminho que preferir no servidor)
  - **Destination**: `/app/dados`

### 4. Deploy
- Clique em **Deploy**
- Aguarde o build (primeira vez leva ~2min)
- Acesse pela URL gerada pelo Coolify

### 5. Primeiro acesso
- Ao abrir o dashboard, ele mostrara a tela de upload
- Envie os 4 arquivos:
  - entrada_pedido.xlsx
  - followup.xlsx
  - mata010.xlsx
  - sciozmq0.csv
- O dashboard sera gerado automaticamente

## Para atualizar os dados
- No dashboard, clique em **"Atualizar Arquivos"** na barra lateral
- Envie os 4 arquivos atualizados
- O dashboard reprocessa tudo automaticamente

## Dominio customizado (opcional)
No Coolify, va em **Settings** do servico:
- Adicione seu dominio (ex: dashboard.suaempresa.com)
- O Coolify configura SSL automaticamente

## Recursos recomendados
- CPU: 0.5 core
- RAM: 512MB (minimo) / 1GB (recomendado)
