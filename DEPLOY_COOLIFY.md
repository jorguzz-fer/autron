# Deploy no Coolify - Dashboard de Pedidos

## Opcao 1: Google Drive (Automatico)

### Passo 1: Criar Service Account no Google Cloud
1. Acesse https://console.cloud.google.com
2. Crie um projeto (ou use um existente)
3. Ative a **Google Drive API**
4. Va em **IAM & Admin > Service Accounts**
5. Crie uma service account
6. Gere uma chave JSON e baixe o arquivo

### Passo 2: Compartilhar pasta do Google Drive
1. No Google Drive, crie uma pasta (ex: "Dashboard Pedidos")
2. Compartilhe a pasta com o email da service account
   (ex: dashboard@seu-projeto.iam.gserviceaccount.com)
3. Copie o ID da pasta (esta na URL: drive.google.com/drive/folders/**ID_AQUI**)

### Passo 3: Colocar os 4 arquivos na pasta
- entrada_pedido.xlsx
- followup.xlsx
- mata010.xlsx
- sciozmq0.csv

### Passo 4: Deploy no Coolify
1. No Coolify, crie um novo servico **Dockerfile**
2. Aponte para o repositorio Git com estes arquivos
   (ou faca upload da pasta dashboard-web)
3. Configure as variaveis de ambiente:

```
GDRIVE_ENABLED=true
GDRIVE_FOLDER_ID=seu_id_da_pasta_aqui
GDRIVE_CREDENTIALS={"type":"service_account","project_id":"...todo o JSON da chave..."}
```

4. Porta: **8501**
5. Deploy!

### Para atualizar dados:
- Basta substituir os arquivos na pasta do Google Drive
- O dashboard recarrega automaticamente a cada 5 minutos
- Ou clique no botao "Atualizar Dados" na sidebar


## Opcao 2: Upload Manual (Sem Google Drive)

### Passo 1: Deploy no Coolify
1. No Coolify, crie um novo servico **Dockerfile**
2. Aponte para o repositorio ou faca upload
3. NAO configure variaveis de Google Drive
4. Porta: **8501**
5. Deploy!

### Para atualizar dados:
- Acesse o dashboard pelo navegador
- Na primeira vez, ele mostra a tela de upload
- Envie os 4 arquivos pelo navegador
- Os dados ficam salvos no container

**IMPORTANTE:** Se o container for recriado, os arquivos precisam ser enviados novamente.
Para persistir, configure um volume no Coolify:
- Volume: `/app/dados`
- Monte em um diretorio persistente do servidor


## Opcao 3: Volume Montado (Avanado)

### Passo 1: No servidor VPS, crie a pasta dos dados
```bash
mkdir -p /opt/dashboard-pedidos/dados
```

### Passo 2: Copie os arquivos para o servidor
```bash
scp entrada_pedido.xlsx followup.xlsx mata010.xlsx sciozmq0.csv \
    usuario@seu-servidor:/opt/dashboard-pedidos/dados/
```

### Passo 3: No Coolify, configure o volume
- Source: `/opt/dashboard-pedidos/dados`
- Destination: `/app/dados`

### Para atualizar:
- Substitua os arquivos no servidor via SCP/SFTP
- Clique "Atualizar Dados" no dashboard


## Configuracao Coolify

### Porta
- Container: 8501
- Healthcheck: GET /_stcore/health

### Variaveis de Ambiente (Google Drive)
```
GDRIVE_ENABLED=true
GDRIVE_FOLDER_ID=id_da_pasta
GDRIVE_CREDENTIALS=json_da_chave
```

### Variaveis de Ambiente (Sem Google Drive)
Nenhuma variavel necessaria.

### Recursos Recomendados
- CPU: 0.5 core
- RAM: 512MB (minimo) / 1GB (recomendado)
