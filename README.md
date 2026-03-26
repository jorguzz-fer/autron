# Dashboard de Pedidos - Autron

Dashboard web para gestao de pedidos, estoque e follow-up.

## Deploy no Coolify

1. Conecte este repositorio no Coolify como **Dockerfile**
2. Porta: **8501**
3. Configure as variaveis de ambiente (ver abaixo)

## Variaveis de Ambiente

### Google Drive (opcional)
```
GDRIVE_ENABLED=true
GDRIVE_FOLDER_ID=id_da_pasta_no_drive
GDRIVE_CREDENTIALS={"type":"service_account",...}
```

### Sem Google Drive
Sem variaveis necessarias. Faca upload dos arquivos pela interface web.

## Arquivos de Dados Necessarios
- `entrada_pedido.xlsx`
- `followup.xlsx`
- `mata010.xlsx`
- `sciozmq0.csv`
