# Usiminas Taubaté · ponte rolante · apresentação técnica

Apresentação animada do sistema de posicionamento de duas pontes rolantes
(Symeo 291, OndoSense D300, encoder T58, remotas, rede sem fio, mesa de comando e PLC).

Página estática única (`index.html`), com o modelo 3D em three.js carregado do jsDelivr.

## Deploy no Coolify (recurso separado do dashboard)

1. **New Resource** › Public Repository › `https://github.com/jorguzz-fer/autron.git`
2. Branch: `main`
3. Build Pack: **Dockerfile**
4. **Base Directory:** `/usiminas-taubate`
5. Porta do container: **80**
6. Deploy e, se quiser, adicione um domínio em Settings.

Navegação: setas ou espaço · F para tela cheia · cada seção tem link próprio (ex.: `#ponte`, `#itens`).
