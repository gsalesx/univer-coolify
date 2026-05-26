# Univer Coolify

Deploy do Univer para Coolify com backend oficial versionado e uma UI demo baseada no `dream-num/univer-pro-sheet-start-kit`.

## Endpoints

- UI: `https://univer.guilhermesales.com`
- API direta: `https://api.univer.guilhermesales.com`

A UI também encaminha `/universer-api/*` internamente para o backend, então o cliente pode usar `window.location.origin`.

## Estrutura

- `docker-compose.yml`: stack completa para Coolify.
- `backend/`: configs oficiais necessárias para Universer, Temporal, Envoy, Nginx e Worker.
- `demo-ui/`: app Vite/React do Univer Pro Sheet Start Kit, servido por Nginx.

## Observações

- As senhas padrão são apenas para ambiente isolado/demo. Para produção, configure variáveis no Coolify.
- O compose usa infra interna do pacote oficial: Postgres, Redis, RabbitMQ, MinIO e Temporal.
- Não depende de `curl https://get.univer.ai/product | bash` em runtime.
