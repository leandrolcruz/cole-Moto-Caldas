# Inventário Q900 - Moto Caldas Ltda

Sistema de inventário para coletor de código de barras NETUM Q900, desenvolvido para a **Moto Caldas Ltda** (Caldas Novas/GO).

Integra com o **Microwork Cloud DMS**.

## Módulos

| Módulo | Descrição |
|--------|-----------|
| **Contagem Livre** | Bipe códigos livremente. Cada bipe = +1 unidade. |
| **Contagem Diária** | Importe a planilha de saída do Microwork. Confira saldo por localização. Duas contagens com conferente. |
| **Contagem por Locação** | Importe planilha de localização. Conte por corredor, estante e prateleira. Duas contagens com conferente. |
| **Contagem de Roupas** | Bipe o código e registre cor e tamanho. Valores ficam salvos para reuso. Duas contagens com conferente. |

## Funcionalidades

- Catálogo de 134 mil peças via IndexedDB
- Validação de código (etiquetas com `-` são rejeitadas)
- Aviso quando peça não está no catálogo
- Sons diferenciados (confirmação, alerta, erro) e vibração
- Exportação em Excel (.xlsx) com compartilhamento direto (Google Drive)
- Duas contagens com conferente nomeado e coluna de Auditoria
- Backup e restauração em JSON
- Funciona 100% offline (PWA)

## Como instalar no NETUM Q900

1. Acesse `https://leandrolcruz.github.io/cole-Moto-Caldas/` no Chrome
2. Menu (⋮) → **Instalar app**
3. O app aparece na tela inicial

## Arquivos

| Arquivo | Função |
|---------|--------|
| `index.html` | App completo (HTML + CSS + JS) |
| `manifest.json` | Configuração PWA |
| `sw.js` | Service Worker para funcionamento offline |

## Desenvolvido por

Leandro Leite
