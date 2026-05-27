# Painel Google Ads · HVSB

Dashboard custom em HTML + Chart.js com identidade visual do Hospital Veterinário São Bernardo.
Lê dados de um JSON e exibe: saldo da conta, alertas, conversões, ranking de campanhas e tendência diária.

## Estrutura

```
hvsb-dashboard/
├── index.html              # Dashboard (HTML + JS + Chart.js)
├── data.json               # Dados que alimentam o dashboard
├── google-ads-script.js    # Script pra Google Ads (etapa 2)
└── README.md
```

## Como funciona hoje (estado atual)

O `index.html` faz `fetch('data.json')` e renderiza todos os componentes.
Atualização: a cada 5 minutos (refresh do client). Hoje os dados são manuais.

## Como automatizar (próximos passos)

Pra ter atualização horária com dados reais do Google Ads, seguir este fluxo:

### 1. Criar uma Google Sheet com 4 abas

| Aba | Conteúdo |
|---|---|
| `KPIs` | Saldo, conversões MTD, custo MTD, meta, etc. |
| `Campanhas` | Lista de campanhas com conv, CPA, custo, status |
| `Diario` | Tendência diária dos últimos 30 dias |
| `Tipos` | Conversões por tipo (WhatsApp, Ligação, Rota) |

### 2. Colar o Google Ads Script

1. Logar no [Google Ads](https://ads.google.com/) → **Ferramentas** → **Scripts** → **+ Novo script**
2. Colar o conteúdo de `google-ads-script.js`
3. Trocar `SPREADSHEET_URL` pela URL da Sheet criada no passo 1
4. **Salvar** e autorizar
5. **Configurar trigger:** Frequência → **De hora em hora**

### 3. Publicar a Sheet como JSON

Opção A — **Web App (recomendado):** criar um Google Apps Script Web App que retorna a planilha em JSON. Mais flexível.

Opção B — **Google Sheets API + chave pública:** publicar a Sheet e ler via gviz endpoint:
```
https://docs.google.com/spreadsheets/d/SHEET_ID/gviz/tq?tqx=out:json&sheet=KPIs
```

### 4. Apontar o dashboard pra Sheet

No `index.html`, trocar o fetch atual:
```js
const res = await fetch('data.json?t=' + Date.now());
```

Por algo como:
```js
const res = await fetch('URL_DO_WEB_APP_OU_GVIZ');
```

E ajustar o parsing pra montar o mesmo formato do `data.json` atual.

## Saldo da conta (caso especial)

Google Ads API não expõe o saldo de conta pré-paga de forma direta.
Alternativa: criar célula manual na Sheet (`B5` na aba `KPIs`) onde o gestor atualiza
o saldo após cada recarga. O script preserva esse valor; só os outros KPIs vêm da API.

## Frequência de atualização

- **Google Ads Script** roda 1x/hora
- **HTML** faz fetch a cada 5 minutos
- **Importante:** Google Ads tem 3-24h de delay em conversões offline (chamadas, importações).
  Por isso o painel mostra "última atualização" no header e tem o disclaimer
  "atualização horária" no footer — não é real-time.

## Personalização

- Cores: variáveis CSS no topo do `<style>` (`--gold`, `--navy`, `--green`, etc.)
- Meta do mês: trocar `MES_META_PERCENT` no Google Ads Script
- Limites de alerta: ajustar nas funções `render()` dentro do `<script>` do index.html
  (atualmente: ≤7 dias = amarelo, ≤3 dias = vermelho, ≤1 dia = crítico)

## Deploy

Push pro GitHub → Vercel deploya automático.
URL: `hvsb-dashboard.vercel.app` (configurar no painel da Vercel)
