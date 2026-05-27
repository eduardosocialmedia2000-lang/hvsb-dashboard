/**
 * GOOGLE ADS SCRIPT · Painel HVSB
 * Conta: 192-048-4711 · Hospital Veterinário São Bernardo · Curitiba/PR
 *
 * COMO INSTALAR:
 * 1. Google Ads → Ferramentas e configurações → Scripts → Novo script
 * 2. Cole este código completo
 * 3. Preencha SPREADSHEET_ID abaixo com o ID da sua Google Planilha
 * 4. Salve → Autorizar → Executar uma vez para testar
 * 5. Configure agendamento: Triggers → Diariamente (horário fixo, ex: 04:00)
 *
 * ABAS que este script cria/atualiza na planilha:
 *   diario    — data | conv | custo | cliques (últimos 90 dias)
 *   campanhas — nome | tipo | conv | cpa | custo | is | isLostBudget | status
 *   kpis      — chave | valor (KPIs do mês atual + anterior)
 *   meta      — chave | valor (lastUpdate, account, accountName, etc.)
 */

// =============================================================
// CONFIGURACAO — altere apenas este bloco
// =============================================================

// ID da Google Planilha (tudo entre /d/ e /edit na URL)
const SPREADSHEET_ID = '1ao1S84eBniYIvzdgNFQZ6poCIXFJkdtUZnlUIa6L9NQ';

// Meta de conversoes: porcentagem sobre o mês anterior
const META_PERCENT = 1.10; // +10%

// Cidade exibida no dashboard
const CITY = 'Curitiba/PR';

// =============================================================


function main() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoje = new Date();

  // ---------- Intervalos de datas ----------
  const inicioMes    = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const inicioMesAnt = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  const fimMesAnt    = new Date(hoje.getFullYear(), hoje.getMonth(), 0);
  const noventaDias  = diasAtras(hoje, 90);
  const seteDias     = diasAtras(hoje, 7);

  // ---------- Queries ----------
  const mtd         = getStats(inicioMes, hoje);
  const mesAnterior = getStats(inicioMesAnt, fimMesAnt);
  const ultimos7d   = getStats(seteDias, hoje);
  const gastoMedio7d = ultimos7d.cost / 7;

  const searchIS    = getSearchIS(inicioMes, hoje);
  const campanhas   = getCampanhas(inicioMes, hoje);
  const diario      = getDiario(noventaDias, hoje);
  const tiposConv   = getConversoesPorTipo(inicioMes, hoje);

  const account = AdsApp.currentAccount();
  const now     = new Date().toISOString();
  const meta    = Math.round(mesAnterior.conversions * META_PERCENT);

  // ---------- Aba: meta ----------
  const sMeta = getOrCreate(ss, 'meta');
  sMeta.clear();
  sMeta.appendRow(['chave', 'valor']);
  [
    ['lastUpdate',   now],
    ['currency',     'BRL'],
    ['account',      account.getCustomerId()],
    ['accountName',  account.getName()],
    ['city',         CITY],
    ['source',       'google-ads-script'],
  ].forEach(r => sMeta.appendRow(r));

  // ---------- Aba: kpis ----------
  const sKpis = getOrCreate(ss, 'kpis');
  sKpis.clear();
  sKpis.appendRow(['chave', 'valor']);
  [
    ['conversoesMTD',         round2(mtd.conversions)],
    ['conversoesMesAnterior', round2(mesAnterior.conversions)],
    ['custoMTD',              round2(mtd.cost)],
    ['custoMesAnterior',      round2(mesAnterior.cost)],
    ['cpaMTD',                mtd.conversions ? round2(mtd.cost / mtd.conversions) : 0],
    ['cpaMesAnterior',        mesAnterior.conversions ? round2(mesAnterior.cost / mesAnterior.conversions) : 0],
    ['cliquesMTD',            mtd.clicks],
    ['cliquesMesAnterior',    mesAnterior.clicks],
    ['meta',                  meta],
    ['metaOrigem',            'Mes anterior (' + Math.round(mesAnterior.conversions) + ' conv) + ' + Math.round((META_PERCENT - 1) * 100) + '%'],
    ['gastoMedio7d',          round2(gastoMedio7d)],
    ['searchIS',              searchIS],
    // Saldo: atualize manualmente na planilha na célula B14 (linha 14, coluna valor)
    // O script não sobrescreve saldoAtual nem ultimaRecarga — isso fica manual
  ].forEach(r => sKpis.appendRow(r));

  // ---------- Aba: campanhas ----------
  const sCamp = getOrCreate(ss, 'campanhas');
  sCamp.clear();
  sCamp.appendRow(['nome', 'tipo', 'conv', 'cpa', 'custo', 'is', 'isLostBudget', 'status']);
  campanhas.forEach(c => sCamp.appendRow([
    c.nome, c.tipo, c.conv, c.cpa, c.custo, c.is, c.isLostBudget, c.status
  ]));

  // ---------- Aba: diario ----------
  const sDiario = getOrCreate(ss, 'diario');
  sDiario.clear();
  sDiario.appendRow(['data', 'conv', 'custo', 'cliques']);
  diario.forEach(d => sDiario.appendRow([d.data, d.conv, d.custo, d.cliques]));

  // ---------- Aba: tipos ----------
  const sTipos = getOrCreate(ss, 'tipos');
  sTipos.clear();
  sTipos.appendRow(['tipo', 'total']);
  Object.keys(tiposConv).forEach(t => sTipos.appendRow([t, tiposConv[t]]));

  Logger.log('OK · Script HVSB concluido: ' + now);
}


// =============================================================
// HELPERS DE QUERY
// =============================================================

function getStats(dateFrom, dateTo) {
  const start = fmt(dateFrom);
  const end   = fmt(dateTo);
  const query = `
    SELECT metrics.clicks, metrics.conversions, metrics.cost_micros
    FROM customer
    WHERE segments.date BETWEEN '${start}' AND '${end}'
  `;
  const rows = AdsApp.search(query);
  let clicks = 0, conversions = 0, costMicros = 0;
  while (rows.hasNext()) {
    const r = rows.next();
    clicks       += parseInt(r.metrics.clicks) || 0;
    conversions  += parseFloat(r.metrics.conversions) || 0;
    costMicros   += parseFloat(r.metrics.costMicros) || 0;
  }
  return {
    clicks,
    conversions,
    cost: costMicros / 1e6,
  };
}

function getSearchIS(dateFrom, dateTo) {
  const start = fmt(dateFrom);
  const end   = fmt(dateTo);
  // IS de busca no nível da conta (media ponderada por impressoes)
  const query = `
    SELECT metrics.search_impression_share,
           metrics.search_budget_lost_impression_share,
           metrics.impressions
    FROM campaign
    WHERE segments.date BETWEEN '${start}' AND '${end}'
      AND campaign.advertising_channel_type = 'SEARCH'
      AND campaign.status = 'ENABLED'
  `;
  const rows = AdsApp.search(query);
  let totalImp = 0, totalIS = 0;
  while (rows.hasNext()) {
    const r = rows.next();
    const imp = parseFloat(r.metrics.impressions) || 0;
    const is  = parseFloat(r.metrics.searchImpressionShare) || 0;
    if (is > 0 && imp > 0) {
      totalImp += imp;
      totalIS  += is * imp;
    }
  }
  return totalImp > 0 ? Math.round((totalIS / totalImp) * 100) : 0;
}

function getCampanhas(dateFrom, dateTo) {
  const start = fmt(dateFrom);
  const end   = fmt(dateTo);
  const query = `
    SELECT campaign.name,
           campaign.advertising_channel_type,
           campaign.status,
           metrics.conversions,
           metrics.cost_micros,
           metrics.clicks,
           metrics.search_impression_share,
           metrics.search_budget_lost_impression_share
    FROM campaign
    WHERE segments.date BETWEEN '${start}' AND '${end}'
      AND campaign.status != 'REMOVED'
    ORDER BY metrics.cost_micros DESC
  `;
  const rows = AdsApp.search(query);
  const out  = [];
  while (rows.hasNext()) {
    const r    = rows.next();
    const conv = parseFloat(r.metrics.conversions) || 0;
    const cost = (parseFloat(r.metrics.costMicros) || 0) / 1e6;
    const isLostBudget = parseFloat(r.metrics.searchBudgetLostImpressionShare) || 0;
    const isValue      = parseFloat(r.metrics.searchImpressionShare) || 0;
    const status       = r.campaign.status;

    let statusOut = 'ok';
    if (status === 'PAUSED')          statusOut = 'paused';
    else if (isLostBudget > 0.05)    statusOut = 'limited';
    else if (conv === 0 && cost > 0) statusOut = 'no-conv';

    let tipo = 'Search';
    const ch = r.campaign.advertisingChannelType;
    if (ch === 'PERFORMANCE_MAX') tipo = 'PMax';
    if (ch === 'DISPLAY')         tipo = 'Display';
    if (ch === 'DEMAND_GEN')      tipo = 'DemandGen';
    if (ch === 'LOCAL_SERVICES')  tipo = 'LSA';

    out.push({
      nome:        r.campaign.name,
      tipo,
      conv:        round2(conv),
      cpa:         conv ? round2(cost / conv) : 0,
      custo:       round2(cost),
      is:          Math.round(isValue * 100),
      isLostBudget: Math.round(isLostBudget * 100),
      status:      statusOut,
    });
  }
  return out;
}

function getDiario(dateFrom, dateTo) {
  const start = fmt(dateFrom);
  const end   = fmt(dateTo);
  const query = `
    SELECT segments.date,
           metrics.conversions,
           metrics.cost_micros,
           metrics.clicks
    FROM customer
    WHERE segments.date BETWEEN '${start}' AND '${end}'
    ORDER BY segments.date ASC
  `;
  const rows = AdsApp.search(query);
  const out  = [];
  while (rows.hasNext()) {
    const r = rows.next();
    out.push({
      data:    r.segments.date,
      conv:    Math.round(parseFloat(r.metrics.conversions) || 0),
      custo:   round2((parseFloat(r.metrics.costMicros) || 0) / 1e6),
      cliques: parseInt(r.metrics.clicks) || 0,
    });
  }
  return out;
}

function getConversoesPorTipo(dateFrom, dateTo) {
  const start = fmt(dateFrom);
  const end   = fmt(dateTo);
  // Segmenta por nome da acao de conversao
  // Configure os nomes no Google Ads para conter: whats, liga/call/phone, rota/direc/maps
  const query = `
    SELECT segments.conversion_action_name, metrics.conversions
    FROM customer
    WHERE segments.date BETWEEN '${start}' AND '${end}'
  `;
  const rows = AdsApp.search(query);
  const out  = { whatsapp: 0, ligacao: 0, rota: 0 };
  while (rows.hasNext()) {
    const r    = rows.next();
    const nome = (r.segments.conversionActionName || '').toLowerCase();
    const conv = parseFloat(r.metrics.conversions) || 0;
    if      (nome.includes('whats'))                                         out.whatsapp += conv;
    else if (nome.includes('liga') || nome.includes('call') || nome.includes('phone')) out.ligacao  += conv;
    else if (nome.includes('rota') || nome.includes('direc') || nome.includes('maps'))  out.rota     += conv;
  }
  out.whatsapp = Math.round(out.whatsapp);
  out.ligacao  = Math.round(out.ligacao);
  out.rota     = Math.round(out.rota);
  return out;
}


// =============================================================
// HELPERS UTILITARIOS
// =============================================================

function getOrCreate(ss, name) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function fmt(d) {
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function diasAtras(ref, n) {
  const d = new Date(ref);
  d.setDate(d.getDate() - n);
  return d;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}
