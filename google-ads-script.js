/**
 * GOOGLE ADS SCRIPT · Painel HVSB
 *
 * Como usar:
 * 1. Logar no Google Ads → Ferramentas → Scripts → Novo script
 * 2. Colar este código
 * 3. Trocar SPREADSHEET_URL pela URL da sua Google Sheet
 * 4. Salvar e autorizar
 * 5. Configurar agendamento: De hora em hora (Trigger horário)
 *
 * O que faz: lê dados da conta + campanhas + tendência diária e escreve numa Sheet,
 * que o dashboard HTML consome.
 */

const SPREADSHEET_URL = 'COLE_AQUI_A_URL_DA_SUA_SHEET';
const MES_META_PERCENT = 1.10; // meta = mês anterior + 10%

function main() {
  const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

  // === 1. SALDO E DADOS DA CONTA ===
  const account = AdsApp.currentAccount();
  const accountBudget = getAccountBudget();

  // === 2. CONVERSÕES E CUSTO MTD ===
  const today = new Date();
  const inicioMes = new Date(today.getFullYear(), today.getMonth(), 1);
  const mtd = getStats(inicioMes, today);

  // Mês anterior
  const inicioMesAnt = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const fimMesAnt = new Date(today.getFullYear(), today.getMonth(), 0);
  const mesAnt = getStats(inicioMesAnt, fimMesAnt);

  // Gasto médio últimos 7 dias
  const seteDiasAtras = new Date(today);
  seteDiasAtras.setDate(today.getDate() - 7);
  const ultimos7d = getStats(seteDiasAtras, today);
  const gastoMedio7d = ultimos7d.cost / 7;

  // === 3. ESCREVER KPIs ===
  const sheetKpis = getOrCreateSheet(ss, 'KPIs');
  sheetKpis.clear();
  sheetKpis.appendRow(['campo', 'valor']);
  sheetKpis.appendRow(['ultimaAtualizacao', new Date().toISOString()]);
  sheetKpis.appendRow(['account', AdsApp.currentAccount().getCustomerId()]);
  sheetKpis.appendRow(['accountName', AdsApp.currentAccount().getName()]);
  sheetKpis.appendRow(['saldoAtual', accountBudget]);
  sheetKpis.appendRow(['gastoMedio7d', gastoMedio7d]);
  sheetKpis.appendRow(['conversoesMTD', mtd.conversions]);
  sheetKpis.appendRow(['custoMTD', mtd.cost]);
  sheetKpis.appendRow(['cpaMTD', mtd.conversions ? mtd.cost / mtd.conversions : 0]);
  sheetKpis.appendRow(['cliquesMTD', mtd.clicks]);
  sheetKpis.appendRow(['conversoesMesAnterior', mesAnt.conversions]);
  sheetKpis.appendRow(['custoMesAnterior', mesAnt.cost]);
  sheetKpis.appendRow(['cpaMesAnterior', mesAnt.conversions ? mesAnt.cost / mesAnt.conversions : 0]);
  sheetKpis.appendRow(['cliquesMesAnterior', mesAnt.clicks]);
  sheetKpis.appendRow(['meta', Math.round(mesAnt.conversions * MES_META_PERCENT)]);
  sheetKpis.appendRow(['metaOrigem', `Mês anterior (${mesAnt.conversions}) + ${(MES_META_PERCENT - 1) * 100}%`]);

  // === 4. CAMPANHAS ===
  const sheetCamp = getOrCreateSheet(ss, 'Campanhas');
  sheetCamp.clear();
  sheetCamp.appendRow(['nome', 'tipo', 'conv', 'cpa', 'custo', 'isLostBudget', 'status']);

  const campanhas = getCampanhas(inicioMes, today);
  campanhas.forEach(c => sheetCamp.appendRow([c.nome, c.tipo, c.conv, c.cpa, c.custo, c.isLostBudget, c.status]));

  // === 5. TENDÊNCIA DIÁRIA · ÚLTIMOS 30 DIAS ===
  const sheetDiario = getOrCreateSheet(ss, 'Diario');
  sheetDiario.clear();
  sheetDiario.appendRow(['data', 'conv', 'custo']);

  const trintaDiasAtras = new Date(today);
  trintaDiasAtras.setDate(today.getDate() - 30);
  const diario = getDiario(trintaDiasAtras, today);
  diario.forEach(d => sheetDiario.appendRow([d.data, d.conv, d.custo]));

  // === 6. CONVERSÕES POR TIPO ===
  const sheetTipo = getOrCreateSheet(ss, 'Tipos');
  sheetTipo.clear();
  sheetTipo.appendRow(['tipo', 'total']);
  const tipos = getConversoesPorTipo(inicioMes, today);
  Object.keys(tipos).forEach(t => sheetTipo.appendRow([t, tipos[t]]));

  Logger.log('✓ Painel atualizado: ' + new Date().toISOString());
}

// ============ HELPERS ============

function getAccountBudget() {
  // Pega o saldo da conta (pré-pago)
  const budgets = AdsApp.budgets().get();
  let total = 0;
  while (budgets.hasNext()) total += budgets.next().getAmount();
  return total;

  /*
   * Para conta pré-paga real, o saldo precisa vir de outra API.
   * Alternativa: cliente atualiza manualmente o saldo numa célula da Sheet,
   * o script só lê esse valor.
   * Ex: sheetKpis.getRange('B5').setValue(saldoManual);
   */
}

function getStats(dateFrom, dateTo) {
  const start = formatDate(dateFrom);
  const end = formatDate(dateTo);
  const report = AdsApp.report(`
    SELECT metrics.clicks, metrics.conversions, metrics.cost_micros
    FROM customer
    WHERE segments.date BETWEEN '${start}' AND '${end}'
  `);
  const row = report.rows().next();
  return {
    clicks: parseInt(row['metrics.clicks']) || 0,
    conversions: parseFloat(row['metrics.conversions']) || 0,
    cost: (parseFloat(row['metrics.cost_micros']) || 0) / 1e6,
  };
}

function getCampanhas(dateFrom, dateTo) {
  const start = formatDate(dateFrom);
  const end = formatDate(dateTo);
  const report = AdsApp.report(`
    SELECT campaign.name, campaign.advertising_channel_type, campaign.status,
           metrics.clicks, metrics.conversions, metrics.cost_micros,
           metrics.search_budget_lost_impression_share
    FROM campaign
    WHERE segments.date BETWEEN '${start}' AND '${end}'
  `);

  const out = [];
  const rows = report.rows();
  while (rows.hasNext()) {
    const r = rows.next();
    const conv = parseFloat(r['metrics.conversions']) || 0;
    const cost = (parseFloat(r['metrics.cost_micros']) || 0) / 1e6;
    const isLostBudget = parseFloat(r['metrics.search_budget_lost_impression_share']) || 0;
    const status = r['campaign.status'];

    let statusOut = 'ok';
    if (status === 'PAUSED') statusOut = 'paused';
    else if (isLostBudget > 0.05) statusOut = 'limited';
    else if (conv === 0) statusOut = 'no-conv';

    let tipo = 'Search';
    if (r['campaign.advertising_channel_type'] === 'PERFORMANCE_MAX') tipo = 'PMax';
    if (r['campaign.advertising_channel_type'] === 'DISPLAY') tipo = 'Display';

    out.push({
      nome: r['campaign.name'],
      tipo,
      conv: Math.round(conv),
      cpa: conv ? cost / conv : 0,
      custo: cost,
      isLostBudget: Math.round(isLostBudget * 100),
      status: statusOut,
    });
  }
  return out;
}

function getDiario(dateFrom, dateTo) {
  const start = formatDate(dateFrom);
  const end = formatDate(dateTo);
  const report = AdsApp.report(`
    SELECT segments.date, metrics.conversions, metrics.cost_micros
    FROM customer
    WHERE segments.date BETWEEN '${start}' AND '${end}'
  `);

  const out = [];
  const rows = report.rows();
  while (rows.hasNext()) {
    const r = rows.next();
    out.push({
      data: r['segments.date'],
      conv: Math.round(parseFloat(r['metrics.conversions']) || 0),
      custo: (parseFloat(r['metrics.cost_micros']) || 0) / 1e6,
    });
  }
  return out.sort((a, b) => a.data.localeCompare(b.data));
}

function getConversoesPorTipo(dateFrom, dateTo) {
  const start = formatDate(dateFrom);
  const end = formatDate(dateTo);
  // Agrupa por nome da ação de conversão (configurar nomes no Google Ads)
  const report = AdsApp.report(`
    SELECT segments.conversion_action_name, metrics.conversions
    FROM customer
    WHERE segments.date BETWEEN '${start}' AND '${end}'
  `);

  const out = { whatsapp: 0, ligacao: 0, rota: 0 };
  const rows = report.rows();
  while (rows.hasNext()) {
    const r = rows.next();
    const nome = (r['segments.conversion_action_name'] || '').toLowerCase();
    const conv = parseFloat(r['metrics.conversions']) || 0;
    if (nome.includes('whats')) out.whatsapp += conv;
    else if (nome.includes('liga') || nome.includes('call') || nome.includes('phone')) out.ligacao += conv;
    else if (nome.includes('rota') || nome.includes('direc') || nome.includes('maps')) out.rota += conv;
  }
  out.whatsapp = Math.round(out.whatsapp);
  out.ligacao = Math.round(out.ligacao);
  out.rota = Math.round(out.rota);
  return out;
}

function getOrCreateSheet(ss, name) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function formatDate(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}
