/**
 * APPS SCRIPT · HVSB Dashboard
 * Script standalone vinculado à planilha abaixo.
 *
 * COMO INSTALAR:
 * 1. script.google.com → Novo projeto (standalone, fora do Sheets)
 * 2. Cole este código completo
 * 3. Preencha SHEET_ID com o ID da sua planilha (entre /d/ e /edit na URL)
 * 4. Implantações → Nova implantação → Aplicativo da web
 *    - Executar como: Eu mesmo
 *    - Quem tem acesso: Qualquer pessoa (mesmo sem login)
 * 5. Copie a URL gerada e cole como DATA_URL no index.html do dashboard
 *
 * FUNCOES disponiveis:
 *   doGet(e)     — endpoint JSON consumido pelo dashboard
 *   setupSheets  — cria as 5 abas com cabecalhos formatados
 *   seedData     — popula dados de exemplo para testar sem o Google Ads Script
 *   setupAndSeed — executa as duas acoes acima em sequencia
 */

// =============================================================
// CONFIGURACAO — altere apenas este bloco
// =============================================================

const SHEET_ID = '1ao1S84eBniYIvzdgNFQZ6poCIXFJkdtUZnlUIa6L9NQ';

// Campos de saldo atualizados manualmente na aba "saldo"
// (o Google Ads Script nao consegue ler saldo pre-pago pela API)

// =============================================================


// =============================================================
// doGet — endpoint principal (retorna JSON para o dashboard)
// =============================================================

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // --- meta ---
    const metaRaw = ss.getSheetByName('meta').getDataRange().getValues().slice(1);
    const metaObj = {};
    metaRaw.forEach(function(row) { if (row[0]) metaObj[row[0]] = row[1]; });

    // --- kpis ---
    const kpisRaw = ss.getSheetByName('kpis').getDataRange().getValues().slice(1);
    const kpisObj = {};
    kpisRaw.forEach(function(row) { if (row[0]) kpisObj[row[0]] = row[1]; });

    // --- saldo (atualizado manualmente) ---
    const saldoSheet = ss.getSheetByName('saldo');
    const saldoRaw   = saldoSheet ? saldoSheet.getDataRange().getValues().slice(1) : [];
    const saldoObj   = {};
    saldoRaw.forEach(function(row) { if (row[0]) saldoObj[row[0]] = row[1]; });

    // --- campanhas ---
    const campSheet  = ss.getSheetByName('campanhas');
    const campRaw    = campSheet.getDataRange().getValues();
    const campHeaders = campRaw[0];
    const campanhas  = campRaw.slice(1).filter(function(r) { return r[0]; }).map(function(r) {
      const obj = {};
      campHeaders.forEach(function(h, i) { obj[h] = r[i]; });
      return obj;
    });

    // --- diario ---
    const diarioSheet   = ss.getSheetByName('diario');
    const diarioRaw     = diarioSheet.getDataRange().getValues();
    const diarioHeaders = diarioRaw[0];
    const diario        = diarioRaw.slice(1).filter(function(r) { return r[0]; }).map(function(r) {
      const obj = {};
      diarioHeaders.forEach(function(h, i) {
        if (h === 'data' && r[i] instanceof Date) {
          obj[h] = Utilities.formatDate(r[i], 'America/Sao_Paulo', 'yyyy-MM-dd');
        } else {
          obj[h] = r[i];
        }
      });
      return obj;
    });

    // --- conversoesPorTipo ---
    const tiposSheet = ss.getSheetByName('tipos');
    const tiposRaw   = tiposSheet ? tiposSheet.getDataRange().getValues().slice(1) : [];
    const conversoesPorTipo = {};
    tiposRaw.forEach(function(row) { if (row[0]) conversoesPorTipo[row[0]] = row[1]; });

    // --- montagem do JSON ---
    const hoje = new Date();

    const out = {
      _meta: {
        lastUpdate:  metaObj.lastUpdate  || new Date().toISOString(),
        currency:    metaObj.currency    || 'BRL',
        account:     metaObj.account     || '192-048-4711',
        accountName: metaObj.accountName || 'Hospital Veterinário São Bernardo',
        city:        metaObj.city        || 'Curitiba/PR',
        source:      'sheet-webapp',
      },
      saldo: {
        atual:              Number(saldoObj.saldoAtual)         || 0,
        gastoMedio7d:       Number(kpisObj.gastoMedio7d)        || 0,
        ultimaRecarga:      saldoObj.ultimaRecarga              || '',
        ultimaRecargaValor: Number(saldoObj.ultimaRecargaValor) || 0,
      },
      mes: {
        ano:           hoje.getFullYear(),
        mes:           hoje.getMonth() + 1,
        diasDoMes:     new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getDate(),
        diasDecorridos: hoje.getDate(),
        meta:          Number(kpisObj.meta)        || 0,
        metaOrigem:    kpisObj.metaOrigem          || '',
      },
      kpis: {
        conversoesMTD:         Number(kpisObj.conversoesMTD)         || 0,
        conversoesMesAnterior: Number(kpisObj.conversoesMesAnterior) || 0,
        custoMTD:              Number(kpisObj.custoMTD)              || 0,
        custoMesAnterior:      Number(kpisObj.custoMesAnterior)      || 0,
        cpaMTD:                Number(kpisObj.cpaMTD)                || 0,
        cpaMesAnterior:        Number(kpisObj.cpaMesAnterior)        || 0,
        cliquesMTD:            Number(kpisObj.cliquesMTD)            || 0,
        cliquesMesAnterior:    Number(kpisObj.cliquesMesAnterior)    || 0,
      },
      campanhas: campanhas,
      diario:    diario,
      conversoesPorTipo: conversoesPorTipo,
      searchIS:  Number(kpisObj.searchIS) || 0,
    };

    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message, stack: err.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// =============================================================
// setupSheets — cria abas com cabecalhos
// =============================================================

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const abas = {
    'meta':     ['chave', 'valor'],
    'kpis':     ['chave', 'valor'],
    'saldo':    ['chave', 'valor'],
    'campanhas':['nome', 'tipo', 'conv', 'cpa', 'custo', 'is', 'isLostBudget', 'status'],
    'diario':   ['data', 'conv', 'custo', 'cliques'],
    'tipos':    ['tipo', 'total'],
  };

  for (var nome in abas) {
    var s = ss.getSheetByName(nome);
    if (!s) s = ss.insertSheet(nome);
    else s.clear();
    s.appendRow(abas[nome]);
    s.getRange(1, 1, 1, abas[nome].length)
      .setFontWeight('bold')
      .setBackground('#0E5C8C')
      .setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }

  Logger.log('OK · 6 abas criadas/atualizadas.');
}


// =============================================================
// seedData — popula dados de exemplo para teste
// (substitui pelo Google Ads Script quando estiver ao vivo)
// =============================================================

function seedData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // --- meta ---
  var sMeta = ss.getSheetByName('meta');
  if (sMeta.getLastRow() > 1) sMeta.getRange(2, 1, sMeta.getLastRow() - 1, 2).clearContent();
  sMeta.getRange(2, 1, 6, 2).setValues([
    ['lastUpdate',   new Date().toISOString()],
    ['currency',     'BRL'],
    ['account',      '192-048-4711'],
    ['accountName',  'Hospital Veterinário São Bernardo'],
    ['city',         'Curitiba/PR'],
    ['source',       'seed-data'],
  ]);

  // --- kpis ---
  var sKpis = ss.getSheetByName('kpis');
  if (sKpis.getLastRow() > 1) sKpis.getRange(2, 1, sKpis.getLastRow() - 1, 2).clearContent();
  sKpis.getRange(2, 1, 12, 2).setValues([
    ['conversoesMTD',         390],
    ['conversoesMesAnterior', 281],
    ['custoMTD',              3145.47],
    ['custoMesAnterior',      3070.00],
    ['cpaMTD',                8.07],
    ['cpaMesAnterior',        10.91],
    ['cliquesMTD',            2615],
    ['cliquesMesAnterior',    2970],
    ['meta',                  310],
    ['metaOrigem',            'Abril 281 conv + 10%'],
    ['gastoMedio7d',          117.30],
    ['searchIS',              48],
  ]);

  // --- saldo (manual — nao sobrescrito pelo Google Ads Script) ---
  var sSaldo = ss.getSheetByName('saldo');
  if (sSaldo.getLastRow() > 1) sSaldo.getRange(2, 1, sSaldo.getLastRow() - 1, 2).clearContent();
  sSaldo.getRange(2, 1, 3, 2).setValues([
    ['saldoAtual',          1849.68],
    ['ultimaRecarga',       '2026-05-26'],
    ['ultimaRecargaValor',  2000.00],
  ]);

  // --- campanhas ---
  var sCamp = ss.getSheetByName('campanhas');
  if (sCamp.getLastRow() > 1) sCamp.getRange(2, 1, sCamp.getLastRow() - 1, 8).clearContent();
  sCamp.getRange(2, 1, 7, 8).setValues([
    ['[SEARCH] [CPC M] Especialidades Veterinárias', 'Search',  91.98, 9.21,   847.20,  38, 52, 'limited'],
    ['[SEARCH] [CPC M] Hospital Veterinário',        'Search',  57.85, 11.94,  690.98,  41, 47, 'limited'],
    ['Display | Remarketing | Leads - HVSB',         'Display', 115.0, 3.84,   441.56,   0, 18, 'limited'],
    ['[SEARCH] [MAX Q] Domicílio',                   'Search',  43.50, 13.70,  595.86,   0, 38, 'limited'],
    ['[SEARCH] [CPC M] Branding',                    'Search',  65.99, 4.00,   264.10,  88,  7, 'ok'],
    ['[SEARCH] [SITE] Institucional | HVSB',         'Search',  14.67, 9.26,   135.78,  60, 30, 'limited'],
    ['[SEARCH] [CPC M] Hospital Veterinário Call',   'Search',   1.00, 169.67, 169.98,  55, 40, 'limited'],
  ]);

  // --- diario (ultimos 36 dias como exemplo) ---
  var sDiario = ss.getSheetByName('diario');
  if (sDiario.getLastRow() > 1) sDiario.getRange(2, 1, sDiario.getLastRow() - 1, 4).clearContent();
  sDiario.getRange(2, 1, 36, 4).setValues([
    ['2026-04-22', 10, 108.00,  91], ['2026-04-23', 14, 123.00, 102], ['2026-04-24', 12, 116.00,  98],
    ['2026-04-25',  9, 100.00,  84], ['2026-04-26',  8,  81.00,  76], ['2026-04-27', 10, 101.00,  88],
    ['2026-04-28', 15, 135.00, 115], ['2026-04-29', 18, 143.00, 124], ['2026-04-30', 13, 127.53, 108],
    ['2026-05-01', 11, 105.70,  92], ['2026-05-02', 13, 118.20, 101], ['2026-05-03',  9,  96.40,  83],
    ['2026-05-04',  8,  87.10,  75], ['2026-05-05', 14, 122.30, 104], ['2026-05-06', 16, 130.55, 112],
    ['2026-05-07', 15, 128.40, 109], ['2026-05-08', 17, 133.20, 114], ['2026-05-09', 12, 110.85,  96],
    ['2026-05-10', 10,  99.30,  87], ['2026-05-11', 18, 138.20, 117], ['2026-05-12', 19, 142.55, 121],
    ['2026-05-13', 16, 130.80, 111], ['2026-05-14', 14, 121.40, 103], ['2026-05-15', 17, 134.20, 114],
    ['2026-05-16', 15, 128.55, 109], ['2026-05-17', 11, 104.30,  91], ['2026-05-18', 13, 112.80, 97],
    ['2026-05-19', 18, 139.40, 118], ['2026-05-20', 21, 152.20, 129], ['2026-05-21', 17, 133.55, 113],
    ['2026-05-22', 15, 100.10,  89], ['2026-05-23', 17, 108.30,  96], ['2026-05-24', 10,  82.40,  73],
    ['2026-05-25',  9,  76.20,  67], ['2026-05-26', 18, 112.00,  98], ['2026-05-27', 17,  96.52,  85],
  ]);

  // --- tipos ---
  var sTipos = ss.getSheetByName('tipos');
  if (sTipos.getLastRow() > 1) sTipos.getRange(2, 1, sTipos.getLastRow() - 1, 2).clearContent();
  sTipos.getRange(2, 1, 3, 2).setValues([
    ['whatsapp', 237],
    ['ligacao',  109],
    ['rota',      44],
  ]);

  Logger.log('OK · Dados de exemplo populados.');
}


function setupAndSeed() {
  setupSheets();
  seedData();
  Logger.log('OK · Tudo pronto. Abra a planilha para conferir.');
}
