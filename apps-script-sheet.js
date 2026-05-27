/**
 * APPS SCRIPT · HVSB Dashboard (standalone)
 * Vinculado à Sheet via SHEET_ID abaixo.
 */

const SHEET_ID = '1H1h33ebtPfWGW3Hh9Aoi-iaKwtTV2n7ZbcZdAR3omhw';

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const abas = {
    'KPIs': ['campo', 'valor'],
    'Campanhas': ['nome', 'tipo', 'conv', 'cpa', 'custo', 'isLostBudget', 'status'],
    'Diario': ['data', 'conv', 'custo'],
    'Tipos': ['tipo', 'total']
  };
  for (const nome in abas) {
    let s = ss.getSheetByName(nome);
    if (!s) s = ss.insertSheet(nome);
    else s.clear();
    s.appendRow(abas[nome]);
    s.getRange(1, 1, 1, abas[nome].length)
      .setFontWeight('bold').setBackground('#0E5C8C').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
  }
  Logger.log('✓ 4 abas criadas/atualizadas.');
}

function seedData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const kpis = ss.getSheetByName('KPIs');
  if (kpis.getLastRow() > 1) kpis.getRange(2, 1, kpis.getLastRow() - 1, 2).clearContent();
  kpis.getRange(2, 1, 14, 2).setValues([
    ['ultimaAtualizacao', new Date().toISOString()],
    ['accountName', 'Hospital Veterinário São Bernardo'],
    ['saldoAtual', 1850.00],
    ['gastoMedio7d', 117.50],
    ['conversoesMTD', 308],
    ['custoMTD', 2480.00],
    ['cpaMTD', 8.07],
    ['cliquesMTD', 2060],
    ['conversoesMesAnterior', 281],
    ['custoMesAnterior', 3070.00],
    ['cpaMesAnterior', 10.91],
    ['cliquesMesAnterior', 2970],
    ['meta', 310],
    ['metaOrigem', 'Abril 281 conv + 10%']
  ]);

  const camp = ss.getSheetByName('Campanhas');
  if (camp.getLastRow() > 1) camp.getRange(2, 1, camp.getLastRow() - 1, 7).clearContent();
  camp.getRange(2, 1, 7, 7).setValues([
    ['Especialidades · Search', 'Search', 84, 6.42, 539.28, 52, 'limited'],
    ['Hospital Vet · Search', 'Search', 62, 9.85, 610.70, 47, 'limited'],
    ['Emergência 24h · Search', 'Search', 58, 7.21, 418.18, 58, 'limited'],
    ['Branding HVSB · Search', 'Search', 42, 4.18, 175.56, 7, 'ok'],
    ['PMax Geral', 'PMax', 31, 11.40, 353.40, 38, 'limited'],
    ['PMax Maps', 'PMax', 18, 9.85, 177.30, 22, 'limited'],
    ['Display Remarketing', 'Display', 13, 3.05, 39.65, 18, 'limited']
  ]);

  const diario = ss.getSheetByName('Diario');
  if (diario.getLastRow() > 1) diario.getRange(2, 1, diario.getLastRow() - 1, 3).clearContent();
  diario.getRange(2, 1, 30, 3).setValues([
    ['2026-04-22', 8, 95.20], ['2026-04-23', 11, 108.60], ['2026-04-24', 9, 102.10],
    ['2026-04-25', 7, 88.40], ['2026-04-26', 6, 71.80], ['2026-04-27', 8, 89.55],
    ['2026-04-28', 12, 119.30], ['2026-04-29', 14, 125.85], ['2026-04-30', 10, 112.40],
    ['2026-05-01', 11, 105.70], ['2026-05-02', 13, 118.20], ['2026-05-03', 9, 96.40],
    ['2026-05-04', 8, 87.10], ['2026-05-05', 14, 122.30], ['2026-05-06', 16, 130.55],
    ['2026-05-07', 15, 128.40], ['2026-05-08', 17, 133.20], ['2026-05-09', 12, 110.85],
    ['2026-05-10', 10, 99.30], ['2026-05-11', 18, 138.20], ['2026-05-12', 19, 142.55],
    ['2026-05-13', 16, 130.80], ['2026-05-14', 14, 121.40], ['2026-05-15', 17, 134.20],
    ['2026-05-16', 15, 128.55], ['2026-05-17', 11, 104.30], ['2026-05-18', 13, 112.80],
    ['2026-05-19', 18, 139.40], ['2026-05-20', 21, 152.20], ['2026-05-21', 17, 133.55]
  ]);

  const tipos = ss.getSheetByName('Tipos');
  if (tipos.getLastRow() > 1) tipos.getRange(2, 1, tipos.getLastRow() - 1, 2).clearContent();
  tipos.getRange(2, 1, 3, 2).setValues([
    ['whatsapp', 187],
    ['ligacao', 89],
    ['rota', 32]
  ]);

  Logger.log('✓ Dados de exemplo populados.');
}

function setupAndSeed() {
  setupSheets();
  seedData();
  Logger.log('✓ Tudo pronto.');
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const kpisRaw = ss.getSheetByName('KPIs').getDataRange().getValues().slice(1);
    const kpis = {};
    kpisRaw.forEach(function(row) { if (row[0]) kpis[row[0]] = row[1]; });

    const campRaw = ss.getSheetByName('Campanhas').getDataRange().getValues();
    const campHeaders = campRaw[0];
    const campanhas = campRaw.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      const obj = {};
      campHeaders.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    });

    const diarioRaw = ss.getSheetByName('Diario').getDataRange().getValues();
    const diarioHeaders = diarioRaw[0];
    const diario = diarioRaw.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      const obj = {};
      diarioHeaders.forEach(function(h, i){
        if (h === 'data' && r[i] instanceof Date) {
          obj[h] = Utilities.formatDate(r[i], 'America/Sao_Paulo', 'yyyy-MM-dd');
        } else {
          obj[h] = r[i];
        }
      });
      return obj;
    });

    const tiposRaw = ss.getSheetByName('Tipos').getDataRange().getValues().slice(1);
    const conversoesPorTipo = {};
    tiposRaw.forEach(function(row){ if (row[0]) conversoesPorTipo[row[0]] = row[1]; });

    const hoje = new Date();
    const out = {
      _meta: {
        lastUpdate: kpis.ultimaAtualizacao || new Date().toISOString(),
        currency: 'BRL',
        account: '192-048-4711',
        accountName: kpis.accountName || 'Hospital Veterinário São Bernardo',
        city: 'Curitiba/PR',
        source: 'sheet-webapp'
      },
      saldo: {
        atual: Number(kpis.saldoAtual) || 0,
        gastoMedio7d: Number(kpis.gastoMedio7d) || 0
      },
      mes: {
        ano: hoje.getFullYear(),
        mes: hoje.getMonth() + 1,
        diasDoMes: new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getDate(),
        diasDecorridos: hoje.getDate(),
        meta: Number(kpis.meta) || 0,
        metaOrigem: kpis.metaOrigem || ''
      },
      kpis: {
        conversoesMTD: Number(kpis.conversoesMTD) || 0,
        conversoesMesAnterior: Number(kpis.conversoesMesAnterior) || 0,
        custoMTD: Number(kpis.custoMTD) || 0,
        custoMesAnterior: Number(kpis.custoMesAnterior) || 0,
        cpaMTD: Number(kpis.cpaMTD) || 0,
        cpaMesAnterior: Number(kpis.cpaMesAnterior) || 0,
        cliquesMTD: Number(kpis.cliquesMTD) || 0,
        cliquesMesAnterior: Number(kpis.cliquesMesAnterior) || 0
      },
      campanhas: campanhas,
      diario: diario,
      conversoesPorTipo: conversoesPorTipo
    };

    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
