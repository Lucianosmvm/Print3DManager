// ═══════════════════════════════════════════════════════════
// PRECIFICADOR 3D — Google Apps Script
// Cole este código em: Extensions > Apps Script > Code.gs
// Depois: Deploy > New deployment > Web app
//   Execute as: Me
//   Who has access: Anyone
// Copie a URL gerada e cole no app (campo APPS_SCRIPT_URL)
// ═══════════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const SHEETS = {
  produtos:      'Produtos',
  filamentos:    'Filamentos',
  solicitacoes:  'Solicitacoes',
  config:        'Config',
  dashboard:     'Dashboard'
};

const HEADERS = {
  produtos: [
    'ID','Nome','Emoji','Cliente','Data','DataImpressao',
    'Qtd','CustoTotal','Preco','LucroLiq','PesoUsado',
    'Horas','Material','Filamento','Impresso','Link'
  ],
  filamentos: [
    'ID','Material','Nome','Cor','TotalG','RestanteG','PrecoKg'
  ],
  solicitacoes: [
    'ID','Produto','Cliente','Qtd','Prioridade','Status','DataCriacao','Obs','Link'
  ],
  config: [
    'Chave','Valor'
  ]
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'sync') return syncAll(payload.data);
    if (action === 'load') return loadAll();

    return respond({ ok: false, error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function doGet(e) {
  if (e.parameter.action === 'load') return loadAll();
  return respond({ ok: true, message: 'Precificador 3D Apps Script ativo.' });
}

// ── SYNC: app → sheets ──────────────────────────────────────
function syncAll(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  writeSheet(ss, SHEETS.produtos, HEADERS.produtos,
    (data.products || []).map(produtoToRow));

  writeSheet(ss, SHEETS.filamentos, HEADERS.filamentos,
    (data.filaments || []).map(filamentoToRow));

  writeSheet(ss, SHEETS.solicitacoes, HEADERS.solicitacoes,
    (data.requests || []).map(solicitacaoToRow));

  writeConfig(ss, data.config || {});
  writeDashboard(ss, data.products || []);

  return respond({ ok: true, message: 'Sincronizado com sucesso!' });
}

// ── LOAD: sheets → app ──────────────────────────────────────
function loadAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const products   = readSheet(ss, SHEETS.produtos,     rowToProduto);
  const filaments  = readSheet(ss, SHEETS.filamentos,   rowToFilamento);
  const requests   = readSheet(ss, SHEETS.solicitacoes, rowToSolicitacao);
  const config     = readConfig(ss);

  return respond({ ok: true, data: { products, filaments, requests, config } });
}

// ── Helpers: write / read ────────────────────────────────────
function writeSheet(ss, name, headers, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();

  const all = [headers, ...rows];
  sheet.getRange(1, 1, all.length, headers.length).setValues(all);

  // Bold header
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

function readSheet(ss, name, mapper) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).filter(r => r[0] !== '').map(mapper);
}

function writeConfig(ss, cfg) {
  let sheet = ss.getSheetByName(SHEETS.config);
  if (!sheet) sheet = ss.insertSheet(SHEETS.config);
  sheet.clearContents();

  const rows = [HEADERS.config, ...Object.entries(cfg).map(([k, v]) => [k, v])];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
}

function readConfig(ss) {
  const sheet = ss.getSheetByName(SHEETS.config);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const cfg = {};
  data.slice(1).forEach(([k, v]) => { if (k) cfg[k] = v; });
  return cfg;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Mappers: objeto → linha ──────────────────────────────────
function produtoToRow(p) {
  return [
    p.id || '',
    p.nome || '',
    p.emoji || '',
    p.cliente || '',
    p.data || '',
    p.dataImpressao || '',
    p.qtd || 1,
    p.custoTotal || 0,
    p.preco || 0,
    p.lucroLiq || 0,
    p.pesoUsado || 0,
    p.horas || 0,
    p.filamentos ? p.filamentos.map(f => f.filamentoNome).join(', ') : (p.filamentoNome || ''),
    p.filamentoNome || '',
    p.impresso ? 'Sim' : 'Não',
    p.link || ''
  ];
}

function filamentoToRow(f) {
  return [
    f.id || '',
    f.material || '',
    f.nome || '',
    f.color || '',
    f.total || 0,
    f.restante || 0,
    f.precoKg || 0
  ];
}

function solicitacaoToRow(r) {
  return [
    r.id || '',
    r.produto || '',
    r.cliente || '',
    r.qtd || 1,
    r.prioridade || 'normal',
    r.status || 'pendente',
    r.dataCriacao || '',
    r.obs || '',
    r.link || ''
  ];
}

// ── Mappers: linha → objeto ──────────────────────────────────
function rowToProduto(r) {
  return {
    id: r[0],
    nome: r[1],
    emoji: r[2],
    cliente: r[3],
    data: r[4],
    dataImpressao: r[5] || null,
    qtd: Number(r[6]) || 1,
    custoTotal: Number(r[7]) || 0,
    preco: Number(r[8]) || 0,
    lucroLiq: Number(r[9]) || 0,
    pesoUsado: Number(r[10]) || 0,
    horas: Number(r[11]) || 0,
    filamentoNome: r[13] || r[12] || '',
    impresso: r[14] === 'Sim',
    link: r[15] || null,
    filamentos: []
  };
}

function rowToFilamento(r) {
  return {
    id: r[0],
    material: r[1],
    nome: r[2],
    color: r[3],
    total: Number(r[4]) || 0,
    restante: Number(r[5]) || 0,
    precoKg: Number(r[6]) || 0
  };
}

function rowToSolicitacao(r) {
  return {
    id: r[0],
    produto: r[1],
    cliente: r[2],
    qtd: Number(r[3]) || 1,
    prioridade: r[4] || 'normal',
    status: r[5] || 'pendente',
    dataCriacao: r[6],
    obs: r[7] || '',
    link: r[8] || ''
  };
}

// ── Dashboard: KPIs por mês ──────────────────────────────────
function writeDashboard(ss, products) {
  let sheet = ss.getSheetByName(SHEETS.dashboard);
  if (!sheet) sheet = ss.insertSheet(SHEETS.dashboard);
  sheet.clearContents();
  sheet.clearFormats();

  const printed = products.filter(p => p.impresso === true || p.impresso === 'Sim');

  // Group by "mm/yyyy"
  const byMonth = {};
  printed.forEach(p => {
    const dateStr = p.dataImpressao || p.data || '';
    const parts = dateStr.split('/');
    if (parts.length < 3) return;
    const key = parts[1] + '/' + parts[2]; // mm/yyyy
    if (!byMonth[key]) byMonth[key] = { receita: 0, custo: 0, lucro: 0, pecas: 0, peso: 0 };
    byMonth[key].receita += (Number(p.preco) || 0) * (Number(p.qtd) || 1);
    byMonth[key].custo   += Number(p.custoTotal) || 0;
    byMonth[key].lucro   += Number(p.lucroLiq)   || 0;
    byMonth[key].pecas   += Number(p.qtd) || 1;
    byMonth[key].peso    += Number(p.pesoUsado)  || 0;
  });

  // Sort months chronologically
  const sorted = Object.keys(byMonth).sort((a, b) => {
    const [ma, ya] = a.split('/').map(Number);
    const [mb, yb] = b.split('/').map(Number);
    return ya !== yb ? ya - yb : ma - mb;
  });

  const headers = ['Mês', 'Peças Impressas', 'Receita (R$)', 'Custo (R$)', 'Lucro Líquido (R$)', 'Margem (%)', 'Filamento (g)'];
  const rows = sorted.map(key => {
    const d = byMonth[key];
    const margem = d.receita > 0 ? ((d.lucro / d.receita) * 100).toFixed(1) : 0;
    return [key, d.pecas, d.receita, d.custo, d.lucro, Number(margem), d.peso];
  });

  // Total row
  const totReceita = rows.reduce((a, r) => a + r[2], 0);
  const totCusto   = rows.reduce((a, r) => a + r[3], 0);
  const totLucro   = rows.reduce((a, r) => a + r[4], 0);
  const totPecas   = rows.reduce((a, r) => a + r[1], 0);
  const totPeso    = rows.reduce((a, r) => a + r[6], 0);
  const totMargem  = totReceita > 0 ? Number(((totLucro / totReceita) * 100).toFixed(1)) : 0;
  rows.push(['TOTAL', totPecas, totReceita, totCusto, totLucro, totMargem, totPeso]);

  // Write title
  sheet.getRange(1, 1).setValue('Dashboard — Produtos Impressos');
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sheet.getRange(2, 1).setValue('Atualizado em: ' + new Date().toLocaleString('pt-BR'));
  sheet.getRange(2, 1).setFontColor('#888888');

  // Write table starting at row 4
  const tableStart = 4;
  const allRows = [headers, ...rows];
  sheet.getRange(tableStart, 1, allRows.length, headers.length).setValues(allRows);

  // Format header row
  const headerRange = sheet.getRange(tableStart, 1, 1, headers.length);
  headerRange.setFontWeight('bold').setBackground('#4A3890').setFontColor('#FFFFFF');

  // Format currency columns (Receita, Custo, Lucro)
  if (rows.length > 0) {
    const dataRows = rows.length;
    sheet.getRange(tableStart + 1, 3, dataRows, 3).setNumberFormat('R$ #,##0.00');
    sheet.getRange(tableStart + 1, 6, dataRows, 1).setNumberFormat('0.0"%"');
    sheet.getRange(tableStart + 1, 7, dataRows, 1).setNumberFormat('#,##0"g"');

    // Bold total row
    sheet.getRange(tableStart + dataRows, 1, 1, headers.length).setFontWeight('bold').setBackground('#EDE8F8');
  }

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // KPI summary block (above table, columns I-J)
  const kpiLabels = [
    ['KPI', 'Valor Total'],
    ['Receita Total', totReceita],
    ['Custo Total', totCusto],
    ['Lucro Total', totLucro],
    ['Margem Geral (%)', totMargem],
    ['Peças Impressas', totPecas],
    ['Filamento Usado (g)', totPeso]
  ];
  sheet.getRange(1, 9, kpiLabels.length, 2).setValues(kpiLabels);
  sheet.getRange(1, 9, 1, 2).setFontWeight('bold').setBackground('#4A3890').setFontColor('#FFFFFF');
  sheet.getRange(2, 10, 4, 1).setNumberFormat('R$ #,##0.00');
  sheet.getRange(5, 10).setNumberFormat('0.0"%"');
  sheet.autoResizeColumn(9);
  sheet.autoResizeColumn(10);
}
