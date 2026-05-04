/**
 * ============================================================
 * Dashboard Turma B — Google Apps Script (Versão 2.0)
 * ============================================================
 *
 * INSTALAÇÃO:
 *  1. script.google.com → Novo projeto → cole este código
 *  2. Ative "Drive API" em: Serviços → + → Drive API
 *  3. Salve → Implantar → Nova implantação
 *     - Tipo: Aplicativo da web
 *     - Executar como: Usuário que instalou
 *     - Quem tem acesso: Qualquer pessoa
 *  4. Copie a URL e cole em APPS_SCRIPT_URL no index.html
 *
 * IDs DOS ARQUIVOS:
 *  - horas.xls:  1Sog9TjfIbq17VJyZKlD5Wj53JW4ZfZjm
 *  - status.xls: 1vsJdRzwO7XvAuQLTkbuQ2VSFUTzU0PeG
 * ============================================================
 */

const HORAS_FILE_ID  = '1Sog9TjfIbq17VJyZKlD5Wj53JW4ZfZjm';
const STATUS_FILE_ID = '1vsJdRzwO7XvAuQLTkbuQ2VSFUTzU0PeG';

const HORAS_COLS = {
  operador:    'Mão de obra',
  tipoServico: 'Tipo de Serviço',
  os:          'Ordem de Serviço',
  data:        'Data de Início',
  horas:       'Horas Normais'
};

const STATUS_COLS = {
  os:     'Ordem de Serviço',
  desc:   'Descrição',
  local:  'Local',
  status: 'Status'
};

const TIPOS_OPER  = { OPER: { nome: 'Operacional',  cor: 'var(--op)' } };
const TIPOS_MANUT = {
  CORR: { nome: 'Corretiva',    cor: 'var(--m-corr)' },
  EM:   { nome: 'Emergencial',  cor: 'var(--m-em)'   },
  PRED: { nome: 'Preditiva',    cor: 'var(--m-pred)' },
  PREV: { nome: 'Preventiva',   cor: 'var(--m-prev)' },
  SERV: { nome: 'Serviços',     cor: 'var(--m-serv)' },
  MODI: { nome: 'Modificativa', cor: 'var(--m-modi)' }
};
const TIPO_KEYS = ['OPER','CORR','EM','PRED','PREV','SERV','MODI'];

function doGet(e) {
  try {
    const result = buildDashboardData();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function buildDashboardData() {
  const toCleanup = [];

  function ensureSpreadsheet(fileId) {
    const file = DriveApp.getFileById(fileId);
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) return fileId;
    const blob = file.getBlob();
    const newFile = Drive.Files.insert(
      { title: file.getName() + '_temp', mimeType: MimeType.GOOGLE_SHEETS },
      blob,
      { convert: true }
    );
    toCleanup.push(newFile.id);
    return newFile.id;
  }

  function readSheet(ssId) {
    const ss = SpreadsheetApp.openById(ssId);
    return ss.getSheets()[0].getDataRange().getValues();
  }

  function normalizeName(name) {
    if (!name) return '';
    return name.toString().trim()
      .normalize('NFD').replace(/\p{Diacritic}/gu, '')
      .toUpperCase().replace(/\s+/g, '.');
  }

  function parseDate(value) {
    if (value instanceof Date && !isNaN(value)) return value;
    if (!value) return null;
    const parts = value.toString().trim().split(/[\/\-]/);
    if (parts.length < 3) return null;
    let d = parseInt(parts[0], 10);
    let m = parseInt(parts[1], 10) - 1;
    let y = parseInt(parts[2], 10);
    if (y < 100) y = 2000 + y;
    return new Date(Date.UTC(y, m, d));
  }

  function calcMeta(year, monthIdx) {
    const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
    let businessDays = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      const dow = new Date(Date.UTC(year, monthIdx, d)).getUTCDay();
      if (dow >= 1 && dow <= 5) businessDays++;
    }
    return { dias_corridos: daysInMonth, dias_uteis_calendario: businessDays };
  }

  function parseHours(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number') return value;
    const n = parseFloat(value.toString().trim().replace(',', '.'));
    return isNaN(n) ? 0 : n;
  }

  function initAgg() {
    const tiposObj = {}, osSets = {}, osListSets = {}, porDia = {};
    TIPO_KEYS.forEach(k => { tiposObj[k] = 0; osSets[k] = {}; osListSets[k] = {}; });
    return {
      horas_oper: 0, horas_manut: 0, horas_total: 0,
      tipos: tiposObj, porDia: porDia,
      osPorTipo: osSets, osListPorTipo: osListSets,
      totalOsSet: {}, diasSet: {}
    };
  }

  function updateAgg(agg, typeKey, dayKey, hours, osId) {
    if (typeKey === 'OPER') agg.horas_oper += hours;
    else agg.horas_manut += hours;
    agg.horas_total += hours;
    agg.tipos[typeKey] = (agg.tipos[typeKey] || 0) + hours;
    if (!agg.porDia[dayKey]) {
      const obj = { dia: dayKey };
      TIPO_KEYS.forEach(k => { obj[k] = 0; });
      obj.oper_total = 0; obj.manut_total = 0;
      agg.porDia[dayKey] = obj;
    }
    agg.porDia[dayKey][typeKey] = (agg.porDia[dayKey][typeKey] || 0) + hours;
    if (typeKey === 'OPER') agg.porDia[dayKey].oper_total += hours;
    else agg.porDia[dayKey].manut_total += hours;
    if (osId) {
      if (!agg.osPorTipo[typeKey]) agg.osPorTipo[typeKey] = {};
      if (!agg.osListPorTipo[typeKey]) agg.osListPorTipo[typeKey] = {};
      agg.osPorTipo[typeKey][osId] = true;
      agg.osListPorTipo[typeKey][osId] = true;
      agg.totalOsSet[osId] = true;
    }
    agg.diasSet[dayKey] = true;
  }

  function finalizeAgg(agg) {
    const result = {
      horas_oper:     parseFloat(agg.horas_oper.toFixed(2)),
      horas_manut:    parseFloat(agg.horas_manut.toFixed(2)),
      horas_total:    parseFloat(agg.horas_total.toFixed(2)),
      tipos:          {},
      por_dia:        [],
      os_por_tipo:    {},
      os_lists:       {},
      total_os:       0,
      dias_trabalhados: 0
    };
    TIPO_KEYS.forEach(k => {
      result.tipos[k] = parseFloat((agg.tipos[k] || 0).toFixed(2));
    });
    Object.keys(agg.porDia).sort().forEach(dk => {
      const obj = agg.porDia[dk];
      TIPO_KEYS.forEach(k => { if (obj[k] === undefined) obj[k] = 0; });
      obj.oper_total  = parseFloat((obj.oper_total  || 0).toFixed(2));
      obj.manut_total = parseFloat((obj.manut_total || 0).toFixed(2));
      result.por_dia.push(obj);
    });
    TIPO_KEYS.forEach(k => {
      result.os_por_tipo[k] = Object.keys(agg.osPorTipo[k] || {}).length;
      // Lista real de OS — numérica quando possível
      result.os_lists[k] = Object.keys(agg.osListPorTipo[k] || {}).map(s => {
        const n = parseFloat(s);
        return isNaN(n) ? s : n;
      }).sort((a, b) => {
        if (typeof a === 'number' && typeof b === 'number') return a - b;
        return String(a).localeCompare(String(b));
      });
    });
    result.total_os          = Object.keys(agg.totalOsSet).length;
    result.dias_trabalhados  = Object.keys(agg.diasSet).length;
    return result;
  }

  // ---- Leitura ----
  const horasSheetId  = ensureSpreadsheet(HORAS_FILE_ID);
  const statusSheetId = ensureSpreadsheet(STATUS_FILE_ID);
  const horasData  = readSheet(horasSheetId);
  const statusData = readSheet(statusSheetId);

  // Índices de colunas — horas
  const horasHeader = horasData[0].map(h => h.toString().trim());
  const idxOp    = horasHeader.indexOf(HORAS_COLS.operador);
  const idxTp    = horasHeader.indexOf(HORAS_COLS.tipoServico);
  const idxOs    = horasHeader.indexOf(HORAS_COLS.os);
  const idxDt    = horasHeader.indexOf(HORAS_COLS.data);
  const idxHr    = horasHeader.indexOf(HORAS_COLS.horas);
  if (idxOp < 0 || idxTp < 0 || idxOs < 0 || idxDt < 0 || idxHr < 0) {
    throw new Error('Colunas não encontradas em horas.xls. Cabeçalho: ' + horasHeader.join(' | '));
  }

  // Índices de colunas — status
  const statusHeader = statusData[0].map(h => h.toString().trim());
  const idxOsSt = statusHeader.indexOf(STATUS_COLS.os);
  const idxDeSt = statusHeader.indexOf(STATUS_COLS.desc);
  const idxLoSt = statusHeader.indexOf(STATUS_COLS.local);
  const idxStSt = statusHeader.indexOf(STATUS_COLS.status);

  // Dicionário OS → detalhes
  const osDetails = {};
  for (let r = 1; r < statusData.length; r++) {
    const row = statusData[r];
    if (idxOsSt < 0 || !row[idxOsSt]) continue;
    const key = row[idxOsSt].toString().trim();
    if (!key) continue;
    osDetails[key] = {
      desc:   idxDeSt >= 0 ? String(row[idxDeSt] || '').trim() : '—',
      local:  idxLoSt >= 0 ? String(row[idxLoSt] || '').trim() : '—',
      status: idxStSt >= 0 ? String(row[idxStSt] || '').trim() : '—'
    };
  }

  // Processamento das horas
  const mesesSet = {}, operadoresSet = {};
  const dadosPorMes = {}, dadosAnual = {};

  for (let r = 1; r < horasData.length; r++) {
    const row = horasData[r];
    if (!row[idxOp] || !row[idxTp] || !row[idxDt] || !row[idxHr]) continue;
    const operador = normalizeName(row[idxOp]);
    const tipo     = row[idxTp].toString().trim().toUpperCase();
    if (TIPO_KEYS.indexOf(tipo) < 0) continue;
    const horasNum = parseHours(row[idxHr]);
    if (horasNum <= 0) continue;
    const dateObj = parseDate(row[idxDt]);
    if (!dateObj) continue;
    const mesKey      = Utilities.formatDate(dateObj, 'GMT', 'yyyy-MM');
    const diaKey      = Utilities.formatDate(dateObj, 'GMT', 'dd');
    const anualDiaKey = Utilities.formatDate(dateObj, 'GMT', 'MM-dd');
    const osId        = row[idxOs] ? row[idxOs].toString().trim() : '';

    operadoresSet[operador] = true;
    mesesSet[mesKey] = true;

    if (!dadosPorMes[mesKey]) dadosPorMes[mesKey] = {};
    if (!dadosPorMes[mesKey][operador]) dadosPorMes[mesKey][operador] = initAgg();
    if (!dadosPorMes[mesKey]['_ALL'])   dadosPorMes[mesKey]['_ALL']   = initAgg();
    if (!dadosAnual[operador])          dadosAnual[operador]           = initAgg();
    if (!dadosAnual['_ALL'])            dadosAnual['_ALL']             = initAgg();

    updateAgg(dadosPorMes[mesKey][operador], tipo, diaKey,      horasNum, osId);
    updateAgg(dadosPorMes[mesKey]['_ALL'],   tipo, diaKey,      horasNum, osId);
    updateAgg(dadosAnual[operador],          tipo, anualDiaKey, horasNum, osId);
    updateAgg(dadosAnual['_ALL'],            tipo, anualDiaKey, horasNum, osId);
  }

  // Montar estrutura final
  const dadosFinal = {};
  const mesesList  = Object.keys(mesesSet).sort();

  mesesList.forEach(ym => {
    const monthAggs  = dadosPorMes[ym];
    const finalMonth = {};
    const [yrStr, moStr] = ym.split('-');
    finalMonth['_meta'] = calcMeta(parseInt(yrStr, 10), parseInt(moStr, 10) - 1);
    for (const op in monthAggs) {
      finalMonth[op] = finalizeAgg(monthAggs[op]);
    }
    dadosFinal[ym] = finalMonth;
  });

  const anualFinal = {};
  for (const op in dadosAnual) anualFinal[op] = finalizeAgg(dadosAnual[op]);
  dadosFinal['ANUAL'] = anualFinal;

  // Limpeza
  toCleanup.forEach(id => {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  });

  return {
    operadores:  Object.keys(operadoresSet).sort(),
    meses:       mesesList,
    tipos_oper:  TIPOS_OPER,
    tipos_manut: TIPOS_MANUT,
    dados:       dadosFinal,
    os_details:  osDetails,
    _gerado_em:  new Date().toISOString()
  };
}
