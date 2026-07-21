// ═══════════════════════════════════════════════════════
// Your Vacation — Qualificador de Captação
// Google Apps Script Backend — v8
// Cole este código em: Planilha > Extensões > Apps Script
// Depois: Implantar > Gerenciar implantações > Editar > Nova versão > Implantar
// ═══════════════════════════════════════════════════════

const SALA_SHEETS = {
  'Alta Vista':       'Alta Vista',
  'Atrium':           'Alta Vista',
  'Marina':           'Alta Vista',
  'Toulon':           'Toulon',       // aba dedicada (regras iguais ao Alta Vista no frontend)
  'Externo':          'Alta Vista',
  'Externo CN':       'Alta Vista',
  'Thermas SP':       'São Pedro',
  'SPTR':             'São Pedro',
  'São Pedro':        'São Pedro',   // nome de aba legado
  'São Pedro Resort': 'São Pedro',   // legado
  'Externo SP':       'São Pedro',
  'Atibaia':          'Atibaia',
  'Porta da sala':    'Atibaia',
  'Vest Casa':        'Atibaia',
  'Entrada Outlet':   'Atibaia',
  'Corredor Outlet':  'Atibaia',
};
const DEFAULT_SHEET = 'Alta Vista'; // fallback — evita criação de abas novas

const COLUNAS = [
  'ID','Data','Hora','Captador','Sala','Nome','Telefone','Cidade','Resultado','Tipo','Renda','Modo','EmSala','TipoSala','Venda',
  // Campos detalhados das respostas
  'PontoCaptacao','IdadeTitular','ProfissaoTitular','IdadeConjuge','ProfissaoConjuge',
  'Carro','Casa','Cartao','Viagens','Modalidade',
  // v2.21 — input numérico do carro (ano de 4 dígitos)
  'AnoCarro',
  // v2.23 — pergunta sobre apresentação de multipropriedade
  'ConheceuMultiprop',
];
const N_COLS  = COLUNAS.length;
const COL_WIDTHS = [180,100,80,150,150,150,130,130,90,150,130,80,80,100,80,150,120,150,120,150,120,80,80,80,120,100,140];

function aplicarCabecalho(sheet) {
  const header = sheet.getRange(1, 1, 1, N_COLS);
  header.setValues([COLUNAS]);
  header.setBackground('#1a2340');
  header.setFontColor('#c9a84c');
  header.setFontWeight('bold');
  sheet.setFrozenRows(1);
  COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function getOrCreateSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    aplicarCabecalho(sheet);
    return sheet;
  }
  const cabecalhoAtual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cabecalhoCorreto = COLUNAS.every((col, i) => cabecalhoAtual[i] === col);
  if (!cabecalhoCorreto) {
    sheet.getRange(1, 1, 1, N_COLS).setValues([COLUNAS]);
    const header = sheet.getRange(1, 1, 1, N_COLS);
    header.setBackground('#1a2340');
    header.setFontColor('#c9a84c');
    header.setFontWeight('bold');
    sheet.setFrozenRows(1);
    COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

function resolverSheet(sala) {
  return getOrCreateSheetByName(SALA_SHEETS[sala] || DEFAULT_SHEET);
}

function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const pos = ids.map(String).indexOf(String(id));
  return pos >= 0 ? pos + 2 : -1;
}

// ── PI — armazenado em PropertiesService (invisível na planilha) ──
function getPIs() {
  const raw = PropertiesService.getScriptProperties().getProperty('pi_store');
  return raw ? JSON.parse(raw) : [];
}
function savePIs(arr) {
  // PropertiesService suporta até 500KB por propriedade — suficiente para meses de PI
  PropertiesService.getScriptProperties().setProperty('pi_store', JSON.stringify(arr));
}

// ── DDP — armazenado na aba "DDP" do Sheets ──
const DDP_SHEET_NAME = 'DDP';
const DDP_COLS = ['hotel','dateISO','ciPool','ciCot','ciConv','ciGrupos','salaPool','salaCot','salaConv','vendPool','vendCot','vendConv','visitantes'];

function getDDPSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DDP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DDP_SHEET_NAME);
    const header = sheet.getRange(1, 1, 1, DDP_COLS.length);
    header.setValues([DDP_COLS]);
    header.setBackground('#1a2340');
    header.setFontColor('#c9a84c');
    header.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // hotel
    sheet.setColumnWidth(2, 120); // dateISO
    for (let c = 3; c <= DDP_COLS.length; c++) sheet.setColumnWidth(c, 90);
  }
  return sheet;
}

function getDDPStore() {
  const sheet = getDDPSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, DDP_COLS.length).getValues();

  function _normDate(v) {
    if (!v) return '';
    if (v instanceof Date) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, '0');
      const d = String(v.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + d;
    }
    const s = String(v).trim().replace(/^DDP_/, '');
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const dt = new Date(s);
    if (!isNaN(dt)) {
      return dt.getFullYear() + '-' + String(dt.getMonth()+1).padStart(2,'0') + '-' + String(dt.getDate()).padStart(2,'0');
    }
    return s;
  }

  return rows
    .filter(r => r[0] && r[1]) // hotel e dateISO preenchidos
    .map(r => {
      const obj = {};
      DDP_COLS.forEach((col, i) => {
        obj[col] = (i >= 2) ? (Number(r[i]) || 0) : String(r[i] || '').trim();
      });
      obj.dateISO = _normDate(r[1]);
      return obj;
    })
    .filter(r => r.dateISO); // descarta linhas com data inválida
}

function saveDDPEntry(entry) {
  const sheet = getDDPSheet();
  const lastRow = sheet.getLastRow();
  const dateISO = String(entry.dateISO || '').replace(/^DDP_/, '');

  // Normaliza valor de data vindo do Sheets (pode ser Date object ou string)
  function _normSheetDate(v) {
    if (!v) return '';
    if (v instanceof Date) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, '0');
      const d = String(v.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + d;
    }
    const s = String(v).trim().replace(/^DDP_/, '');
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const dt = new Date(s);
    if (!isNaN(dt)) {
      const y = dt.getFullYear();
      const mo = String(dt.getMonth() + 1).padStart(2, '0');
      const d = String(dt.getDate()).padStart(2, '0');
      return y + '-' + mo + '-' + d;
    }
    return s;
  }

  // Busca linha existente com mesmo hotel+dateISO
  let targetRow = -1;
  if (lastRow >= 2) {
    const hotelCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const dateCol  = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    for (let i = 0; i < hotelCol.length; i++) {
      if (String(hotelCol[i]) === entry.hotel && _normSheetDate(dateCol[i]) === dateISO) {
        targetRow = i + 2;
        break;
      }
    }
  }

  // Monta rowData forçando dateISO como string (apóstrofo evita parse automático pelo Sheets)
  const rowData = DDP_COLS.map(col => {
    if (col === 'dateISO') return dateISO; // string pura
    if (col === 'hotel')   return entry.hotel || '';
    return entry[col] !== undefined ? Number(entry[col]) : 0;
  });

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, DDP_COLS.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  // Força formato texto na célula dateISO para evitar conversão futura
  const writtenRow = targetRow > 0 ? targetRow : sheet.getLastRow();
  sheet.getRange(writtenRow, 2).setNumberFormat('@STRING@');
}

// ── CRUZAMENTO CHECK-INS × DDP ──
// Aba 'Checkins' (preenchida manualmente): Data | Hotel | Pool | Conv
//   Hotel deve ser 'Alta Vista', 'Atrium' ou 'Marina' (Caldas Novas)
// Cruza com a aba 'DDP' (gerada automaticamente) por data + hotel
// Grava o resultado em 'Cruzamento' (cria se não existir)
const CRUZAMENTO_SHEET = 'Cruzamento';
const CRUZAMENTO_COLS = ['dateISO','hotel','ciPool','ciCot','checkinsPool','checkinsConv','penetracao','atualizadoEm'];
function getCruzamentoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CRUZAMENTO_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CRUZAMENTO_SHEET);
    const header = sheet.getRange(1, 1, 1, CRUZAMENTO_COLS.length);
    header.setValues([CRUZAMENTO_COLS]);
    header.setBackground('#1a2340').setFontColor('#c9a84c').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 110);
    sheet.setColumnWidth(7, 110);
    sheet.setColumnWidth(8, 160);
  }
  return sheet;
}

// Lê a aba 'Checkins' e retorna dict {(dateISO, hotel): {pool, conv}}
function lerCheckins() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Checkins');
  if (!sheet) return {};
  const last = sheet.getLastRow();
  if (last < 2) return {};
  const values = sheet.getRange(2, 1, last - 1, 4).getValues();
  const map = {};
  values.forEach(r => {
    const data = String(r[0] || '').trim();
    const hotel = String(r[1] || '').trim();
    const pool = Number(r[2]) || 0;
    const conv = Number(r[3]) || 0;
    if (data && hotel) {
      const key = data + '|' + hotel;
      map[key] = { pool, conv };
    }
  });
  return map;
}

function atualizarCruzamentoCheckins() {
  const checkins = lerCheckins();
  if (Object.keys(checkins).length === 0) {
    Logger.log('⚠️  Aba "Checkins" vazia ou não existe. Preencha ela primeiro (Data | Hotel | Pool | Conv).');
    return;
  }
  const ddp = getDDPStore();
  const sheet = getCruzamentoSheet();
  // Apaga dados antigos
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, CRUZAMENTO_COLS.length).clearContent();
  // Monta saída
  const out = [];
  const agora = new Date();
  Object.keys(checkins).forEach(key => {
    const [data, hotel] = key.split('|');
    const ddpEntry = ddp.find(d => d.dateISO === data && d.hotel === hotel);
    const ciPool  = ddpEntry ? Number(ddpEntry.ciPool  || 0) : 0;
    const ciCot   = ddpEntry ? Number(ddpEntry.ciCot   || 0) : 0;
    const c       = checkins[key];
    const totalCI = ciPool + ciCot;
    const totalCk = c.pool + c.conv;
    const penetracao = totalCI > 0 ? Math.round(totalCk / totalCI * 1000) / 10 : 0;
    out.push([
      data, hotel,
      ciPool, ciCot,
      c.pool, c.conv,
      penetracao,
      Utilities.formatDate(agora, 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss')
    ]);
  });
  // Ordena por data desc
  out.sort((a, b) => String(b[0]).localeCompare(String(a[0])));
  if (out.length > 0) {
    sheet.getRange(2, 1, out.length, CRUZAMENTO_COLS.length).setValues(out);
  }
  Logger.log('Cruzamento atualizado: ' + out.length + ' linhas');
}

// Cria gatilho diário que roda às 22h (atualiza o cruzamento automaticamente)
function criarGatilhoCruzamento() {
  // Remove gatilhos antigos desta função
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'atualizarCruzamentoCheckins') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('atualizarCruzamentoCheckins')
    .timeBased()
    .atHour(22)
    .everyDays(1)
    .create();
  Logger.log('Gatilho diário criado: atualiza o cruzamento às 22h');
}

// ── ADICIONAR COLUNAS FALTANTES — sem sobrescrever dados ──
// Adiciona na aba especificada (ou em todas as abas da planilha) as colunas
// que estão no array COLUNAS mas faltam no cabeçalho da aba.
// Uso: executar uma vez após reimplantar o backend, para garantir que
// abas existentes tenham todas as 27 colunas antes do primeiro lead novo.
function adicionarColunasFaltantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsAlvo = ['Alta Vista', 'São Pedro', 'Atibaia', 'Thermas SP (Parque)'];
  let totalAdicoes = 0;
  sheetsAlvo.forEach(nomeAba => {
    const sheet = ss.getSheetByName(nomeAba);
    if (!sheet) return;
    const headerAtual = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    const faltam = COLUNAS.filter(col => headerAtual.indexOf(col) < 0);
    if (faltam.length === 0) {
      Logger.log('[' + nomeAba + '] já tem todas as ' + COLUNAS.length + ' colunas');
      return;
    }
    const startCol = sheet.getLastColumn() + 1;
    const newHeaders = faltam.map((col, i) => [col]);
    sheet.getRange(1, startCol, faltam.length, 1).setValues(newHeaders);
    // Aplica cor de cabeçalho pra ficar consistente
    const range = sheet.getRange(1, startCol, faltam.length, 1);
    range.setBackground('#1a2340').setFontColor('#c9a84c').setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('[' + nomeAba + '] adicionadas ' + faltam.length + ' colunas: ' + faltam.join(', '));
    totalAdicoes += faltam.length;
  });
  Logger.log('===== TOTAL: ' + totalAdicoes + ' colunas adicionadas em todas as abas =====');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // saveDDP: salva/atualiza registro de DDP na aba DDP do Sheets
    if (data.action === 'saveDDP') {
      let dateISO = data.dateISO || '';
      if (dateISO.startsWith('DDP_')) dateISO = dateISO.slice(4);
      const hotel = data.hotel || '';
      if (!hotel || !dateISO) return jsonResponse({ status: 'error', message: 'hotel e dateISO obrigatórios' });
      const entry = {
        hotel, dateISO,
        ciPool:     Number(data.ciPool)     || 0,
        ciCot:      Number(data.ciCot)      || 0,
        ciConv:     Number(data.ciConv)     || 0,
        ciGrupos:   Number(data.ciGrupos)   || 0,
        salaPool:   Number(data.salaPool)   || 0,
        salaCot:    Number(data.salaCot)    || 0,
        salaConv:   Number(data.salaConv)   || 0,
        vendPool:   Number(data.vendPool)   || 0,
        vendCot:    Number(data.vendCot)    || 0,
        vendConv:   Number(data.vendConv)   || 0,
        visitantes: Number(data.visitantes) || 0,
      };
      saveDDPEntry(entry);
      return jsonResponse({ status: 'ok' });
    }

    // getDDP via POST (alternativa ao GET)
    if (data.action === 'getDDP') {
      const hotel = data.hotel || '';
      const store = getDDPStore();
      const rows = hotel ? store.filter(r => r.hotel === hotel) : store;
      return jsonResponse({ status: 'ok', rows });
    }

    // dedup: não mais necessário (Sheets já garante upsert por hotel+dateISO), mantido por compatibilidade
    if (data.action === 'dedup') {
      return jsonResponse({ status: 'ok', removed: 0, message: 'DDP agora usa aba no Sheets — dedup automático no upsert' });
    }

    // update_sala: atualiza EmSala, TipoSala e Venda de um lead existente
    if (data._action === 'update_sala') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetsAlvo = Object.values(SALA_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean);
      for (const sheet of sheetsAlvo) {
        const row = findRowById(sheet, data.id);
        if (row > 0) {
          const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const set = (col, val) => {
            const i = cabecalho.indexOf(col);
            if (i >= 0) sheet.getRange(row, i + 1).setValue(val);
          };
          set('EmSala',   data.emSala   ? 'SIM' : '');
          set('TipoSala', data.tipoSala || '');
          set('Venda',    data.venda    ? 'SIM' : '');
          if (data.sala !== undefined) set('Sala', data.sala);
          return jsonResponse({ status: 'ok', id: data.id });
        }
      }
      return jsonResponse({ status: 'not_found', id: data.id });
    }

    // register_pi: salva em PropertiesService (sem criar aba na planilha)
    if (data._action === 'register_pi') {
      const pis = getPIs();
      if (pis.some(p => String(p.id) === String(data.id))) {
        return jsonResponse({ status: 'duplicate', id: data.id });
      }
      pis.unshift({
        id:       String(data.id),
        date:     data.date     || '',
        time:     data.time     || '',
        captador: data.captador || '',
        sala:     data.sala     || '',
        nomeParc: data.nomeParc || '',
        dateISO:  data.dateISO  || '',
      });
      // Mantém últimos 1000 PIs
      if (pis.length > 1000) pis.pop();
      savePIs(pis);
      return jsonResponse({ status: 'ok', id: data.id });
    }

    // novo lead — salva na aba da sala correspondente
    const sheet = resolverSheet(data.sala || '');
    const existingRow = findRowById(sheet, data.id);
    if (existingRow > 0) {
      // Lead já existe — sincroniza EmSala, TipoSala e Venda se presentes no request
      if (data.emSala || data.tipoSala || data.venda) {
        const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const set = (col, val) => {
          const i = cabecalho.indexOf(col);
          if (i >= 0) sheet.getRange(existingRow, i + 1).setValue(val);
        };
        if (data.emSala)   set('EmSala',   'SIM');
        if (data.tipoSala) set('TipoSala', data.tipoSala);
        if (data.venda)    set('Venda',    'SIM');
      }
      return jsonResponse({ status: 'duplicate', id: data.id });
    }

    const ans = data._ans || {};
    // Normaliza data para DD/MM/YYYY fixo — evita que o Sheets converta para Date object
    function _fmtDate(v) {
      if (!v) return '';
      const s = String(v).trim();
      // Já está em DD/MM/YYYY
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
      // YYYY-MM-DD
      const m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m1) return m1[3] + '/' + m1[2] + '/' + m1[1];
      // Qualquer outro formato JS Date string
      const dt = new Date(s);
      if (!isNaN(dt)) {
        return String(dt.getDate()).padStart(2,'0') + '/' +
               String(dt.getMonth()+1).padStart(2,'0') + '/' + dt.getFullYear();
      }
      return s;
    }
    const dataFmt = _fmtDate(data.date);
    sheet.appendRow([
      data.id || '', dataFmt, "'" + (data.time || ''),
      data.captador || '', data.sala || '', data.nome || '',
      data.tel || '', data.cidade || '', data.verdict || '',
      data.tipo || '', data.renda || '', data.mode || '',
      data.emSala   ? 'SIM' : '',          // EmSala
      data.tipoSala || '',                 // TipoSala
      data.venda    ? 'SIM' : '',          // Venda
      // Campos detalhados
      data.sala || '',                     // PontoCaptacao (mesmo que Sala — filtrável)
      ans.idadeMarido      || '',          // IdadeTitular
      ans.profissaoTitular || '',          // ProfissaoTitular
      ans.idadeConjuge     || '',          // IdadeConjuge
      ans.profissaoConjuge || '',          // ProfissaoConjuge
      ans.carro            || '',          // Carro (faixa — retrocompat com regras)
      ans.casa             || '',          // Casa
      ans.cartao           || '',          // Cartao
      ans.viagens          || '',          // Viagens
      data.modalidade      || '',          // Modalidade (pool/convidados/proprietarios)
      // v2.21 — input numérico do carro
      (typeof ans.anoCarro === 'number' ? ans.anoCarro : (ans.anoCarro || '')),  // AnoCarro
      // v2.23 — apresentação prévia de multipropriedade
      ans.conheceuMultiprop || '',         // ConheceuMultiprop
    ]);

    // Aplica cor na linha recém inserida e trava formato da coluna Data como texto
    const row = sheet.getLastRow();
    const nCols = sheet.getLastColumn();
    const idxData = COLUNAS.indexOf('Data');
    if (idxData >= 0) sheet.getRange(row, idxData + 1).setNumberFormat('@STRING@');
    const range = sheet.getRange(row, 1, 1, nCols);
    if (data.verdict === 'Q')       range.setBackground('#d4edda');
    if (data.verdict === 'PARCIAL') range.setBackground('#fff3cd');
    if (data.verdict === 'NQ')      range.setBackground('#f8d7da');

    return jsonResponse({ status: 'ok', id: data.id });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const params = e && e.parameter ? e.parameter : {};
    const salaFiltro = params.sala || null;

    // ?_action=update_sala&id=...&sala=... — atualiza coluna Sala via GET
    if (params._action === 'update_sala') {
      const sheetsAlvo = Object.values(SALA_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean);
      for (const sheet of sheetsAlvo) {
        const row = findRowById(sheet, params.id);
        if (row > 0) {
          const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const set = (col, val) => {
            const i = cabecalho.indexOf(col);
            if (i >= 0) sheet.getRange(row, i + 1).setValue(val);
          };
          if (params.sala !== undefined) set('Sala', params.sala);
          if (params.emSala !== undefined) set('EmSala', params.emSala === 'true' ? 'SIM' : '');
          if (params.tipoSala !== undefined) set('TipoSala', params.tipoSala);
          if (params.venda !== undefined) set('Venda', params.venda === 'true' ? 'SIM' : '');
          return jsonResponse({ status: 'ok', id: params.id });
        }
      }
      return jsonResponse({ status: 'not_found', id: params.id });
    }

    // ?_type=pi — retorna pesquisas incompletas do PropertiesService
    if (params._type === 'pi') {
      return jsonResponse({ status: 'ok', pis: getPIs() });
    }

    // ?action=getAuditores — retorna CSV da planilha de auditores como JSON
    if (params.action === 'getAuditores') {
      try {
        const AUDITORES_SS_ID = '1rNvoycg3S6PdIyVZtfECVEsBTobJXzeX';
        const AUDITORES_GID   = 1550622683;
        const auditSS   = SpreadsheetApp.openById(AUDITORES_SS_ID);
        const sheets    = auditSS.getSheets();
        let auditSheet  = sheets.find(s => s.getSheetId() === AUDITORES_GID) || sheets[0];
        const allValues = auditSheet.getDataRange().getValues();
        // Converter para array de arrays de strings (igual ao CSV)
        const rows = allValues.map(row => row.map(cell => {
          if (cell instanceof Date) {
            // Formatar como DD/MM/YYYY
            const d = cell.getDate(), m = cell.getMonth()+1, y = cell.getFullYear();
            return (d<10?'0'+d:d)+'/'+(m<10?'0'+m:m)+'/'+y;
          }
          return cell === null || cell === undefined ? '' : String(cell);
        }));
        return jsonResponse({ status: 'ok', rows });
      } catch(err) {
        return jsonResponse({ status: 'error', message: err.message });
      }
    }

    // ?action=getDDP&hotel=SPTR — retorna registros DDP
    if (params.action === 'getDDP') {
      const hotel = params.hotel || '';
      const store = getDDPStore();
      const rows = hotel ? store.filter(r => r.hotel === hotel) : store;
      return jsonResponse({ status: 'ok', rows });
    }

    // ?action=saveDDP&hotel=...&dateISO=...&ciPool=...  — salva via GET (POST não sobrevive ao redirect)
    if (params.action === 'saveDDP') {
      const hotel   = params.hotel   || '';
      const dateISO = (params.dateISO || '').replace(/^DDP_/, '');
      if (!hotel || !dateISO) return jsonResponse({ status: 'error', message: 'hotel e dateISO obrigatórios' });
      const entry = {
        hotel, dateISO,
        ciPool:     Number(params.ciPool)     || 0,
        ciCot:      Number(params.ciCot)      || 0,
        ciConv:     Number(params.ciConv)     || 0,
        ciGrupos:   Number(params.ciGrupos)   || 0,
        salaPool:   Number(params.salaPool)   || 0,
        salaCot:    Number(params.salaCot)    || 0,
        salaConv:   Number(params.salaConv)   || 0,
        vendPool:   Number(params.vendPool)   || 0,
        vendCot:    Number(params.vendCot)    || 0,
        vendConv:   Number(params.vendConv)   || 0,
        visitantes: Number(params.visitantes) || 0,
      };
      saveDDPEntry(entry);
      return jsonResponse({ status: 'ok' });
    }

    const sheetsAlvo = salaFiltro
      ? [ss.getSheetByName(SALA_SHEETS[salaFiltro] || salaFiltro)].filter(Boolean)
      : Object.values(SALA_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean);

    // ?dateISO=2026-04-24 filtra por dia exato; ?dateFrom=...&dateTo=... filtra por range
    const dateISOFiltro = params.dateISO || null;
    const dateFrom = params.dateFrom || null;
    const dateTo   = params.dateTo   || null;

    // Normaliza qualquer formato de data para YYYY-MM-DD
    function _normRowDate(v) {
      if (!v) return '';
      if (v instanceof Date) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
      }
      const s = String(v).trim();
      // Já está em YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      // DD/MM/YYYY
      const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m1) return m1[3] + '-' + m1[2] + '-' + m1[1];
      // "Thu Apr 24 2026 ..." ou qualquer formato JS Date string
      const dt = new Date(s);
      if (!isNaN(dt)) {
        const y = dt.getFullYear();
        const mo = String(dt.getMonth() + 1).padStart(2, '0');
        const d = String(dt.getDate()).padStart(2, '0');
        return y + '-' + mo + '-' + d;
      }
      return s;
    }

    const leads = [];
    sheetsAlvo.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const idx = {};
      COLUNAS.forEach(col => { idx[col] = cabecalho.indexOf(col); });
      const nCols = sheet.getLastColumn();
      const rows = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
      rows.forEach(r => {
        if (!r[0]) return;
        // Filtro por data — dia exato ou range
        // Leads sem data (rowDateISO='') são sempre incluídos para não perder registros
        if (dateISOFiltro || dateFrom || dateTo) {
          const rowDateISO = idx['Data'] >= 0 ? _normRowDate(r[idx['Data']]) : '';
          if (dateISOFiltro && rowDateISO && rowDateISO !== dateISOFiltro) return;
          if (dateFrom && rowDateISO && rowDateISO < dateFrom) return;
          if (dateTo   && rowDateISO && rowDateISO > dateTo)   return;
        }
        const get = (col) => idx[col] >= 0 ? r[idx[col]] : '';
        leads.push({
          id: String(get('ID')), date: String(get('Data')), time: (function(){ var v = get('Hora'); if (!v) return ''; if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'HH:mm'); return String(v); })(),
          captador: get('Captador'), sala: get('Sala'), nome: get('Nome'),
          tel: get('Telefone'), cidade: get('Cidade'), verdict: get('Resultado'),
          tipo: get('Tipo'), renda: get('Renda'), mode: get('Modo'),
          emSala:   get('EmSala')   === 'SIM',
          tipoSala: get('TipoSala') || null,
          venda:    get('Venda')    === 'SIM',
          _ans: {
            idadeMarido:      get('IdadeTitular')      || null,
            profissaoTitular: get('ProfissaoTitular')  || null,
            idadeConjuge:     get('IdadeConjuge')      || null,
            profissaoConjuge: get('ProfissaoConjuge')  || null,
            carro:            get('Carro')             || null,
            casa:             get('Casa')              || null,
            cartao:           get('Cartao')            || null,
            viagens:          get('Viagens')           || null,
            // v2.21 — input numérico do carro (vazio para leads antigos)
            anoCarro:         (function(){ var v = get('AnoCarro'); if (v === '' || v == null) return null; var n = parseInt(v, 10); return isNaN(n) ? null : n; })(),
            // v2.23 — apresentação prévia de multipropriedade
            conheceuMultiprop: get('ConheceuMultiprop') || null,
          },
        });
      });
    });

    return jsonResponse({ status: 'ok', leads });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function migrarTodasAsAbas() {
  Object.values(SALA_SHEETS).forEach(nome => getOrCreateSheetByName(nome));
  Logger.log('Migração concluída.');
}

// ── AUDITORIA DE ABAS — apenas loga o que tem em cada aba ──
// Rodar manualmente antes de migrarAbasPorSala() para entender o cenário.
function auditarAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('========== AUDITORIA DE ABAS ==========');
  Logger.log('Mapa oficial SALA_SHEETS:');
  Object.keys(SALA_SHEETS).forEach(sala => {
    Logger.log('  ' + sala + ' → ' + SALA_SHEETS[sala]);
  });
  Logger.log('');
  Logger.log('Abas presentes na planilha:');
  const sheets = ss.getSheets();
  sheets.forEach(s => {
    const last = s.getLastRow();
    if (last < 2) {
      Logger.log('  [' + s.getName() + '] vazia (ou só cabeçalho)');
      return;
    }
    const cabecalho = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    const idxSala = cabecalho.indexOf('Sala');
    if (idxSala < 0) {
      Logger.log('  [' + s.getName() + '] sem coluna Sala — pulando');
      return;
    }
    // Conta distribuição de salas
    const nCols = s.getLastColumn();
    const rows = s.getRange(2, 1, last - 1, nCols).getValues();
    const contagem = {};
    rows.forEach(r => {
      const sala = String(r[idxSala] || '').trim() || '(vazio)';
      contagem[sala] = (contagem[sala] || 0) + 1;
    });
    Logger.log('  [' + s.getName() + '] ' + rows.length + ' linhas');
    Object.keys(contagem).sort().forEach(sala => {
      const esperado = SALA_SHEETS[sala] || '???';
      const flag = esperado === '???' ? ' ⚠️  NÃO MAPEADA' : (esperado !== s.getName() ? ' ⚠️  ABA ERRADA (esperado: ' + esperado + ')' : '');
      Logger.log('      - ' + sala + ': ' + contagem[sala] + flag);
    });
  });
  Logger.log('========== FIM AUDITORIA ==========');
}

// ── AUDITORIA DDP — diagnostica a taxa de marcação de EmSala / TipoSala / Venda ──
// Roda uma vez para entender quanto do DDP automático está faltando.
// Considera apenas abas que alimentam DDP (Alta Vista, Atrium, Marina, SPTR, Thermas SP)
function auditarDDP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SALAS_DDP = ['Alta Vista', 'Atrium', 'Marina', 'SPTR', 'Thermas SP'];
  // Calcula janela: últimos 30 dias
  const hoje = new Date();
  const de30 = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
  const isoDe = de30.toISOString().slice(0, 10);
  const isoAte = hoje.toISOString().slice(0, 10);

  Logger.log('========== AUDITORIA DDP ==========');
  Logger.log('Janela: ' + isoDe + ' a ' + isoAte + ' (últimos 30 dias)');
  Logger.log('');

  const totais = { total: 0, emSala: 0, semEmSala: 0, tipoPool: 0, tipoCot: 0, tipoConv: 0, semTipo: 0, venda: 0 };
  const porSala = {};

  SALAS_DDP.forEach(salaNome => {
    // O SALA_SHEETS mapeia tudo — para SPTR/Thermas SP, fica na aba 'São Pedro'
    const abaNome = salaNome === 'SPTR' || salaNome === 'Thermas SP' ? 'São Pedro' : 'Alta Vista';
    const sheet = ss.getSheetByName(abaNome);
    if (!sheet) { Logger.log('⚠️  Aba "' + abaNome + '" não encontrada'); return; }
    const last = sheet.getLastRow();
    if (last < 2) return;
    const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxData = cabecalho.indexOf('Data');
    const idxSala = cabecalho.indexOf('Sala');
    const idxEmSala = cabecalho.indexOf('EmSala');
    const idxTipoSala = cabecalho.indexOf('TipoSala');
    const idxVenda = cabecalho.indexOf('Venda');
    if (idxData < 0 || idxSala < 0) return;
    const nCols = sheet.getLastColumn();
    const rows = sheet.getRange(2, 1, last - 1, nCols).getValues();

    const s = { total: 0, emSala: 0, semEmSala: 0, tipoPool: 0, tipoCot: 0, tipoConv: 0, semTipo: 0, venda: 0 };
    rows.forEach(r => {
      // Filtra por sala (considera nomes legados também)
      const sala = String(r[idxSala] || '').trim();
      const salaNormalizada = (sala === 'São Pedro Resort') ? 'SPTR' : sala;
      if (salaNormalizada !== salaNome) return;
      // Filtra por data
      const dataStr = String(r[idxData] || '').trim();
      const isoMatch = dataStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!isoMatch) return;
      const dateISO = isoMatch[3] + '-' + isoMatch[2] + '-' + isoMatch[1];
      if (dateISO < isoDe || dateISO > isoAte) return;
      s.total++;
      const emSala = String(r[idxEmSala] || '').trim().toUpperCase() === 'SIM';
      if (emSala) s.emSala++; else s.semEmSala++;
      const tipoSala = String(r[idxTipoSala] || '').trim().toLowerCase();
      if (tipoSala === 'pool') s.tipoPool++;
      else if (tipoSala === 'cotista') s.tipoCot++;
      else if (tipoSala === 'conv' || tipoSala === 'convidado') s.tipoConv++;
      else s.semTipo++;
      const venda = String(r[idxVenda] || '').trim().toUpperCase() === 'SIM';
      if (venda) s.venda++;
    });
    porSala[salaNome] = s;
    totais.total      += s.total;
    totais.emSala     += s.emSala;
    totais.semEmSala  += s.semEmSala;
    totais.tipoPool   += s.tipoPool;
    totais.tipoCot    += s.tipoCot;
    totais.tipoConv   += s.tipoConv;
    totais.semTipo    += s.semTipo;
    totais.venda      += s.venda;
  });

  Logger.log('--- Por sala (últimos 30 dias) ---');
  Object.keys(porSala).forEach(nome => {
    const s = porSala[nome];
    const taxaEmSala  = s.total ? Math.round(s.emSala / s.total * 100) : 0;
    const taxaTipo    = s.emSala ? Math.round((s.tipoPool + s.tipoCot + s.tipoConv) / s.emSala * 100) : 0;
    Logger.log('  [' + nome + '] ' + s.total + ' leads no período');
    Logger.log('      EmSala: ' + s.emSala + ' / ' + s.total + ' (' + taxaEmSala + '%)');
    Logger.log('      Sem EmSala: ' + s.semEmSala + ' (leads que entraram mas não foram marcados como em sala)');
    Logger.log('      TipoSala (só dos marcados em sala): Pool=' + s.tipoPool + ' Cot=' + s.tipoCot + ' Conv=' + s.tipoConv + ' Sem=' + s.semTipo + ' (' + taxaTipo + '% preenchido)');
    Logger.log('      Vendas: ' + s.venda);
  });
  Logger.log('');
  Logger.log('--- TOTAL GERAL ---');
  const taxaEmSalaTotal = totais.total ? Math.round(totais.emSala / totais.total * 100) : 0;
  const taxaTipoTotal   = totais.emSala ? Math.round((totais.tipoPool + totais.tipoCot + totais.tipoConv) / totais.emSala * 100) : 0;
  Logger.log('  Leads no período: ' + totais.total);
  Logger.log('  Com EmSala: ' + totais.emSala + ' (' + taxaEmSalaTotal + '%)');
  Logger.log('  Sem EmSala (gargalo principal): ' + totais.semEmSala);
  Logger.log('  Com TipoSala (dos marcados): ' + (totais.tipoPool + totais.tipoCot + totais.tipoConv) + ' / ' + totais.emSala + ' (' + taxaTipoTotal + '%)');
  Logger.log('  Vendas marcadas: ' + totais.venda);
  Logger.log('');
  Logger.log('--- DIAGNÓSTICO ---');
  if (totais.semEmSala / Math.max(totais.total, 1) > 0.5) {
    Logger.log('  ❌ Mais de 50% dos leads sem EmSala — os captadores não estão marcando (ou o app não está salvando).');
    Logger.log('  Próximo passo: investigar UX do botão "Entrou em sala" no app.');
  } else if (totais.semEmSala / Math.max(totais.total, 1) > 0.2) {
    Logger.log('  ⚠️  20-50% dos leads sem EmSala — problema de adesão parcial.');
    Logger.log('  Próximo passo: backfill dos vazios + melhorias de UX.');
  } else {
    Logger.log('  ✓ Maioria dos leads marcados. Backfill cobre só os antigos.');
  }
  if (totais.emSala > 0 && (totais.tipoPool + totais.tipoCot + totais.tipoConv) / totais.emSala < 0.5) {
    Logger.log('  ❌ Entre os marcados, menos de 50% têm TipoSala.');
  }
  Logger.log('========== FIM AUDITORIA DDP ==========');
}

// ── BACKFILL DDP — seta EmSala=SIM e TipoSala=pool nos leads sem flag ──
// Rodar SOMENTE após auditarDDP() confirmar que vale a pena.
// Afeta apenas os últimos 30 dias. Não toca em leads já marcados.
function backfillEmSala() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SALAS_DDP = ['Alta Vista', 'Atrium', 'Marina', 'SPTR', 'Thermas SP'];
  const hoje = new Date();
  const de30 = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
  const isoDe = de30.toISOString().slice(0, 10);
  const isoAte = hoje.toISOString().slice(0, 10);

  Logger.log('========== BACKFILL DDP ==========');
  Logger.log('Janela: ' + isoDe + ' a ' + isoAte);
  Logger.log('Vai setar EmSala=SIM e TipoSala=pool em leads sem flag.');
  Logger.log('Leads já marcados (EmSala=SIM ou TipoSala preenchido) NÃO serão tocados.');
  Logger.log('');

  let totalAlterados = 0;
  const porSala = {};

  SALAS_DDP.forEach(salaNome => {
    const abaNome = salaNome === 'SPTR' || salaNome === 'Thermas SP' ? 'São Pedro' : 'Alta Vista';
    const sheet = ss.getSheetByName(abaNome);
    if (!sheet) return;
    const last = sheet.getLastRow();
    if (last < 2) return;
    const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxData = cabecalho.indexOf('Data');
    const idxSala = cabecalho.indexOf('Sala');
    const idxEmSala = cabecalho.indexOf('EmSala');
    const idxTipoSala = cabecalho.indexOf('TipoSala');
    if (idxData < 0 || idxSala < 0) return;
    const nCols = sheet.getLastColumn();
    const rows = sheet.getRange(2, 1, last - 1, nCols).getValues();
    const alteracoes = [];
    rows.forEach((r, i) => {
      const sala = String(r[idxSala] || '').trim();
      const salaNormalizada = (sala === 'São Pedro Resort') ? 'SPTR' : sala;
      if (salaNormalizada !== salaNome) return;
      const dataStr = String(r[idxData] || '').trim();
      const isoMatch = dataStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (!isoMatch) return;
      const dateISO = isoMatch[3] + '-' + isoMatch[2] + '-' + isoMatch[1];
      if (dateISO < isoDe || dateISO > isoAte) return;
      const emSalaAtual = String(r[idxEmSala] || '').trim().toUpperCase();
      const tipoSalaAtual = String(r[idxTipoSala] || '').trim();
      // Só altera se AMBOS estiverem vazios (não toca em lead parcialmente marcado)
      if (!emSalaAtual && !tipoSalaAtual) {
        alteracoes.push({ rowIdx: i + 2, captador: r[3], sala, dateISO });
      }
    });
    if (alteracoes.length === 0) {
      Logger.log('  [' + salaNome + '] nenhum lead precisou de backfill');
      return;
    }
    // Aplica as alterações: seta EmSala=SIM e TipoSala=pool nas células
    alteracoes.forEach(a => {
      if (idxEmSala >= 0) sheet.getRange(a.rowIdx, idxEmSala + 1).setValue('SIM');
      if (idxTipoSala >= 0) sheet.getRange(a.rowIdx, idxTipoSala + 1).setValue('pool');
    });
    totalAlterados += alteracoes.length;
    porSala[salaNome] = alteracoes.length;
    Logger.log('  [' + salaNome + '] ' + alteracoes.length + ' leads alterados');
  });

  Logger.log('');
  Logger.log('Total de leads alterados: ' + totalAlterados);
  Logger.log('Distribuição:');
  Object.keys(porSala).forEach(nome => {
    Logger.log('  [' + nome + '] ' + porSala[nome] + ' leads');
  });
  Logger.log('========== FIM BACKFILL ==========');
}

// ── MIGRAÇÃO DE ABAS POR SALA — separa leads mal-posicionados ──
// 1. Identifica a aba oficial de cada operação (a com MAIS LINHAS VÁLIDAS
//    — desconsidera vazias e salas não-mapeadas)
// 2. Move linhas mal-posicionadas para a aba oficial correta baseado na coluna 'Sala'
// 3. Renomeia abas duplicadas (ex: 2x "São Pedro") — a perdedora vira "Nome (legada — vazia)"
// Rodar manualmente UMA VEZ após auditarAbas() confirmar o cenário.
function migrarAbasPorSala() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 1. Identifica a aba oficial de cada operação (a com MAIS LINHAS VÁLIDAS)
  //    Linhas válidas = com sala que existe no SALA_SHEETS
  // Considera só abas cujo nome bate com algum dos valores únicos de SALA_SHEETS
  const valoresUnicos = [...new Set(Object.values(SALA_SHEETS))]; // ['Alta Vista', 'São Pedro', 'Atibaia']
  const contagemPorNome = {};
  valoresUnicos.forEach(nome => { contagemPorNome[nome] = []; });
  sheets.forEach(s => {
    if (contagemPorNome.hasOwnProperty(s.getName())) {
      contagemPorNome[s.getName()].push(s);
    }
  });

  // Helper: conta linhas válidas (sala existe no SALA_SHEETS) de uma sheet
  function _countValidas(sheet) {
    const last = sheet.getLastRow();
    if (last < 2) return 0;
    const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxSala = cabecalho.indexOf('Sala');
    if (idxSala < 0) return 0;
    const nCols = sheet.getLastColumn();
    const rows = sheet.getRange(2, 1, last - 1, nCols).getValues();
    let count = 0;
    rows.forEach(r => {
      const sala = String(r[idxSala] || '').trim();
      if (SALA_SHEETS[sala]) count++;
    });
    return count;
  }

  const abaOficial = {}; // operacao (nome único) → Sheet oficial
  Object.keys(contagemPorNome).forEach(nome => {
    const candidatos = contagemPorNome[nome];
    if (candidatos.length === 0) {
      Logger.log('⚠️  Aba oficial "' + nome + '" não existe — criando vazia');
      abaOficial[nome] = getOrCreateSheetByName(nome);
    } else if (candidatos.length === 1) {
      abaOficial[nome] = candidatos[0];
      Logger.log('✓ Aba oficial "' + nome + '": única existente (' + _countValidas(candidatos[0]) + ' linhas válidas de ' + candidatos[0].getLastRow() + ')');
    } else {
      // Mais de uma — escolhe a com mais linhas VÁLIDAS; renomeia as outras
      candidatos.forEach(c => { c._linhasValidas = _countValidas(c); });
      candidatos.sort((a, b) => b._linhasValidas - a._linhasValidas);
      abaOficial[nome] = candidatos[0];
      Logger.log('✓ Aba oficial "' + nome + '": escolhida com ' + candidatos[0]._linhasValidas + ' linhas válidas (de ' + candidatos[0].getLastRow() + ' totais)');
      for (let i = 1; i < candidatos.length; i++) {
        const duplicada = candidatos[i];
        const novoNome = nome + ' (legada — vazia)';
        try {
          duplicada.setName(novoNome);
          Logger.log('  ↻ Aba duplicada renomeada para "' + novoNome + '" (' + duplicada._linhasValidas + ' válidas de ' + duplicada.getLastRow() + ')');
        } catch (e) {
          Logger.log('  ⚠️  Não consegui renomear aba (já existe "' + novoNome + '"?): ' + e.message);
        }
      }
    }
  });

  // 2. Para cada aba (oficial ou não), move linhas para a aba oficial correta
  let totalMovidas = 0;
  const logMovidas = [];
  sheets.forEach(s => {
    const nomeAba = s.getName();
    if (nomeAba.endsWith('(legada — vazia)')) {
      Logger.log('⏭  Pulando aba legada: ' + nomeAba);
      return;
    }
    if (!valoresUnicos.includes(nomeAba)) {
      Logger.log('⏭  Pulando aba fora do escopo: ' + nomeAba);
      return;
    }
    const last = s.getLastRow();
    if (last < 2) return;
    const cabecalho = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    const idxSala = cabecalho.indexOf('Sala');
    if (idxSala < 0) return;
    const nCols = s.getLastColumn();
    const dados = s.getRange(2, 1, last - 1, nCols).getValues();

    // Coleta linhas que precisam sair desta aba
    const linhasParaMover = []; // { rowIdx (1-based na planilha), sala, destinoAba, valores }
    dados.forEach((r, i) => {
      const sala = String(r[idxSala] || '').trim();
      if (!sala) return;
      const destino = SALA_SHEETS[sala];
      if (!destino) {
        Logger.log('  ⚠️  Linha ' + (i + 2) + ' tem sala "' + sala + '" não mapeada — IGNORANDO');
        return;
      }
      if (destino !== nomeAba) {
        linhasParaMover.push({ rowIdx: i + 2, sala, destinoAba: destino, valores: r });
      }
    });

    if (linhasParaMover.length === 0) {
      Logger.log('✓ Aba "' + nomeAba + '": todas as ' + dados.length + ' linhas estão no lugar certo');
      return;
    }

    Logger.log('⤴  Aba "' + nomeAba + '": ' + linhasParaMover.length + ' linhas para mover');
    linhasParaMover.forEach(m => {
      logMovidas.push({ de: nomeAba, para: m.destinoAba, sala: m.sala, rowIdx: m.rowIdx });
    });
  });

  // 3. Executa as movimentações de baixo para cima em cada aba de origem
  // Agrupa por aba de origem
  const porOrigem = {};
  logMovidas.forEach(m => {
    if (!porOrigem[m.de]) porOrigem[m.de] = [];
    porOrigem[m.de].push(m);
  });
  Object.keys(porOrigem).forEach(nomeOrigem => {
    const sheetOrigem = ss.getSheetByName(nomeOrigem);
    if (!sheetOrigem) return;
    const cabecalho = sheetOrigem.getRange(1, 1, 1, sheetOrigem.getLastColumn()).getValues()[0];
    const idxSala = cabecalho.indexOf('Sala');
    const nCols = sheetOrigem.getLastColumn();
    // Agrupa por destino
    const porDestino = {};
    porOrigem[nomeOrigem].forEach(m => {
      if (!porDestino[m.destinoAba]) porDestino[m.destinoAba] = [];
      porDestino[m.destinoAba].push(m);
    });
    Object.keys(porDestino).forEach(nomeDestino => {
      const sheetDestino = abaOficial[nomeDestino];
      if (!sheetDestino) return;
      // Coleta os valores de cada linha (de baixo para cima para não deslocar índice ao deletar)
      const rowsParaInserir = [];
      const rowIdxs = porDestino[nomeDestino].map(m => m.rowIdx).sort((a, b) => b - a);
      rowIdxs.forEach(rowIdx => {
        const valores = sheetOrigem.getRange(rowIdx, 1, 1, nCols).getValues()[0];
        rowsParaInserir.unshift(valores);
      });
      // Insere no destino (abaixo do cabeçalho)
      sheetDestino.getRange(sheetDestino.getLastRow() + 1, 1, rowsParaInserir.length, nCols).setValues(rowsParaInserir);
      // Reaplica cor baseada em verdict
      rowsParaInserir.forEach((r, i) => {
        const row = sheetDestino.getLastRow() - rowsParaInserir.length + i + 1;
        const verdict = r[cabecalho.indexOf('Resultado')];
        if (typeof verdict === 'string') {
          const range = sheetDestino.getRange(row, 1, 1, nCols);
          if (verdict === 'Q')       range.setBackground('#d4edda');
          else if (verdict === 'PARCIAL') range.setBackground('#fff3cd');
          else if (verdict === 'NQ') range.setBackground('#f8d7da');
        }
      });
      // Remove da origem (de baixo para cima)
      rowIdxs.forEach(rowIdx => {
        sheetOrigem.deleteRow(rowIdx);
      });
      totalMovidas += rowIdxs.length;
      Logger.log('  ⤴  ' + rowIdxs.length + ' linhas: "' + nomeOrigem + '" → "' + nomeDestino + '"');
    });
  });

  Logger.log('========== MIGRAÇÃO CONCLUÍDA ==========');
  Logger.log('Total de linhas movidas: ' + totalMovidas);
  Logger.log('Abas oficiais finais:');
  Object.keys(abaOficial).forEach(nome => {
    Logger.log('  [' + nome + '] ' + abaOficial[nome].getLastRow() + ' linhas');
  });
}

// ── Remove linhas duplicadas por ID, mantendo apenas a ÚLTIMA ocorrência de cada ID ──
// Rodar manualmente no Apps Script quando houver duplicatas
function dedupTodasAsAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAbas = [...new Set(Object.values(SALA_SHEETS))]; // ['Alta Vista', 'São Pedro', 'Atibaia']
  let totalRemovidas = 0;

  nomeAbas.forEach(nome => {
    const sheet = ss.getSheetByName(nome);
    if (!sheet) { Logger.log(nome + ': aba não encontrada, pulando.'); return; }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log(nome + ': vazia, pulando.'); return; }

    const nCols = sheet.getLastColumn();
    const allRows = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();

    // Mapeia id → índice da última ocorrência (mais recente = maior índice no array)
    const lastIdx = {};
    allRows.forEach((r, i) => {
      const id = String(r[0]).trim();
      if (id) lastIdx[id] = i;
    });

    // Marca linhas a remover (qualquer linha que NÃO é a última ocorrência do seu ID)
    const linhasParaRemover = []; // índices no array (base 0)
    allRows.forEach((r, i) => {
      const id = String(r[0]).trim();
      if (id && lastIdx[id] !== i) linhasParaRemover.push(i);
    });

    if (!linhasParaRemover.length) {
      Logger.log(nome + ': sem duplicatas.');
      return;
    }

    // Remove de baixo para cima para não deslocar índices
    linhasParaRemover.sort((a, b) => b - a);
    linhasParaRemover.forEach(i => {
      sheet.deleteRow(i + 2); // +2 porque array é base-0 e linha 1 é cabeçalho
    });

    totalRemovidas += linhasParaRemover.length;
    Logger.log(nome + ': ' + linhasParaRemover.length + ' duplicatas removidas. Restam ' + (allRows.length - linhasParaRemover.length) + ' linhas.');
  });

  Logger.log('=== Dedup concluído. Total removidas: ' + totalRemovidas + ' linhas. ===');
}

// ── Migra dados do PropertiesService (ddp_store antigo) para a aba DDP ──
// Rodar UMA VEZ manualmente no Apps Script após implantar nova versão
function migrarDDPParaSheets() {
  const raw = PropertiesService.getScriptProperties().getProperty('ddp_store');
  if (!raw) { Logger.log('Nada para migrar — ddp_store vazio.'); return; }
  const store = JSON.parse(raw);
  if (!store || !store.length) { Logger.log('Nada para migrar — lista vazia.'); return; }
  getDDPSheet(); // garante que a aba existe
  let migrados = 0;
  store.forEach(entry => {
    if (!entry.hotel || !entry.dateISO) return;
    let d = entry.dateISO;
    if (d.startsWith('DDP_')) d = d.slice(4);
    saveDDPEntry(Object.assign({}, entry, { dateISO: d }));
    migrados++;
  });
  Logger.log('Migração DDP concluída: ' + migrados + ' registros transferidos para a aba DDP.');
  // NÃO apaga o ddp_store antigo automaticamente — confirme visualmente na planilha antes de limpar
}

// Reaplica cores em todas as linhas existentes (rodar uma vez manualmente)
function reaplCarCores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SALA_SHEETS).forEach(nome => {
    const sheet = ss.getSheetByName(nome);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const resultIdx = cabecalho.indexOf('Resultado');
    if (resultIdx < 0) return;
    const nCols = sheet.getLastColumn();
    const rows = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
    rows.forEach((r, i) => {
      const verdict = r[resultIdx];
      const range = sheet.getRange(i + 2, 1, 1, nCols);
      if (verdict === 'Q')       range.setBackground('#d4edda');
      else if (verdict === 'PARCIAL') range.setBackground('#fff3cd');
      else if (verdict === 'NQ') range.setBackground('#f8d7da');
    });
    Logger.log(`${nome}: ${lastRow - 1} linhas processadas`);
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── AUDITORIA DE CIDADES THERMAS SP — % de São Paulo Capital + top 10 ──
// Lê TODA a aba "São Pedro" (que contém leads de Thermas SP + SPTR + Externo SP),
// filtra só Thermas SP (cobre nomes legados), e calcula:
//   - Total de leads de Thermas SP no histórico
//   - % que vem de "São Paulo" capital exato
//   - Top 10 cidades (em contagem absoluta + % relativo)
function auditarCidadesThermas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('São Pedro');
  if (!sheet) { Logger.log('❌ Aba "São Pedro" não encontrada'); return; }
  const last = sheet.getLastRow();
  if (last < 2) { Logger.log('❌ Aba "São Pedro" vazia'); return; }
  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxSala = cabecalho.indexOf('Sala');
  const idxCidade = cabecalho.indexOf('Cidade');
  if (idxSala < 0 || idxCidade < 0) {
    Logger.log('❌ Coluna "Sala" ou "Cidade" não encontrada');
    return;
  }
  const nCols = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, last - 1, nCols).getValues();

  // Nomes que devem cair em "Thermas SP" (cobre legado)
  const SALAS_THERMAS = new Set(['Thermas SP', 'Thermas São Pedro', 'THERMAS SÃO PEDRO', 'THERMAS SÃO PEDRO - HOTEL', 'Thermas - UH AP']);

  let total = 0;
  let spCapital = 0;
  const cidades = {}; // contagem de cidades

  rows.forEach(r => {
    const sala = String(r[idxSala] || '').trim();
    if (!SALAS_THERMAS.has(sala)) return;
    total++;
    const cidade = String(r[idxCidade] || '').trim() || '(vazio)';
    cidades[cidade] = (cidades[cidade] || 0) + 1;
    if (cidade.toLowerCase() === 'são paulo') spCapital++;
  });

  Logger.log('========== AUDITORIA THERMAS SP × CIDADE ==========');
  Logger.log('Total de leads do Thermas SP (todo o histórico): ' + total);
  Logger.log('Leads de São Paulo capital: ' + spCapital);
  const pct = total ? Math.round(spCapital / total * 1000) / 10 : 0;
  Logger.log('% de São Paulo capital: ' + pct + '%');
  Logger.log('');
  Logger.log('Top 15 cidades (com % do total):');
  const top = Object.entries(cidades).sort((a, b) => b[1] - a[1]).slice(0, 15);
  top.forEach(([cid, qtd]) => {
    const pctCid = total ? Math.round(qtd / total * 1000) / 10 : 0;
    const flag = cid.toLowerCase() === 'são paulo' ? ' ← SP CAPITAL' : '';
    Logger.log('  ' + cid.padEnd(30) + ' ' + String(qtd).padStart(6) + '  (' + pctCid + '%)' + flag);
  });
  Logger.log('========== FIM AUDITORIA ==========');
}
