// ═══════════════════════════════════════════════════════
// Your Vacation — Qualificador de Captação
// Google Apps Script Backend — v6
// Cole este código em: Planilha > Extensões > Apps Script
// Depois: Implantar > Gerenciar implantações > Editar > Nova versão > Implantar
// ═══════════════════════════════════════════════════════

const SALA_SHEETS = {
  'Alta Vista':       'Alta Vista',
  'São Pedro Resort': 'São Pedro',
  'Atibaia':          'Atibaia',
};
const DEFAULT_SHEET = 'Leads';

const COLUNAS = ['ID','Data','Hora','Captador','Sala','Nome','Telefone','Cidade','Resultado','Tipo','Renda','Modo','EmSala'];
const N_COLS  = COLUNAS.length;
const COL_WIDTHS = [180, 100, 80, 150, 150, 150, 130, 130, 90, 150, 130, 80, 80];

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

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // update_sala: atualiza coluna EmSala de um lead existente
    if (data._action === 'update_sala') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetsAlvo = Object.values(SALA_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean);
      for (const sheet of sheetsAlvo) {
        const row = findRowById(sheet, data.id);
        if (row > 0) {
          const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          const emSalaCol = cabecalho.indexOf('EmSala');
          if (emSalaCol >= 0) {
            sheet.getRange(row, emSalaCol + 1).setValue(data.emSala ? 'SIM' : '');
          }
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
      return jsonResponse({ status: 'duplicate', id: data.id });
    }

    sheet.appendRow([
      data.id || '', data.date || '', data.time || '',
      data.captador || '', data.sala || '', data.nome || '',
      data.tel || '', data.cidade || '', data.verdict || '',
      data.tipo || '', data.renda || '', data.mode || '',
      data.emSala ? 'SIM' : '',
    ]);

    // Aplica cor na linha recém inserida
    const row = sheet.getLastRow();
    const nCols = sheet.getLastColumn();
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

    // ?_type=pi — retorna pesquisas incompletas do PropertiesService
    if (params._type === 'pi') {
      return jsonResponse({ status: 'ok', pis: getPIs() });
    }

    const sheetsAlvo = salaFiltro
      ? [ss.getSheetByName(SALA_SHEETS[salaFiltro] || salaFiltro)].filter(Boolean)
      : Object.values(SALA_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean);

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
        const get = (col) => idx[col] >= 0 ? r[idx[col]] : '';
        leads.push({
          id: String(get('ID')), date: String(get('Data')), time: String(get('Hora')),
          captador: get('Captador'), sala: get('Sala'), nome: get('Nome'),
          tel: get('Telefone'), cidade: get('Cidade'), verdict: get('Resultado'),
          tipo: get('Tipo'), renda: get('Renda'), mode: get('Modo'),
          emSala: get('EmSala') === 'SIM',
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
