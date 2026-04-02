// ═══════════════════════════════════════════════════════
// Your Vacation — Qualificador de Captação
// Google Apps Script Backend — v2 (abas por sala)
// Cole este código em: Planilha > Extensões > Apps Script
// Depois: Implantar > Gerenciar implantações > Editar > Nova versão > Implantar
// ═══════════════════════════════════════════════════════

// Mapeamento sala → nome da aba na planilha
const SALA_SHEETS = {
  'Alta Vista':       'Alta Vista',
  'São Pedro Resort': 'São Pedro',
  'Atibaia':          'Atibaia',
};
const DEFAULT_SHEET = 'Leads'; // fallback para salas não mapeadas

const COLUNAS = ['ID','Data','Hora','Captador','Sala','Nome','Telefone','Cidade','Resultado','Tipo','Renda','Modo'];
const N_COLS  = COLUNAS.length;

function getOrCreateSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(COLUNAS);
    sheet.setFrozenRows(1);
    const header = sheet.getRange(1, 1, 1, N_COLS);
    header.setBackground('#1a2340');
    header.setFontColor('#c9a84c');
    header.setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 150);
    sheet.setColumnWidth(7, 130);
    sheet.setColumnWidth(8, 130);
    sheet.setColumnWidth(9, 90);
    sheet.setColumnWidth(10, 150);
    sheet.setColumnWidth(11, 130);
    sheet.setColumnWidth(12, 80);
  }
  return sheet;
}

function resolverSheet(sala) {
  const nome = SALA_SHEETS[sala] || DEFAULT_SHEET;
  return getOrCreateSheetByName(nome);
}

// POST — recebe um lead e salva na aba da sala correspondente
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = resolverSheet(data.sala || '');

    // Evita duplicatas pelo ID
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      if (ids.map(String).includes(String(data.id))) {
        return jsonResponse({ status: 'duplicate', id: data.id });
      }
    }

    sheet.appendRow([
      data.id       || '',
      data.date     || '',
      data.time     || '',
      data.captador || '',
      data.sala     || '',
      data.nome     || '',
      data.tel      || '',
      data.cidade   || '',
      data.verdict  || '',
      data.tipo     || '',
      data.renda    || '',
      data.mode     || '',
    ]);

    const row = sheet.getLastRow();
    const range = sheet.getRange(row, 1, 1, N_COLS);
    if (data.verdict === 'Q')       range.setBackground('#d4edda');
    if (data.verdict === 'PARCIAL') range.setBackground('#fff3cd');
    if (data.verdict === 'NQ')      range.setBackground('#f8d7da');

    return jsonResponse({ status: 'ok', id: data.id });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// GET — retorna leads de todas as abas consolidados (para CDP/sync)
// Aceita ?sala=Nome para filtrar por sala específica
function doGet(e) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const params = e && e.parameter ? e.parameter : {};
    const salaFiltro = params.sala || null;

    const sheetsAlvo = salaFiltro
      ? [ss.getSheetByName(SALA_SHEETS[salaFiltro] || salaFiltro)].filter(Boolean)
      : Object.values(SALA_SHEETS).map(n => ss.getSheetByName(n)).filter(Boolean);

    const leads = [];
    sheetsAlvo.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      const rows = sheet.getRange(2, 1, lastRow - 1, N_COLS).getValues();
      rows.forEach(r => {
        if (!r[0]) return; // ignora linha vazia
        leads.push({
          id:       String(r[0]),
          date:     r[1],
          time:     r[2],
          captador: r[3],
          sala:     r[4],
          nome:     r[5],
          tel:      r[6],
          cidade:   r[7],
          verdict:  r[8],
          tipo:     r[9],
          renda:    r[10],
          mode:     r[11],
        });
      });
    });

    return jsonResponse({ status: 'ok', leads });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
