// ═══════════════════════════════════════════════════════
// Your Vacation — Qualificador de Captação
// Google Apps Script Backend
// Cole este código em: Planilha > Extensões > Apps Script
// Depois: Implantar > Nova implantação > Web App
//   - Executar como: EU (sua conta)
//   - Quem pode acessar: Qualquer pessoa
// ═══════════════════════════════════════════════════════

const SHEET_NAME = 'Leads';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Cabeçalho
    sheet.appendRow([
      'ID', 'Data', 'Hora', 'Captador', 'Sala',
      'Nome', 'Telefone', 'Resultado', 'Tipo', 'Renda', 'Modo'
    ]);
    sheet.setFrozenRows(1);
    // Formatação do cabeçalho
    const header = sheet.getRange(1, 1, 1, 11);
    header.setBackground('#1a2340');
    header.setFontColor('#c9a84c');
    header.setFontWeight('bold');
    sheet.setColumnWidth(1, 180);  // ID
    sheet.setColumnWidth(2, 100);  // Data
    sheet.setColumnWidth(3, 80);   // Hora
    sheet.setColumnWidth(4, 150);  // Captador
    sheet.setColumnWidth(5, 150);  // Sala
    sheet.setColumnWidth(6, 150);  // Nome
    sheet.setColumnWidth(7, 130);  // Telefone
    sheet.setColumnWidth(8, 90);   // Resultado
    sheet.setColumnWidth(9, 150);  // Tipo
    sheet.setColumnWidth(10, 130); // Renda
    sheet.setColumnWidth(11, 80);  // Modo
  }
  return sheet;
}

// POST — recebe um lead e salva na planilha
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    // Evita duplicatas pelo ID
    const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
    if (ids.includes(data.id)) {
      return jsonResponse({ status: 'duplicate', id: data.id });
    }

    sheet.appendRow([
      data.id        || '',
      data.date      || '',
      data.time      || '',
      data.captador  || '',
      data.sala      || '',
      data.nome      || '',
      data.tel       || '',
      data.verdict   || '',
      data.tipo      || '',
      data.renda     || '',
      data.mode      || '',
    ]);

    // Cor de linha por resultado
    const lastRow = sheet.getLastRow();
    const rowRange = sheet.getRange(lastRow, 1, 1, 11);
    if (data.verdict === 'Q')       rowRange.setBackground('#d4edda');
    if (data.verdict === 'PARCIAL') rowRange.setBackground('#fff3cd');
    if (data.verdict === 'NQ')      rowRange.setBackground('#f8d7da');

    return jsonResponse({ status: 'ok', id: data.id });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// GET — retorna todos os leads (para CDP/sync)
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse({ status: 'ok', leads: [] });

    const rows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const leads = rows
      .filter(r => r[0]) // ignora linhas vazias
      .map(r => ({
        id:       r[0],
        date:     r[1],
        time:     r[2],
        captador: r[3],
        sala:     r[4],
        nome:     r[5],
        tel:      r[6],
        verdict:  r[7],
        tipo:     r[8],
        renda:    r[9],
        mode:     r[10],
      }));

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
