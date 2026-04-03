// ═══════════════════════════════════════════════════════
// Your Vacation — Qualificador de Captação
// Google Apps Script Backend — v3 (cabeçalho auto-migração)
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

// Larguras das colunas (em pixels)
const COL_WIDTHS = [180, 100, 80, 150, 150, 150, 130, 130, 90, 150, 130, 80];

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
    // Aba não existe — cria do zero
    sheet = ss.insertSheet(sheetName);
    aplicarCabecalho(sheet);
    return sheet;
  }

  // Aba existe — verifica se o cabeçalho está atualizado
  const cabecalhoAtual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cabecalhoCorreto = COLUNAS.every((col, i) => cabecalhoAtual[i] === col);

  if (!cabecalhoCorreto) {
    // Migra: garante que a linha 1 tem todas as colunas certas
    // Preserva dados existentes — só reescreve o cabeçalho
    const nColsAtual = cabecalhoAtual.length;
    if (nColsAtual < N_COLS) {
      // Expande cabeçalho para ter as novas colunas
      sheet.getRange(1, 1, 1, N_COLS).setValues([COLUNAS]);
    } else {
      sheet.getRange(1, 1, 1, N_COLS).setValues([COLUNAS]);
    }
    // Reaplicar formatação
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

      // Descobre índice de cada coluna pelo cabeçalho (tolerante a ordem diferente)
      const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const idx = {};
      COLUNAS.forEach(col => { idx[col] = cabecalho.indexOf(col); });

      const nCols = sheet.getLastColumn();
      const rows  = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
      rows.forEach(r => {
        if (!r[0]) return;
        const get = (col) => idx[col] >= 0 ? r[idx[col]] : '';
        leads.push({
          id:       String(get('ID')),
          date:     String(get('Data')),
          time:     String(get('Hora')),
          captador: get('Captador'),
          sala:     get('Sala'),
          nome:     get('Nome'),
          tel:      get('Telefone'),
          cidade:   get('Cidade'),
          verdict:  get('Resultado'),
          tipo:     get('Tipo'),
          renda:    get('Renda'),
          mode:     get('Modo'),
        });
      });
    });

    return jsonResponse({ status: 'ok', leads });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// Função utilitária — pode rodar manualmente para forçar migração de todas as abas
function migrarTodasAsAbas() {
  Object.values(SALA_SHEETS).forEach(nome => getOrCreateSheetByName(nome));
  Logger.log('Migração concluída.');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
