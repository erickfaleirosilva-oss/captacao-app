// ═══════════════════════════════════════════════════════
// Your Vacation — Qualificador de Captação
// Google Apps Script Backend — v7
// Cole este código em: Planilha > Extensões > Apps Script
// Depois: Implantar > Gerenciar implantações > Editar > Nova versão > Implantar
// ═══════════════════════════════════════════════════════

const SALA_SHEETS = {
  'Alta Vista':       'Alta Vista',
  'Atrium':           'Alta Vista',
  'Marina':           'Alta Vista',
  'Externo':          'Alta Vista',
  'Thermas SP':       'São Pedro',
  'SPTR':             'São Pedro',
  'São Pedro Resort': 'São Pedro',   // legado
  'Atibaia':          'Atibaia',
};
const DEFAULT_SHEET = 'Alta Vista'; // fallback — evita criação de abas novas

const COLUNAS = [
  'ID','Data','Hora','Captador','Sala','Nome','Telefone','Cidade','Resultado','Tipo','Renda','Modo','EmSala',
  // Campos detalhados das respostas
  'PontoCaptacao','IdadeTitular','ProfissaoTitular','IdadeConjuge','ProfissaoConjuge',
  'Carro','Casa','Cartao','Viagens',
];
const N_COLS  = COLUNAS.length;
const COL_WIDTHS = [180,100,80,150,150,150,130,130,90,150,130,80,80,150,120,150,120,150,120,80,80,80];

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
const DDP_COLS = ['hotel','dateISO','ciPool','ciCot','salaPool','salaCot','salaConv','vendPool','vendCot','vendConv'];

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
        ciPool:   Number(data.ciPool)   || 0,
        ciCot:    Number(data.ciCot)    || 0,
        salaPool: Number(data.salaPool) || 0,
        salaCot:  Number(data.salaCot)  || 0,
        salaConv: Number(data.salaConv) || 0,
        vendPool: Number(data.vendPool) || 0,
        vendCot:  Number(data.vendCot)  || 0,
        vendConv: Number(data.vendConv) || 0,
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

    const ans = data._ans || {};
    sheet.appendRow([
      data.id || '', data.date || '', data.time || '',
      data.captador || '', data.sala || '', data.nome || '',
      data.tel || '', data.cidade || '', data.verdict || '',
      data.tipo || '', data.renda || '', data.mode || '',
      data.emSala ? 'SIM' : '',
      // Campos detalhados
      data.sala || '',                     // PontoCaptacao (mesmo que Sala — filtrável)
      ans.idadeMarido      || '',          // IdadeTitular
      ans.profissaoTitular || '',          // ProfissaoTitular
      ans.idadeConjuge     || '',          // IdadeConjuge
      ans.profissaoConjuge || '',          // ProfissaoConjuge
      ans.carro            || '',          // Carro
      ans.casa             || '',          // Casa
      ans.cartao           || '',          // Cartao
      ans.viagens          || '',          // Viagens
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

    // ?action=getDDP&hotel=SPTR — retorna registros DDP
    if (params.action === 'getDDP') {
      const hotel = params.hotel || '';
      const store = getDDPStore();
      const rows = hotel ? store.filter(r => r.hotel === hotel) : store;
      return jsonResponse({ status: 'ok', rows });
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
          _ans: {
            idadeMarido:      get('IdadeTitular')      || null,
            profissaoTitular: get('ProfissaoTitular')  || null,
            idadeConjuge:     get('IdadeConjuge')      || null,
            profissaoConjuge: get('ProfissaoConjuge')  || null,
            carro:            get('Carro')             || null,
            casa:             get('Casa')              || null,
            cartao:           get('Cartao')            || null,
            viagens:          get('Viagens')           || null,
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
