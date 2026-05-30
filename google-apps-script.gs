const SHEET_NAME = 'Mensagens';
const SHEET_PRESENCA = 'Presenças';

function doGet() {
  getSheet_();

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Web App ativo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = readData_(e);

  if (data.tipo === 'presenca') {
    const sheet = getPresencaSheet_();
    sheet.appendRow([
      new Date(),
      data.nome || '',
      parseInt(data.pessoas) || 0,
      parseFloat(data.total) || 0
    ]);
  } else {
    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      data.data || '',
      data.presente || '',
      data.valor || '',
      data.nome || '',
      data.mensagem || '',
      data.contato || ''
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function readData_(e) {
  if (e && e.parameter && Object.keys(e.parameter).length) {
    return e.parameter;
  }

  try {
    return JSON.parse(e.postData.contents || '{}');
  } catch (err) {
    return {};
  }
}

function getPresencaSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_PRESENCA);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_PRESENCA);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Recebido em', 'Nome', 'Pessoas', 'Total (R$)']);
  }

  return sheet;
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Recebido em',
      'Data do site',
      'Presente',
      'Valor',
      'Nome',
      'Mensagem',
      'Contato/observacao'
    ]);
  }

  return sheet;
}
