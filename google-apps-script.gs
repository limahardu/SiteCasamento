const SHEET_NAME = 'Mensagens';

function doGet() {
  getSheet_();

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Web App ativo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = getSheet_();
  const data = readData_(e);

  sheet.appendRow([
    new Date(),
    data.data || '',
    data.presente || '',
    data.valor || '',
    data.nome || '',
    data.mensagem || '',
    data.contato || ''
  ]);

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
