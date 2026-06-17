// Google Apps Script proxy for ERP dashboard
// Deploy: Extensions > Apps Script > Deploy > New Deployment
//   Type: Web App
//   Execute as: Me (your Google account)
//   Who has access: Anyone
// Copy the Web App URL and set it as PROXY_URL in index.html

const SPREADSHEET_ID = '1N87NXPJuRNXyuONZOJyIN5iyMdpgmykx_iTGWAZSgW0';

function doGet(e) {
  const sheet = e.parameter.sheet || '';
  const range = e.parameter.range || '';

  if (!sheet) {
    return jsonResponse({ error: 'missing sheet parameter' });
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(sheet);
    if (!sh) return jsonResponse({ error: 'sheet not found: ' + sheet });

    let values;
    if (range) {
      values = sh.getRange(range).getValues();
    } else {
      values = sh.getDataRange().getValues();
    }

    return jsonResponse({ values: values });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
