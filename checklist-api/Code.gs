// Mom Checklist State - Apps Script Web App
// Spreadsheet: https://docs.google.com/spreadsheets/d/18K23ByRflscbI5EejwXrZOf8dYzVB-frX5oOQMhXuEA

var SPREADSHEET_ID = '18K23ByRflscbI5EejwXrZOf8dYzVB-frX5oOQMhXuEA';

function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
}

function doGet(e) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var result = {};

  for (var i = 1; i < data.length; i++) {
    var itemId = data[i][0];
    var checked = data[i][1];
    if (itemId) {
      result[itemId] = (checked === true || checked === 'TRUE' || checked === 'true');
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var itemId = payload.itemId;
    var checked = payload.checked;

    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var found = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === itemId) {
        sheet.getRange(i + 1, 2).setValue(checked);
        sheet.getRange(i + 1, 3).setValue(new Date().toISOString());
        found = true;
        break;
      }
    }

    if (!found) {
      sheet.appendRow([itemId, checked, new Date().toISOString()]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
