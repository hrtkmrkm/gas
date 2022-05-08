
var postUrl = 'url';
var username = 'Notification';
var icon = ':sun_with_face:';

function myFunction() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getRange("A1");
  var text = range.getValue();

  const payload = {
    "username" : username,
    "icon_emoji": icon,
    'text'       : text,
  };
  const params = {
    'method'     : 'post',
    'contentType': 'application/json',
    'payload'    : JSON.stringify(payload)
  }; 
  UrlFetchApp.fetch(postUrl, params);
}