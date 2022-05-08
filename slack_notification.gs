var postUrl = 'url';
var username = 'Notification';
var icon = ':sun_with_face:';
var message = 'TEST!';

function myFunction() {
  const payload = {
    "username" : username,
    "icon_emoji": icon,
    'text'       : message,
  };
  const params = {
    'method'     : 'post',
    'contentType': 'application/json',
    'payload'    : JSON.stringify(payload)
  };
  UrlFetchApp.fetch(postUrl, params);
}