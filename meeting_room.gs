const CID = 'カレンダーID';
const FROM = "メールアドレス";

function onFormSubmit(e) {
  var itemResponses = getFormResponses(e);
  var room = itemResponses['会議室'];
  var start = itemResponses['開始日時'];
  var end = itemResponses['終了日時'];
  var name = itemResponses['名前'];
  var email = itemResponses['メールアドレス'];
  var calendar = CalendarApp.getCalendarById(CID);
  calendar.setTimeZone("Asia/Tokyo");
  var title = name + '@' + room;
  var startTime = new Date(start);
  var endTime = new Date(end);
  options = {location: room}

  if (start > end && email) {
    sendDateErrorEmail(name, email, start, end);
    return;
  }
  
  events = calendar.getEvents(startTime, endTime);
  location = false;
  for (var i in events) {
    var event = events[i];
    if (event.getLocation() == room) {
      location = true;
    }
  }

  if (events.length && location && email) {
    sendOverlappingErrorEmail(name, email, room, start, end);
    return;
  }

  if (email) {
    sendConfirmationEmail(name, email, room, start, end)
  }
  calendar.createEvent(title, startTime, endTime, options);

}

function getFormResponses(e) {
  var itemResponses = e.response.getItemResponses();
  var obj = {};
  for (var i = 0; i < itemResponses.length; i++) { 
    var itemResponse = itemResponses[i]; 
    var question = itemResponse.getItem().getTitle(); 
    var answer = itemResponse.getResponse();
    obj[question] = answer;
  }
  return obj;
}

function sendConfirmationEmail(name, email, room, start, end){
  var body =
      name + " 様\n\n" +
      "下記の日時で" + room + "を予約しました。\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "会議室: " + room + "\n" +
      "開始日時: "  + start + "\n" +
      "終了日時: " + end + "\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  
  GmailApp.sendEmail(
    email,
    room + 'の予約が完了しました。',
    body,
    {
      from: FROM,
      name:'***'
    });
}

function sendDateErrorEmail(name, email, start, end){
  var body =
      name + " 様\n\n" +
      "下記のように開始日時が終了日時より前に設定されています。再度予約を行って下さい。\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "開始日時: "  + start + "\n" +
      "終了日時: " + end + "\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  
  GmailApp.sendEmail(
    email,
    '指定の日時では会議室の予約が取れませんでした。',
    body,
    {
      from: FROM,
      name:'***'
    });
}

function sendOverlappingErrorEmail(name, email, room, start, end){
  var body =
      name + " 様\n\n" +
      "下記の日時の" + room + "は既に予約が入っています。\n" +
      "カレンダーでご確認後、会議室か日時を変更して再度予約を行って下さい。\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "会議室: " + room + "\n" +
      "開始日時: "  + start + "\n" +
      "終了日時: " + end + "\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  
  GmailApp.sendEmail(
    email,
    '指定の日時の' + room + 'は既に予約が入っております。',
    body,
    {
      from: FROM,
      name:'***'
    });
}
