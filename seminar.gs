const CAPACITY = 10;
const NOTIFICATION_EMAIL = "通知用メールアドレス";
const FROM = "返信用メールアドレス";
const AMOUNT = 2000;
const Q1 = "氏名";
const Q2 = "メールアドレス";
const Q3 = "参加時間";
const Q4 = "自由記入欄";
const T1 = "①・②両方";
const T2 = "①のみ";
const T3 = "②のみ";
const ATTENDANCE = {"①・②両方": [1, 1],
                    "①のみ": [1, 0],
                    "②のみ": [0, 1]};
const YEAR = 22

function onFormSubmit(e) { 
  var form = FormApp.getActiveForm();
  var title = form.getTitle();
  var description = form.getDescription();
  
  var itemResponses = getFormResponses(e);
  var name = itemResponses[Q1].replace(/\s+/g, "");
  var email = itemResponses[Q2];
  var time = itemResponses[Q3];
  var comments = itemResponses[Q4];
　
  ss = SpreadsheetApp.openById('17IxPj4VRl-T2dQ3rNEbUmgx0Y3RSV6qnZlX7Gu-ZX-M')
  SpreadsheetApp.setActiveSpreadsheet(ss);
  // Create a sheet
  var sheet_name = '2022/01/30';
  var sheet = ss.getSheetByName(sheet_name);
  if (sheet === null) {
    sheet = ss.insertSheet(sheet_name);
    sheet.getRange("J1:K2").setValues([["①部合計", "=SUM(D:D)"],
                                       ["②部合計", "=SUM(E:E)"]]);
  }
  
  // Check the duplicate application
  var lastRow = getLastRowInACol(sheet, "A");
  var row = lastRow + 1;
  var duplicated = 0
  if (lastRow != 0) {
    var applicants = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
    var duplicatedRow = applicants.indexOf(name);
    if (duplicatedRow != -1) {
      row = duplicatedRow + 1;
      duplicated = 1
    }
  }

  // Send a confirmation email.
  var remainingDailyQuota = MailApp.getRemainingDailyQuota();
  var delivered = 0
  var subject = "ソフトウェアトレーニング申し込み完了のお知らせ"; 
  if (email && remainingDailyQuota > 0){
    delivered = 1
    var args = {
      subject: subject,
      title: title,
      description: description, 
      name: name,
      email: email,
      time: time,
      comments: comments
    };
    duplicated ? sendDuplicateEmail(args) : sendConfirmationEmail(args);
  }
  
  // Register Calender
  registerCalender()

  // Write the registration information in the sheet
  var count = ATTENDANCE[time]
  var values = [name, time, comments, count[0], count[1], sumOfArray(count) * AMOUNT, email, remainingDailyQuota, delivered]; 
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  SpreadsheetApp.flush();
  
  // Count the application
  var lastRow = getLastRowInACol(sheet, "A");
  var arr1 = sheet.getRange(1, 4, lastRow, 1).getValues().flat();
  var arr2 = sheet.getRange(1, 5, lastRow, 1).getValues().flat();
  c1 = sumOfArray(arr1);
  c2 = sumOfArray(arr2); 
  
  // close the form
  if (c1 >= CAPACITY && c2 >= CAPACITY) {
    form.setAcceptingResponses(false);
    sendNotificationEmail(title, description);
  }
  // Change the form
  var items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE); 
  var item =  items[0];
  var multipleChoiceItem = item.asMultipleChoiceItem();
  item.setHelpText("①残席:" + String(CAPACITY - c1) + " ②残席:" + String(CAPACITY - c2));
  
  // Change the choices in the form
  if (c1 >= CAPACITY) {
    item.setHelpText("①残席:0 ②残席:" + String(CAPACITY - c2));
    multipleChoiceItem.setChoiceValues(["②のみ"]);  
  }
  if (c2 >= CAPACITY) {
    item.setHelpText("①残席:" + String(CAPACITY - c1) + " ②残席:0");
    multipleChoiceItem.setChoiceValues(["①のみ"]);  
  }
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

function getFileIdByName(fn) {
  var files  = DriveApp.getFilesByName(fn);
  if(files.hasNext()) {
     return files.next().getId();
  }
  return null;
}

function sumOfArray(arr) {
  return arr.reduce((a, b) => a + b, 0);
}

function getLastRowInACol(sheet, col) {
  var values = sheet.getRange(col + ":" + col).getValues();
  return values.filter(String).length;
}

function registerCalender() {
  let calendar = CalendarApp.getDefaultCalendar();
  let title = "テストイベントです！";
  let startTime = new Date();
  let endTime = new Date();
  new Date(endTime.setHours(endTime.getHours() + 2));
  calendar.createEvent(title, startTime, endTime);
}

function sendConfirmationEmail(args){
  var body =
      args.name + " 様\n\n" +
      "下記の内容でお申し込みを受け付けいたしました。\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      args.title + "\n\n" + 
      args.description + "\n\n" +
      "氏名: "  + args.name + "\n" +
      "参加時間: " + args.time + "\n" +
      "自由記入欄: " + args.comments + "\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  
  GmailApp.sendEmail(
    args.email,
    args.subject,
    body,
    {
      from: FROM,
      name:'***'
    });
}

function sendDuplicateEmail(args){
  var body =
      args.name + " 様\n\n" +
      "この日程のソフトウェアトレーニングは既に申し込みをされています。\n" +        
      "下記の内容で再度お申し込みを受け付けいたしました。\n" +
      "他の日程とお間違いないかご確認下さい。\n" +        
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      args.title + "\n\n" + 
      args.description + "\n\n" +
      "氏名: "  + args.name + "\n" +
      "参加時間: " + args.time + "\n" +
      "自由記入欄: " + args.comments + "\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  
  GmailApp.sendEmail(
    args.email,
    "(重複申込) " + args.subject,
    body,
    {
      from: FROM,
      name:'***'
    });
}

function sendNotificationEmail(title, description) {
  var subject = "満席:「" + title + "」";
  var body =
      "下記のソフトウェアトレーニングが定員に達したため、受付を終了いたしました。\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      title + "\n\n" + 
      description + "\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
    
  GmailApp.sendEmail(
    NOTIFICATION_EMAIL,
    subject,
    body,
    {
      from: FROM,
      name:'***'
    });
  
}