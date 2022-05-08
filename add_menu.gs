function onOpen() { 
  var ui = SpreadsheetApp.getUi(); 
  var menu = ui.createMenu("追加機能"); //メニュー名 
  menu.addItem('メール送信','sendEmail'); //表示名、スクリプト名 
  menu.addItem('名簿作成', 'makeList'); //表示名、スクリプト名 
  menu.addToUi(); 
}

function sendEmail() {
  Browser.msgBox('未送信の返信メールはありません。');
}

function makeList() {
  var fileName = '名簿_ソフトウェアトレーニング(2022/1/30)'
  var ss = SpreadsheetApp.create(fileName);
  Browser.msgBox('参加者名簿を作成しました。-> ソフトウェアトレーニング(2022/1/30)_名簿');
}