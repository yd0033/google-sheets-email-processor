function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('実行', 'showEmailAddresses')
    .addItem('値をクリア', 'clearValues')
    .addToUi();
}

function showEmailAddresses() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var inputCell = sheet.getRange("B3");
  var outputCell = sheet.getRange("A6");
  var inputData = inputCell.getValue();

  // カンマ区切りのメールアドレスを配列に分割する
  var emailAddresses = inputData.split(",");

  // エラーチェック
  var invalidEmails = emailAddresses.filter(function (email) {
    return !isValidEmail(email);
  });

  if (invalidEmails.length > 0) {
    // エラーメッセージを表示
    var errorMessage = "以下のメールアドレスが不正です:\n" + invalidEmails.join("\n");
    SpreadsheetApp.getUi().alert(errorMessage);
    return;
  }

  // 結果をセルに表示
  for (var i = 0; i < emailAddresses.length; i++) {
    var email = emailAddresses[i].trim();
    if (email !== "") {
      outputCell.offset(i, 0).setValue(email);
    }
  }
}

function isValidEmail(email) {
  // シンプルなメールアドレスの正規表現パターン
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

function clearValues() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var b3Cell = sheet.getRange("B3");
  var a6Range = sheet.getRange("A6:A");

  b3Cell.clearContent();
  a6Range.clearContent();
}
