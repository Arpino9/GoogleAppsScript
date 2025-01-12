function deleteOldGmails() {
  // 1年以上前のメールを検索（件数を制限）
  var deleteThreads = GmailApp.search('older_than:1y -is:starred -is:important'); // 最大10件取得

  if (deleteThreads.length === 0) {
    Logger.log("削除対象のメールはありませんでした。");
    return;
  }

  Logger.log('該当スレッド: ' + deleteThreads.length + '件');

  for (var i = 0; i < deleteThreads.length; i++) {
    try {
      Logger.log('処理中のスレッド: ' + i);
      var messages = deleteThreads[i].getMessages();

      for (var j = 0; j < messages.length; j++) {
        var receivedDate = Utilities.formatDate(messages[j].getDate(), Session.getScriptTimeZone(), 'yyyy/MM/dd'); // 受信日付を取得
        Logger.log(receivedDate + ": " + messages[j].getSubject() + "を削除しました。");
      }

      deleteThreads[i].moveToTrash();
    } catch (e) {
      Logger.log('エラー: ' + e.message);
    }
  }
}
