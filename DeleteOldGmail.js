function deleteOldGmails() {
  // 5年以上前のメールを削除
  var deleteThreads = GmailApp.search('older_than:5y -is:starred -is:important');
  Logger.log('該当スレッド: ' + deleteThreads.length + '件');
  for (var i = 0; i < deleteThreads.length; i++) 
  {
    messages[j].moveToTrash();
  }
}
