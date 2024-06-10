const deleteInfo = ["no-reply@accounts.google.com",
                    "noreply-apps-scripts-notifications@google.com",
                    "verify@x.com"];

//=============================
// Function：DeleteUnnecessaryGmails
//  Outline：指定されたメールアドレスを削除する
//-----------------------------
//   Create：2024/06/10
//   Update：  
//=============================
function DeleteUnnecessaryGmails() {
  // 削除対象期間を設定したい場合、たとえば「newer_than:3d」のようにすると直近3日が対象になる。
  var deleteThreads = GmailApp.search('newer_than:1d -is:starred -is:important');
  Logger.log('該当スレッド: ' + deleteThreads.length + '件');
  for (var i = 0; i < deleteThreads.length; i++) 
  {
    var messages = deleteThreads[i].getMessages();

    for (var j = 0; j < messages.length; j++){
      let slicedFrom = DivideMailAddress(messages[j].getFrom());

       if (deleteInfo.includes(slicedFrom)){
          console.log(messages[j].getSubject() + "を削除しました。");
          deleteThreads[j].moveToTrash();
       }
    }
  }
}

//=============================
// Function：DivideMailAddress
//  Outline：送信者アドレスが<>内にあれば切り出す
//    Param：fromAddress     (受信アドレス)
//   Return：slicedFrom (送信者の正式なアドレス)
//-----------------------------
//   Create：2024/06/10
//   Update：  
//=============================
function DivideMailAddress(fromAddress){
  const startPosition = SearchCharactorPosition(fromAddress, "<") + 1;
  const endPosition   = SearchCharactorPosition(fromAddress, ">");

  let slicedFrom;

  if(startPosition > -1 && endPosition > -1 ){
      slicedFrom = fromAddress.substring(startPosition, endPosition);
  } else {
      slicedFrom = fromAddress;
  }

  return slicedFrom;
}

//=============================
// Function：searchCharactorPosition
//  Outline：指定した文字が先頭から何文字目か判定する
//    Param：value            (検索対象の値)
//           charOfSearch     (判定したい文字)
//   Return：numSlicedFromTop (先頭からの位置)
//-----------------------------
//   Create：2021/01/02
//   Update：  
//=============================
function SearchCharactorPosition(value, charOfSearch ){
    numSlicedFromTop = value.indexOf(charOfSearch);
    return numSlicedFromTop;
}
