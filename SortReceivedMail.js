//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// 要設定項目
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
const INITIAL_DATE = "2021/01/01 00:00:00";   // 1.日付の初期値(初回起動時のみに使う変数なので、変えなくてもOK)
const MAX_THREADS  = 20;                      // 2.スレッドの最大取得数(処理が重い場合は調整する)

// 3.ラベルとアドレスの紐づけ
//   形式：[Gmailに設定したラベル名, 送信者アドレス]
//        ex)["Amazon","@amazon.co.jp"]
const sortInfo = [["Adobe", "@mail.adobe.com"],
                ["Amazon","@amazon.co.jp"]];

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//↓ここから処理

//=============================
// Function：SortReceivedMail
//  Outline：新着メールを取得してラベルを付与するメイン処理
//-----------------------------
//   Create：2021/01/02
//   Update：2023/04/12 
//=============================
function SortReceivedMail() {
let query = "";                                 // スレッド仕分け用のクエリ
let SearchSenderQuery = "";                     // 送信者検索用のクエリ
let numSlicedFromTop                            // 先頭から何文字目か
let startPosition;                              // 先頭からの開始位置
let endPosition;                                // 先頭からの終了位置
let slicedFrom;                                 // <>内のアドレス
let curSortInfoLabel                            // 現在のsortInfoのラベル名(デバッグ用)

  // 最終のメール日時を取得
  let lastDateText = PropertiesService.getScriptProperties().getProperty("lastDateText");
  
  // 最終のメール日時がなければ初期値を代入する
  if(lastDateText === null) lastDateText = INITIAL_DATE;
  
  // 最終メール日時を取得
  const lastDate = new Date(lastDateText);
  let newDate = lastDate;

  // 日付が最終更新日以降の受信メールを検索
 let date = Utilities.formatDate(lastDate, 'JST','yyyy/MM/dd');
  query += "after:" + date;
  query += " has:nouserlabels";
  let threads = GmailApp.search(query, 0, MAX_THREADS);
  
  // マッチしなければ終了
  if(threads.length === 0){
    return;
  }

  // 指定された期間内のスレッドを取得する
  let allMessages = GmailApp.getMessagesForThreads(threads);
  allMessages.reverse();

  // スレッド>メールの順に検索して判定する
  for(const thread of allMessages){
    for(const mail of thread){
      // 日付が古いor送信者が自分なら無視する
      const targetDate = mail.getDate();
      
      let from = mail.getFrom();
      let myEmail     = Session.getActiveUser().getEmail();
      
      if( targetDate<=lastDate || from.match(myEmail) ) continue;
      
      // 日付が最終更新日以降なら、日付を書き換える
      if(targetDate > newDate) newDate = targetDate;
      
      // 放り込むラベル名とマッチしていればラベルを付与する
      for (var i = 0; i < sortInfo.length; i++) {
        
        // ※デバッグ用
        //let sortInfoLen = sortInfo.length;
        //let curInfo     = sortInfo[i][1];
        //let subject     = mail.getSubject();
        //let sendDate    = mail.getDate();
        
        // 送信者アドレスが<>内にあれば切り出す
        startPosition = SearchCharactorPosition(from, "<") + 1;
        endPosition   = SearchCharactorPosition(from, ">");

        if(startPosition > -1 && endPosition > -1 ){
           slicedFrom = from.substring(startPosition, endPosition);
        } else {
           slicedFrom = from;
        }
        
        // 設定した配列の中に、対象のアドレスがあるか確認する
        if(slicedFrom.search(sortInfo[i][1]) > 0) {
          curSortInfoLabel = sortInfo[i][0];

          // 指定した送信アドレスを検索してラベルをつける
          SearchSenderQuery = "from:" + from;          

          let updThread = GmailApp.search(SearchSenderQuery, 0, MAX_THREADS);

          if (updThread == null){
            continue;
          }
          console.log(from);
          console.log(sortInfo[i][1]);
          console.log(slicedFrom.search(sortInfo[i][1]));

          var targetLabel = GmailApp.getUserLabelByName(sortInfo[i][0]);
          targetLabel.addToThreads(updThread);
        }
      }
    }
  }  

  // 8.受信日が最新の日付データをプロパティストアに書き込む
  const newDateText = Utilities.formatDate( newDate, 'Asia/Tokyo', "yyyy/MM/dd HH:mm:ss");
  PropertiesService.getScriptProperties().setProperty("lastDateText",newDateText);

  //------------------------------------------------------

  
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
}