// 【カレンダー設定】--------------------------------------------------------
// 登録・取得するイベントの色
const EventColor = CalendarApp.EventColor.ORANGE;
// 取得開始日付
const LoadingStartDate = "2008/01/01 00:00";

// 【SpreadSheet設定】--------------------------------------------------------
// シートID
const SheetID = "191fTeVKET2K5yZ6trFewRV3_8GJ80s8qC92-NtgNvv0";
// 登録用シート名
const RegisteringSheetName = "登録(ゲーム)"
// 取得用シート名
const FetchingSheetName    = "取得(ゲーム)"
// ---------------------------------------------------------------

// 一時格納用のリスト
const Games = [];

/*
  メイン処理
 */
function myFunction() {
  // 登録
  RegisterGame();
  ClearSheet(RegisteringSheetName);

  // 取得
  ClearSheet(FetchingSheetName);  
  GetEvents();
}

function RegisterGame(){
  let sheet = GetSheet(RegisteringSheetName);

  for (let row = 2; row < sheet.getLastRow() + 1; row++){
    let sheet     = GetSheet(RegisteringSheetName);
    let purchased = FormatDate(sheet.getRange(row,1).getValue());
    let title     = sheet.getRange(row,2).getValue();
    let store     = sheet.getRange(row,3).getValue();
    let maker     = sheet.getRange(row,4).getValue();
    let platform  = sheet.getRange(row,5).getValue();
    let released  = FormatDate(sheet.getRange(row,6).getValue());
    let summary   = sheet.getRange(row,7).getValue();

    const calendar = CalendarApp.getDefaultCalendar();
    calendar.setTimeZone = 'Asia／Tokyo';

    var event = calendar.createAllDayEvent(title,new Date(purchased));
    event.setColor(EventColor);

    let description = '【購入先】\n' + store + '\n' 
                    + '\n【メーカー】\n' + maker + '\n' 
                    + '\n【プラットフォーム】\n' + platform + '\n' 
                    + '\n【発売日】\n' + released + '\n'
                    + '\n【概要】\n' + summary
    event.setDescription(description);
  }
}

/*
  日付の書式変換
 */
function FormatDate(date){
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

/*
  シートのクリア
 */
function ClearSheet(sheetName){
  let sheet = GetSheet(sheetName);
  sheet.getRange(2, 1, sheet.getLastRow() + 1, sheet.getLastColumn()).clearContent();
}

/*
  シートを取得
 */
function GetSheet(sheetName){
  const spreadSheet = SpreadsheetApp.openById(SheetID); 
  return spreadSheet.getSheetByName(sheetName);
}

/* 
 イベントを取得
 */
function GetEvents(){
  const calendar  = CalendarApp.getDefaultCalendar();  
  const events = calendar.getEvents(new Date(LoadingStartDate), new Date(), { search: "【プラットフォーム】" });

  for (const event of events) { 
    event.setColor(EventColor);
    PushRecordToList(event); 
  }
 
  Games.sort();

  let sheet = GetSheet(FetchingSheetName);
  sheet.getRange(2,1, Games.length, Games[0].length).setValues(Games);
}

/* 
 取得したレコードをリストに入れる
 */
function PushRecordToList(event){
  let purchased = FormatDate(event.getStartTime());
  let title     = event.getTitle();

  let descriptions = event.getDescription().split("【");

  let store    = RemoveNewLineCode(descriptions[1].replace("購入先】", ""));
  let maker    = RemoveNewLineCode(descriptions[2].replace("メーカー】", ""));
  let platform = RemoveNewLineCode(descriptions[3].replace("プラットフォーム】", ""));
  let released = RemoveNewLineCode(descriptions[4].replace("発売日】", ""));
  let summary  = RemoveNewLineCode(descriptions[5].replace("概要】", ""));

  Games.push([purchased, title, store, maker, platform, released, summary]);
}

/* 
 改行コードを削除する
 */
function RemoveNewLineCode(text){
  return text.replace(/\r?\n/g, "");
}