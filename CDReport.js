// 【カレンダー設定】--------------------------------------------------------
// 登録・取得するイベントの色
const EventColor = CalendarApp.EventColor.YELLOW;
// 取得開始日付
const LoadingStartDate = "2008/01/01 00:00";

// 【SpreadSheet設定】--------------------------------------------------------
// シートID
const SheetID = "191fTeVKET2K5yZ6trFewRV3_8GJ80s8qC92-NtgNvv0";
// 登録用シート名
const RegisteringSheetName = "登録(CD)"
// 取得用シート名
const FetchingSheetName    = "取得(CD)"
// ---------------------------------------------------------------

// 一時格納用のリスト
const CDs = [];

/*
  メイン処理
 */
function myFunction() {
  // 登録
  RegisterCD();
  ClearSheet(RegisteringSheetName);

  // 取得
  ClearSheet(FetchingSheetName);  
  GetEvents();
}

function RegisterCD(){
  let sheet = GetSheet(RegisteringSheetName);

  for (let row = 2; row < sheet.getLastRow() + 1; row++){
    let sheet     = GetSheet(RegisteringSheetName);
    let purchased = sheet.getRange(row,1).getValue();
    let title     = sheet.getRange(row,2).getValue();
    let store     = sheet.getRange(row,3).getValue();
    let label     = sheet.getRange(row,4).getValue();
    let artist    = sheet.getRange(row,5).getValue();
    let released  = FormatDate(sheet.getRange(row,6).getValue());
    let songs     = sheet.getRange(row,7).getValue();
    let summary   = sheet.getRange(row,8).getValue();

    const calendar = CalendarApp.getDefaultCalendar();
    calendar.setTimeZone = 'Asia／Tokyo';

    var event = calendar.createAllDayEvent(title,new Date(purchased));
    event.setColor(EventColor);

    let description = '【入手先】\n' + store + '\n' 
                    + '\n【レーベル】\n' + label + '\n' 
                    + '\n【アーティスト】\n' + artist + '\n' 
                    + '\n【発売日】\n' + released + '\n'
                    + '\n【曲目】\n' + songs + '\n' 
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
  const events = calendar.getEvents(new Date(LoadingStartDate), new Date(), { search: "【レーベル】" });

  for (const event of events) { 
    event.setColor(EventColor);
    PushRecordToList(event); 
  }
 
  CDs.sort();

  let sheet = GetSheet(FetchingSheetName);
  sheet.getRange(2,1, CDs.length, CDs[0].length).setValues(CDs);
}

/* 
 取得したレコードをリストに入れる
 */
function PushRecordToList(event){
  let purchased = FormatDate(event.getStartTime());
  let title     = event.getTitle();

  let descriptions = event.getDescription().split("【");
  
  let store    = RemoveNewLineCode(descriptions[1].replace("入手先】", ""), "");
  let label    = RemoveNewLineCode(descriptions[2].replace("レーベル】", ""), "");
  let artist   = RemoveNewLineCode(descriptions[3].replace("アーティスト】", ""), "");
  let released = RemoveNewLineCode(descriptions[4].replace("発売日】", ""), "");
  let songs    = RemoveNewLineCode(descriptions[5].replace("曲目】", ""), " ");
  let summary  = RemoveNewLineCode(descriptions[6].replace("概要】", ""), "");

  CDs.push([purchased, title, store, label, artist, released, songs, summary]);
}

/* 
 改行コードを削除する
 */
function RemoveNewLineCode(text, replacedText){
  return text.replace(/\r?\n/g, replacedText);
}