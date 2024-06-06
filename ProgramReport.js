// 【カレンダー設定】--------------------------------------------------------
// 登録・取得するイベントの色
const EventColor = CalendarApp.EventColor.PALE_RED;
// 取得開始日付(デフォルト値)
const LoadingStartDate = "2008/01/01 00:00";

// 【SpreadSheet設定】--------------------------------------------------------
// シートID
const SheetID = "191fTeVKET2K5yZ6trFewRV3_8GJ80s8qC92-NtgNvv0";
// 登録用シート名
const RegisteringSheetName = "登録(番組)"
// 取得用シート名
const FetchingSheetName    = "取得(番組)"
// ---------------------------------------------------------------

// 一時格納用のリスト
const Programs = [];

/*
  メイン処理
 */
function myFunction() {
  // 登録
  RegisterPrograms();
  //ClearSheet(RegisteringSheetName);

  // 取得
  ClearSheet(FetchingSheetName);  
  GetEvents();
}

function RegisterPrograms(){
  let sheet = GetSheet(RegisteringSheetName);

  for (let row = 2; row < sheet.getLastRow() + 1; row++){
    let sheet       = GetSheet(RegisteringSheetName);
    let watchedDate = sheet.getRange(row,1).getValue();
    let title       = sheet.getRange(row,2).getValue();
    let episode     = sheet.getRange(row,3).getValue();
    let subTitle    = sheet.getRange(row,4).getValue();
    let site        = sheet.getRange(row,5).getValue();
    //let airdate     = FormatDate(sheet.getRange(row,6).getValue());
    let summary     = sheet.getRange(row,7).getValue();

    const calendar = CalendarApp.getDefaultCalendar();
    calendar.setTimeZone = 'Asia／Tokyo';

    let event;
    
    if (episode != ""){
      title += " 第" + episode + "話";      
    }

    event = calendar.createAllDayEvent(title, watchedDate);
    event.setColor(EventColor);
    
    let description = '\n【サブタイトル】\n' + subTitle + '\n' 
                    + '\n【視聴先】\n' + site + '\n' 
                    //+ '\n【放送日】\n' + airdate + '\n'
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
  const events = calendar.getEvents(new Date(LoadingStartDate), new Date(), { search: "【視聴先】" });

  for (const event of events) { 
    event.setColor(EventColor);
    PushRecordToList(event); 
  }
 
  Programs.sort();

  let sheet = GetSheet(FetchingSheetName);
  sheet.getRange(2,1, Programs.length, Programs[0].length).setValues(Programs);
}

/* 
 取得したレコードをリストに入れる
 */
function PushRecordToList(event){
  let watchedDate = FormatDate(event.getStartTime());
  let title     = event.getTitle();

  let descriptions = event.getDescription().split("【");
  
  let subTitle    = RemoveNewLineCode(descriptions[1].replace("サブタイトル】", ""), "");
  let site        = RemoveNewLineCode(descriptions[2].replace("視聴先】", ""), "");
  //let airdate     = RemoveNewLineCode(descriptions[3].replace("放送日】", ""), "");
  //let summary     = RemoveNewLineCode(descriptions[4].replace("概要】", ""), "");
  let summary     = RemoveNewLineCode(descriptions[3].replace("概要】", ""), "");

  //Programs.push([watchedDate, title, subTitle, site, airdate, summary]);
  Programs.push([watchedDate, title, subTitle, site, summary]);
}

/* 
 改行コードを削除する
 */
function RemoveNewLineCode(text, replacedText){
  return text.replace(/\r?\n/g, replacedText);
}