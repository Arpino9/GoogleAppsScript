// @ts-nocheck
/*============================
 ここでスケジュール情報を定義する
==============================*/
const strSheetName          = '勤怠記録用'      // シート名
const strTimeZone           = 'Asia／Tokyo'    // タイムゾーン
const strLunchName          = '昼食';          // 昼食時間のタイトル名
const strYear               = '2023';         // 対象年度
const strNoonStartTime      = '09:00:00';     // 始業時間(午前)
const strNoonEndTime        = '12:00:00';     // 就業時間(午後)
const strLunchStartTime     = '12:00:00';     // 昼食時間(開始)
const strLunchEndTime       = '13:00:00';     // 昼食時間(終了)
const strAfternoonStartTime = '13:00:00';     // 始業時間(午前)
//============================
function myFunction() {
  let startTime,endTime;
  let arrOption;

  // 1.シートを定義する
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName(strSheetName); 
  const range = sheet.getRange(1,1,sheet.getLastRow(), sheet.getLastColumn());
  const values = range.getValues();

  for(let i = 1; i < values.length; i++){  
    // 2.書き込む情報を取得
    let strMonth            = values[i][1].toString();
    let strDay              = values[i][2].toString();
    let strAfternoonEndTime = values[i][3] + ':' + values[i][4] + ':00';
    let strPlace            = values[i][5].toString();
    let strNoonTitle        = values[i][6].toString();
    let strAfternoonTitle   = values[i][7].toString();
    let strNoonRemarks      = values[i][8].toString();
    let strLunchRemarks     = values[i][9].toString();
    let strAfternoonRemarks = values[i][10].toString();
    let strSameTitle        = values[i][11].toString();
    let strPrevTitleNoon    = PropertiesService.getScriptProperties().getProperty("strNoonTitle");
    let strPrevTitleAfterN  = PropertiesService.getScriptProperties().getProperty("strAfternoonTitle");

    // 3.スケジュールに書き込む
    const calendar = CalendarApp.getDefaultCalendar();
    calendar.setTimeZone = strTimeZone;

    /* 
      午前 
    */
    startTime = new Date(strYear + '/' + strMonth + '/' + strDay + ' ' + strNoonStartTime); 
    endTime   = new Date(strYear + '/' + strMonth + '/' + strDay + ' ' + strNoonEndTime); 

    arrOption = {
      description : strNoonRemarks,
      location    : strPlace
    };
    
    if(strSameTitle == "前回と同じ(午前, 午後)")
    {
       strNoonTitle = strPrevTitleNoon;
       strAfternoonTitle = strPrevTitleAfterN;
    }
    
    calendar.createEvent(strNoonTitle,startTime,endTime,arrOption);

    // プロパティに書き込む
    PropertiesService.getScriptProperties().setProperty("strNoonTitle", strNoonTitle);

    /* 
      昼
    */
    startTime = new Date(strYear + '/' + strMonth + '/' + strDay + ' ' + strLunchStartTime); 
    endTime   = new Date(strYear + '/' + strMonth + '/' + strDay + ' ' + strLunchEndTime); 

    arrOption = {
      description : strLunchRemarks,
      location    : strPlace
    };

　　calendar.createEvent(strLunchName,startTime,endTime,arrOption);

    /* 
      午後
    */
    startTime = new Date(strYear + '/' + strMonth + '/' + strDay + ' ' + strAfternoonStartTime); 
    endTime   = new Date(strYear + '/' + strMonth + '/' + strDay + ' ' + strAfternoonEndTime); 

    arrOption = {
      description : strAfternoonRemarks,
      location    : strPlace
    };

    calendar.createEvent(strAfternoonTitle,startTime,endTime,arrOption);

    // プロパティに書き込む
    PropertiesService.getScriptProperties().setProperty("strAfternoonTitle", strAfternoonTitle); 
  }

  // 4.登録が終わった行を削除
  for(let i = 1; i < values.length; i++){  
    sheet.deleteRow(2);
  }
}