/*=============================
   Function：getGaDataPrevDay
    Outline：Google AnalyticsとAdsenseからデータ取得する
    Remarks：参考にさせて頂いたソース
             https://tekito-style.me/columns/gas-analytics-basic
-------------------------------
     Create：2021/01/23
     Update：2021/01/24  
=============================*/
function getGaDataPrevDay() {
// 昨日の日付を取得する
var date = new Date();
date.setDate(date.getDate() - 1);

//　年度を取得する
var year = date.getFullYear();

//　曜日を取得する
var ary = ['日', '月', '火', '水', '木', '金', '土'];
var weekDayNum = date.getDay();
var weekDay = ary[weekDayNum] ;

//-------GAデータ-----------
// GAからデータを取得する関数(メモ)
/*
   ga:date               : 日付
   ga:yearMonth          : 年月
   ga:users              : ユーザー (新規とリピーター) 
   ga:newUsers           : 新規ユーザー  
   ga:sourceMedium       : 参照元／メディア
   ga:landingPagePath    : ランディングページ
   ga:pagePath           : ページ
   ga:pageTitle          : ページタイトル
   ga:avgSessionDuration : 平均セッション時間
   ga:avgTimeOnPage      : 平均ページ滞在時間
   ga:bounceRate         : 直帰率
   ga:sessions           : セッション数
   ga:pageviews          : ページビュー数
   ga:percentNewSessions : 新規セッション率
   ga:goalXXCompletions  : 目標XXの完了数
   ga:pageValue          : ページの価値
 */

function getGaData(startDate, endDate) {
var gaData = Analytics.Data.Ga.get(
'ga:XXXXXXXXXXX', //アナリティクスで使っているID名
startDate,
endDate,
'ga:pageviews, ga:sessions, ga:users, ga:newUsers,ga:bounceRate,ga:exitRate,ga:avgSessionDuration,ga:pageviewsPerSession, ga:totalEvents, ga:adsenseRevenue'
).rows;

return gaData;
}

// 戻り値（指定期間）を指定し、関数を実効
var gaData = getGaData('yesterday', 'yesterday');

//-------スプレッドシートへの書き込み-----------
// 書き込むシートを選択する
var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadSheet.getSheetByName('Daily'); //集計するシートタブ名を指定

//最終行+1行目の取得
var lastRow = sheet.getLastRow()+1;
//最終行+1行目を追加
sheet.insertRows(lastRow);

//追加した行に書き込む
sheet.getRange(lastRow, 1).setValue(year);
sheet.getRange(lastRow, 2).setValue(date);
sheet.getRange(lastRow, 3).setValue(weekDay);
sheet.getRange(lastRow, 4).setValue(gaData[0][0]);
sheet.getRange(lastRow, 5).setValue(gaData[0][2]);
sheet.getRange(lastRow, 6).setValue(gaData[0][6]/60);


sheet.getRange(lastRow, 7).setValue(gaData[0][3]);
sheet.getRange(lastRow, 8).setValue(gaData[0][4] + "%\n");
sheet.getRange(lastRow, 9).setValue(gaData[0][5] + "%\n");

sheet.getRange(lastRow, 10).setValue(gaData[0][1]);
sheet.getRange(lastRow, 11).setValue(gaData[0][7]);
sheet.getRange(lastRow, 12).setValue(gaData[0][8]);
sheet.getRange(lastRow, 13).setValue(gaData[0][9]);
  
//-----セルのフォーマット指定-------
//日時データのフォーマット
var FormatDate = sheet.getRange("B:B").setNumberFormat("MM/DD");

//直帰率
var FormatBounceRate = sheet.getRange("H:H").setNumberFormat("0%");

//離脱率
var FormatExitRate = sheet.getRange("I:I").setNumberFormat("0%");

//滞在時間
var FormatDuration = sheet.getRange("F:F").setNumberFormat("#,##00.00");

//ページ/セッション
var FormatPageSession = sheet.getRange("K:K").setNumberFormat("#,##0.000");

// AdsenseフォーマットをGoogle Financeからドル→円変換して表示させる
let strFormula = 'GOOGLEFINANCE("CURRENCY:USDJPY") * M' + lastRow;
strFormula = 'ROUND(' + strFormula + ')';
var FormulaAdsense = sheet.getRange(lastRow, 14).setFormula(strFormula);

}