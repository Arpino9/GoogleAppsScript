/**
 * 編集用
 */
const Questions = [];

// シート名
const SheetName = "タスク一覧";
// シートID(シートURLの「https://docs.google.com/spreadsheets/d/」から「/edit#gid=0」までの文字列)
const SpreadSheetId = "";
// カラム数
const ColumnCount = 2;
// 出力開始行
let OutPutStartRow = 2;

/**
 * タスクリストの取得
 */
// シート定義
let Sheet       = SpreadsheetApp.openById(SpreadSheetId).getSheetByName(SheetName);

function GetTaskLists() {
  var taskLists = Tasks.Tasklists.list();

  if (!taskLists.items)
  {
    Logger.log('No task lists found.');
  }

  // シートのクリア
  Sheet.getRange(OutPutStartRow, 1, Sheet.getLastRow() + 1, ColumnCount).clearContent();

  var list = Tasks.Tasklists.list();
  list.items.forEach(function(item) 
  {
    Questions.push([item.title, item.id]);
    Logger.log(item.title + " : " + item.id);
  });

  taskLists.items.sort();

  Questions.sort();
  Sheet.getRange(OutPutStartRow,1, Questions.length, Questions[0].length).setValues(Questions);
}
