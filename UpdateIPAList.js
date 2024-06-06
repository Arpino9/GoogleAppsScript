/**
 * 編集用
 */
const Questions = [];

// nヶ月前から取得
const CompletedMin = 30;
// nヶ月前まで取得
const CompletedMax = 12;

// シート名
const SheetName = "ToDo";
const qSheetName = "IPA問題一覧";

// シートID(シートURLの「https://docs.google.com/spreadsheets/d/」から「/edit#gid=0」までの文字列)
const SpreadSheetId = "1tc5uFTh09PBVVnV2OYmGZ3svY6C-6SwCAF6KIUO8l9c";
// カラム数
const ColumnCount = 3;
// 出力開始行
let OutPutStartRow = 2;

/**
 * タスクリストの取得
 */
// シート定義
let Sheet       = SpreadsheetApp.openById(SpreadSheetId).getSheetByName(SheetName);

function GetTaskLists() {
  const taskLists = Tasks.Tasklists.list();

  if (!taskLists.items)
  {
    Logger.log('No task lists found.');
  }

  // 昇順ソート
  const sortedLists   = taskLists.items.sort((a, b) => a.title.localeCompare(b.title));

  // 読み込み済みの試験名がある場合は、読込中の試験から読み込む
  const readTitle = PropertiesService.getScriptProperties().getProperty('ReadingTitle');

  let filteredLists;
  if(readTitle != "None"){
    filteredLists = sortedLists.filter(list => list.title.localeCompare(readTitle) >= 0);
  }else{
    filteredLists = sortedLists;
  }

  var length = length = filteredLists.length;

  for (var i = 0; i < length; i++) 
  {
    var taskList = filteredLists[i];

    if (taskList.title.match(/基本情報/) || 
        taskList.title.match(/応用情報/) ||
        taskList.title.match(/データベーススペシャリスト/) ||
        taskList.title.match(/ネットワークスペシャリスト/) ||
        taskList.title.match(/プロジェクトマネージャ/) ||
        taskList.title.match(/登録セキスペ/))
    {
      PropertiesService.getScriptProperties().setProperty("ReadingTitle", taskList.title); 

      GetTasks(taskList.id, taskList.title)

      console.log("【完了】" + taskList.title);
  
      PropertiesService.getScriptProperties().setProperty("LoadingQIndex", "0");
    }
  }

  PropertiesService.getScriptProperties().setProperty("ReadingTitle", "None"); 
}

/**
 * タスクの取得
 */
function GetTasks(ToDoListID, TaskTitle) {
  const now = new Date();
  const completedMax = new Date();
  const completedMin = new Date();
  completedMax.setMonth(now.getMonth() - CompletedMax);
  completedMin.setMonth(now.getMonth() - CompletedMin);

  // タスク取得オプション
  const options = 
  {
    showCompleted: true,
    showDeleted  : false,
    showHidden   : true,
    //completedMax : completedMax.toISOString(),
    //completedMin : completedMin.toISOString(),
    //dueMax : completedMax.toISOString(),
    //dueMin : completedMin.toISOString(),
    maxResults   : 100,
  };

  const tasks = Tasks.Tasks.list(ToDoListID,options).getItems().sort((a, b) => a.title.localeCompare(b.title));;

  if (!tasks) 
  {
    console.log('登録されているタスクはありません');
  }

  const loadingIndex = PropertiesService.getScriptProperties().getProperty('LoadingQIndex');

  let i = 0;
  if (Number(loadingIndex) != 0){
      i = Number(loadingIndex)
  }

  for (i; i < tasks.length; i++) 
  {
    const gotTask = tasks[i];

    const text = TaskTitle + "_" + gotTask.title;
    const arr = text.split('_');

    const examName = arr[0];
    const year     = arr[1].slice(0, 3);
    const period   = arr[1].slice(3);
    let examType = "";
    let qNo      = "";
    let result   = ""
    let dueDate  = ""

    if (examName == "基本情報" || 
        examName == "応用情報")
    {
      // 基本・応用情報
      examType = arr[2];
      qNo      = arr[3];
    }
    else
    {
      // 高度情報
      examType = arr[2].slice(1, 4);
      qNo      = arr[2].slice(5);
    }

    if (gotTask.due == undefined)
    {
      // 不正解
      result = "×"
    }
    else
    {
      // 正解
      result = "〇"
      dueDate   = new Date(gotTask.due.slice(0, 10));
    }

    UpdateSheet(examName, year, period, examType, qNo, result, dueDate);

    SpreadsheetApp.flush();

    PropertiesService.getScriptProperties().setProperty("LoadingQIndex", i); 
  }
}

function UpdateSheet(examName, year, period, examType, qNo, result, dueDate) {
  let Sheet = SpreadsheetApp.openById(SpreadSheetId).getSheetByName(qSheetName);

  const values = Sheet.getRange("B:F").getValues();

  let filteredCellRow = [];

  let column = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (examName == row[0] && 
        year     == row[1] && 
        period   == row[2] && 
        examType == row[3] && 
        qNo      == row[4]) {
      var cellAddress = Sheet.getRange(i + 1, column + 2).getRow();

      filteredCellRow.push(cellAddress);

      column += column;
    }
  }

  console.log("試験名："  + examName + '\n' +
              "年度  ："  + year + '\n' +
              "時期  ："  + period + '\n' +
              "試験区分：" + examType + '\n' +
              "設問  ："  + qNo + '\n' +
              "回答結果：" + result + '\n' +
              "回答日時：" + dueDate + '\n\n' +
              "Row：" + filteredCellRow);

  Sheet.getRange("K" + filteredCellRow).setValue(result);
  Sheet.getRange("L" + filteredCellRow).setValue(dueDate);
  Sheet.getRange("M" + filteredCellRow).setValue(Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'));
}
