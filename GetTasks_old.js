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
// シートID(シートURLの「https://docs.google.com/spreadsheets/d/」から「/edit#gid=0」までの文字列)
const SpreadSheetId = "";
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
  var taskLists = Tasks.Tasklists.list();

  if (!taskLists.items)
  {
    Logger.log('No task lists found.');
  }

  // シートのクリア
  Sheet.getRange(OutPutStartRow, 1, Sheet.getLastRow() + 1, ColumnCount).clearContent();
  taskLists.items.sort();

  for (var i = 0; i < taskLists.items.length; i++) 
  {
    var taskList = taskLists.items[i];

    if (taskList.title.match(/基本情報/) || 
        taskList.title.match(/応用情報/) ||
        taskList.title.match(/データベーススペシャリスト/) ||
        taskList.title.match(/ネットワークスペシャリスト/) ||
        taskList.title.match(/プロジェクトマネージャ/) ||
        taskList.title.match(/登録セキスペ/))
    {
      /*
      if (taskList.title == "登録セキスペ_H24春"){
          console.log("aa");
      }
      */

      GetTasks(taskList.id, taskList.title)
    }
  }

  // 該当する問題リストをソートして出力
  Questions.sort();
  Sheet.getRange(OutPutStartRow,1, Questions.length, Questions[0].length).setValues(Questions);
}

/**
 * タスクの取得
 */
function GetTasks(ToDoListID, TaskTitle) {
  var now = new Date();
  var completedMax = new Date();
  var completedMin = new Date();
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

  const tasks = Tasks.Tasks.list(ToDoListID,options).getItems().sort();

  if (!tasks) 
  {
    console.log('登録されているタスクはありません');
  }

  for (let i = 0; i < tasks.length; i++) 
  {
    let gotTask = tasks[i];

    if (gotTask.due == undefined)
    {
      continue;
    }

    let dueDate   = new Date(gotTask.due.slice(0, 10));

    if (dueDate < completedMax)
    {
      Questions.push([TaskTitle, gotTask.title, dueDate]);
    }
  }
}
