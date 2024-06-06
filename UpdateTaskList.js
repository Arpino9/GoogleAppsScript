//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// 要設定項目
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
const Title      = "Java SE Silver 紫本";   // 取得するリスト名
const SubTitle   = '【1章】';            // 追加するタスクのサブタイトル
const TaskCount  = 6;                     // 追加するタスク数

function Add(taskId){
  for (var i = TaskCount; i >= 1; i--){
    var task = {
                  title: SubTitle + '問' + i.toString().padStart(2, '0')
                }

    task = Tasks.Tasks.insert(task, taskId);
  }
}

function UpdateTaskList() {
  // タスクリストからタスクを取得
  let taskLists = Tasks.Tasklists.list();

  if (taskLists.items) {
    for (var i = 0; i < taskLists.items.length; i++) {
      if (taskLists.items[i].title == Title){
          Add(taskLists.items[i].id);
      }
    }
  } else {
    Logger.log('No task lists found.');
  }
}