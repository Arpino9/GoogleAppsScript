//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
// 要設定項目
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
const Title = "登録セキスペ_R06春_午前";   // 1.リスト名
const TaskCount  = 80;                // 2.タスク数

function AddTaskList() {
  var newList ={
    "title" : Title
  };

  var list = Tasks.Tasklists.insert(newList);

  for (var i = TaskCount; i >= 1; i--){
  var task = {
                title: '問' + i.toString().padStart(2, '0')
              }
    console.log(i);
    task = Tasks.Tasks.insert(task, list.id);
  }
}