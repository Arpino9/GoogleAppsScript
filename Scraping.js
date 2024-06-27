function myFunction() {
  const url = "https://www.ap-siken.com/kakomon/06_haru/";
  const res = UrlFetchApp.fetch(url);
  const SheetID = "";

  const context = res.getContentText();
  console.log(context);

  const title = context.substring(context.search("<h2>") + 4, 
                                  context.search("</h2>"));

  const text = context.substring(context.search("qtable start"), 
                                 context.search("qtable end"));

  // シート作成
  const sheet = SpreadsheetApp.openById(SheetID);
  const newSheet = sheet.insertSheet();
  newSheet.setName(title);

  // ヘッダ
  newSheet.getRange((1,1),(1,1)).setValue("No");
  newSheet.getRange((1,1),(2,2)).setValue("論点");
  newSheet.getRange((1,1),(3,3)).setValue("分類");

  // 新しいシートを一番最後に移動
  sheet.moveActiveSheet(sheet.getSheets().length);

  let rowCnt = 2;

  for (var row of text.split('<tr>')){
   var column = row.split('<td>');

   if (column.length < 3){
    continue;
   }

    var qNo   = column[1].substring(column[1].search(">") + 1, column[1].search("</a>"));
    var qName = column[2];
    var qSort = column[3];

    newSheet.getRange((rowCnt,rowCnt),(1,1)).setValue(qNo);
    newSheet.getRange((rowCnt,rowCnt),(2,2)).setValue(qName);
    newSheet.getRange((rowCnt,rowCnt),(3,3)).setValue(qSort);

    rowCnt++;
  }

  newSheet.autoResizeColumn(1);
  newSheet.autoResizeColumn(2);
  newSheet.autoResizeColumn(3);
}
