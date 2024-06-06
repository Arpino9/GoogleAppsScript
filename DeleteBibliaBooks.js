// Spreadsheet
const SheetId = "1wztZPdsMW8dFwit3VJ-g0VhT8eTZA9gggnrMPDpxxpI";
const SheetName = 'シート1';

function myFunction() {
  const sheet = GetSheet();
  const lastRow = sheet.getLastRow();

  const calendar = CalendarApp.getDefaultCalendar();

  for (let i=2; i <= lastRow; i++){
    let bookId  = sheet.getRange(i,1).getValue();

    console.log(bookId)
    const book = calendar.getEventById(bookId);

    if (book == undefined){
      continue;
    }

    try{
      book.deleteEvent();
      console.log("DELETED:" + book.getTitle())
    }catch(e){
      console.log("FAILED:" + book.getTitle())
      continue;
    }finally{
      DeleteSpreadSheetRecord(bookId);
    }
  }
}

function GetSheet(){
  const spreadSheet = SpreadsheetApp.openById(SheetId); 
  return spreadSheet.getSheetByName(SheetName);
}

function DeleteSpreadSheetRecord(bookId){
  if (bookId == ""){
    return;
  }

  const sheet = GetSheet();
  var textFinder = sheet.createTextFinder(bookId);
  if (textFinder == undefined){
    return;
  }

  var cells = textFinder.findAll();

  sheet.deleteRow(cells[0].getRow());
}

function DeleteBookOnCalendar(bookId){
  console.log("Deleting：" + bookId);
  const calendar = CalendarApp.getDefaultCalendar();
  const book = calendar.getEventById(bookId);
  book.deleteEvent();
}