// CSV
const FolderId_CSV = '1R4PViHvJ2bg0zW7nKJSrvxF7s4ZocAgp';

// Spreadsheet
const SheetId = "1wztZPdsMW8dFwit3VJ-g0VhT8eTZA9gggnrMPDpxxpI";
const SheetName = 'シート1';

function myFunction() {
  const values = GetCSVValues(FolderId_CSV);
  
  SortSpreadSheet();
  Execute(values);

  // プロパティに書き込む
  PropertiesService.getScriptProperties().setProperty("TemporaryID", "None"); 
}

/*
  本の評価を取得
 */
function GetRating(rating){
  switch (rating){
    case "0": return "☆☆☆☆☆";

    case "1": return "★☆☆☆☆";

    case "2": return "★★☆☆☆";

    case "3": return "★★★☆☆";

    case "4": return "★★★★☆";

    case "5": return "★★★★★";

    default:  return "☆☆☆☆☆";
  }
}

/*
  シートを取得
 */
function GetSheet(){
  const spreadSheet = SpreadsheetApp.openById(SheetId); 
  return spreadSheet.getSheetByName(SheetName);
}

/*
  イベントのIDを取得
 */
function SearchEventID(title){
  const sheet = GetSheet();
  const lastRow = sheet.getLastRow();

  var range = sheet.getRange(1,1, lastRow, 5)

  var values = range.getValues().map(value => {
                 return {
                   id       : value[0],
                   title    : value[1],
                   author   : value[2],
                   publisher: value[3],
                   date     : value[4]
                 }
               });
 var results = values.filter(value => value.title == title);
 
  if (results    == undefined ||
      results[0] == undefined){
    return undefined;
  }

  return results[0].id;
}

function ExistsID(id){
  const sheet = GetSheet();
  var cells = sheet.createTextFinder(id).findAll();

  return cells[0] != undefined;
}

/*
  IDとマッチする本の日付を取得
 */
function SearchBookDate(id){
  const sheet = GetSheet();
  var cells = sheet.createTextFinder(id).findAll();

  return (sheet.getRange(cells[0].getRow(),5).getValue());
}

/*
  CSVの情報を取得
 */
function GetCSVValues(folderId){
  const folder = DriveApp.getFolderById(folderId);//csvを格納したフォルダのID
  const files  = folder.getFiles();
  const file   = files.next();
  const fileId = file.getId();

  const blob   = DriveApp.getFileById(fileId).getBlob();
  const csv    = blob.getDataAsString();
  return Utilities.parseCsv(csv);
}

/*
  スプレッドシートの読了記録をIDから検索して削除する
 */
function DeleteSpreadSheetRecord(title, author){
  let id = SearchEventID(title, author);

  const sheet = GetSheet();
  var textFinder = sheet.createTextFinder(id);
  var cells      = textFinder.findAll();

  sheet.deleteRow(cells[0].getRow());
}

/*
 過去に登録された書籍データと比較
*/
function UpdateSpreadSheetRecord(bookID, title, author, publisher, date, details, caption, thumbnail, rating){
  let id = SearchEventID(title, author);

  if (id == undefined){
    return true;
  }

  console.log(id);
  const sheet = GetSheet();
  var textFinder = sheet.createTextFinder(id);
  var cells = textFinder.findAll();

  let foundTitle     = sheet.getRange(cells[0].getRow(),2).getValue();
  let foundAuthor    = sheet.getRange(cells[0].getRow(),3).getValue();
  let foundPublisher = sheet.getRange(cells[0].getRow(),4).getValue();
  let foundDate      = sheet.getRange(cells[0].getRow(),5).getDisplayValue();
  let foundDetails   = sheet.getRange(cells[0].getRow(),6).getValue();
  let foundCaption   = sheet.getRange(cells[0].getRow(),7).getValue();
  let foundThumbnail = sheet.getRange(cells[0].getRow(),8).getValue();
  let foundRating    = sheet.getRange(cells[0].getRow(),9).getValue();

  if((foundTitle     == title) &&
     (foundAuthor    == author) &&
     (foundPublisher == publisher) &&
     (foundDate      == date) &&
     (foundDetails   == details) &&
     (foundCaption   == caption) &&
     (GetRating(foundRating)    == rating)){
      // 変更なし
      console.log("既存の読書データと一致：" + title);
      return true;
     }

  if((foundTitle      == title) &&
     (foundAuthor     == author)){
        console.log("既存の読書データを修正：" + title);

        sheet.getRange(cells[0].getRow(),1).setValue = bookID;

      if (foundPublisher == publisher){
        sheet.getRange(cells[0].getRow(),4).setValue = publisher;
      }
      if (foundDate == date){
        sheet.getRange(cells[0].getRow(),5).setValue = date;
      }
      if (foundDetails == details){
        sheet.getRange(cells[0].getRow(),6).setValue = details;
      }
      if (foundCaption == caption){
        sheet.getRange(cells[0].getRow(),7).setValue = caption;
      }
      if (foundThumbnail == thumbnail){
        sheet.getRange(cells[0].getRow(),8).setValue = thumbnail;
      }
      if (foundRating == rating){
        sheet.getRange(cells[0].getRow(),9).setValue = rating;
      }
      // タイトル・著者以外が変更されている
      return false;
  }

  return true;
}

/*
  スプレッドシートを取得
 */
function SortSpreadSheet(){
  const sheet = GetSheet();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  var range = sheet.getRange(2,1, lastRow, lastColumn)

  range.sort({column: 5, ascending: false});
}

/*
  メイン処理(カレンダーの本情報と異なっている場合は更新)
 */
function Execute(values){
  var books  = values.map(value => {
                return {
                  title    : value[0],
                  author   : value[2],
                  publisher: value[4],
                  date     : value[6],
                  details  : value[7],
                  caption  : value[8],
                  thumbnail: value[9],
                  read     : value[12],
                  rating   : value[13]
                }
             })
  books = books.sort(value => value.date);
  
  var temporaryID = PropertiesService.getScriptProperties().getProperty('TemporaryID')

  if ((temporaryID != "None") && 
      (temporaryID != null)){
    if (ExistsID(temporaryID)){
      var date = Utilities.formatDate(SearchBookDate(temporaryID), "JST", "yyyy/MM/dd")
      books = books.filter(value => value.date <= date);
    }
  }

  for (book of books){
    let title     = book.title;
    let author    = book.author;
    let publisher = book.publisher;
    let date      = book.date;
    let details   = book.details;
    let caption   = book.caption;
    let thumbnail = book.thumbnail;
    let read      = book.read;
    let rating    = GetRating(book.rating);

    if (read == 1){
      continue;
    }

    // タイトルと著者名で検索
    let bookID = SearchEventID(title);

    if (bookID == undefined){
      // 未登録
      DebugSearchingBook(book);
      var id = CreateEvent(book);
      ReportSpreadSheet(id, book.title, book.author, book.publisher, book.date, book.details, book.caption, book.thumbnail, book.rating);
      continue;
    }

    let isSame = UpdateSpreadSheetRecord(bookID, title, author, publisher, date, details, caption, thumbnail, rating);

    if (isSame == true){
      // 更新不要
      PropertiesService.getScriptProperties().setProperty("TemporaryID", bookID); 
      continue;
    }else{
      // 要更新
      let canRegister = DeleteBookOnCalendar(bookID);
      
      if (canRegister == true){
        CreateEvent(book);
      }
    }
  }
}

function DeleteBookOnCalendar(bookId){
  const calendar = CalendarApp.getDefaultCalendar();
  const book = calendar.getEventById(bookId);

  try{
     book.deleteEvent();
     console.log("DELETED SUCCESSED:" + book.getTitle())

     return true;
  }catch(e){
     console.log("DELETE FAILED:" + book.getTitle())

     return false;
  }
}

/* 
 デバッグ用(不要ならコメントアウト) 
 */
function DebugSearchingBook(book){
  console.log('日付：'    + book.date);
  console.log('タイトル：' + book.title);
  console.log('著者'      + book.author);
  console.log('出版社'    + book.publisher);
  console.log('詳細情報'  + book.details);
  console.log('本の概要'  + book.caption);
  console.log('本の評価'  + book.rating);
}

/* 
 イベントの登録 
 */
function CreateEvent(book){
  let title          = book.title;
  let author         = book.author;
  let publisher      = book.publisher;
  let registeredDate = book.date;
  let details        = book.details;
  let caption        = book.caption;  // デバッグすると重い
  let thumbnail      = book.thumbnail;
  let rating         = GetRating(book.rating);

  const calendar = CalendarApp.getDefaultCalendar();
  calendar.setTimeZone = 'Asia／Tokyo';

  // タグを作成
  var event = calendar.createAllDayEvent(title,new Date(registeredDate));
  var id = event.getId();

  event.setColor(CalendarApp.EventColor.PALE_GREEN);

  let description = '【著者】\n' + author + '\n' 
                  + '\n【出版社】\n' + publisher + '\n' 
                  + '\n【詳細情報】\n' + details + '\n' 
                  + '\n【本の概要】\n' + caption + '\n'
                  + '\n【本の評価】\n' + rating + '\n' 
                  + '\n【登録ID】\n' + id + '\n'
                  + '\n【サムネイル】\n' + thumbnail
  event.setDescription(description);

  // プロパティに書き込む
  PropertiesService.getScriptProperties().setProperty("TemporaryID", id);

  console.log("ADDED CALENDAR：" + title + "　ID：" + id);

  return id; 
}

/* 
 Excelに保存 
 */
function ReportSpreadSheet(id, title, author, publisher, date, details, caption, thumbnail, rating){
  const spreadSheet = SpreadsheetApp.openById(SheetId); 
  const sheet = spreadSheet.getSheetByName(SheetName);
  const lastRow = sheet.getLastRow();
  console.log("LASTROW:" + lastRow)

  sheet.getRange(lastRow + 1, 1).setValue(id);
  sheet.getRange(lastRow + 1, 2).setValue(title);
  sheet.getRange(lastRow + 1, 3).setValue(author);
  sheet.getRange(lastRow + 1, 4).setValue(publisher);
  sheet.getRange(lastRow + 1, 5).setValue(date);
  sheet.getRange(lastRow + 1, 6).setValue(details);
  sheet.getRange(lastRow + 1, 7).setValue(caption);
  sheet.getRange(lastRow + 1, 8).setValue(thumbnail);
  sheet.getRange(lastRow + 1, 9).setValue(rating);
}