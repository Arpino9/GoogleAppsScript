// Spreadsheet
const SheetId = "1JFRj_QV6xyKOQxYW2toAw_FBtoJj_Sg-6fBHindsiF4";

function myFunction() {
  const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

  FetchChildPages(SCRIPT_PROPERTIES.getProperty('PARENT_PAGE_ID_BLACK'), '黒本');
  FetchChildPages(SCRIPT_PROPERTIES.getProperty('PARENT_PAGE_ID_PURPLE'), '紫本');
  FetchChildPages(SCRIPT_PROPERTIES.getProperty('PARENT_PAGE_ID_WHITE'), '白本');
}

/*
  親ページから子ページを全て抽出する
*/
function FetchChildPages(pageID, sheetName) {
  const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
  const NOTION_TOKEN   = SCRIPT_PROPERTIES.getProperty('NOTION_TOKEN');

  const url = `https://api.notion.com/v1/blocks/${pageID}/children?page_size=100`;

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${NOTION_TOKEN}`,
      'Notion-Version': '2022-06-28'
    }
  });

  const json = JSON.parse(res.getContentText());

  // child_page だけ抽出
  const childPages = json.results
    .filter(b => b.type === 'child_page')
    .map(b => ({
      id: b.id,
      title: b.child_page.title
    }));

  const sheet = GetSheet(sheetName);
  ClearSheet(sheetName)

  for(const childPage of childPages){    
    const blocks = FetchAllBlocks(childPage.id);
    ExtractTextFromBlocks(sheet, childPage.title, blocks);
  }
}

/*
  1ページのブロックを全て抽出する
*/
function FetchAllBlocks(pageId) {
  const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
  const NOTION_TOKEN = SCRIPT_PROPERTIES.getProperty('NOTION_TOKEN');

  let allBlocks = [];
  let cursor = null;

  do {
    let url = `https://api.notion.com/v1/blocks/${pageId}/children?page_size=100`;
    if (cursor) {
      url += `&start_cursor=${cursor}`;
    }

    const res = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': `Bearer ${NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28'
      }
    });

    const json = JSON.parse(res.getContentText());

    allBlocks = allBlocks.concat(json.results);
    cursor = json.next_cursor;

  } while (cursor);

  return allBlocks;
}

/*
  Notionの1ブロックごとの書き込み用データを作る
*/
function ExtractTextFromBlocks(sheet, title, blocks) {
  const rows = [];
 
  let currentRow = [
      title, // A列
      '',    // B列
      '',    // C列
      '',    // D列
      '',    // E列
      ''     // F列
    ];

  for (const block of blocks) {  
    if (
      block.type !== 'paragraph' &&
      block.type !== 'heading_2' &&
      block.type !== 'to_do'
    ) continue;

    switch (block.type) {
      case 'paragraph':
        let text = RichTextToText(block.paragraph.rich_text);

        const values = text.split(' ');

        currentRow[3] = values[0];

        if (values.length >= 3){
          currentRow[4] = values[1];
        }

        if (values.length >= 4){
          currentRow[5] = values[2];
        }

        break;

      case 'heading_2':
        currentRow[1] = RichTextToText(block.heading_2.rich_text);
        break;

      case 'to_do':
        if (block.to_do.checked === true){
          currentRow[2] = '〇';
        }else{
          currentRow[2] = '×';
        }

        // ここまででcurrentRowが出そろったので、配列に入れる
        rows.push(currentRow);
        currentRow = [title, '', '', '', '', ''];

        break;

      default:
        // 未対応ブロックは無視
        break;
    }
  }

  const startRow = sheet.getLastRow() + 1;

  sheet.getRange(startRow, 1, rows.length, rows[0].length)
       .setValues(rows);
}

/*
  リッチテキストからテキストを抽出する
*/
function RichTextToText(richText) {
  return (richText || []).map(rt => rt.plain_text || '').join('');
}

/*
  シートを取得
 */
function GetSheet(sheetName){
  const spreadSheet = SpreadsheetApp.openById(SheetId); 
  return spreadSheet.getSheetByName(sheetName);
}

/*
  シートのクリア
 */
function ClearSheet(sheetName){
  let sheet = GetSheet(sheetName);
  sheet.getRange(2, 1, sheet.getLastRow() + 1, sheet.getLastColumn()).clearContent();
}