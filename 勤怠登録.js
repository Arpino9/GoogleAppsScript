// ============================================================
// 勤怠登録 GAS ウェブアプリ - Code.gs
// ============================================================
// 【設定】自分のカレンダーIDに変更してください
const CALENDAR_ID = 'primary'; // 'primary' = デフォルトカレンダー
                                // 専用カレンダーなら 'xxxx@group.calendar.google.com'

// 昼食時間の固定設定
const LUNCH_START = '12:00';
const LUNCH_END   = '13:00';

// イベントカラー
const COLOR_WORK  = CalendarApp.EventColor.CYAN;  // 午前・午後勤務
const COLOR_LUNCH = CalendarApp.EventColor.YELLOW; // 昼食

// ------------------------------------------------------------
// doGet: ウェブアプリのエントリーポイント
// ------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('勤怠登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------------------------------------------------
// registerAttendance: フォームから呼ばれる登録処理
// 3イベントに分割して登録する：
//   [1] 午前勤務  : startTime 〜 12:00
//   [2] 昼食      : 12:00    〜 13:00
//   [3] 午後勤務  : 13:00    〜 endTime
// ------------------------------------------------------------
function registerAttendance(data) {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) throw new Error('カレンダーが見つかりません。CALENDAR_IDを確認してください。');

    const dateStr  = data.date;       // 'YYYY-MM-DD'
    const startStr = data.startTime;  // 'HH:MM'
    const endStr   = data.endTime;    // 'HH:MM'

    const startDate     = parseDateTime(dateStr, startStr);
    const endDate       = parseDateTime(dateStr, endStr);
    const lunchStart    = parseDateTime(dateStr, LUNCH_START);
    const lunchEnd      = parseDateTime(dateStr, LUNCH_END);

    // バリデーション
    if (endDate <= startDate) {
      return { success: false, message: '退勤時刻は出勤時刻より後にしてください。' };
    }
    if (startDate >= lunchStart) {
      return { success: false, message: `出勤時刻は昼食開始（${LUNCH_START}）より前にしてください。` };
    }
    if (endDate <= lunchEnd) {
      return { success: false, message: `退勤時刻は昼食終了（${LUNCH_END}）より後にしてください。` };
    }

    // 勤務時間の計算（昼食60分を除く）
    const amMin    = (lunchStart - startDate) / 60000;
    const pmMin    = (endDate - lunchEnd) / 60000;
    const totalMin = amMin + pmMin;
    const memoSuffix = data.memo ? `（${data.memo}）` : '';

    // ① 午前勤務イベント
    const amTitle = `午前勤務 ${formatDuration(amMin)}${memoSuffix}`;
    const amEvent = calendar.createEvent(amTitle, startDate, lunchStart, {
      description: buildDescription(data, startStr, LUNCH_START, amMin, totalMin),
    });
    amEvent.setColor(COLOR_WORK);

    // ② 昼食イベント
    const lunchEvent = calendar.createEvent('昼食', lunchStart, lunchEnd);
    lunchEvent.setColor(COLOR_LUNCH);

    // ③ 午後勤務イベント
    const pmTitle = `午後勤務 ${formatDuration(pmMin)}${memoSuffix}`;
    const pmEvent = calendar.createEvent(pmTitle, lunchEnd, endDate, {
      description: buildDescription(data, LUNCH_END, endStr, pmMin, totalMin),
    });
    pmEvent.setColor(COLOR_WORK);

    const dateLabel = formatDate(startDate);
    return {
      success: true,
      message: `登録しました！\n${dateLabel} ${startStr}〜${endStr}\n実働 ${formatDuration(totalMin)}（昼食除く）`,
    };

  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// ------------------------------------------------------------
// ヘルパー関数
// ------------------------------------------------------------

function parseDateTime(dateStr, timeStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const [h, min]  = timeStr.split(':').map(Number);
  return new Date(y, m - 1, d, h, min, 0);
}

function formatDuration(totalMin) {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return m === 0 ? `${h}h` : `${h}h${m}m`;
}

function formatDate(date) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return `${date.getMonth() + 1}/${date.getDate()}(${days[date.getDay()]})`;
}

function buildDescription(data, segStart, segEnd, segMin, totalMin) {
  const lines = [
    `${segStart}〜${segEnd}（${formatDuration(segMin)}）`,
    '',
    `出勤: ${data.startTime}`,
    `退勤: ${data.endTime}`,
    `実働合計: ${formatDuration(totalMin)}（昼食除く）`,
  ];
  if (data.memo) lines.push(`備考: ${data.memo}`);
  lines.push('', '※ 勤怠登録アプリより自動作成');
  return lines.join('\n');
}
