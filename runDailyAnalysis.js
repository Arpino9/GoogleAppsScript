// ============================================================
// 株式買い時判断システム — GAS スケルトン
// ============================================================
// 【事前準備】
//   1. setup() を手動実行してスクリプトプロパティを設定
//   2. J-Quants: https://jpx-jquants.com で無料アカウント登録
//   3. Claude API Key: https://console.anthropic.com で取得
//   4. トリガー設定: runDailyAnalysis を毎朝8:00に設定
// ============================================================

const PROPS  = PropertiesService.getScriptProperties();
const SS     = SpreadsheetApp.getActiveSpreadsheet();
const LIST   = '銘柄リスト';    // Sheet1 の名前
const REPORT = '分析レポート';  // Sheet2 の名前

// ============================================================
// メインエントリーポイント（タイマートリガーで毎朝実行）
// ============================================================
function runDailyAnalysis() {
  const tickers = getActiveTickers();
  if (tickers.length === 0) return;

  // getJQuantsToken() の呼び出しを丸ごと削除
  const results = [];

  for (const ticker of tickers) {
    try {
      const stockData = fetchStockData(ticker.code); // ← 引数1つに
      if (!stockData) continue;
      const analysis = analyzeWithClaude(ticker, stockData);
      appendToReport(buildReportRow(ticker, stockData, analysis));
      results.push({ ticker, stockData, analysis });
      Utilities.sleep(1500);
    } catch (e) {
      Logger.log(`${ticker.code} エラー: ${e.message}`);
    }
  }

  if (results.length > 0) sendGmailReport(results);
}

// ============================================================
// ① 銘柄リストの読み込み
// ============================================================
function getActiveTickers() {
  const sheet = SS.getSheetByName(LIST);
  const data  = sheet.getDataRange().getValues();

  // ヘッダー行（1行目）をスキップ。D列（index=3）が TRUE/✓ の行のみ対象
  return data.slice(1)
    .filter(row => row[0] && (row[3] === true || row[3] === 'TRUE' || row[3] === '✓'))
    .map(row => ({
      code:   String(row[0]).padStart(4, '0'), // 銘柄コード（4桁）
      name:   row[1],
      sector: row[2],
      memo:   row[4] || ''
    }));
}

// ============================================================
// ③ 株価データ取得 & テクニカル指標計算
// ============================================================
function fetchStockData(code) {
  // 東証の銘柄コードに .T を付けると Yahoo Finance が認識する
  // 例: 7203 → 7203.T（トヨタ）
  const ticker = code + '.T';
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${ticker}`
    + `?interval=1d&range=3mo`; // 3ヶ月分の日足を取得

  const res = UrlFetchApp.fetch(url, {
    headers:            { 'User-Agent': 'Mozilla/5.0' },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log(`${code} HTTPエラー: ${res.getResponseCode()} ${res.getContentText()}`);
    return null;
  }

  const result = JSON.parse(res.getContentText()).chart?.result?.[0];
  if (!result) return null;

  const quote      = result.indicators.quote[0];
  const timestamps = result.timestamp;
  const closes     = quote.close.filter(v => v != null);
  const volumes    = quote.volume.filter(v => v != null);

  if (closes.length < 15) {
    Logger.log(`${code}: データ不足（${closes.length}件）`);
    return null;
  }

  // タイムスタンプ（UNIX秒）→ 日付文字列に変換
  const latestTs  = timestamps[timestamps.length - 1];
  const latestDate = Utilities.formatDate(
    new Date(latestTs * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'
  );

  return {
    date:      latestDate,
    close:     Math.round(closes[closes.length - 1] * 10) / 10,
    prevClose: Math.round(closes[closes.length - 2] * 10) / 10,
    volume:    Math.round(volumes[volumes.length - 1]),
    avgVolume: avg(volumes.slice(-20)),
    rsi:       calcRSI(closes, 14),
    ma25:      avg(closes.slice(-25)),
    ma75:      closes.length >= 75 ? avg(closes.slice(-75)) : avg(closes)
  };
}

// testSingleStock も apiKey を使うよう修正
function testSingleStock() {
  const apiKey = PROPS.getProperty('J_QUANTS_API_KEY');
  const data   = fetchStockData(apiKey, '7203');
  Logger.log('株価データ: ' + JSON.stringify(data));

  const analysis = analyzeWithClaude(
    { code: '7203', name: 'トヨタ自動車', sector: '自動車', memo: '' },
    data
  );
  Logger.log('AI分析結果: ' + JSON.stringify(analysis));
}

// RSI（単純平均方式）
function calcRSI(closes, period) {
  if (closes.length < period + 1) return null;
  let gains = 0, losses = 0;
  for (let i = closes.length - period; i < closes.length; i++) {
    const d = closes[i] - closes[i - 1];
    if (d > 0) gains += d; else losses += Math.abs(d);
  }
  const avgGain = gains / period;
  const avgLoss = losses / period;
  if (avgLoss === 0) return 100;
  return Math.round(100 - 100 / (1 + avgGain / avgLoss));
}

function avg(arr) {
  if (!arr || arr.length === 0) return null;
  return Math.round(arr.reduce((a, b) => a + b, 0) / arr.length * 10) / 10;
}

function fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// ============================================================
// ④ Claude API で AI 分析
//    Haiku モデルを使用（高速・低コスト）
// ============================================================
function analyzeWithClaude(ticker, data) {
  const changeRate    = data.prevClose
    ? ((data.close - data.prevClose) / data.prevClose * 100).toFixed(2)
    : 'N/A';
  const volumeRatio   = data.avgVolume
    ? (data.volume / data.avgVolume).toFixed(2)
    : 'N/A';
  const vs25 = data.ma25 ? (data.close > data.ma25 ? '上' : '下') : '-';
  const vs75 = data.ma75 ? (data.close > data.ma75 ? '上' : '下') : '-';

  const prompt = `あなたは日本株のテクニカル分析の専門家です。
以下のデータをもとに、短期（数日〜1週間）の買い時かどうかを判断してください。

【銘柄】${ticker.name}（${ticker.code}）  業種: ${ticker.sector}
【直近データ（${data.date}）】
- 終値: ${data.close}円  前日比: ${changeRate}%
- 出来高: ${data.volume}（20日平均比 ${volumeRatio}倍）
- RSI(14): ${data.rsi}
- 25日MA: ${data.ma25}円（終値は${vs25}）
- 75日MA: ${data.ma75}円（終値は${vs75}）
${ticker.memo ? `- 備考: ${ticker.memo}` : ''}

以下の JSON 形式のみで返してください（前後の文章や \`\`\` 不要）:
{
  "verdict": "買い寄り" | "様子見" | "見送り",
  "score": 1〜10の整数（10が最も強い買いシグナル）,
  "comment": "分析コメント（80字以内）",
  "risk": "主なリスク（30字以内）"
}`;

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'x-api-key':            PROPS.getProperty('CLAUDE_API_KEY'),
      'anthropic-version':    '2023-06-01'
    },
    payload: JSON.stringify({
      model:      'claude-haiku-4-5-20251001', // 高速・低コストモデル
      max_tokens: 256,
      messages:   [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  try {
    const raw  = JSON.parse(res.getContentText()).content[0].text;
    const json = raw.replace(/```json|```/g, '').trim();
    return JSON.parse(json);
  } catch (e) {
    Logger.log(`Claude パースエラー: ${e.message}`);
    return { verdict: '様子見', score: 5, comment: '分析エラー（要確認）', risk: '-' };
  }
}

// ============================================================
// ⑤ レポートシートに追記
// ============================================================
function buildReportRow(ticker, data, analysis) {
  const now        = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const changeRate = data.prevClose
    ? ((data.close - data.prevClose) / data.prevClose * 100).toFixed(2) + '%'
    : '-';
  return [
    now,               // A: 実行日時
    ticker.code,       // B: 銘柄コード
    ticker.name,       // C: 銘柄名
    data.close,        // D: 終値
    changeRate,        // E: 前日比(%)
    data.rsi,          // F: RSI
    analysis.verdict,  // G: 判定
    analysis.comment,  // H: AI分析コメント
    analysis.risk,     // I: 主なリスク
    analysis.score     // J: スコア(1-10)
  ];
}

function appendToReport(row) {
  SS.getSheetByName(REPORT).appendRow(row);
}

// ============================================================
// ⑥ Gmail 送信
// ============================================================
function sendGmailReport(results) {
  const today   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M月d日(E)');
  const subject = `[株式分析] ${today} の買い時レポート`;

  const buy   = results.filter(r => r.analysis.verdict === '買い寄り');
  const watch = results.filter(r => r.analysis.verdict === '様子見');
  const pass  = results.filter(r => r.analysis.verdict === '見送り');

  const section = (label, items) => {
    if (items.length === 0) return '';
    let s = `■ ${label}（${items.length}件）\n${'─'.repeat(40)}\n`;
    items.sort((a, b) => b.analysis.score - a.analysis.score)
      .forEach(({ ticker, stockData, analysis }) => {
        const chg = stockData.prevClose
          ? ((stockData.close - stockData.prevClose) / stockData.prevClose * 100).toFixed(2)
          : '-';
        s += `▶ ${ticker.name}（${ticker.code}）スコア: ${analysis.score}/10\n`;
        s += `   終値 ${stockData.close}円（${chg}%）  RSI: ${stockData.rsi}\n`;
        s += `   ${analysis.comment}\n`;
        s += `   リスク: ${analysis.risk}\n\n`;
      });
    return s;
  };

  const body = [
    `${'━'.repeat(44)}`,
    `  ${today} 株式買い時分析レポート`,
    `${'━'.repeat(44)}`,
    '',
    section('買い寄り', buy),
    section('様子見',   watch),
    section('見送り',   pass),
    '─'.repeat(44),
    '※ 本メールはGASで自動生成されています。',
    '※ 投資判断は必ずご自身の責任でお願いします。'
  ].join('\n');

  GmailApp.sendEmail(
    PROPS.getProperty('NOTIFY_EMAIL'),
    subject,
    body
  );
  Logger.log(`Gmailを送信しました: ${subject}`);
}

// ============================================================
// セットアップ用ヘルパー（初回のみ手動で実行する）
// ============================================================
function setup() {
  PROPS.setProperties({
    'CLAUDE_API_KEY': 'sk-ant-XXXXX',
    'NOTIFY_EMAIL':   'XXXXX@XXXX.com'
  });
  Logger.log('セットアップ完了');
}

// ------------------------------------------------------------
// 以下テスト用
// ============================================================
// 手動で1銘柄を設定してテストする
// ============================================================
function testSingleStock() {
  const data = fetchStockData('7203');
  Logger.log('株価データ: ' + JSON.stringify(data));

  if (!data) return;
  const analysis = analyzeWithClaude(
    { code: '7203', name: 'トヨタ自動車', sector: '自動車', memo: '' },
    data
  );
  Logger.log('AI分析結果: ' + JSON.stringify(analysis));
}

// ============================================================
// Claude APIの応答テスト(ステータスが200なら成功、400なら失敗)
// ============================================================
function debugClaude() {
  const ticker = { code: '7203', name: 'トヨタ自動車', sector: '自動車', memo: '' };
  const data   = fetchStockData('7203');

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         PROPS.getProperty('CLAUDE_API_KEY'),
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 256,
      messages:   [{ role: 'user', content: 'テスト: 「OK」とだけ返してください' }]
    }),
    muteHttpExceptions: true
  });

  Logger.log('ステータス: ' + res.getResponseCode());
  Logger.log('レスポンス: ' + res.getContentText()); // ← ここを確認
}