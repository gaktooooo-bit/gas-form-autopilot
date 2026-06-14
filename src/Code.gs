/**
 * Code.gs  ―  メイン処理
 * ------------------------------------------------------------------
 * Googleフォームに回答が送信されるたびに onFormSubmit が自動で動き、
 *   1. 回答を「問い合わせ管理」台帳に受付番号付きで整形して追記
 *   2. 申込者へ自動返信メールを送信
 *   3. 社内へ Slack（またはメール）で通知
 * を一気通貫で行います。
 *
 * トリガーの登録は setupTriggers() を一度実行するだけ（README参照）。
 * ------------------------------------------------------------------
 */

/**
 * 【初回のみ実行】必要なトリガーをまとめて登録します。
 * - フォーム送信時トリガー（onFormSubmit）
 * - 日次レポートトリガー（sendDailyReport / Report.gs）
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 二重登録を防ぐため、同名の既存トリガーを一旦削除
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'onFormSubmit' || fn === 'sendDailyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // フォーム送信トリガー
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // 日次レポート（毎朝8時台）
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  Logger.log('トリガーを登録しました。');
}

/**
 * フォーム送信時に自動実行されるメイン関数。
 * @param {Object} e フォーム送信イベント（e.namedValues に回答が入る）
 */
function onFormSubmit(e) {
  try {
    // 1) 回答を取り出す（手動テスト時は readLatestResponse_ で代用）
    const answers = e && e.namedValues
      ? flattenNamedValues_(e.namedValues)
      : readLatestResponse_();

    // 2) 受付番号を採番
    const ticket = issueTicketNumber_();

    // 3) 管理台帳へ整形して追記
    const record = appendToManagementSheet_(ticket, answers);

    // 4) 申込者へ自動返信
    if (CONFIG.AUTO_REPLY.ENABLED) {
      sendAutoReply_(ticket, record);
    }

    // 5) 社内通知
    notifyStaff_(ticket, record);

    Logger.log('処理完了: ' + ticket);
  } catch (err) {
    // どこかで失敗したら管理者にメールで知らせる（取りこぼし防止）
    handleError_('onFormSubmit', err);
  }
}

/**
 * 受付番号を発行します（例: REQ-20260614-001）。
 * その日の連番はスクリプトプロパティで管理し、日付が変われば 001 に戻ります。
 */
function issueTicketNumber_() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const lastDate = props.getProperty('TICKET_DATE');
  let seq = Number(props.getProperty('TICKET_SEQ') || '0');

  if (lastDate !== today) {
    seq = 0; // 日付が変わったらリセット
  }
  seq += 1;

  props.setProperty('TICKET_DATE', today);
  props.setProperty('TICKET_SEQ', String(seq));

  const seqStr = ('000' + seq).slice(-3);
  return CONFIG.TICKET_PREFIX + '-' + today + '-' + seqStr;
}

/**
 * 管理台帳シートへ1行追記します。無ければヘッダー付きで自動作成。
 * @return {Object} 追記した内容（後続のメール/通知で使う）
 */
function appendToManagementSheet_(ticket, answers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.MANAGEMENT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.MANAGEMENT_SHEET_NAME);
    sheet.appendRow(['受付番号', '受付日時', '氏名', 'メール', '分類', '対応状況', '内容（元データ）']);
    sheet.setFrozenRows(1);
  }

  const name = answers[CONFIG.COLUMNS.NAME] || '（氏名未入力）';
  const email = answers[CONFIG.COLUMNS.EMAIL] || '';
  const category = answers[CONFIG.COLUMNS.CATEGORY] || '未分類';
  const now = new Date();

  sheet.appendRow([
    ticket,
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    name,
    email,
    category,
    '未対応', // 初期ステータス（プルダウンにしておくと運用しやすい）
    JSON.stringify(answers),
  ]);

  return { ticket: ticket, name: name, email: email, category: category, receivedAt: now };
}

/**
 * フォームの namedValues（{質問:[回答]} 形式）を {質問:回答} に平坦化。
 */
function flattenNamedValues_(namedValues) {
  const out = {};
  Object.keys(namedValues).forEach(function (key) {
    const v = namedValues[key];
    out[key] = Array.isArray(v) ? v.join(' / ') : String(v);
  });
  return out;
}

/**
 * 手動テスト用：回答シートの最終行を読み、{質問:回答} で返します。
 * （エディタから onFormSubmit を手動実行したときに使われます）
 */
function readLatestResponse_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.RESPONSE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('回答シート「' + CONFIG.RESPONSE_SHEET_NAME + '」に読み取れる回答がありません。');
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const last = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues()[0];

  const out = {};
  headers.forEach(function (h, i) {
    out[h] = last[i] == null ? '' : String(last[i]);
  });
  return out;
}
