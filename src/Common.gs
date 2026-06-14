/**
 * Common.gs  ―  共通処理（エラーハンドリング）
 * ------------------------------------------------------------------
 * 処理が失敗したとき、取りこぼさないよう管理者へメールで知らせます。
 * 自動化は「動かなくなったことに気づけない」のが一番怖いので、
 * 失敗の通知をきちんと入れておくのが実運用のコツです。
 * ------------------------------------------------------------------
 */

/**
 * 例外を記録し、管理者へ通知します。
 * @param {string} where  どの関数で失敗したか
 * @param {Error}  err    発生した例外
 */
function handleError_(where, err) {
  const message = '[' + where + '] ' + (err && err.stack ? err.stack : err);
  Logger.log(message);

  try {
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: '【要確認】問い合わせ自動処理でエラーが発生しました',
      body: [
        '自動処理中にエラーが発生しました。内容をご確認ください。',
        '',
        '発生箇所: ' + where,
        '日時　　: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
        '',
        '--- 詳細 ---',
        message,
      ].join('\n'),
    });
  } catch (mailErr) {
    // メール通知自体が失敗した場合はログに残すだけ
    Logger.log('エラー通知メールの送信にも失敗: ' + mailErr);
  }
}
