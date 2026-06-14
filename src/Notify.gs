/**
 * Notify.gs  ―  通知まわり（自動返信メール / Slack / 社内メール）
 * ------------------------------------------------------------------
 * 通知の「送り方」をここに集約。文面や宛先は Config.gs 側で変えられます。
 * ------------------------------------------------------------------
 */

/**
 * 申込者あての自動返信メールを送ります。
 * メールアドレスが取れない場合はスキップ（エラーにはしない）。
 */
function sendAutoReply_(ticket, record) {
  if (!record.email || !isValidEmail_(record.email)) {
    Logger.log('自動返信スキップ: 有効なメールアドレスがありません。');
    return;
  }

  const subject = fillTemplate_(CONFIG.AUTO_REPLY.SUBJECT, record);
  const body = fillTemplate_(CONFIG.AUTO_REPLY.BODY, record);

  MailApp.sendEmail({
    to: record.email,
    subject: subject,
    body: body,
    name: CONFIG.AUTO_REPLY.SENDER_NAME || undefined,
  });
}

/**
 * 社内へ通知します。Slack優先、無効ならメール、両方OFFなら何もしません。
 */
function notifyStaff_(ticket, record) {
  if (CONFIG.SLACK.ENABLED) {
    notifySlack_(record);
  }
  if (CONFIG.STAFF_EMAIL.ENABLED) {
    notifyStaffEmail_(record);
  }
}

/**
 * Slack（Incoming Webhook）へ通知します。
 */
function notifySlack_(record) {
  const url = getSlackWebhookUrl_();
  if (!url) {
    Logger.log('Slack通知スキップ: Webhook URL が未設定です（スクリプトプロパティ SLACK_WEBHOOK_URL）。');
    return;
  }

  const payload = {
    blocks: [
      { type: 'header', text: { type: 'plain_text', text: CONFIG.SLACK.TITLE, emoji: true } },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: '*受付番号:*\n' + record.ticket },
          { type: 'mrkdwn', text: '*分類:*\n' + record.category },
          { type: 'mrkdwn', text: '*お名前:*\n' + record.name },
          { type: 'mrkdwn', text: '*メール:*\n' + (record.email || '（なし）') },
        ],
      },
      {
        type: 'context',
        elements: [{ type: 'mrkdwn', text: '受付: ' + Utilities.formatDate(record.receivedAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') }],
      },
    ],
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    Logger.log('Slack通知に失敗: ' + res.getResponseCode() + ' / ' + res.getContentText());
  }
}

/**
 * 社内メール通知（Slackを使わない運用向け）。
 */
function notifyStaffEmail_(record) {
  const subject = fillTemplate_(CONFIG.STAFF_EMAIL.SUBJECT, record);
  const body = [
    '新しい問い合わせが届きました。',
    '',
    '受付番号: ' + record.ticket,
    'お名前　: ' + record.name,
    'メール　: ' + (record.email || '（なし）'),
    '分類　　: ' + record.category,
    '受付日時: ' + Utilities.formatDate(record.receivedAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
  ].join('\n');

  MailApp.sendEmail({ to: CONFIG.STAFF_EMAIL.TO, subject: subject, body: body });
}

/* ── 補助関数 ───────────────────────────────────────────── */

/** {{name}} {{ticket}} などのプレースホルダを置換します。 */
function fillTemplate_(template, record) {
  return String(template)
    .replace(/\{\{ticket\}\}/g, record.ticket)
    .replace(/\{\{name\}\}/g, record.name)
    .replace(/\{\{category\}\}/g, record.category || '');
}

/** ごく簡易なメールアドレス形式チェック。 */
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
}
