/**
 * Report.gs  ―  サブ機能：定期レポート自動配信
 * ------------------------------------------------------------------
 * 日次トリガー（setupTriggers で毎朝登録）で動き、
 * 「問い合わせ管理」台帳から前日分を集計して、
 * サマリーを Slack と管理者メールに自動配信します。
 *
 *   ・前日の受付件数
 *   ・分類ごとの内訳
 *   ・未対応の残件数（全期間）
 * ------------------------------------------------------------------
 */

/**
 * 日次レポートを配信します（timeBasedトリガーから呼ばれる）。
 */
function sendDailyReport() {
  try {
    const summary = buildDailySummary_();
    const text = formatReportText_(summary);

    // 管理者へメール
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: '【日次レポート】問い合わせ ' + summary.dateLabel + '（受付 ' + summary.total + '件）',
      body: text,
    });

    // Slackへも投稿（任意）
    if (CONFIG.SLACK.ENABLED) {
      postReportToSlack_(summary, text);
    }

    Logger.log('日次レポートを配信しました: ' + summary.dateLabel);
  } catch (err) {
    handleError_('sendDailyReport', err);
  }
}

/**
 * 台帳を読み、前日分を集計します。
 * @return {Object} 集計結果
 */
function buildDailySummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MANAGEMENT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { dateLabel: yesterdayLabel_(), total: 0, byCategory: {}, openCount: 0 };
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  // 列: 0=受付番号 1=受付日時 2=氏名 3=メール 4=分類 5=対応状況

  const target = yesterdayLabel_();      // 例: 2026/06/13
  const byCategory = {};
  let total = 0;
  let openCount = 0;

  values.forEach(function (row) {
    const receivedLabel = String(row[1]).slice(0, 10); // yyyy/MM/dd
    const category = row[4] || '未分類';
    const status = row[5] || '';

    if (receivedLabel === target) {
      total += 1;
      byCategory[category] = (byCategory[category] || 0) + 1;
    }
    if (status !== '対応済' && status !== '完了') {
      openCount += 1; // 全期間の未対応残件
    }
  });

  return { dateLabel: target, total: total, byCategory: byCategory, openCount: openCount };
}

/** レポート本文（プレーンテキスト）を組み立てます。 */
function formatReportText_(s) {
  const lines = [];
  lines.push('■ 問い合わせ日次レポート（' + s.dateLabel + '）');
  lines.push('');
  lines.push('前日の受付件数: ' + s.total + ' 件');

  const cats = Object.keys(s.byCategory);
  if (cats.length > 0) {
    lines.push('');
    lines.push('【分類ごとの内訳】');
    cats.forEach(function (c) {
      lines.push('　・' + c + ': ' + s.byCategory[c] + ' 件');
    });
  }

  lines.push('');
  lines.push('未対応の残件（全期間）: ' + s.openCount + ' 件');
  lines.push('');
  lines.push('※「問い合わせ管理」シートの対応状況を「対応済」にすると残件から外れます。');
  return lines.join('\n');
}

/** Slackへレポートを投稿します。 */
function postReportToSlack_(summary, text) {
  const url = getSlackWebhookUrl_();
  if (!url) return;

  const payload = {
    blocks: [
      { type: 'header', text: { type: 'plain_text', text: ':bar_chart: 問い合わせ日次レポート', emoji: true } },
      { type: 'section', text: { type: 'mrkdwn', text: '```' + text + '```' } },
    ],
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

/** 「昨日」の yyyy/MM/dd ラベルを返します。 */
function yesterdayLabel_() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}
