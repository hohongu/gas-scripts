/**
 * Gmailの「要処理」ラベル付き未読メールをClaude APIで分類し、
 * スプレッドシートへの記録とSlack通知を行うスクリプト。
 *
 * 事前にスクリプトプロパティへ以下を設定すること：
 *   CLAUDE_API_KEY    : Claude APIのAPIキー
 *   SLACK_WEBHOOK_URL : Slack Incoming WebhookのURL
 */

// ラベル名・シート名の定義
var TODO_LABEL_NAME = '要処理';
var DONE_LABEL_NAME = '処理済み';
var MAIL_LOG_SHEET_NAME = 'メールログ';
var ERROR_LOG_SHEET_NAME = 'エラーログ';

// Claude API設定（コスト削減のためHaikuを使用）
var CLAUDE_MODEL = 'claude-haiku-5-5';
var CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
var CLAUDE_API_VERSION = '2023-06-01';

// 分類カテゴリ
var CATEGORIES = ['クレーム', '質問', '注文', 'その他'];

// Claudeに送る本文の最大文字数
var MAX_BODY_LENGTH = 4000;

/**
 * メイン処理：未読の要処理メールを分類・記録・通知し、ラベルを付け替える。
 */
function processEmails() {
  // 前回の実行が長引いた場合の二重実行を防ぐ
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return;
  }

  try {
    var todoLabel = GmailApp.getUserLabelByName(TODO_LABEL_NAME);
    if (!todoLabel) {
      logError_('ラベル確認', '「' + TODO_LABEL_NAME + '」ラベルが見つかりません。');
      return;
    }
    var doneLabel = GmailApp.getUserLabelByName(DONE_LABEL_NAME) ||
      GmailApp.createLabel(DONE_LABEL_NAME);

    var threads = GmailApp.search('label:' + TODO_LABEL_NAME + ' is:unread');
    threads.forEach(function (thread) {
      try {
        processThread_(thread, todoLabel, doneLabel);
      } catch (e) {
        // 1件の失敗で全体を止めない。ラベルは残すので次回再試行される
        logError_('メール処理: ' + thread.getFirstMessageSubject(), e);
      }
    });
  } catch (e) {
    logError_('processEmails', e);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 5分おきに processEmails を実行するトリガーを登録する。
 * 既存の同名トリガーは削除してから登録し直す（重複防止）。
 */
function setEmailTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'processEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}

/**
 * 1スレッド分の処理を行う。
 * @param {GoogleAppsScript.Gmail.GmailThread} thread 対象スレッド
 * @param {GoogleAppsScript.Gmail.GmailLabel} todoLabel 「要処理」ラベル
 * @param {GoogleAppsScript.Gmail.GmailLabel} doneLabel 「処理済み」ラベル
 */
function processThread_(thread, todoLabel, doneLabel) {
  // スレッド内の未読メールのうち最新の1通を対象とする
  var unread = thread.getMessages().filter(function (m) {
    return m.isUnread();
  });
  if (unread.length === 0) {
    return;
  }
  var message = unread[unread.length - 1];

  var subject = message.getSubject();
  var sender = message.getFrom();
  var received = message.getDate();
  var body = message.getPlainBody().substring(0, MAX_BODY_LENGTH);

  // 2. Claude APIで分類・要約
  var result = classifyEmail_(subject, body);

  // 3. スプレッドシートに記録（受信日時・送信者・件名・分類・要約）
  var logSheet = getSheetWithHeader_(MAIL_LOG_SHEET_NAME,
    ['受信日時', '送信者', '件名', '分類', '要約']);
  logSheet.appendRow([received, sender, subject, result.category, result.summary]);

  // 4. Slack通知。失敗しても記録済みなので、ラベル付け替えは続行する（重複記録を防ぐ）
  try {
    notifySlack_(subject, result.category, result.summary);
  } catch (e) {
    logError_('Slack通知: ' + subject, e);
  }

  // 5. ラベルの付け替え
  thread.addLabel(doneLabel);
  thread.removeLabel(todoLabel);
}

/**
 * メールをClaude APIで分類し、要約を取得する。
 * @param {string} subject 件名
 * @param {string} body 本文
 * @return {{category: string, summary: string}}
 */
function classifyEmail_(subject, body) {
  var apiKey = getProperty_('CLAUDE_API_KEY');

  var prompt =
    '次のメールを分類し、要約してください。\n' +
    '分類は「' + CATEGORIES.join('」「') + '」のいずれか1つです。\n' +
    '要約は日本語で100文字以内にしてください。\n' +
    '以下のJSON形式のみを出力し、他の文章は出力しないでください。\n' +
    '{"category": "分類", "summary": "要約"}\n\n' +
    '件名: ' + subject + '\n本文:\n' + body;

  var response = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': CLAUDE_API_VERSION
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 500,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Claude APIエラー (HTTP ' + code + '): ' + response.getContentText());
  }

  // 応答の先頭が text ブロックとは限らない（thinking 等が先に来る場合がある）ため、
  // type が text のブロックを連結して取り出す
  var content = JSON.parse(response.getContentText()).content || [];
  var text = content.filter(function (block) {
    return block.type === 'text' && block.text;
  }).map(function (block) {
    return block.text;
  }).join('');
  if (!text) {
    throw new Error('Claudeの応答にテキストが含まれていません: ' + response.getContentText());
  }
  // 余計な文字が付いていても最初のJSONオブジェクトを取り出す
  var match = text.match(/\{[\s\S]*\}/);
  if (!match) {
    throw new Error('Claudeの応答をJSONとして解釈できません: ' + text);
  }
  var parsed = JSON.parse(match[0]);

  // 想定外の分類は「その他」に丸める
  var category = CATEGORIES.indexOf(parsed.category) >= 0 ? parsed.category : 'その他';
  return { category: category, summary: String(parsed.summary || '') };
}

/**
 * Slack Incoming Webhookで通知する。
 * @param {string} subject 件名
 * @param {string} category 分類
 * @param {string} summary 要約
 */
function notifySlack_(subject, category, summary) {
  var webhookUrl = getProperty_('SLACK_WEBHOOK_URL');

  var text =
    ':email: *新しいメールを処理しました*\n' +
    '*件名:* ' + subject + '\n' +
    '*分類:* ' + category + '\n' +
    '*要約:* ' + summary;

  var response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Slack通知エラー (HTTP ' + code + '): ' + response.getContentText());
  }
}

/**
 * スクリプトプロパティから値を取得する。未設定ならエラーにする。
 * @param {string} key プロパティ名
 * @return {string}
 */
function getProperty_(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('スクリプトプロパティ「' + key + '」が設定されていません。');
  }
  return value;
}

/**
 * 指定名のシートを取得する。なければヘッダー付きで作成する。
 * @param {string} name シート名
 * @param {Array<string>} header ヘッダー行
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheetWithHeader_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * エラー内容を「エラーログ」シートに記録する。
 * ログ記録自体の失敗で処理が止まらないよう、Loggerにも出力する。
 * @param {string} where エラー発生箇所
 * @param {Error|string} error エラー
 */
function logError_(where, error) {
  var message = (error && error.message) ? error.message : String(error);
  var stack = (error && error.stack) ? error.stack : '';
  Logger.log('[' + where + '] ' + message);
  try {
    var sheet = getSheetWithHeader_(ERROR_LOG_SHEET_NAME,
      ['発生日時', '発生箇所', 'エラー内容', 'スタックトレース']);
    sheet.appendRow([new Date(), where, message, stack]);
  } catch (e) {
    Logger.log('エラーログの記録に失敗しました: ' + e.message);
  }
}
