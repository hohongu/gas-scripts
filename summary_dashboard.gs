/**
 * 売上データを月ごとに集計し、「月次サマリー」シートに書き出して
 * 月次推移の棒グラフを表示するダッシュボードスクリプト。
 */

// シート名の定義
var SOURCE_SHEET_NAME = '売上データ';
var SUMMARY_SHEET_NAME = '月次サマリー';

/**
 * メイン処理：売上データを月次集計し、サマリーシートとグラフを更新する。
 */
function runDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  // 出力先シートがなければ作成する
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }

  var rows = readSalesRows_(sourceSheet);
  var monthly = aggregateByMonth_(rows);
  writeSummary_(summarySheet, monthly);
  updateChart_(summarySheet, monthly.length);
}

/**
 * 毎朝9時に runDashboard を実行するトリガーを登録する。
 * 既存の同名トリガーは削除してから登録し直す（重複防止）。
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

/**
 * 「売上データ」シートの全データ行を読み込む（ヘッダー行は除く）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 売上データシート
 * @return {Array<Array>} A〜D列の値の配列
 */
function readSalesRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, 4).getValues();
}

/**
 * 売上行を月ごとに集計する。
 * 日付が不正な行や金額が数値でない行はスキップする。
 * @param {Array<Array>} rows 売上データの行
 * @return {Array<{key: string, label: string, total: number, count: number}>}
 *   月の昇順に並んだ集計結果
 */
function aggregateByMonth_(rows) {
  var map = {};

  rows.forEach(function (row) {
    var date = toDate_(row[0]);
    var amount = Number(row[3]);
    if (!date || row[3] === '' || isNaN(amount)) {
      return;
    }

    // 並べ替え用のキー（例：2026-01）
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var key = year + '-' + ('0' + month).slice(-2);

    if (!map[key]) {
      map[key] = { key: key, label: year + '年' + month + '月', total: 0, count: 0 };
    }
    map[key].total += amount;
    map[key].count += 1;
  });

  return Object.keys(map).sort().map(function (key) {
    return map[key];
  });
}

/**
 * セルの値を Date に変換する。変換できない場合は null を返す。
 * 「2026/01/05」のような文字列にも対応する。
 * @param {Date|string} value セルの値
 * @return {Date|null}
 */
function toDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'string' && value !== '') {
    var date = new Date(value);
    return isNaN(date.getTime()) ? null : date;
  }
  return null;
}

/**
 * 「月次サマリー」シートをクリアし、集計結果を書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 出力先シート
 * @param {Array} monthly 月次集計結果
 */
function writeSummary_(sheet, monthly) {
  sheet.clear();

  var values = [['月', '合計売上', '件数']];
  monthly.forEach(function (m) {
    values.push([m.label, m.total, m.count]);
  });
  sheet.getRange(1, 1, values.length, 3).setValues(values);

  // 見た目の整形
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  if (monthly.length > 0) {
    sheet.getRange(2, 2, monthly.length, 1).setNumberFormat('#,##0');
  }
}

/**
 * 既存のグラフを削除し、月次推移の棒グラフを作成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 出力先シート
 * @param {number} monthCount 集計した月数
 */
function updateChart_(sheet, monthCount) {
  // 毎回作り直すため既存のグラフを削除
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });

  // データがない場合はグラフを作らない
  if (monthCount === 0) {
    return;
  }

  // 月（A列）と合計売上（B列）を使って棒グラフを作成
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, monthCount + 1, 2))
    .setNumHeaders(1)
    .setPosition(2, 5, 0, 0)
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上', format: '#,##0' })
    .build();

  sheet.insertChart(chart);
}
