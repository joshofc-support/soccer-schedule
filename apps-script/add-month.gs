/**
 * サッカー部スケジュール ― 「新しい月を追加」ツール（Google Apps Script）
 *
 * 使い方：
 *   スプレッドシート上部メニュー「📅 カレンダー」→「新しい月を追加…」を選び、
 *   年月（例：2026年11月）を入力すると、新しいシートが自動作成されます。
 *
 * 仕組み：
 *   直近の月シート（「◯◯年◯月」形式）をテンプレートとして複製し、
 *   見出し・列構成・書式・セル結合はそのまま引き継いだうえで、
 *   予定データだけを消去し、その月の「日にち」と「曜日」を自動入力します。
 *
 * ※ このスクリプトはスプレッドシートに紐づきます（バインド型）。
 *   予定の入力や日々の運用に、実行のたびの承認は不要です（初回のみ承認）。
 */

var WEEK = ['日', '月', '火', '水', '木', '金', '土'];

/** スプレッドシートを開いたときにメニューを追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 カレンダー')
    .addItem('新しい月を追加…', 'addMonthSheet')
    .addToUi();
}

/** メニュー本体：年月を入力 → 新しい月シートを作成 */
function addMonthSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1) 年月を入力してもらう
  var res = ui.prompt(
    '新しい月を追加',
    '追加する年月を入力してください。\n例：2026年11月 ／ 2026-11 ／ 202611',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var ym = parseYearMonth(res.getResponseText());
  if (!ym) {
    ui.alert('入力エラー', '年月を「2026年11月」の形式で入力してください。', ui.ButtonSet.OK);
    return;
  }
  var year = ym.year, month = ym.month;
  var name = year + '年' + month + '月';

  // 2) 同名シートの重複チェック
  var exist = ss.getSheetByName(name);
  if (exist) {
    ui.alert('すでに存在します', '「' + name + '」シートは既にあります。', ui.ButtonSet.OK);
    ss.setActiveSheet(exist);
    return;
  }

  // 3) テンプレート（最新の月シート）を複製
  var tpl = findTemplateSheet(ss);
  if (!tpl) {
    ui.alert('テンプレートが見つかりません',
      '「◯◯年◯月」形式の月シートが1つも見つかりませんでした。\n手本となる月シートを1つ用意してから、もう一度実行してください。',
      ui.ButtonSet.OK);
    return;
  }
  var sheet = tpl.copyTo(ss).setName(name);

  // 4) 見出し行を検出し、データ行を消去（書式・結合・入力規則は残す）
  var head = findHeader(sheet);
  if (!head) {
    ui.alert('注意',
      'テンプレートの見出し（「日にち」）が見つからなかったため、複製のみ行いました。\n「' + name + '」シートを手動でご確認ください。',
      ui.ButtonSet.OK);
    ss.setActiveSheet(sheet);
    return;
  }
  var dataStart = head.headerRow + 1;
  var days = new Date(year, month, 0).getDate();   // その月の日数

  // 必要な行数を確保（足りなければ追加）
  var maxRows = sheet.getMaxRows();
  var available = maxRows - dataStart + 1;
  if (available < days) sheet.insertRowsAfter(maxRows, days - available);

  // 既存の予定データを消去（見出しより下すべて。書式は保持）
  sheet.getRange(dataStart, 1, sheet.getMaxRows() - dataStart + 1, head.lastCol).clearContent();

  // 5) 日付（実データ＝年込み）と曜日を書き込む
  var dateVals = [], weekVals = [];
  for (var d = 1; d <= days; d++) {
    var dt = new Date(year, month - 1, d);
    dateVals.push([dt]);
    weekVals.push([WEEK[dt.getDay()]]);
  }
  sheet.getRange(dataStart, head.dateCol, days, 1).setValues(dateVals).setNumberFormat('m/d');
  if (head.weekCol) sheet.getRange(dataStart, head.weekCol, days, 1).setValues(weekVals);

  // 6) 作成した月を表示
  ss.setActiveSheet(sheet);
  ui.alert('作成しました',
    '「' + name + '」シートを作成しました（' + days + '日分）。\n予定を入力してください。カレンダーには自動で反映されます。',
    ui.ButtonSet.OK);
}

/** 「2026年11月」「2026-11」「2026/11」「202611」などを {year, month} に変換 */
function parseYearMonth(s) {
  if (!s) return null;
  s = String(s).trim().replace(/\s+/g, '');
  var m = s.match(/^(\d{4})[年\-\/\.](\d{1,2})月?$/) || s.match(/^(\d{4})(\d{2})$/);
  if (!m) return null;
  var year = +m[1], month = +m[2];
  if (year < 2000 || year > 2100) return null;
  if (month < 1 || month > 12) return null;
  return { year: year, month: month };
}

/** 「◯◯年◯月」形式のシートのうち、最も新しい月のものをテンプレートとして返す */
function findTemplateSheet(ss) {
  var best = null, bestKey = -1;
  ss.getSheets().forEach(function (sh) {
    var m = sh.getName().match(/^(\d{4})年(\d{1,2})月$/);
    if (!m) return;
    var key = (+m[1]) * 100 + (+m[2]);
    if (key > bestKey) { bestKey = key; best = sh; }
  });
  return best;
}

/** 見出し行（「日にち」を含む行）と、日付列・曜日列・最終列を探す */
function findHeader(sheet) {
  var rows = Math.min(10, sheet.getMaxRows());
  var cols = Math.max(1, sheet.getLastColumn());
  var vals = sheet.getRange(1, 1, rows, cols).getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (String(vals[r][c]).indexOf('日にち') >= 0) {
        var weekCol = 0;
        for (var cc = 0; cc < vals[r].length; cc++) {
          if (String(vals[r][cc]).indexOf('曜日') >= 0) { weekCol = cc + 1; break; }
        }
        return { headerRow: r + 1, dateCol: c + 1, weekCol: weekCol, lastCol: cols };
      }
    }
  }
  return null;
}
