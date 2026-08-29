/**
 * 西軽精機ナレッジ（仮） バックエンド
 * DBスプレッドシート: 1wmPkFZ-EdtUQchWlVGspeclP3yHbO9YyViAM5AraQj0
 *
 * デプロイ後の確認手順は加工ナレッジ/CLAUDE.mdの「デプロイ・動作確認の手順」に準拠。
 * setupSheets_ はタブ構成を作る一回限りの関数。GASエディタで手動実行すること。
 */

var SHEET_QUALITY_LOG = '品質情報記録ログ';
var SHEET_TOOL_MEMO = 'ツール配置メモ';
var SHEET_TOOL_POSITIONS = 'ツール配置ポジション';
var SHEET_SHIPPING_SPEC = '出荷仕様';

// I-PRO同期データ（進捗状況照会と共有）。2026-08-14に列構成を確認済み：
// 工場番号／製造番号／材料手配区分／得意先コード／品番(図番）／品名／担当者／…
var IPRO_SOURCE_SPREADSHEET_ID = '1g-NnnSgGyS_5oIINuUfvNi7_o7iWPik6aL0HmoL5VO4';

/** タブとヘッダー行を作る。既存タブがあれば何もしない。GASエディタで1回だけ手動実行。 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEET_QUALITY_LOG, [
    '投稿ID', 'タイムスタンプ', '図番', '部署', '投稿者メール', '投稿者名',
    '外観ランク', '内容', '写真URL', '共有フラグ'
  ]);

  ensureSheet_(ss, SHEET_TOOL_MEMO, [
    '投稿ID', 'タイムスタンプ', '図番', '投稿者メール', '投稿者名',
    '内容', '写真URL', '共有フラグ'
  ]);

  ensureSheet_(ss, SHEET_TOOL_POSITIONS, [
    '図番', '列区分', '順番', 'Tナンバー', '工具説明',
    '正面チャック径', '背面チャック径', 'サイクルタイム',
    '専用ツール保管', '前進端位置', '最終更新者メール', '最終更新日時'
  ]);

  ensureSheet_(ss, SHEET_SHIPPING_SPEC, [
    '図番',
    '検査記録の添付', 'ミルシート', 'トレー梱包', 'NG限度見本',
    'カット品', 'テストピース', '借用ゲージ有無', '借用ゲージ種類',
    '梱包方法', 'その他必要事項',
    '最終更新者メール', '最終更新日時'
  ]);

  // デフォルトのSheet1が残っていれば削除（タブ構成を綺麗に保つ）
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('セットアップ完了: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * GET API
 * ?action=zubanInfo&zuban=XXXX  … 図番に紐づく過去トラ・品質情報・出荷仕様等をまとめて返す
 */
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'zubanInfo') {
      return jsonResponse_(getZubanInfo_(e.parameter.zuban));
    }
    if (action === 'resolveZuban') {
      return jsonResponse_(resolveZubanFromSeiban_(e.parameter.seiban));
    }
    return jsonResponse_({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}

/**
 * POST API（JSON body、{action: '...', ...} の形）
 * action: postQualityLog / postToolMemo / saveToolPositions / saveShippingSpec
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    if (action === 'postQualityLog') return jsonResponse_(postQualityLog_(payload));
    if (action === 'postToolMemo') return jsonResponse_(postToolMemo_(payload));
    if (action === 'saveToolPositions') return jsonResponse_(saveToolPositions_(payload));
    if (action === 'saveShippingSpec') return jsonResponse_(saveShippingSpec_(payload));
    return jsonResponse_({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 図番に紐づく情報一式を集めて返す。
 * TODO: 品名・得意先の取得は進捗状況照会のI-PRO同期データを参照する想定だが、
 *       実シートの列構成を未確認のため lookupZubanMaster_ はスタブのまま。
 *       実装前に対象スプレッドシートの列構成をユーザーと確認すること。
 */
function getZubanInfo_(zuban) {
  if (!zuban) return { error: 'zuban is required' };

  var master = lookupZubanMaster_(zuban);
  var pastTrouble = findPastTrouble_(zuban);
  var qualityLog = readQualityLog_(zuban);
  var toolMemo = readToolMemo_(zuban);
  var toolPositions = readToolPositions_(zuban);
  var shippingSpec = readShippingSpec_(zuban);

  return {
    zuban: zuban,
    hinmei: master.hinmei,
    tokuisaki: master.tokuisaki,
    pastTrouble: pastTrouble,       // 品質情報／不具合改善計画書（都度Drive検索）＋共有された品質情報記録／ツール配置メモ
    qualityLog: qualityLog,         // 自部署の品質情報（共有フラグOFFのもの）
    toolMemo: toolMemo,             // ツール配置メモ（一次・二次加工向け）
    toolPositions: toolPositions,   // ツール配置ポジション（正面／背面／サイクルタイム）
    shippingSpec: shippingSpec      // 出荷仕様（検査・生産管理・出荷担当が入力）
  };
}

/**
 * QRから取り出した製番（例: I-PROのURL末尾の値）から、対応する図番を探す。
 * I-PRO同期データの「製造番号」列と一致する行を探し、「品番(図番）」「品名」「得意先コード」を返す。
 * 1製番に図番が複数見つかった場合はcandidatesに全件返す（呼び出し側で選択画面を出す想定）。
 */
function resolveZubanFromSeiban_(seiban) {
  if (!seiban) return { error: 'seiban is required' };
  var rows = findIproRowsByColumn_('製造番号', seiban);
  if (rows.length === 0) return { found: false, seiban: seiban };
  var candidates = rows.map(function (row) {
    return {
      zuban: row['品番(図番）'] || row['品番(図番)'],
      hinmei: row['品名'],
      tokuisakiCode: row['得意先コード']
    };
  }).filter(function (c) {
    // 「製造番号」列を持つが図番の列名が違う／無い別タブ（進捗管理など）がヒットして
    // 空の候補が混ざることがあるため、図番が取れているものだけ残す
    return !!c.zuban;
  });
  // 図番が重複している場合（同じ図番が別行にある等）は1つにまとめる
  var uniqueByZuban = {};
  candidates.forEach(function (c) { uniqueByZuban[c.zuban] = c; });
  var uniqueCandidates = Object.keys(uniqueByZuban).map(function (z) { return uniqueByZuban[z]; });
  return {
    found: true,
    seiban: seiban,
    // 1製番に図番が複数見つかった場合、呼び出し側（フロント）で選択画面を出す
    multiple: uniqueCandidates.length > 1,
    candidates: uniqueCandidates,
    zuban: uniqueCandidates.length === 1 ? uniqueCandidates[0].zuban : null,
    hinmei: uniqueCandidates.length === 1 ? uniqueCandidates[0].hinmei : null,
    tokuisakiCode: uniqueCandidates.length === 1 ? uniqueCandidates[0].tokuisakiCode : null
  };
}

/** 図番から品名・得意先コードを引く（②画面のヘッダー表示用）。 */
function lookupZubanMaster_(zuban) {
  var row = findIproRowByColumn_('品番(図番）', zuban) || findIproRowByColumn_('品番(図番)', zuban);
  if (!row) return { hinmei: null, tokuisaki: null };
  return { hinmei: row['品名'] || null, tokuisaki: row['得意先コード'] || null };
}

/**
 * I-PRO同期データ（進捗状況照会と共有のスプレッドシート）から、指定列が指定値と一致する最初の行を返す。
 * どのタブに目的の列があるか固定できていないため、全タブを順に探す。
 */
function findIproRowByColumn_(columnName, value) {
  var rows = findIproRowsByColumn_(columnName, value);
  return rows.length > 0 ? rows[0] : null;
}

/** findIproRowByColumn_ の複数件版。1製番に図番が複数ある場合の検出に使う。 */
function findIproRowsByColumn_(columnName, value) {
  if (!value) return [];
  var ss = SpreadsheetApp.openById(IPRO_SOURCE_SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var out = [];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (sheet.getLastRow() < 2) continue;
    var values = sheet.getDataRange().getValues();
    var header = values[0];
    var col = header.indexOf(columnName);
    if (col === -1) continue;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][col]) === String(value)) {
        out.push(rowToObject_(header, values[i]));
      }
    }
  }
  return out;
}

/**
 * 過去トラ：{図番} 品質情報 スプレッドシート／{得意先} {図番} {品名} 不具合改善計画書 フォルダを
 * Driveのタイトル部分一致で検索し、内容を要約して返す。
 * 品質情報はそのまま、不具合改善計画書は一行要約＋元ファイルへのリンクのみ返す（本文は開かない）。
 */
function findPastTrouble_(zuban) {
  var items = [];

  // 品質情報スプレッドシート（タイトルに図番を含むもの）
  var qiFiles = DriveApp.searchFiles(
    'title contains "' + zuban.replace(/"/g, '') + '" and title contains "品質情報" and mimeType = "application/vnd.google-apps.spreadsheet"'
  );
  while (qiFiles.hasNext()) {
    var f = qiFiles.next();
    items.push({
      source: '品質情報',
      title: f.getName(),
      url: f.getUrl(),
      // TODO: シート本文の読み取り・日付ごとのコメント抽出・Sheet.getImages()での写真検出は未実装
      note: 'シート本文の要約読み取りは未実装。まずはリンクのみ。'
    });
  }

  // 不具合改善計画書フォルダ（タイトルに図番を含むもの）
  var fkFolders = DriveApp.searchFiles(
    'title contains "' + zuban.replace(/"/g, '') + '" and title contains "不具合改善計画書" and mimeType = "application/vnd.google-apps.folder"'
  );
  while (fkFolders.hasNext()) {
    var folder = fkFolders.next();
    items.push({
      source: '不具合改善計画書',
      title: folder.getName(),
      url: folder.getUrl(),
      note: '一行要約（不適合事象）の自動抽出は未実装。まずはフォルダへのリンクのみ。'
    });
  }

  // 新システム内で「共有する」を選んだ品質情報記録・ツール配置メモ
  items = items.concat(findSharedEntries_(SHEET_QUALITY_LOG, zuban));
  items = items.concat(findSharedEntries_(SHEET_TOOL_MEMO, zuban));

  return items;
}

function findSharedEntries_(sheetName, zuban) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var header = rows[0];
  var zubanCol = header.indexOf('図番');
  var shareCol = header.indexOf('共有フラグ');
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][zubanCol] === zuban && rows[i][shareCol] === true) {
      out.push({ source: sheetName, row: rowToObject_(header, rows[i]) });
    }
  }
  return out;
}

function readQualityLog_(zuban) {
  return readRowsByZuban_(SHEET_QUALITY_LOG, zuban).filter(function (r) { return r['共有フラグ'] !== true; });
}

function readToolMemo_(zuban) {
  return readRowsByZuban_(SHEET_TOOL_MEMO, zuban);
}

function readToolPositions_(zuban) {
  return readRowsByZuban_(SHEET_TOOL_POSITIONS, zuban);
}

function readShippingSpec_(zuban) {
  var rows = readRowsByZuban_(SHEET_SHIPPING_SPEC, zuban);
  return rows.length > 0 ? rows[rows.length - 1] : null; // 図番ごとに1件・上書き想定
}

function readRowsByZuban_(sheetName, zuban) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getDataRange().getValues();
  var header = rows[0];
  var zubanCol = header.indexOf('図番');
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][zubanCol] === zuban) out.push(rowToObject_(header, rows[i]));
  }
  return out;
}

function rowToObject_(header, row) {
  var obj = {};
  for (var i = 0; i < header.length; i++) obj[header[i]] = row[i];
  return obj;
}

/** 品質情報記録の投稿（⑤画面）。承認フローなし、送信したら即座に反映。 */
function postQualityLog_(payload) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUALITY_LOG);
  var id = Utilities.getUuid();
  sheet.appendRow([
    id, new Date(), payload.zuban, payload.department,
    payload.userEmail, payload.userName,
    payload.rank || '', payload.content || '', payload.photoUrl || '',
    !!payload.shared
  ]);
  return { id: id };
}

/** ツール配置メモの投稿（③画面）。承認フローなし。 */
function postToolMemo_(payload) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TOOL_MEMO);
  var id = Utilities.getUuid();
  sheet.appendRow([
    id, new Date(), payload.zuban,
    payload.userEmail, payload.userName,
    payload.content || '', payload.photoUrl || '',
    !!payload.shared
  ]);
  return { id: id };
}

/** ツール配置ポジションの保存（③編集画面）。図番の既存行を全削除してから書き直す（上書き）。 */
function saveToolPositions_(payload) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TOOL_POSITIONS);
  deleteRowsByZuban_(sheet, payload.zuban);
  var now = new Date();
  (payload.positions || []).forEach(function (p) {
    sheet.appendRow([
      payload.zuban, p.column, p.order, p.tNumber, p.description,
      payload.frontChuck || '', payload.backChuck || '', payload.cycleTime || '',
      payload.toolStorage || '', payload.forwardPosition || '',
      payload.userEmail, now
    ]);
  });
  return { ok: true };
}

/** 出荷仕様の保存（⑤画面の各セクション）。図番ごとに1行、上書き更新。 */
function saveShippingSpec_(payload) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHIPPING_SPEC);
  var rows = sheet.getDataRange().getValues();
  var header = rows[0];
  var zubanCol = header.indexOf('図番');
  var rowIndex = -1;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][zubanCol] === payload.zuban) { rowIndex = i; break; }
  }
  var existing = rowIndex >= 0 ? rowToObject_(header, rows[rowIndex]) : { '図番': payload.zuban };
  var merged = Object.assign({}, existing, payload.fields || {});
  merged['最終更新者メール'] = payload.userEmail;
  merged['最終更新日時'] = new Date();
  var newRow = header.map(function (h) { return merged[h] !== undefined ? merged[h] : ''; });
  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 1, 1, 1, header.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
  return { ok: true };
}

function deleteRowsByZuban_(sheet, zuban) {
  var rows = sheet.getDataRange().getValues();
  var header = rows[0];
  var zubanCol = header.indexOf('図番');
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][zubanCol] === zuban) sheet.deleteRow(i + 1);
  }
}
