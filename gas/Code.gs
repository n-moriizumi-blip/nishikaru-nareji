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
var SHEET_ZUBAN_INDEX = '図番インデックス';
var SHEET_SEIBAN_INDEX = '製番インデックス';

// I-PRO同期データ。2026-08-14に列構成を確認済み：
// 工場番号／製造番号／材料手配区分／得意先コード／品番(図番）／品名／担当者／…（60列、約1.4MB、全件）
var IPRO_SOURCE_SPREADSHEET_ID = '1g-NnnSgGyS_5oIINuUfvNi7_o7iWPik6aL0HmoL5VO4';
// 「進捗状況照会」共有スプレッドシート。製造番号・品番(図番)・品名・得意先コード・得意先名が
// 同じ行に揃っており、I-Pro Sourceの約1/10のサイズ（約145KB）。ユーザー提案(2026-08-30)により、
// こちらを優先的に検索し、見つからない場合のみI-Pro Sourceにフォールバックする。
var IPRO_PROGRESS_SPREADSHEET_ID = '1F9Iu5t62WDW5lg_eeEa6XW9ngCUJ2DmXTKjqd5oXrac';

// 社員マスタ「組織図マスタ」。氏名・Mail Address・課名・工程名の列を持つ。部署ごとの機能出し分けに使う。
var ORG_MASTER_SPREADSHEET_ID = '1fffjE_bwrzswvRO62U0OHwvqrs5b_UuSV5IbudUMxec';

// 動作確認用：全部署の画面にアクセスできるアカウント（本来の役割による出し分けとは別に、確認のため常時allAccess:trueを返す）。
var ALL_ACCESS_EMAILS = ['n-moriizumi@nishikaru.co.jp'];

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

  ensureSheet_(ss, SHEET_ZUBAN_INDEX, [
    '図番', '品名', '得意先コード', '検査記録フォルダURL', '品質情報リンク', '改善計画書リンク', '更新日時'
  ]);

  ensureSheet_(ss, SHEET_SEIBAN_INDEX, [
    '製造番号', '図番', '品名', '得意先コード', '更新日時'
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
    if (action === 'scanZuban') {
      return jsonResponse_(scanZuban_(e.parameter.seiban));
    }
    if (action === 'role') {
      return jsonResponse_(getRole_(e.parameter.email));
    }
    if (action === 'inspectionFolder') {
      return jsonResponse_(getInspectionFolderUrl_(e.parameter.zuban));
    }
    return jsonResponse_({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}

/**
 * POST API（JSON body、{action: '...', ...} の形）
 * action: postQualityLog / postToolMemo / updateQualityLog / deleteQualityLog / updateToolMemo / deleteToolMemo /
 *         saveToolPositions / saveShippingSpec / uploadPhoto
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    if (action === 'postQualityLog') return jsonResponse_(postQualityLog_(payload));
    if (action === 'postToolMemo') return jsonResponse_(postToolMemo_(payload));
    if (action === 'updateQualityLog') return jsonResponse_(updateQualityLog_(payload));
    if (action === 'deleteQualityLog') return jsonResponse_(deleteQualityLog_(payload));
    if (action === 'updateToolMemo') return jsonResponse_(updateToolMemo_(payload));
    if (action === 'deleteToolMemo') return jsonResponse_(deleteToolMemo_(payload));
    if (action === 'saveToolPositions') return jsonResponse_(saveToolPositions_(payload));
    if (action === 'saveShippingSpec') return jsonResponse_(saveShippingSpec_(payload));
    if (action === 'uploadPhoto') return jsonResponse_(uploadPhoto_(payload));
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
 * I-PRO同期データの参照（lookupZubanMaster_）とDrive全体検索（findPastTrouble_）が重いため、
 * 5分間キャッシュする（CacheService）。品質情報記録・ツール配置メモ等を投稿/保存した際は
 * invalidateZubanCache_でその図番のキャッシュを即座に破棄し、「送信したら即座に反映」を保つ。
 */
function getZubanInfo_(zuban, knownMaster) {
  if (!zuban) return { error: 'zuban is required' };

  var cache = CacheService.getScriptCache();
  var cacheKey = 'zubanInfo:' + zuban;
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  // knownMaster（{hinmei, tokuisaki}）が渡された場合はI-PROスキャンを省略する（scanZuban_から使用）。
  var master = knownMaster || lookupZubanMaster_(zuban);
  var pastTrouble = findPastTrouble_(zuban);
  var qualityLog = readQualityLog_(zuban);
  var toolMemo = readToolMemo_(zuban);
  var toolPositions = readToolPositions_(zuban);
  var shippingSpec = readShippingSpec_(zuban);

  var result = {
    zuban: zuban,
    hinmei: master.hinmei,
    tokuisaki: master.tokuisaki,
    pastTrouble: pastTrouble,       // 品質情報／不具合改善計画書（都度Drive検索）＋共有された品質情報記録／ツール配置メモ
    qualityLog: qualityLog,         // 自部署の品質情報（共有フラグOFFのもの）
    toolMemo: toolMemo,             // ツール配置メモ（一次・二次加工向け）
    toolPositions: toolPositions,   // ツール配置ポジション（正面／背面／サイクルタイム）
    shippingSpec: shippingSpec      // 出荷仕様（検査・生産管理・出荷担当が入力）
  };

  try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} // 5分。サイズ超過等は無視して素通りさせる

  return result;
}

/** 図番のzubanInfoキャッシュを破棄する。投稿・保存系のAPIが完了した直後に呼ぶこと。 */
function invalidateZubanCache_(zuban) {
  if (!zuban) return;
  try { CacheService.getScriptCache().remove('zubanInfo:' + zuban); } catch (e) {}
}

/**
 * QRから取り出した製番（例: I-PROのURL末尾の値）から、対応する図番を探す。
 * まず「製番インデックス」シートを見て、あれば即座に返す（I-PROスキャン省略）。
 * 無ければI-PRO同期データの「製造番号」列と一致する行を探し、見つかった単一候補の場合のみ
 * インデックスに書き込む（1製番に図番が複数見つかるケースはインデックス化しない。稀なケースであり、
 * 常にライブスキャンで正しく複数候補を出したいため）。
 */
function resolveZubanFromSeiban_(seiban) {
  if (!seiban) return { error: 'seiban is required' };

  var indexed = findSeibanIndexRow_(seiban);
  if (indexed) {
    return {
      found: true, seiban: seiban, multiple: false,
      candidates: [{ zuban: indexed['図番'], hinmei: indexed['品名'] || null, tokuisakiCode: indexed['得意先コード'] || null }],
      zuban: indexed['図番'], hinmei: indexed['品名'] || null, tokuisakiCode: indexed['得意先コード'] || null
    };
  }

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

  var result = {
    found: true,
    seiban: seiban,
    // 1製番に図番が複数見つかった場合、呼び出し側（フロント）で選択画面を出す
    multiple: uniqueCandidates.length > 1,
    candidates: uniqueCandidates,
    zuban: uniqueCandidates.length === 1 ? uniqueCandidates[0].zuban : null,
    hinmei: uniqueCandidates.length === 1 ? uniqueCandidates[0].hinmei : null,
    tokuisakiCode: uniqueCandidates.length === 1 ? uniqueCandidates[0].tokuisakiCode : null
  };

  if (!result.multiple && result.zuban) {
    upsertSeibanIndex_(seiban, result.zuban, result.hinmei, result.tokuisakiCode);
  }
  return result;
}

/**
 * QRスキャン専用の統合エンドポイント。①製番→図番変換と②その図番の情報一式取得を1回のリクエストに
 * まとめる。resolveZuban→zubanInfoを別々に呼ぶと、どちらも内部でI-PRO同期データ（大きいスプレッドシート）
 * をスキャンするため、QRスキャン1回あたりの待ち時間がほぼ倍になっていた（2026-08-29発見）。
 * resolveZubanFromSeiban_の時点で品名・得意先は既に分かっているため、getZubanInfo_に渡してスキャンを省略する。
 */
function scanZuban_(seiban) {
  var resolved = resolveZubanFromSeiban_(seiban);
  if (resolved.error || !resolved.found || resolved.multiple) {
    // 見つからない／複数候補の場合はフロント側で処理する（追加のzubanInfo取得はしない）
    return resolved;
  }
  // 製造番号で見つかる行は品名が入っていないシートのこともあるため（2026-08-29実データで確認）、
  // hinmeiが取れている場合だけ渡して省略し、取れていない場合はgetZubanInfo_側の
  // lookupZubanMaster_（図番での再検索）にフォールバックさせて正しさを優先する。
  var knownMaster = resolved.hinmei ? { hinmei: resolved.hinmei, tokuisaki: resolved.tokuisakiCode } : null;
  var info = getZubanInfo_(resolved.zuban, knownMaster);
  info.seiban = seiban;
  info.found = true;
  info.multiple = false;
  return info;
}

/**
 * 図番から品名・得意先コードを引く（②画面のヘッダー表示用）。
 * まず「図番インデックス」シートを見て、あれば即座に返す（I-PROスキャン省略）。
 * 無ければライブスキャン（scanZubanMaster_）し、結果をインデックスに書き込む。
 */
function lookupZubanMaster_(zuban) {
  var indexed = findZubanIndexRow_(zuban);
  if (indexed && indexed['品名']) {
    return { hinmei: indexed['品名'] || null, tokuisaki: indexed['得意先コード'] || null };
  }
  var master = scanZubanMaster_(zuban);
  if (master.hinmei) {
    upsertZubanIndex_(zuban, { '品名': master.hinmei, '得意先コード': master.tokuisaki || '' });
  }
  return master;
}

/** lookupZubanMaster_のインデックスを使わない版。インデックス自体の再構築（refreshOneZuban_）で使う。 */
function scanZubanMaster_(zuban) {
  var row = findIproRowByColumn_('品番(図番）', zuban) || findIproRowByColumn_('品番(図番)', zuban);
  if (!row) return { hinmei: null, tokuisaki: null };
  return { hinmei: row['品名'] || null, tokuisaki: row['得意先コード'] || null };
}

/**
 * I-PRO関連データから、指定列が指定値と一致する最初の行を返す。
 * どのタブに目的の列があるか固定できていないため、全タブを順に探す。
 */
function findIproRowByColumn_(columnName, value) {
  var rows = findIproRowsByColumn_(columnName, value);
  return rows.length > 0 ? rows[0] : null;
}

/**
 * findIproRowByColumn_ の複数件版。1製番に図番が複数ある場合の検出に使う。
 * まず小さく速い「進捗状況照会」（製造番号・品番(図番)・品名・得意先が同じ行に揃っている）を探し、
 * 見つからない場合のみ大きい「I-Pro Source」（全件、約10倍のサイズ）にフォールバックする
 * （ユーザー提案、2026-08-30）。
 */
function findIproRowsByColumn_(columnName, value) {
  if (!value) return [];
  var fromProgress = findRowsInSpreadsheet_(IPRO_PROGRESS_SPREADSHEET_ID, columnName, value);
  if (fromProgress.length > 0) return fromProgress;
  return findRowsInSpreadsheet_(IPRO_SOURCE_SPREADSHEET_ID, columnName, value);
}

/**
 * 指定スプレッドシート内で、指定列が指定値と一致する行をすべて返す（全タブ対象）。
 * パフォーマンス上の注意：目的の列を持たないタブの全データを読み込まない（ヘッダー行だけ先に確認し、
 * 列が無ければ即スキップする）。これが無いと1回の呼び出しに30秒近くかかっていた（2026-08-29計測）。
 */
function findRowsInSpreadsheet_(spreadsheetId, columnName, value) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheets = ss.getSheets();
  var out = [];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var col = header.indexOf(columnName);
    if (col === -1) continue; // 目的の列が無いタブは全データを読み込まずスキップ
    var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][col]) === String(value)) {
        out.push(rowToObject_(header, values[i]));
      }
    }
  }
  return out;
}

/**
 * DriveApp.searchFiles() / Folder.searchFiles() / Folder.getFoldersByName() は、
 * 共有ドライブ（Team Drive）上のファイルを検索対象に含められない既知の制限があり、
 * 検査記録フォルダ等が共有ドライブ「ドライブ」配下にあるため常に0件になっていた
 * （2026-08-29、DriveApp.getFolderByIdでの直接アクセスは成功するのに検索系だけ0件になる現象で発覚）。
 * 代わりにAdvanced Drive Service（appsscript.jsonでenabledAdvancedServicesに追加済み）の
 * Drive.Files.listをsupportsAllDrives/includeItemsFromAllDrives付きで使う。
 */
function driveFilesList_(query) {
  var res = Drive.Files.list({
    q: query,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    corpora: 'allDrives',
    fields: 'files(id, name, mimeType, webViewLink)'
  });
  return res.files || [];
}

/**
 * 過去トラ：{図番} 品質情報 スプレッドシート／改善計画書フォルダを検索し、内容を要約して返す。
 * 品質情報はそのまま、改善計画書は一行要約＋元ファイルへのリンクのみ返す（本文は開かない）。
 *
 * まず「図番インデックス」シートを見て、既に調べたことがあれば（更新日時があれば）そのリンクを
 * そのまま使い、Drive検索を省略する。無ければscanPastTroubleFiles_でライブ検索し、結果をインデックスに書き込む。
 * refreshZubanIndex()が毎日1回、インデックス済みの図番だけ再スキャンして最新化する。
 */
function findPastTrouble_(zuban) {
  var items = [];
  var seenKeys = {};
  function pushUnique(item, key) {
    if (seenKeys[key]) return;
    seenKeys[key] = true;
    items.push(item);
  }

  // 「更新日時」は品名等インデックス（lookupZubanMaster_）と共用の列のため、そちらだけが先に
  // 書き込まれているケースがある（getZubanInfo_内でlookupZubanMaster_→findPastTrouble_の順に呼ばれるため）。
  // 更新日時だけでなく、この機能が実際に書き込む品質情報リンク・改善計画書リンク自体が
  // 存在するか（JSON.stringifyされた文字列は結果が0件でも"[]"という空でない文字列になる）で判定する。
  var indexed = findZubanIndexRow_(zuban);
  var qiFiles, fkFiles;
  if (indexed && indexed['品質情報リンク'] && indexed['改善計画書リンク']) {
    try { qiFiles = JSON.parse(indexed['品質情報リンク'] || '[]'); } catch (e) { qiFiles = []; }
    try { fkFiles = JSON.parse(indexed['改善計画書リンク'] || '[]'); } catch (e) { fkFiles = []; }
  } else {
    var scanned = scanPastTroubleFiles_(zuban);
    qiFiles = scanned.qi;
    fkFiles = scanned.fk;
    upsertZubanIndex_(zuban, {
      '品質情報リンク': JSON.stringify(qiFiles),
      '改善計画書リンク': JSON.stringify(fkFiles)
    });
  }

  qiFiles.forEach(function (f) {
    pushUnique({
      source: '品質情報',
      title: f.name,
      url: f.url,
      // TODO: シート本文の読み取り・日付ごとのコメント抽出・Sheet.getImages()での写真検出は未実装
      note: 'シート本文の要約読み取りは未実装。まずはリンクのみ。'
    }, f.url);
  });
  fkFiles.forEach(function (f) {
    pushUnique({
      source: '改善計画書',
      title: f.name,
      url: f.url,
      note: '一行要約（不適合事象）の自動抽出は未実装。まずはフォルダへのリンクのみ。'
    }, f.url);
  });

  // 新システム内で「共有する」を選んだ品質情報記録・ツール配置メモ
  items = items.concat(findSharedEntries_(SHEET_QUALITY_LOG, zuban));
  items = items.concat(findSharedEntries_(SHEET_TOOL_MEMO, zuban));

  return items;
}

/**
 * findPastTrouble_のライブスキャン部分（インデックスを使わない）。
 * 改善計画書フォルダの命名・置き場所は実データ上ゆれが大きいことが判明済み（2026-08-29）：
 * 「不具合改善計画書」「不適合改善計画書」の両方の表記があり、置き場所も
 * 少なくとも3パターン確認済み（①「(KP)社内不良改善計画書」フォルダ直下、
 * ②「(CC)客先クレーム改善計画書」フォルダ直下、③図番フォルダ自体の直下）。
 * 集約フォルダが今後も増える／変わる可能性があるため、特定の親フォルダに絞り込む
 * アプローチはやめ、両方の表記に共通する「改善計画書」で緩く一致させたうえで
 * 共有ドライブ全体を対象に検索する（Advanced Drive Serviceで共有ドライブ対応済みのため可能）。
 */
function scanPastTroubleFiles_(zuban) {
  var qi = [];
  var zubanFolder = findZubanFolder_(zuban);
  if (zubanFolder) {
    driveFilesList_(
      "name contains '品質情報' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false and '" + zubanFolder.getId() + "' in parents"
    ).forEach(function (f) { qi.push({ name: f.name, url: f.webViewLink }); });
  }

  var zubanEsc = String(zuban).replace(/'/g, "\\'");
  var fk = driveFilesList_(
    "name contains '" + zubanEsc + "' and name contains '改善計画書' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
  ).map(function (f) { return { name: f.name, url: f.webViewLink }; });

  return { qi: qi, fk: fk };
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

/**
 * 図番インデックス・製番インデックス（2026-08-30追加）。
 * 一度調べた図番・製番の結果（品名・得意先・検査記録フォルダ・品質情報リンク・改善計画書リンク）を
 * 「西軽精機ナレッジ DB」の専用シートに保存し、次回以降はI-PROスキャンやDrive検索を省略して
 * このシートを読むだけで返せるようにする（ユーザー提案、2026-08-30）。refreshZubanIndex()が
 * 毎日1回、既にインデックス済みの分だけ再スキャンして最新化する。
 */
function findSeibanIndexRow_(seiban) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SEIBAN_INDEX);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = header.indexOf('製造番号');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(seiban)) return rowToObject_(header, values[i]);
  }
  return null;
}

function upsertSeibanIndex_(seiban, zuban, hinmei, tokuisakiCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SEIBAN_INDEX);
  var newRow = [seiban, zuban, hinmei || '', tokuisakiCode || '', new Date()];
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = header.indexOf('製造番号');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(seiban)) {
      sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
      return;
    }
  }
  sheet.appendRow(newRow);
}

function findZubanIndexRow_(zuban) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = header.indexOf('図番');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(zuban)) return rowToObject_(header, values[i]);
  }
  return null;
}

/** fieldsに渡したキーだけ更新する（他の既存フィールドはそのまま残す）。図番の行が無ければ新規追加。 */
function upsertZubanIndex_(zuban, fields) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = header.indexOf('図番');
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(zuban)) {
      var existing = rowToObject_(header, values[i]);
      var merged = Object.assign({}, existing, fields, { '更新日時': new Date() });
      var updatedRow = header.map(function (h) { return merged[h] !== undefined ? merged[h] : ''; });
      sheet.getRange(i + 1, 1, 1, header.length).setValues([updatedRow]);
      return;
    }
  }
  var merged2 = Object.assign({ '図番': zuban }, fields, { '更新日時': new Date() });
  var newRow = header.map(function (h) { return merged2[h] !== undefined ? merged2[h] : ''; });
  sheet.appendRow(newRow);
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
  invalidateZubanCache_(payload.zuban);
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
  invalidateZubanCache_(payload.zuban);
  return { id: id };
}

/**
 * 投稿ID(「投稿ID」列)で1件を更新／削除する共通処理。
 * 承認フローなし・投稿者本人に限定しない「誰でも編集・削除できる」設計（2026-08-14合意）のため、
 * 呼び出し元での権限チェックは行わない。
 */
function findRowIndexById_(sheet, idCol, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2 || idCol === -1) return -1;
  var ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // シート上の実際の行番号（1始まり）
  }
  return -1;
}

function updatePostById_(sheetName, postId, fields) {
  if (!postId) return { error: 'postId is required' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowNum = findRowIndexById_(sheet, header.indexOf('投稿ID'), postId);
  if (rowNum === -1) return { error: '投稿が見つかりません' };
  var zuban = sheet.getRange(rowNum, header.indexOf('図番') + 1).getValue();
  Object.keys(fields).forEach(function (key) {
    var col = header.indexOf(key);
    if (col !== -1) sheet.getRange(rowNum, col + 1).setValue(fields[key]);
  });
  invalidateZubanCache_(zuban);
  return { ok: true };
}

function deletePostById_(sheetName, postId) {
  if (!postId) return { error: 'postId is required' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowNum = findRowIndexById_(sheet, header.indexOf('投稿ID'), postId);
  if (rowNum === -1) return { error: '投稿が見つかりません' };
  var zuban = sheet.getRange(rowNum, header.indexOf('図番') + 1).getValue();
  sheet.deleteRow(rowNum);
  invalidateZubanCache_(zuban);
  return { ok: true };
}

/** ツール配置メモの更新／削除（③編集画面）。 */
function updateToolMemo_(payload) {
  return updatePostById_(SHEET_TOOL_MEMO, payload.postId, {
    '内容': payload.content || '', '写真URL': payload.photoUrl || '', '共有フラグ': !!payload.shared
  });
}
function deleteToolMemo_(payload) {
  return deletePostById_(SHEET_TOOL_MEMO, payload.postId);
}

/** 品質情報記録の更新／削除（⑤画面）。 */
function updateQualityLog_(payload) {
  return updatePostById_(SHEET_QUALITY_LOG, payload.postId, {
    '内容': payload.content || '', '外観ランク': payload.rank || '', '写真URL': payload.photoUrl || '', '共有フラグ': !!payload.shared
  });
}
function deleteQualityLog_(payload) {
  return deletePostById_(SHEET_QUALITY_LOG, payload.postId);
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
  invalidateZubanCache_(payload.zuban);
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
  invalidateZubanCache_(payload.zuban);
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

/**
 * ログインユーザーのメールアドレスから、組織図マスタを照合して部署（役割）を判定する。
 * 役割は5種類：seizou（一次・二次加工）／kensa（検査）／seisan（生産管理・出荷担当以外）／
 * shukka（生産管理課・出荷担当）／other（仕上げ・洗浄など、上記以外）。
 * 組織図マスタに見つからない場合はotherを返す（安全側：機能を絞る方向のデフォルト）。
 */
function getRole_(email) {
  if (!email) return { error: 'email is required' };
  var allAccess = ALL_ACCESS_EMAILS.indexOf(String(email).trim().toLowerCase()) !== -1;
  var ss = SpreadsheetApp.openById(ORG_MASTER_SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (sheet.getLastRow() < 2) continue;
    var values = sheet.getDataRange().getValues();
    var header = values[0];
    var mailCol = header.indexOf('Mail Address');
    if (mailCol === -1) mailCol = header.indexOf('メールアドレス');
    var sectionCol = header.indexOf('課名');
    var processCol = header.indexOf('工程名');
    var nameCol = header.indexOf('氏名');
    if (mailCol === -1) continue;
    for (var i = 1; i < values.length; i++) {
      var rowMail = String(values[i][mailCol] || '').trim().toLowerCase();
      if (rowMail && rowMail === String(email).trim().toLowerCase()) {
        var section = sectionCol >= 0 ? values[i][sectionCol] : '';
        var process = processCol >= 0 ? values[i][processCol] : '';
        return {
          found: true,
          name: nameCol >= 0 ? values[i][nameCol] : '',
          section: section,
          process: process,
          role: resolveRole_(section, process),
          allAccess: allAccess
        };
      }
    }
  }
  return { found: false, role: 'other', allAccess: allAccess };
}

/** 課名・工程名の文字列から役割キーを判定する。想定外の組み合わせはother（安全側）にフォールバック。 */
function resolveRole_(section, process) {
  section = String(section || '');
  process = String(process || '');
  if (section.indexOf('製造課') !== -1 && (process.indexOf('一次加工') !== -1 || process.indexOf('二次加工') !== -1)) {
    return 'seizou';
  }
  if (section.indexOf('生産管理課') !== -1 && process.indexOf('出荷') !== -1) {
    return 'shukka';
  }
  if (section.indexOf('生産管理課') !== -1) {
    return 'seisan';
  }
  if (section.indexOf('品質保証課') !== -1 && process.indexOf('検査') !== -1) {
    return 'kensa';
  }
  return 'other';
}

/**
 * 図番に対応する検査記録フォルダ（検査記録の親フォルダ／社名／図番／）を探し、URLを返す。
 * 「図番インデックス」に既にURLがあればそれを返し、無ければfindZubanFolder_で探してインデックスに書き込む。
 */
function getInspectionFolderUrl_(zuban) {
  if (!zuban) return { error: 'zuban is required' };
  var indexed = findZubanIndexRow_(zuban);
  if (indexed && indexed['更新日時'] && indexed['検査記録フォルダURL']) {
    return { found: true, zuban: zuban, url: indexed['検査記録フォルダURL'] };
  }
  var folder = findZubanFolder_(zuban);
  if (!folder) return { found: false, zuban: zuban };
  upsertZubanIndex_(zuban, { '検査記録フォルダURL': folder.getUrl() });
  return { found: true, zuban: zuban, url: folder.getUrl() };
}

/**
 * タイトルが図番と完全一致するフォルダを探す（検査記録／社名／図番／の図番フォルダを想定）。
 * 検索自体はAdvanced Drive Service（driveFilesList_）で行い、見つかったIDをDriveApp.getFolderByIdで
 * 開き直してFolderオブジェクトとして返す（getFolderByIdは共有ドライブ上でも問題なく使えるため、
 * 戻り値を使う側（createFolder/createFile等）はこれまでどおりDriveAppのAPIで操作できる）。
 */
function findZubanFolder_(zuban) {
  var name = String(zuban).replace(/'/g, "\\'");
  var files = driveFilesList_(
    "name = '" + name + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
  );
  if (files.length === 0) return null;
  return DriveApp.getFolderById(files[0].id);
}

/**
 * 写真アップロード（③ツール配置メモ・⑤品質情報記録の写真添付欄）。
 * 図番フォルダ（検査記録／社名／図番／）配下の「写真」サブフォルダに保存する。
 * 図番フォルダがまだ存在しない場合（検査記録が未作成の図番）はエラーを返す
 * （社名フォルダの自動作成は、得意先コード→社名の対応表が未整備のため未実装。TODO）。
 */
function uploadPhoto_(payload) {
  if (!payload.zuban || !payload.dataBase64) return { error: 'zuban and dataBase64 are required' };
  var zubanFolder = findZubanFolder_(payload.zuban);
  if (!zubanFolder) {
    return { error: 'この図番の検査記録フォルダが見つからないため、写真を保存できません（図番: ' + payload.zuban + '）' };
  }
  var photoFiles = driveFilesList_(
    "name = '写真' and mimeType = 'application/vnd.google-apps.folder' and trashed = false and '" + zubanFolder.getId() + "' in parents"
  );
  var photoFolder = photoFiles.length > 0 ? DriveApp.getFolderById(photoFiles[0].id) : zubanFolder.createFolder('写真');

  var mimeType = payload.mimeType || 'image/jpeg';
  var bytes = Utilities.base64Decode(payload.dataBase64);
  var blob = Utilities.newBlob(bytes, mimeType, payload.filename || (Utilities.getUuid() + '.jpg'));
  var file = photoFolder.createFile(blob);
  return { url: file.getUrl() };
}

/**
 * 図番インデックスを最新化する（2026-08-30追加）。時間主導トリガーから1日1回呼ばれる想定。
 * 全件スキャンではなく、既にインデックス済みの図番だけを再チェックする
 * （新規に検索されたことのない図番はここでは追加しない。初回スキャン時に自動でインデックスされるため）。
 * GASの実行時間制限（6分）に収まるよう、時間を超えたら打ち切って次回の実行に委ねる。
 * setupDailyIndexTriggerを1回実行してトリガー登録すること。
 */
function refreshZubanIndex() {
  var startTime = Date.now();
  var maxRunMs = 5 * 60 * 1000;

  var zubanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
  if (!zubanSheet || zubanSheet.getLastRow() < 2) {
    Logger.log('図番インデックスは空のため、更新対象なし');
    return;
  }
  var header = zubanSheet.getRange(1, 1, 1, zubanSheet.getLastColumn()).getValues()[0];
  var zubanCol = header.indexOf('図番');
  var rows = zubanSheet.getDataRange().getValues();
  var refreshed = 0;
  for (var i = 1; i < rows.length; i++) {
    if (Date.now() - startTime > maxRunMs) {
      Logger.log('実行時間の上限に近づいたため打ち切り（' + refreshed + '件更新、残り' + (rows.length - i) + '件は次回に持ち越し）');
      break;
    }
    var zuban = rows[i][zubanCol];
    if (!zuban) continue;
    refreshOneZuban_(zuban);
    refreshed++;
  }
  Logger.log('図番インデックス更新完了: ' + refreshed + '件');
}

/** 指定した図番1件分を、インデックスを使わずライブスキャンして図番インデックスを上書きする。 */
function refreshOneZuban_(zuban) {
  var master = scanZubanMaster_(zuban);
  var scanned = scanPastTroubleFiles_(zuban);
  var zubanFolder = findZubanFolder_(zuban);
  upsertZubanIndex_(zuban, {
    '品名': master.hinmei || '',
    '得意先コード': master.tokuisaki || '',
    '検査記録フォルダURL': zubanFolder ? zubanFolder.getUrl() : '',
    '品質情報リンク': JSON.stringify(scanned.qi),
    '改善計画書リンク': JSON.stringify(scanned.fk)
  });
  invalidateZubanCache_(zuban); // zubanInfoの5分キャッシュも合わせて破棄し、更新をすぐ反映させる
}

/** 図番インデックスの毎日自動更新トリガーを登録する。GASエディタで1回だけ手動実行すること。 */
function setupDailyIndexTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshZubanIndex') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshZubanIndex')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  Logger.log('毎日3時台にrefreshZubanIndexを実行するトリガーを登録しました');
}
