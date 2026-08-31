/**
 * 西軽精機ナレッジ（仮） バックエンド
 * DBスプレッドシート: 1wmPkFZ-EdtUQchWlVGspeclP3yHbO9YyViAM5AraQj0
 *
 * デプロイ後の確認手順は加工ナレッジ/CLAUDE.mdの「デプロイ・動作確認の手順」に準拠。
 * setupSheets はタブ構成を作る一回限りの関数。GASエディタで手動実行すること。
 * gas/への変更はmainにpushすればGitHub Actionsが自動でclasp push/deployする
 * （.github/workflows/deploy-gas.yml、2026-08-30導入）。
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

// Google OAuthクライアントID（index.htmlのGOOGLE_CLIENT_IDと同じ値）。IDトークンのサーバー側検証で、
// 他アプリ向けに発行されたトークンを受け付けないようaudクレームと突き合わせるのに使う。
var OAUTH_CLIENT_ID = '800178947678-t49i9pr40ci70th6dgpuslfr4dldqjqh.apps.googleusercontent.com';

// 動作確認用：全部署の画面にアクセスできるアカウント（本来の役割による出し分けとは別に、確認のため常時allAccess:trueを返す）。
var ALL_ACCESS_EMAILS = ['n-moriizumi@nishikaru.co.jp'];

// ④紙の取込確認（ツールレイアウト表のAI読み取り）で使うGeminiモデル。APIキー自体はコードに書かず
// スクリプトのプロパティ「GEMINI_API_KEY」に保存する（このリポジトリはpublicのため、コードに
// 書くとキーが漏洩する）。2026-08-31のパイロットテストで最も精度が高かったモデルを採用。
var GEMINI_MODEL = 'gemini-3.1-pro-preview';

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
    '検査記録の添付', 'ミルシート', 'トレー梱包', 'NG限度見本', 'キーエンス測定', 'キーエンスプログラム名',
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
    if (action === 'ocrToolLayout') return jsonResponse_(ocrToolLayout_(payload));
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
  // 未登録の図番（I-PROに存在しない等）でzubanInfoが異常に遅くなる事例が報告されたため、
  // どの処理に時間がかかっているか特定できるよう時間計測ログを仕込む（5秒超の場合のみ記録、
  // 通常時のログを埋もれさせないため）。2026-08-31。
  var t0 = Date.now();
  var master = knownMaster || lookupZubanMaster_(zuban);
  var t1 = Date.now();
  var pastTrouble = findPastTrouble_(zuban);
  var t2 = Date.now();
  var qualityLog = readQualityLog_(zuban);
  var toolMemo = readToolMemo_(zuban);
  var toolPositions = readToolPositions_(zuban);
  var shippingSpec = readShippingSpec_(zuban);
  var t3 = Date.now();
  if (t3 - t0 > 5000) {
    Logger.log('getZubanInfo_が遅延（図番:' + zuban + '）: 品名等取得=' + (t1 - t0) + 'ms, 過去トラ検索=' + (t2 - t1) + 'ms, その他読み取り=' + (t3 - t2) + 'ms, 合計=' + (t3 - t0) + 'ms');
  }

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
/**
 * 「進捗状況照会」の品番(図番)列には、社内不良(KP)・客先クレーム(CC)関連の改善計画書が
 * 紐づく作業行だけ、先頭に"KP "/"CC "という注記が付くことがある（実データで確認、2026-08-30。
 * 例：「KP 210-404133-1」）。これは実際のDriveフォルダ名やI-Pro Source側の値には出てこない、
 * 進捗状況照会シート内だけの注記のため、図番として扱う際は取り除く。
 */
function stripZubanPrefix_(value) {
  return String(value || '').replace(/^(KP|CC)\s+/, '');
}

/**
 * 数字だけの図番を照合する際のキーを作る（2026-08-30）。
 * 「進捗状況照会」は10分ごとに外部連携でデータが更新されるため、数字だけの図番
 * （例:"03624100"）のセルがテキスト（0付き）と数値（0落ち、"3624100"）の間で
 * 更新のたびに不安定に変わることが実データで確認された。単純な文字列一致で
 * 図番インデックスの行を照合すると、この0落ちのタイミングで「別の図番」と
 * 誤認識され、seedZubanIndexが同じ図番を何度も新規と判定して重複行を量産して
 * しまう（実例：3624100が50行以上重複）。数字だけの図番は先頭の0の有無に
 * 関わらず同じ図番として扱うため、整数値としてのキーに正規化する。
 * 数字以外を含む図番（ハイフンや英字入り）はそのまま。
 */
function numericZubanKey_(value) {
  var s = String(value || '');
  return /^\d+$/.test(s) ? String(parseInt(s, 10)) : s;
}

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
      zuban: stripZubanPrefix_(row['品番(図番）'] || row['品番(図番)']),
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

/**
 * lookupZubanMaster_のインデックスを使わない版。インデックス自体の再構築（refreshOneZuban_）で使う。
 * 「進捗状況照会」の品番(図番)列には、社内不良(KP)・客先クレーム(CC)関連の作業行だけ
 * 「KP 」「CC 」という接頭辞が付くことがある（stripZubanPrefix_の説明を参照）。図番から逆引きする
 * ここでは、接頭辞付きの値でも見つけられるよう両方の形で探す（2026-08-30、実データで発覚・修正）。
 */
function scanZubanMaster_(zuban) {
  // 表記ゆれ（全角/半角カッコの列名）・接頭辞（KP/CC）ありなしを、以前は6通り総当たりで
  // 1つずつ別呼び出ししていたため、I-PROに存在しない図番だと6回ともI-Pro Source（大きい方）まで
  // フォールバックしてしまい、1回のzubanInfoが2分以上かかる不具合があった（2026-08-31、実機ログで確認）。
  // 列名・値の候補をまとめて渡し、シートの読み込み自体は1回で済むようにした。
  var row = findIproRowByColumn_(
    ['品番(図番）', '品番(図番)'],
    [zuban, 'KP ' + zuban, 'CC ' + zuban]
  );
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
 * 指定スプレッドシート内で、指定列（表記ゆれ等の候補を配列で渡せる）が指定値（複数候補可）の
 * いずれかと一致する行をすべて返す（全タブ対象）。columnName/valueは単一の文字列でも配列でもよい
 * （単一値を渡す既存の呼び出し元とも互換）。
 * パフォーマンス上の注意：目的の列を持たないタブの全データを読み込まない（ヘッダー行だけ先に確認し、
 * 列が無ければ即スキップする）。これが無いと1回の呼び出しに30秒近くかかっていた（2026-08-29計測）。
 * また列名・値の候補が複数ある場合も、シートの読み込み自体は1回で済むようにしている
 * （以前は候補の組み合わせごとに毎回シートを開き直しており、I-PROに存在しない図番だと
 * 数分かかる不具合の原因になっていた。2026-08-31修正）。
 */
function findRowsInSpreadsheet_(spreadsheetId, columnNames, values) {
  var columnNameList = [].concat(columnNames);
  var valueSet = {};
  [].concat(values).forEach(function (v) { valueSet[String(v)] = true; });

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheets = ss.getSheets();
  var out = [];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var col = -1;
    for (var c = 0; c < columnNameList.length; c++) {
      col = header.indexOf(columnNameList[c]);
      if (col !== -1) break;
    }
    if (col === -1) continue; // 目的の列が無いタブは全データを読み込まずスキップ
    var rowsValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    for (var i = 1; i < rowsValues.length; i++) {
      if (valueSet[String(rowsValues[i][col])]) {
        out.push(rowToObject_(header, rowsValues[i]));
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
  var key = numericZubanKey_(zuban);
  for (var i = 1; i < values.length; i++) {
    if (numericZubanKey_(values[i][col]) === key) return rowToObject_(header, values[i]);
  }
  return null;
}

/**
 * fieldsに渡したキーだけ更新する（他の既存フィールドはそのまま残す）。図番の行が無ければ新規追加。
 * 読み取り→書き込みの間にLockServiceで排他制御し、実行が重なっても重複行ができないようにする
 * （2026-08-30、図番インデックスに大量の重複行が発生した問題への対処の一環）。
 */
function upsertZubanIndex_(zuban, fields) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
    var values = sheet.getDataRange().getValues();
    var header = values[0];
    var col = header.indexOf('図番');
    var key = numericZubanKey_(zuban);
    for (var i = 1; i < values.length; i++) {
      if (numericZubanKey_(values[i][col]) === key) {
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
  } finally {
    lock.releaseLock();
  }
}

/**
 * GoogleのIDトークンをサーバー側で検証し、検証済みのメールアドレス・氏名を返す（2026-08-30追加）。
 * これまではクライアント側でJWTをデコードしただけの値（`payload.userEmail`/`userName`）を
 * そのまま信用していたため、理論上は誰でもfetch()を直接叩いて他人になりすまして投稿できた。
 * 投稿者の身元を記録する系のAPI（postQualityLog_/postToolMemo_/saveToolPositions_/
 * saveShippingSpec_）は、この検証を通った値だけを信用するようにする。
 * GAS単体でJWTの署名検証を実装するのは煩雑なため、Google公式のtokeninfoエンドポイントを使う
 * （検証はGoogle側で行われ、結果としてデコード済みのクレームが返る）。
 */
function verifyIdToken_(idToken) {
  if (!idToken) return { error: 'ログインが必要です（IDトークンがありません）' };
  var res;
  try {
    res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
  } catch (e) {
    return { error: 'ログイン情報の確認に失敗しました: ' + String(e) };
  }
  if (res.getResponseCode() !== 200) {
    return { error: 'ログイン情報が無効です。再度サインインしてからお試しください。' };
  }
  var claims = JSON.parse(res.getContentText());
  if (claims.aud !== OAUTH_CLIENT_ID) {
    return { error: 'ログイン情報が無効です（対象アプリが一致しません）' };
  }
  if (claims.email_verified !== 'true' && claims.email_verified !== true) {
    return { error: 'メールアドレスが確認されていません' };
  }
  if (!claims.email || claims.email.toLowerCase().indexOf('@nishikaru.co.jp') === -1) {
    return { error: '社内アカウントでのログインが必要です' };
  }
  return { email: claims.email, name: claims.name || claims.email };
}

/** payload.idTokenを検証し、成功したら検証済みidentityを渡してfnを実行する。失敗時はエラーをそのまま返す。 */
function withVerifiedIdentity_(payload, fn) {
  var identity = verifyIdToken_(payload.idToken);
  if (identity.error) return identity;
  return fn(identity);
}

/** 品質情報記録の投稿（⑤画面）。承認フローなし、送信したら即座に反映。 */
function postQualityLog_(payload) {
  return withVerifiedIdentity_(payload, function (identity) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUALITY_LOG);
    var id = Utilities.getUuid();
    sheet.appendRow([
      id, new Date(), payload.zuban, payload.department,
      identity.email, identity.name,
      payload.rank || '', payload.content || '', payload.photoUrl || '',
      !!payload.shared
    ]);
    invalidateZubanCache_(payload.zuban);
    return { id: id };
  });
}

/** ツール配置メモの投稿（③画面）。承認フローなし。 */
function postToolMemo_(payload) {
  return withVerifiedIdentity_(payload, function (identity) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TOOL_MEMO);
    var id = Utilities.getUuid();
    sheet.appendRow([
      id, new Date(), payload.zuban,
      identity.email, identity.name,
      payload.content || '', payload.photoUrl || '',
      !!payload.shared
    ]);
    invalidateZubanCache_(payload.zuban);
    return { id: id };
  });
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

/**
 * URLからDriveファイルIDを取り出し、ゴミ箱に移動する（完全削除ではなく、誤操作時に復元できるように）。
 * ファイルが既に無い等のエラーは無視する（孤立ファイルの掃除が主目的で、失敗しても投稿の更新・削除自体は
 * 妨げたくないため）。
 */
function trashDriveFileByUrl_(url) {
  if (!url) return;
  var m = /\/d\/([^/]+)/.exec(String(url));
  if (!m) return;
  try {
    DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (e) {
    // ファイルが見つからない等は無視
  }
}

function updatePostById_(sheetName, postId, fields) {
  if (!postId) return { error: 'postId is required' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowNum = findRowIndexById_(sheet, header.indexOf('投稿ID'), postId);
  if (rowNum === -1) return { error: '投稿が見つかりません' };
  var zuban = sheet.getRange(rowNum, header.indexOf('図番') + 1).getValue();

  // 写真が新しいものに差し替えられた場合、Drive上の古い写真ファイルが孤立して残らないようゴミ箱へ移動する。
  if (Object.prototype.hasOwnProperty.call(fields, '写真URL')) {
    var photoCol = header.indexOf('写真URL');
    if (photoCol !== -1) {
      var oldPhotoUrl = sheet.getRange(rowNum, photoCol + 1).getValue();
      if (oldPhotoUrl && oldPhotoUrl !== fields['写真URL']) {
        trashDriveFileByUrl_(oldPhotoUrl);
      }
    }
  }

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

  // 投稿削除時、写真が添付されていればDrive上のファイルもゴミ箱へ移動する（孤立ファイル防止）。
  var photoCol = header.indexOf('写真URL');
  if (photoCol !== -1) {
    var photoUrl = sheet.getRange(rowNum, photoCol + 1).getValue();
    if (photoUrl) trashDriveFileByUrl_(photoUrl);
  }

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
  return withVerifiedIdentity_(payload, function (identity) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TOOL_POSITIONS);
    deleteRowsByZuban_(sheet, payload.zuban);
    var now = new Date();
    (payload.positions || []).forEach(function (p) {
      sheet.appendRow([
        payload.zuban, p.column, p.order, p.tNumber, p.description,
        payload.frontChuck || '', payload.backChuck || '', payload.cycleTime || '',
        payload.toolStorage || '', payload.forwardPosition || '',
        identity.email, now
      ]);
    });
    invalidateZubanCache_(payload.zuban);
    return { ok: true };
  });
}

/** 出荷仕様の保存（⑤画面の各セクション）。図番ごとに1行、上書き更新。 */
function saveShippingSpec_(payload) {
  return withVerifiedIdentity_(payload, function (identity) {
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
    merged['最終更新者メール'] = identity.email;
    merged['最終更新日時'] = new Date();
    var newRow = header.map(function (h) { return merged[h] !== undefined ? merged[h] : ''; });
    if (rowIndex >= 0) {
      sheet.getRange(rowIndex + 1, 1, 1, header.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    invalidateZubanCache_(payload.zuban);
    return { ok: true };
  });
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
 * 図番フォルダが無い場合、得意先名から既存の会社名フォルダを探し、その下に図番フォルダを新規作成する
 * （2026-08-30追加、ユーザー提案）。
 *
 * 会社名の完全一致のみで探す（前方一致は一度実装したが、実データで「松永精密」→
 * 「松永精密工業 KSP75-FP-007479(3)-003-01 マグネットキャッチ 不適合改善計画書」という
 * 別図番の改善計画書フォルダを会社フォルダと誤認識し、無関係な場所に図番フォルダを作ってしまう
 * 事故が発生したため、2026-08-30に完全一致のみに変更した）。
 * 完全一致が見つからない場合は自動作成をあきらめ、エラーを返す（想定外の場所に新しいフォルダを
 * 作って既存の置き場所と混同・分散するより、作成できない方が安全なため）。
 */
function findExistingCompanyFolder_(companyName) {
  var nameEsc = String(companyName).replace(/'/g, "\\'");
  var exact = driveFilesList_(
    "name = '" + nameEsc + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
  );
  return exact.length > 0 ? exact[0] : null;
}

/** 得意先コードから得意先名を引く（「進捗状況照会」に得意先コード・得意先名が同じ行に揃っているため）。 */
function lookupTokuisakiName_(tokuisakiCode) {
  if (!tokuisakiCode) return null;
  var rows = findRowsInSpreadsheet_(IPRO_PROGRESS_SPREADSHEET_ID, '得意先コード', tokuisakiCode);
  return rows.length > 0 ? (rows[0]['得意先名'] || null) : null;
}

/** 図番フォルダが無い場合に、既存の会社名フォルダの下へ新規作成を試みる。作れなければnullを返す。 */
function createZubanFolderIfCompanyKnown_(zuban) {
  var master = scanZubanMaster_(zuban);
  var tokuisakiName = lookupTokuisakiName_(master.tokuisaki);
  if (!tokuisakiName) return null;
  var companyFolderMeta = findExistingCompanyFolder_(tokuisakiName);
  if (!companyFolderMeta) return null;
  var companyFolder = DriveApp.getFolderById(companyFolderMeta.id);
  return companyFolder.createFolder(zuban);
}

/**
 * 写真アップロード（③ツール配置メモ・⑤品質情報記録の写真添付欄）。
 * 図番フォルダ（検査記録／社名／図番／）配下の「写真」サブフォルダに保存する。
 * 図番フォルダがまだ存在しない場合は、既存の会社名フォルダが見つかれば自動作成する
 * （createZubanFolderIfCompanyKnown_）。会社名フォルダ自体が見つからない場合はエラーを返す。
 */
function uploadPhoto_(payload) {
  if (!payload.zuban || !payload.dataBase64) return { error: 'zuban and dataBase64 are required' };
  var zubanFolder = findZubanFolder_(payload.zuban);
  if (!zubanFolder) {
    zubanFolder = createZubanFolderIfCompanyKnown_(payload.zuban);
  }
  if (!zubanFolder) {
    return { error: 'この図番の検査記録フォルダが見つからず、自動作成もできませんでした（得意先が特定できないか、既存の会社名フォルダが見つかりません）。図番: ' + payload.zuban };
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
 * ④紙の取込確認：ツールレイアウト表（手書き含む）の写真をGemini APIに送り、
 * ③ツール配置画面の編集データと同じ形（正面/背面/サイクルチャック径・工具リスト・変更履歴メモ等）で
 * 読み取り結果を返す（2026-08-31、実データでのパイロットテストを経て導入）。
 * この時点ではDB・図番インデックスへは一切書き込まない。返した内容は必ずアプリ側で人が
 * 原本の写真と1件ずつ照合してから、既存のsaveToolPositions/postToolMemoで保存する。
 * パイロットテストで判明した既知の限界：型番・数値は概ね高精度だが、まれに1桁程度の誤読が
 * 起こりうる（例:"#6713"→"#67B"、"φ1.53"→"φ1.55"）。この確認ステップは省略しないこと。
 */
function ocrToolLayout_(payload) {
  if (!payload.imageBase64) return { error: 'imageBase64 is required' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error: 'GEMINI_API_KEYが設定されていません（GASエディタの「プロジェクトの設定」→「スクリプト プロパティ」で設定してください）' };

  var prompt = 'これは工場のNCツールレイアウト表（手書き含む）の写真です。以下の項目をできるだけ正確に読み取り、' +
    '指定したJSON形式のみで出力してください（説明文やコードフェンスは不要です）。読めない・自信がない場合は' +
    '値に"?"を含めてください。数字・型番は特に注意して1文字ずつ確認してください。\n\n' +
    '{\n' +
    '  "zuban": "図番",\n' +
    '  "hinmei": "品名",\n' +
    '  "frontChuck": "正面チャック(GB)径",\n' +
    '  "backChuck": "背面チャック径",\n' +
    '  "cycleTime": "サイクルタイム",\n' +
    '  "toolStorage": "専用ツール保管（有/無など）",\n' +
    '  "forwardPosition": "前進端位置",\n' +
    '  "positions": [ {"column": "front", "tNumber": "Tナンバー", "description": "工具の説明"} ],\n' +
    '  "memo": "上部の変更履歴メモを可能な限りそのまま書き起こしたもの"\n' +
    '}\n\n' +
    'positionsのcolumnは、表の3列（正面チャック径の列はfront、背面チャック径の列はback、右端のサイクルタイムの列はcycle）に' +
    '対応させ、上から並んでいる順番のまま出力してください。';

  var resp;
  try {
    resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{
            parts: [
              { text: prompt },
              { inline_data: { mime_type: payload.mimeType || 'image/jpeg', data: payload.imageBase64 } }
            ]
          }]
        }),
        muteHttpExceptions: true
      }
    );
  } catch (e) {
    return { error: 'Gemini APIへの通信に失敗しました: ' + e };
  }

  var code = resp.getResponseCode();
  if (code !== 200) {
    return { error: 'Gemini APIエラー(' + code + '): ' + resp.getContentText().substring(0, 300) };
  }

  var json;
  try {
    json = JSON.parse(resp.getContentText());
  } catch (e) {
    return { error: 'Gemini応答の解析に失敗しました' };
  }

  var text = json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;
  if (!text) return { error: 'Gemini応答が空でした（画像が不鮮明・読み取り不能の可能性）' };

  var jsonText = text.replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/```\s*$/, '').trim();
  var parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    return { error: '読み取り結果の解析に失敗しました: ' + e };
  }

  return { ok: true, data: parsed };
}

/**
 * 図番インデックスを最新化する（2026-08-30追加）。時間主導トリガーから1日1回呼ばれる想定。
 * 全件スキャンではなく、既にインデックス済みの図番だけを再チェックする
 * （新規に検索されたことのない図番はここでは追加しない。初回スキャン時に自動でインデックスされるため）。
 * 例えば「1回目の検査時点では検査記録フォルダ・品質情報が無かったが、後から作成された」ケースも、
 * この再チェックで拾われて図番インデックスに反映される。
 *
 * GASの実行時間制限（6分）に収まるよう、時間切れなら1分後に自身を再実行して続きから再開する
 * （seedZubanIndexと同じ仕組み）。図番インデックスの件数が多い場合、1日の実行だけでは
 * 全件を再チェックしきれないことがあるが、続きは翌日の実行が同じ再開ロジックで拾う。
 * 完全に再チェックし終えたら、続けてchainSeedZubanIndex_経由でseedZubanIndexを実行する
 * （新規図番の索引化。2つを同時に走らせず、この順で連結することで所要時間を短くしている）。
 * setupDailyIndexTriggerを1回実行してトリガー登録すること。
 */
function refreshZubanIndex() {
  var startTime = Date.now();
  var maxRunMs = 5 * 60 * 1000;
  var props = PropertiesService.getScriptProperties();

  // 前回、時間切れで自分自身を再実行するために予約した一時トリガー（今動いているのでもう不要）だけを消す。
  // 毎日3時台に実行される本来のトリガーはハンドラー名が同じでも別トリガーなので、これでは消えない。
  deleteTriggerById_(props.getProperty('refreshZubanIndexContinuationTriggerId'));
  props.deleteProperty('refreshZubanIndexContinuationTriggerId');

  var zubanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
  if (!zubanSheet || zubanSheet.getLastRow() < 2) {
    Logger.log('図番インデックスは空のため、更新対象なし');
    chainSeedZubanIndex_();
    return;
  }
  var header = zubanSheet.getRange(1, 1, 1, zubanSheet.getLastColumn()).getValues()[0];
  var zubanCol = header.indexOf('図番');
  var rows = zubanSheet.getDataRange().getValues();

  var cursor = Number(props.getProperty('refreshZubanIndexCursor') || '1');
  if (cursor < 1 || cursor >= rows.length) cursor = 1; // 前回で最後まで終わっていたら最初の行から

  var refreshed = 0;
  var i;
  for (i = cursor; i < rows.length; i++) {
    if (Date.now() - startTime > maxRunMs) break;
    var zuban = rows[i][zubanCol];
    if (!zuban) continue;
    try {
      refreshOneZuban_(zuban);
      refreshed++;
    } catch (e) {
      // Drive APIの一時的なエラー等で1件失敗しても、処理全体を止めず次に進む（2026-08-30）。
      // 失敗した図番はここではスキップされるが、次回の全体再実行時に改めて対象になる。
      Logger.log('図番「' + zuban + '」の更新中にエラー（スキップして続行）: ' + e);
    }
  }

  if (i < rows.length) {
    props.setProperty('refreshZubanIndexCursor', String(i));
    var t = ScriptApp.newTrigger('refreshZubanIndex').timeBased().after(60 * 1000).create();
    props.setProperty('refreshZubanIndexContinuationTriggerId', t.getUniqueId());
    Logger.log('実行時間の上限のため中断（今回' + refreshed + '件更新、' + i + '/' + (rows.length - 1) + '。1分後に自動で続きを実行します）');
    return;
  }

  props.deleteProperty('refreshZubanIndexCursor');
  Logger.log('図番インデックス更新完了: ' + refreshed + '件（全' + (rows.length - 1) + '件）');

  // 続けてseedZubanIndex（新規図番の索引化）を実行する。固定時刻の別トリガーにせず、
  // refreshZubanIndexの完了直後につなげることで、2つが同時に走ってDrive/I-PROへの
  // 負荷が競合するのを避け、全体の所要時間を短くする（ユーザー指摘、2026-08-30）。
  chainSeedZubanIndex_();
}

/** refreshZubanIndexの完了直後にseedZubanIndexを実行する一時トリガーを予約する。 */
function chainSeedZubanIndex_() {
  var props = PropertiesService.getScriptProperties();
  deleteTriggerById_(props.getProperty('seedZubanIndexContinuationTriggerId'));
  var t = ScriptApp.newTrigger('seedZubanIndex').timeBased().after(30 * 1000).create();
  props.setProperty('seedZubanIndexContinuationTriggerId', t.getUniqueId());
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
  // refreshZubanIndexだけを固定時刻（深夜0時台）で毎日実行する。完了したらrefreshZubanIndex自身が
  // 続けてseedZubanIndexを実行する（chainSeedZubanIndex_）ため、seedZubanIndex用の固定時刻トリガーは
  // 別途登録しない（2つを同時に走らせて所要時間が伸びるのを防ぐため、2026-08-30変更）。
  // 開始を深夜0時に早めたのも、勤務開始時間までに確実に終わらせるため（ユーザー指摘、2026-08-30）。
  deleteTriggersByHandler_('refreshZubanIndex');
  deleteTriggersByHandler_('seedZubanIndex');
  ScriptApp.newTrigger('refreshZubanIndex')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  Logger.log('毎日0時台にrefreshZubanIndexを実行するトリガーを登録しました（完了後、自動でseedZubanIndexが続けて実行されます）');
}

/**
 * 図番インデックスを事前に一括で埋める（2026-08-30追加、ユーザー提案）。
 * 「進捗状況照会」に載っている図番（＝I-PROに実在する図番の全量）のうち、まだ図番インデックスに
 * 無いものを順番に検索し、検査記録フォルダ・品質情報・改善計画書が実在するかを調べて記録する。
 * これにより、初めてスキャンされる図番でも即座に結果を返せるようになる。
 *
 * GASの実行時間制限（6分）があるため、件数が多いと1回では終わらないことがある。
 * その場合は1分後に自分自身を再実行するトリガーを自動登録するので、GASエディタで1回実行すれば、
 * あとは放置で完了する。完了すると実行ログに「事前作成が完了しました」と出る。
 * 通常はrefreshZubanIndexの完了直後にchainSeedZubanIndex_経由で自動的に呼ばれる
 * （setupDailyIndexTriggerで登録される固定時刻トリガーはrefreshZubanIndexのみ）。
 */
function seedZubanIndex() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('seedZubanIndexPaused') === 'true') {
    Logger.log('seedZubanIndexPaused中のため何もせず終了します（resumeSeedZubanIndex()で再開）');
    return;
  }

  var startTime = Date.now();
  var maxRunMs = 5 * 60 * 1000;

  // 前回、時間切れで自分自身を再実行するために予約した一時トリガー（今動いているのでもう不要）だけを消す。
  // 毎日4時台に実行される本来のトリガーはハンドラー名が同じでも別トリガーなので、これでは消えない。
  deleteTriggerById_(props.getProperty('seedZubanIndexContinuationTriggerId'));
  props.deleteProperty('seedZubanIndexContinuationTriggerId');

  var allZubans = listAllKnownZubans_();
  var already = loadIndexedZubanSet_(); // 図番インデックスを1回だけ読み込んでSetにする（毎回全件チェックしても軽い）

  var processed = 0;
  var skipped = 0;
  var remaining = 0;
  var errored = 0;
  for (var i = 0; i < allZubans.length; i++) {
    var zuban = allZubans[i];
    var key = numericZubanKey_(zuban);
    if (already[key]) { skipped++; continue; }
    if (Date.now() - startTime > maxRunMs) { remaining++; continue; } // 時間切れ後は件数だけ数える
    try {
      refreshOneZuban_(zuban);
      already[key] = true;
      processed++;
    } catch (e) {
      // Drive APIの一時的なエラー（"Empty response"等）で処理全体が止まり、次の自動継続の
      // 予約すらされなくなる不具合があったため（2026-08-30）、1件のエラーで全体を巻き込まない
      // ようにする。already[key]は立てないので、次回の実行で自動的に再試行される。
      errored++;
      Logger.log('図番「' + zuban + '」の処理中にエラー（次回再試行します）: ' + e);
    }
  }

  if (remaining > 0) {
    var t = ScriptApp.newTrigger('seedZubanIndex').timeBased().after(60 * 1000).create();
    props.setProperty('seedZubanIndexContinuationTriggerId', t.getUniqueId());
    Logger.log('実行時間の上限のため中断（今回' + processed + '件処理、エラー' + errored + '件、残り約' + remaining + '件。1分後に自動で続きを実行します）');
    return;
  }

  if (errored > 0) {
    var t2 = ScriptApp.newTrigger('seedZubanIndex').timeBased().after(60 * 1000).create();
    props.setProperty('seedZubanIndexContinuationTriggerId', t2.getUniqueId());
    Logger.log('今回新規' + processed + '件、エラー' + errored + '件（1分後に再試行します）、既存スキップ' + skipped + '件、全' + allZubans.length + '件');
    return;
  }

  Logger.log('図番インデックスの事前作成が完了しました（今回新規' + processed + '件、既存スキップ' + skipped + '件、全' + allZubans.length + '件）');
}

/**
 * seedZubanIndexを一時停止する（2026-08-30、重複行の急増に対する応急処置）。
 * 予約済みの継続トリガーを削除し、以後seedZubanIndexが呼ばれても即座に終了するフラグを立てる。
 * GASエディタで手動実行。再開はresumeSeedZubanIndex()。
 */
function pauseSeedZubanIndex() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('seedZubanIndexPaused', 'true');
  deleteTriggerById_(props.getProperty('seedZubanIndexContinuationTriggerId'));
  props.deleteProperty('seedZubanIndexContinuationTriggerId');
  deleteTriggersByHandler_('seedZubanIndex'); // seedZubanIndex専用の固定時刻トリガーは存在しないため、これで全て消して問題ない
  Logger.log('seedZubanIndexを一時停止しました。再開はresumeSeedZubanIndex()を実行してください。');
}

/** pauseSeedZubanIndex()で立てた一時停止フラグを解除する。GASエディタで手動実行。 */
function resumeSeedZubanIndex() {
  PropertiesService.getScriptProperties().deleteProperty('seedZubanIndexPaused');
  Logger.log('一時停止を解除しました。次はrefreshZubanIndexの完了時か、seedZubanIndexを手動実行した時に動きます。');
}

/**
 * 図番インデックスの重複行を整理する（2026-08-30、進捗状況照会の10分更新による0落ち混在で
 * 発生した大量重複行への対処）。numericZubanKey_で数字だけの図番の0落ちを吸収した上で
 * グループ化し、同じ図番のグループは一番更新日時が新しい行だけを残して他を削除する。
 * GASエディタで手動実行。完了後にresumeSeedZubanIndex()で再開すること。
 */
function dedupeZubanIndex() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = header.indexOf('図番');
  var updatedCol = header.indexOf('更新日時');

  var groups = {}; // numericZubanKey_ -> values配列上のindexのリスト
  for (var i = 1; i < values.length; i++) {
    var raw = values[i][col];
    if (!raw) continue;
    var key = numericZubanKey_(raw);
    if (!groups[key]) groups[key] = [];
    groups[key].push(i);
  }

  var rowsToDelete = [];
  Object.keys(groups).forEach(function (key) {
    var idxList = groups[key];
    if (idxList.length === 1) return;
    idxList.sort(function (a, b) {
      var da = values[a][updatedCol] ? new Date(values[a][updatedCol]).getTime() : 0;
      var db = values[b][updatedCol] ? new Date(values[b][updatedCol]).getTime() : 0;
      return db - da; // 新しい順。先頭（idxList[0]）を残し、残りを削除する
    });
    idxList.slice(1).forEach(function (dupIdx) { rowsToDelete.push(dupIdx + 1); });
  });

  rowsToDelete.sort(function (a, b) { return b - a; }); // 行番号が大きい順に削除（後続行のズレ防止）
  rowsToDelete.forEach(function (r) { sheet.deleteRow(r); });

  Logger.log('図番インデックスの重複整理が完了しました。削除' + rowsToDelete.length + '行（統合後の図番数: ' + Object.keys(groups).length + '）');
}

/**
 * 図番インデックスに既にある図番を、Setとして1回で読み込む（seedZubanIndexの高速化用）。
 * numericZubanKey_で正規化したキーを使うことで、数字だけの図番の0落ち表記ゆれがあっても
 * 「既にインデックス済み」と正しく判定できるようにしている（2026-08-30）。
 */
function loadIndexedZubanSet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ZUBAN_INDEX);
  var set = {};
  if (!sheet || sheet.getLastRow() < 2) return set;
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var col = header.indexOf('図番');
  for (var i = 1; i < values.length; i++) {
    if (values[i][col]) set[numericZubanKey_(values[i][col])] = true;
  }
  return set;
}

/** ハンドラー名が一致するトリガーを全部削除する。setupDailyIndexTriggerでの再登録（意図的な全消し）専用。 */
function deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
}

/**
 * 特定の1つのトリガーだけをユニークIDで削除する。seedZubanIndex/refreshZubanIndexの自己再実行用の
 * 一時トリガーを消す際に使う（deleteTriggersByHandler_だとハンドラー名が同じ毎日実行トリガーまで
 * 巻き込んで消してしまうため、こちらを使う必要がある）。
 */
function deleteTriggerById_(id) {
  if (!id) return;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getUniqueId() === id) ScriptApp.deleteTrigger(t);
  });
}

/**
 * 「進捗状況照会」の全タブから「品番(図番)」列の値を重複なく集める（＝実在する図番の一覧）。
 * numericZubanKey_でのdedupにより、数字だけの図番が0落ち表記で複数回現れても1件に収まる。
 */
function listAllKnownZubans_() {
  var ss = SpreadsheetApp.openById(IPRO_PROGRESS_SPREADSHEET_ID);
  var seen = {};
  var out = [];
  ss.getSheets().forEach(function (sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var col = header.indexOf('品番(図番)');
    if (col === -1) col = header.indexOf('品番(図番）');
    if (col === -1) return;
    var values = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues();
    values.forEach(function (row) {
      var z = stripZubanPrefix_(row[0]);
      var key = numericZubanKey_(z);
      if (z && !seen[key]) { seen[key] = true; out.push(z); }
    });
  });
  return out;
}

/**
 * 既存「{図番} 品質情報」スプレッドシートの品質情報記録ログへの一括移行（2026-08-30）。
 *
 * 対象ファイルの構造（実データ十数件で確認済み）：
 *   1シート目に「得意先／図番／品名」のラベル付きヘッダーと、その下に自由記述の本文が続く。
 *   本文は1件のみのシートもあれば、複数の日付の記録が同じ領域に蓄積されているシートもあり、
 *   日付書式も統一されていない（"2026/08/27" "2026.8.21" "'25.04.04" など）。
 *   → 日付らしき文字列で始まるセルが見つかった場合のみ、そこを区切りとして複数エントリーに
 *     分割する（QUALITY_INFO_DATE_RE_）。見つからなければ全体を1件として扱う（ユーザー合意、2026-08-30）。
 *   写真はセル上配置の画像（OverGridImage）としてtypically 1〜4枚埋め込まれているが、
 *   Apps ScriptのOverGridImageにはgetBlob()が無くバイトを直接取得できないため、
 *   スプレッドシートをxlsx形式でエクスポートし、Utilities.unzip()でxl/media/配下から
 *   画像ファイルそのものを取り出す方式を採る（実データで動作確認済み）。
 *   日付エントリーごとの写真の対応付けはできないため、抽出した写真は全て最後（最新）の
 *   エントリーにのみ紐づける（ユーザー合意）。
 *
 * 実行順序：
 *   1. addMigrationSourceColumn() を1回手動実行（「移行元ファイルID」列を追加）
 *   2. previewMigrateQualityInfo() を実行 → 「移行プレビュー」タブで図番・ランク・本文・写真枚数を確認
 *   3. 問題なければ runMigrateQualityInfo() を実行 → 品質情報記録ログへ実際に登録（共有フラグtrue）
 * いずれもGASの6分制限に収まらない場合は1分後に自動で自分自身を再実行し、続きから再開する
 * （refreshZubanIndex/seedZubanIndexと同じ自己再実行の仕組み）。runMigrateQualityInfoは
 * 「移行元ファイルID」列で処理済みファイルを判定するため、再実行しても重複登録されない。
 */
var QUALITY_INFO_LABEL_TOKENS_ = ['品質情報', '外観ランク', '得意先', '図番', '品名'];
var QUALITY_INFO_RANK_VALUES_ = ['A', 'B', 'C', 'D', 'E'];
var QUALITY_INFO_DATE_RE_ = /^[\s　・\-◆●○□■]*'?(\d{2,4})[.\/](\d{1,2})[.\/](\d{1,2})/;

/** SHEET_QUALITY_LOGに「移行元ファイルID」列を追加する（既存シート用、初回のみ手動実行）。 */
function addMigrationSourceColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUALITY_LOG);
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('移行元ファイルID') !== -1) {
    Logger.log('既に追加済みです');
    return;
  }
  sheet.getRange(1, lastCol + 1).setValue('移行元ファイルID');
  Logger.log('「移行元ファイルID」列を追加しました');
}

/** SHEET_SHIPPING_SPECに「キーエンス測定」列を追加する（既存シート用、初回のみ手動実行）。 */
function addKeyenceColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHIPPING_SPEC);
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('キーエンス測定') !== -1) {
    Logger.log('既に追加済みです');
    return;
  }
  var ngCol = header.indexOf('NG限度見本');
  var insertAt = ngCol !== -1 ? ngCol + 2 : lastCol + 1; // NG限度見本の直後（1始まり列番号）
  sheet.insertColumnAfter(insertAt - 1);
  sheet.getRange(1, insertAt).setValue('キーエンス測定');
  Logger.log('「キーエンス測定」列を追加しました');
}

/** SHEET_SHIPPING_SPECに「キーエンスプログラム名」列を追加する（既存シート用、初回のみ手動実行）。 */
function addKeyenceProgramColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHIPPING_SPEC);
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('キーエンスプログラム名') !== -1) {
    Logger.log('既に追加済みです');
    return;
  }
  var keyenceCol = header.indexOf('キーエンス測定');
  var insertAt = keyenceCol !== -1 ? keyenceCol + 2 : lastCol + 1; // キーエンス測定の直後（1始まり列番号）
  sheet.insertColumnAfter(insertAt - 1);
  sheet.getRange(1, insertAt).setValue('キーエンスプログラム名');
  Logger.log('「キーエンスプログラム名」列を追加しました');
}

function ensureMigrationPreviewSheet_() {
  return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), '移行プレビュー', [
    '元ファイル名', '元URL', '図番', '外観ランク', '分割件数', '内容プレビュー', '写真枚数', '状態'
  ]);
}

/** Drive全体（共有ドライブ含む）から「{図番} 品質情報」候補ファイルを全件（ページング対応）取得する。 */
function listQualityInfoCandidateFiles_() {
  var out = [];
  var pageToken = null;
  do {
    var res = Drive.Files.list({
      q: "name contains '品質情報' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      corpora: 'allDrives',
      pageSize: 1000,
      pageToken: pageToken,
      fields: 'nextPageToken, files(id, name, webViewLink, modifiedTime)'
    });
    (res.files || []).forEach(function (f) { out.push(f); });
    pageToken = res.nextPageToken || null;
  } while (pageToken);
  return out;
}

function extractZubanFromQualityInfoTitle_(title) {
  var t = String(title);
  if (t.trim() === '品質情報テンプレート') return null; // 未使用の雛形そのもの
  t = t.replace(/\s*品質情報テンプレート\s*$/, '');
  t = t.replace(/\s*品質情報\s*$/, '');
  t = t.trim();
  return t || null;
}

function parseQualityInfoDate_(text) {
  var m = QUALITY_INFO_DATE_RE_.exec(String(text).trim());
  if (!m) return null;
  var y = Number(m[1]);
  if (y < 100) y += 2000;
  var date = new Date(y, Number(m[2]) - 1, Number(m[3]));
  return isNaN(date.getTime()) ? null : date;
}

/**
 * 「{図番} 品質情報」スプレッドシート1件を解析する。
 * ラベル・見出し・図番/品名/外観ランクの値そのものを除いた残りのセル文字列を本文候補として集め、
 * 日付らしきセルを区切りにエントリーへ分割する（見つからなければ1件にまとめる）。
 * 戻り値: { error } または { zuban, rank, hinmei, entries: [{ timestamp, content }] }
 *         entries が空配列＝未使用の空テンプレート（本文なし）。
 */
function normalizeQualityInfoToken_(s) {
  return String(s).replace(/[\s　]+/g, '');
}

/**
 * セルの値を比較・格納用の文字列に変換する。日付型セル（Sheets側で日付として認識された値）は
 * Dateオブジェクトで返ってくるため、そのままString()するとDate.toString()の英語表記
 * （"Mon Jul 27 2026 16:00:00 GMT+0900..."）になってしまい、日付らしき行の検出に失敗する。
 * 必ずyyyy/MM/dd形式の文字列に揃える。
 */
function qualityInfoCellToText_(raw) {
  if (raw instanceof Date) return Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy/MM/dd');
  return String(raw).trim();
}

function parseQualityInfoSpreadsheet_(file, lastUpdatedDate) {
  var zuban = extractZubanFromQualityInfoTitle_(file.name);
  if (!zuban) return { error: '図番をファイル名から特定できません' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(file.id);
  } catch (e) {
    return { error: '開けませんでした: ' + e };
  }
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { zuban: zuban, rank: '', hinmei: '', entries: [] };
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var excludeTokens = {};
  QUALITY_INFO_LABEL_TOKENS_.forEach(function (t) { excludeTokens[normalizeQualityInfoToken_(t)] = true; });
  excludeTokens[normalizeQualityInfoToken_(zuban)] = true;

  var rank = '';
  var hinmei = '';
  var labelRows = {}; // 「図番」「得意先」「品名」ラベルが実際に存在する行番号（この行は丸ごと除外する）
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var raw = values[r][c];
      if (raw === '' || raw === null) continue;
      var v = qualityInfoCellToText_(raw);
      if (v === '図番' || v === '得意先' || v === '品名') labelRows[r] = true;
      if (v === '図番' && c + 1 < values[r].length) excludeTokens[normalizeQualityInfoToken_(values[r][c + 1])] = true;
      if (v === '品名' && c + 1 < values[r].length) {
        hinmei = qualityInfoCellToText_(values[r][c + 1]);
        excludeTokens[normalizeQualityInfoToken_(hinmei)] = true;
      }
      if (!rank && r < 6 && QUALITY_INFO_RANK_VALUES_.indexOf(v) !== -1) rank = v;
    }
  }
  if (rank) excludeTokens[normalizeQualityInfoToken_(rank)] = true;

  // 得意先名（会社名）は「図番」「得意先」ラベルの行にまたがってマージされているため、
  // どちらかのラベルが存在する行は、ラベル自体・品名の値以外の文字列をまとめて除外する。
  Object.keys(labelRows).forEach(function (rowIndexStr) {
    var row = values[Number(rowIndexStr)];
    row.forEach(function (x) {
      if (x === '' || x === null) return;
      var v2 = qualityInfoCellToText_(x);
      if (v2 && v2 !== '図番' && v2 !== '得意先' && v2 !== '品名' && v2 !== hinmei) {
        excludeTokens[normalizeQualityInfoToken_(v2)] = true;
      }
    });
  });

  var cells = [];
  for (var r3 = 0; r3 < values.length; r3++) {
    for (var c3 = 0; c3 < values[r3].length; c3++) {
      var raw3 = values[r3][c3];
      if (raw3 === '' || raw3 === null || raw3 === undefined) continue;
      var v3 = qualityInfoCellToText_(raw3);
      if (!v3 || excludeTokens[normalizeQualityInfoToken_(v3)]) continue;
      cells.push(v3);
    }
  }

  if (cells.length === 0) return { zuban: zuban, rank: rank, hinmei: hinmei, entries: [] };

  var entries = [];
  var current = null;
  cells.forEach(function (text) {
    var d = parseQualityInfoDate_(text);
    if (d) {
      current = { timestamp: d, lines: [text] };
      entries.push(current);
    } else if (current) {
      current.lines.push(text);
    } else {
      current = { timestamp: null, lines: [text] };
      entries.push(current);
    }
  });
  entries.forEach(function (e) {
    e.content = e.lines.join('\n');
    if (!e.timestamp) e.timestamp = lastUpdatedDate;
    delete e.lines;
  });

  return { zuban: zuban, rank: rank, hinmei: hinmei, entries: entries };
}

/** xlsxエクスポート+unzipで、埋め込み画像のBlob配列を取得する。 */
function extractQualityInfoPhotoBlobs_(fileId) {
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  var xlsxBlob = resp.getBlob().setContentType('application/zip');
  return Utilities.unzip(xlsxBlob).filter(function (e) { return e.getName().indexOf('xl/media/') === 0; });
}

/** 抽出した写真Blobを、通常のアップロード先（図番フォルダ配下の「写真」サブフォルダ）へ保存しURLを返す。 */
function uploadMigratedPhotos_(zuban, blobs) {
  if (blobs.length === 0) return [];
  var zubanFolder = findZubanFolder_(zuban) || createZubanFolderIfCompanyKnown_(zuban);
  if (!zubanFolder) return [];
  var photoFiles = driveFilesList_(
    "name = '写真' and mimeType = 'application/vnd.google-apps.folder' and trashed = false and '" + zubanFolder.getId() + "' in parents"
  );
  var photoFolder = photoFiles.length > 0 ? DriveApp.getFolderById(photoFiles[0].id) : zubanFolder.createFolder('写真');
  return blobs.map(function (blob) { return photoFolder.createFile(blob).getUrl(); });
}

/**
 * 移行対象の一覧をドライラン解析し、「移行プレビュー」タブに結果を書き出す（品質情報記録ログには一切書き込まない）。
 * 写真は枚数のみ確認し、この時点ではアップロードしない。GASエディタで手動実行。
 */
function previewMigrateQualityInfo() {
  var startTime = Date.now();
  var maxRunMs = 5 * 60 * 1000;
  var props = PropertiesService.getScriptProperties();

  deleteTriggerById_(props.getProperty('migratePreviewContinuationTriggerId'));
  props.deleteProperty('migratePreviewContinuationTriggerId');

  var sheet = ensureMigrationPreviewSheet_();
  var listJson = props.getProperty('migratePreviewFileList');
  var files;
  if (listJson) {
    files = JSON.parse(listJson);
  } else {
    files = listQualityInfoCandidateFiles_();
    props.setProperty('migratePreviewFileList', JSON.stringify(files));
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).clearContent();
    Logger.log('対象候補: ' + files.length + '件');
  }

  var cursor = Number(props.getProperty('migratePreviewCursor') || '0');
  var rowsToAppend = [];
  var i;
  for (i = cursor; i < files.length; i++) {
    if (Date.now() - startTime > maxRunMs) break;
    var f = files[i];
    if (f.name.trim() === '品質情報テンプレート') continue;

    var parsed = parseQualityInfoSpreadsheet_(f, new Date(f.modifiedTime));
    if (parsed.error) {
      rowsToAppend.push([f.name, f.webViewLink, '', '', 0, '', 0, 'スキップ: ' + parsed.error]);
      continue;
    }
    if (parsed.entries.length === 0) {
      rowsToAppend.push([f.name, f.webViewLink, parsed.zuban, parsed.rank, 0, '', 0, 'スキップ: 本文なし（未使用テンプレート）']);
      continue;
    }
    var photoCount = 0;
    try {
      photoCount = extractQualityInfoPhotoBlobs_(f.id).length;
    } catch (e) {
      photoCount = -1;
    }
    var summary = parsed.entries.map(function (e) {
      var d = e.timestamp ? Utilities.formatDate(e.timestamp, 'Asia/Tokyo', 'yyyy/MM/dd') : '(日付不明)';
      return d + ': ' + e.content.replace(/\n/g, ' ').substring(0, 60);
    }).join(' / ');
    rowsToAppend.push([f.name, f.webViewLink, parsed.zuban, parsed.rank, parsed.entries.length, summary, photoCount, 'OK']);
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
  }

  if (i < files.length) {
    props.setProperty('migratePreviewCursor', String(i));
    var t = ScriptApp.newTrigger('previewMigrateQualityInfo').timeBased().after(60 * 1000).create();
    props.setProperty('migratePreviewContinuationTriggerId', t.getUniqueId());
    Logger.log('実行時間の上限のため中断（' + i + '/' + files.length + '。1分後に自動で続きを実行します）');
    return;
  }

  props.deleteProperty('migratePreviewCursor');
  props.deleteProperty('migratePreviewFileList');
  Logger.log('プレビュー作成完了。「移行プレビュー」タブを確認してください（全' + files.length + '件）');
}

/**
 * previewMigrateQualityInfoで内容を確認した後、実際に品質情報記録ログへ登録する（共有フラグtrue）。
 * 「移行元ファイルID」列で処理済みファイルを判定するため、途中で中断されても再実行すれば
 * 未処理分だけ続きから処理する（addMigrationSourceColumn()を先に1回実行しておくこと）。
 */
function runMigrateQualityInfo() {
  var startTime = Date.now();
  var maxRunMs = 5 * 60 * 1000;
  var props = PropertiesService.getScriptProperties();

  deleteTriggerById_(props.getProperty('migrateRunContinuationTriggerId'));
  props.deleteProperty('migrateRunContinuationTriggerId');

  var qlSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUALITY_LOG);
  var header = qlSheet.getRange(1, 1, 1, qlSheet.getLastColumn()).getValues()[0];
  var sourceCol = header.indexOf('移行元ファイルID');
  if (sourceCol === -1) {
    Logger.log('「移行元ファイルID」列がありません。先にaddMigrationSourceColumn()を実行してください。');
    return;
  }

  var listJson = props.getProperty('migrateRunFileList');
  var files;
  if (listJson) {
    files = JSON.parse(listJson);
  } else {
    files = listQualityInfoCandidateFiles_();
    props.setProperty('migrateRunFileList', JSON.stringify(files));
    Logger.log('対象候補: ' + files.length + '件');
  }

  var alreadyDone = {};
  if (qlSheet.getLastRow() >= 2) {
    qlSheet.getRange(2, sourceCol + 1, qlSheet.getLastRow() - 1, 1).getValues().forEach(function (row) {
      if (row[0]) alreadyDone[row[0]] = true;
    });
  }

  var cursor = Number(props.getProperty('migrateRunCursor') || '0');
  var rowsToAppend = [];
  var inserted = 0, skipped = 0;
  var i;
  for (i = cursor; i < files.length; i++) {
    if (Date.now() - startTime > maxRunMs) break;
    var f = files[i];
    if (f.name.trim() === '品質情報テンプレート') continue;
    if (alreadyDone[f.id]) { skipped++; continue; }

    var parsed = parseQualityInfoSpreadsheet_(f, new Date(f.modifiedTime));
    if (parsed.error || parsed.entries.length === 0) { skipped++; continue; }

    var photoUrls = [];
    try {
      photoUrls = uploadMigratedPhotos_(parsed.zuban, extractQualityInfoPhotoBlobs_(f.id));
    } catch (e) {
      Logger.log('写真取得エラー(' + f.name + '): ' + e);
    }

    parsed.entries.forEach(function (entry, idx) {
      var isLast = idx === parsed.entries.length - 1;
      rowsToAppend.push([
        Utilities.getUuid(), entry.timestamp, parsed.zuban, '',
        '', '(移行データ)', parsed.rank, entry.content,
        isLast ? photoUrls.join('\n') : '', true, f.id
      ]);
    });
    invalidateZubanCache_(parsed.zuban);
    inserted++;
  }

  if (rowsToAppend.length > 0) {
    qlSheet.getRange(qlSheet.getLastRow() + 1, 1, rowsToAppend.length, 11).setValues(rowsToAppend);
  }

  if (i < files.length) {
    props.setProperty('migrateRunCursor', String(i));
    var t = ScriptApp.newTrigger('runMigrateQualityInfo').timeBased().after(60 * 1000).create();
    props.setProperty('migrateRunContinuationTriggerId', t.getUniqueId());
    Logger.log('実行時間の上限のため中断（' + i + '/' + files.length + '件処理、今回' + inserted + '件登録。1分後に自動で続きを実行します）');
    return;
  }

  props.deleteProperty('migrateRunCursor');
  props.deleteProperty('migrateRunFileList');
  Logger.log('移行完了。新規登録' + inserted + 'ファイル分、スキップ' + skipped + '件（全' + files.length + '件）');
}
