/************************************************************
 *  顧客管理システム  Code.gs
 *  Phase 1：ログイン / 新規登録 / 顧客一覧
 *
 *  対象スプレッドシート：コピー～ 🟡CW運用管理（テスト用コピー）
 *  ※ このファイルは「完全版」です。Apps Script の Code.gs を
 *     まるごとこの内容に貼り替えてください。
 *
 *  === 実行順（初回） ===
 *   1) rebuildMasters   … 管理シートをまっさらに作り直し（初回のみ推奨）
 *   2) setup            … 見出し整備・サンプル投入・M列入力規則
 *   3) diagnose         … うまくいったか確認（ログを見る）
 ************************************************************/

/* ========== 基本設定 ========== */
const SPREADSHEET_ID     = '1HFrJ-7Ju1zDzm2pNO9Dw-DOjioB68gXXTnyXSfEmjCI';
const CUSTOMER_SHEET      = '顧客リスト';
const CUSTOMER_START_ROW  = 3;          // 1〜2行目は見出し。データは3行目から
const TOKEN_DAYS          = 30;         // 自動ログインの有効日数

// 新規登録の書き込み開始位置（2026-09-12 ユーザー指定：この行から下の
// 「名前もIDも空の行」に詰めて書き込む）。顧客リストは過去データが
// 途中で途切れていたり、離れた場所に別の入力エリアがあったりするため、
// getLastRow() 等で自動判定せず、ここに実際の行番号を指定する方式にした。
// ズレてきたら、この数字を書き換えて setup 不要・そのまま再デプロイでOK。
const APPEND_SEARCH_START_ROW = 7716;

/* ========== 顧客リストの列番号（A列 = 1） ========== */
const COL = {
  NO: 1, 外部申請: 2, Line誘導: 3, 名前: 4, ID: 5,
  年齢: 8, 地域: 9, 応募金額: 10, 応募日: 11, 案件名: 12,
  担当者: 13, 運用アカウント: 14, 媒体: 15, LINE追加: 16,
  ステータス: 19, 備考: 21, 重複: 22, 重複元NO: 23, 次工程: 24
};
const COL_LAST = 24;  // X列まで読み書きする

/* ========== 選択肢（スプシの入力規則と完全一致・2026-09-06 実確認済み） ========== */
// 案件名（L列）：＋ は全角プラス（U+FF0B）。半角 + はスプシでエラーになる
const OPT_PROJECT = ['投稿作成＋SNS運用', 'アカウント作成＋SNS運用'];
// 媒体（O列）
const OPT_MEDIA   = ['CW', 'ランサーズ', 'シュフティ'];
// ステータス（S列）：スラッシュは半角 / 、「ﾌﾞﾛｯｸ」は半角カナ
const OPT_STATUS  = ['垢バン/否認', '公式ライン追加', '面談予約', '面談済み', '辞退/ﾌﾞﾛｯｸ', '対応不可'];
// 「未対応リスト」で対応完了とみなして非表示にするステータス（2026-09-13 追加：対応不可／垢バン・否認も完了扱い）
const MINE_DONE_STATUS = ['面談済み', '辞退/ﾌﾞﾛｯｸ', '対応不可', '垢バン/否認'];

/* ========== 管理シートの見出し定義 ========== */
const MGMT_SHEETS = {
  'メンバーマスタ':          ['userId', 'passwordHash', '氏名', '表示担当者名', '役割', 'チーム', '権限', '有効'],
  'ユーザーアカウントマスタ': ['userId', 'アカウント名', '有効'],
  'ログイントークン':        ['tokenHash', 'userId', '最終アクセス日時', '有効期限', 'userAgent', '有効'],
  'システム実績':            ['日付', '顧客NO', '担当者', 'チーム', '種別', '数値', 'メモ', '更新者', '更新日時'],
  'KPI目標':                ['対象月', 'チーム', 'ユーザー', '役割', '指標', '月次目標', '日次目標']
};

/* ========== サンプルデータ（パスワードは全員 1234） ========== */
const SAMPLE_MEMBERS = [
  ['ajima',  '1234', '安嶋',   '安嶋',   '採用者', '安嶋G',   'admin'],
  ['sasaki', '1234', '佐々木', '佐々木', '採用者', '佐々木G', 'admin'],
  ['wakagi', '1234', '若木',   '若木',   '採用者', '若木G',   'member'],
  ['yamane', '1234', '山根',   '山根',   '運用者', '安嶋G',   'member'],
  ['tanaka', '1234', '田中',   '田中',   '運用者', '佐々木G', 'member'],
  ['suzuki', '1234', '鈴木',   '鈴木',   '運用者', '若木G',   'member']
];
const SAMPLE_ACCOUNTS = [
  ['ajima', 'gdcc'], ['ajima', 'ajima_sns'],
  ['sasaki', 'sasaki_cw'], ['wakagi', 'wakagi01'],
  ['yamane', 'yamane_a'], ['yamane', 'yamane_b'],
  ['tanaka', 'tanaka_cw'], ['suzuki', 'suzuki_cw']
];

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

/**
 * fromRow から下に向かって「名前もIDも空」の行を count 個探して返す（行番号の配列）。
 * 顧客リストの途中に空白があっても、そこへ詰めて書き込める。
 */
function findBlankRows_(sh, fromRow, count) {
  const scanRows = Math.max(count * 20, 2000);   // 十分な余裕を持って読む
  const maxRow = Math.max(sh.getMaxRows(), fromRow + scanRows);
  const readCount = Math.min(scanRows, maxRow - fromRow + 1);
  const vals = sh.getRange(fromRow, COL.名前, readCount, 2).getValues();
  const blanks = [];
  for (let i = 0; i < vals.length && blanks.length < count; i++) {
    if (!String(vals[i][0] || '').trim() && !String(vals[i][1] || '').trim()) blanks.push(fromRow + i);
  }
  let extra = fromRow + readCount;
  while (blanks.length < count) { blanks.push(extra++); }   // 万一見つからない場合の保険
  return blanks;
}


/* ============================================================
 *  Web アプリの入口
 * ============================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('顧客管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ============================================================
 *  rebuildMasters()  ―― 管理シートをまっさらに作り直す（初回のみ推奨）
 *   既存の「メンバーマスタ」等に別データが入っていても強制的に整え直します。
 *   ※ システム実績 / KPI目標 は消しません（見出しだけ用意）
 * ============================================================ */
function rebuildMasters() {
  const ss = ss_();
  const log = [];
  const wipe = ['メンバーマスタ', 'ユーザーアカウントマスタ', 'ログイントークン'];

  wipe.forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); log.push('新規作成: ' + name); }
    unprotect_(sh);
    sh.clear();
    sh.clearFormats();
    const headers = MGMT_SHEETS[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    log.push('初期化: ' + name);
  });

  // メンバー投入
  const mSh = ss.getSheetByName('メンバーマスタ');
  const mRows = SAMPLE_MEMBERS.map(function (r) {
    return [r[0], hash_(r[1]), r[2], r[3], r[4], r[5], r[6], true];
  });
  mSh.getRange(2, 1, mRows.length, 8).setValues(mRows);
  log.push('メンバー ' + mRows.length + '件 投入');

  // アカウント投入
  const aSh = ss.getSheetByName('ユーザーアカウントマスタ');
  const aRows = SAMPLE_ACCOUNTS.map(function (r) { return [r[0], r[1], true]; });
  aSh.getRange(2, 1, aRows.length, 3).setValues(aRows);
  log.push('アカウント名 ' + aRows.length + '件 投入');

  // 実績 / KPI の見出しだけ用意
  ['システム実績', 'KPI目標'].forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    unprotect_(sh);
    const headers = MGMT_SHEETS[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    log.push('見出し整備: ' + name);
  });

  const msg = 'rebuildMasters 完了\n----------------\n' + log.join('\n');
  Logger.log(msg);
  return msg;
}


/* ============================================================
 *  setup()  ―― 見出し整備 / サンプル投入（不足分のみ） / M列入力規則
 * ============================================================ */
function setup() {
  const ss = ss_();
  const log = [];

  // 1) 管理シートを作成／見出しを整備
  Object.keys(MGMT_SHEETS).forEach(function (name) {
    try {
      let sh = ss.getSheetByName(name);
      const created = !sh;
      if (!sh) sh = ss.insertSheet(name);
      unprotect_(sh);
      const headers = MGMT_SHEETS[name];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      log.push((created ? '作成: ' : '確認: ') + name);
    } catch (e) {
      log.push('⚠ ' + name + ' でエラー: ' + e.message + ' → シートの保護を手動で外して再実行してください');
    }
  });

  // 2) サンプルメンバー（'ajima' が居なければ投入）
  try {
    const mSh = ss.getSheetByName('メンバーマスタ');
    const have = getMembers_().some(function (m) { return String(m.userId || '').trim(); });
    if (!have) {
      const rows = SAMPLE_MEMBERS.map(function (r) {
        return [r[0], hash_(r[1]), r[2], r[3], r[4], r[5], r[6], true];
      });
      mSh.getRange(mSh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
      log.push('サンプルメンバー ' + rows.length + '件を投入');
    } else {
      log.push('メンバーマスタに既存データあり → サンプル投入はスキップ');
    }
  } catch (e) { log.push('⚠ メンバー投入でエラー: ' + e.message); }

  // 3) サンプルの運用アカウント名
  try {
    const aSh = ss.getSheetByName('ユーザーアカウントマスタ');
    const have = getAccounts_().some(function (a) { return String(a.userId || '').trim(); });
    if (!have) {
      const rows = SAMPLE_ACCOUNTS.map(function (r) { return [r[0], r[1], true]; });
      aSh.getRange(aSh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
      log.push('サンプルアカウント名 ' + rows.length + '件を投入');
    } else {
      log.push('ユーザーアカウントマスタに既存データあり → サンプル投入はスキップ');
    }
  } catch (e) { log.push('⚠ アカウント投入でエラー: ' + e.message); }

  // 4) 顧客リスト M列（担当者）の入力規則を「表示担当者名」リストで整備
  try {
    refreshOwnerValidation_();
    log.push('顧客リスト M列の入力規則を更新');
  } catch (e) { log.push('⚠ M列入力規則の更新でエラー: ' + e.message); }

  const msg = 'setup 完了\n----------------\n' + log.join('\n');
  Logger.log(msg);
  return msg;
}


/* ============================================================
 *  diagnose()  ―― 状態を確認（ログに出力）
 * ============================================================ */
function diagnose() {
  const ss = ss_();
  const out = [];
  Object.keys(MGMT_SHEETS).forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh) { out.push('✗ ' + name + '：シートが存在しません'); return; }
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].join(' | ') : '(空)';
    const prot = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length;
    out.push('● ' + name + '：' + lastRow + '行  見出し=[' + head + ']' + (prot ? '  ★保護あり' : ''));
  });
  out.push('----------------');
  const members = getMembers_();
  out.push('メンバー件数: ' + members.length);
  members.forEach(function (m) {
    out.push('  userId=' + m.userId + ' / 有効=' + m.有効 + '(' + truthy_(m.有効) + ')'
      + ' / passwordHashが1234のハッシュと一致=' + (String(m.passwordHash) === hash_('1234')));
  });
  out.push('hash_("1234") = ' + hash_('1234'));
  const r = login('ajima', '1234');
  out.push('login("ajima","1234") → ' + JSON.stringify(r));

  const msg = 'diagnose\n----------------\n' + out.join('\n');
  Logger.log(msg);
  return msg;
}


/* ============================================================
 *  補助：シート保護を外す（自分が編集者/オーナーのとき）
 * ============================================================ */
function unprotect_(sh) {
  try {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
      try { if (p.canEdit()) p.remove(); } catch (e) {}
    });
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
      try { if (p.canEdit()) p.remove(); } catch (e) {}
    });
  } catch (e) {}
}

/** メンバーの表示担当者名リストを、顧客リスト M列の入力規則に設定する */
function refreshOwnerValidation_() {
  const ss = ss_();
  const cust = ss.getSheetByName(CUSTOMER_SHEET);
  if (!cust) throw new Error('「' + CUSTOMER_SHEET + '」シートが見つかりません');
  const names = getMembers_().map(function (m) { return String(m.表示担当者名 || '').trim(); })
                             .filter(function (v) { return v; });
  if (!names.length) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)   // 一致しなくても書き込みは通す（UIでは警告のみ）
    .build();
  const rows = cust.getMaxRows() - CUSTOMER_START_ROW + 1;
  cust.getRange(CUSTOMER_START_ROW, COL.担当者, rows, 1).setDataValidation(rule);
}

/** 管理者用：メンバーを増減したあとに手動実行すると M列の候補を更新できる */
function updateOwnerValidation() { return refreshOwnerValidation_(); }

/** 全ユーザーの運用アカウント名を、顧客リスト N列の入力規則に設定する */
function refreshAccountValidation_() {
  const ss = ss_();
  const cust = ss.getSheetByName(CUSTOMER_SHEET);
  if (!cust) throw new Error('「' + CUSTOMER_SHEET + '」シートが見つかりません');
  const names = getAccounts_()
    .filter(function (a) { return truthy_(a.有効); })
    .map(function (a) { return String(a.アカウント名 || '').trim(); })
    .filter(function (v, i, arr) { return v && arr.indexOf(v) === i; });
  if (!names.length) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names, true)
    .setAllowInvalid(true)   // 一致しなくても書き込みは通す（UIでは警告のみ）
    .build();
  const rows = cust.getMaxRows() - CUSTOMER_START_ROW + 1;
  cust.getRange(CUSTOMER_START_ROW, COL.運用アカウント, rows, 1).setDataValidation(rule);
}

/** 管理者用：アカウントを増減したあとに手動実行すると N列の候補を更新できる */
function updateAccountValidation() { return refreshAccountValidation_(); }

/* ============================================================
 *  運用アカウントの追加（「新規登録」タブの「＋追加」から）
 *   自分（ログインユーザー）名義のアカウントとして追加し、
 *   顧客リスト N列のプルダウン候補も同時に更新する。
 * ============================================================ */
function addAccount(token, name) {
  const user = requireUser_(token);
  const accountName = String(name || '').trim();
  if (!accountName) throw new Error('アカウント名を入力してください。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const aSh = ss_().getSheetByName('ユーザーアカウントマスタ');
    const already = getAccounts_().some(function (a) {
      return String(a.userId).trim() === user.userId && String(a.アカウント名).trim() === accountName && truthy_(a.有効);
    });
    if (!already) {
      aSh.appendRow([user.userId, accountName, true]);
    }
    refreshAccountValidation_();
  } finally {
    lock.releaseLock();
  }

  const myAccounts = getAccounts_()
    .filter(function (a) { return String(a.userId).trim() === user.userId && truthy_(a.有効); })
    .map(function (a) { return a.アカウント名; });
  return { ok: true, myAccounts: myAccounts };
}


/* ============================================================
 *  ハッシュ（SHA-256 → 16進文字列）
 * ============================================================ */
function hash_(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(str), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}


/* ============================================================
 *  シート読み取りユーティリティ
 * ============================================================ */
function sheetObjects_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.map(function (row, i) {
    const o = { _row: i + 2 };
    headers.forEach(function (h, c) { o[h] = row[c]; });
    return o;
  });
}
function getMembers_()  { return sheetObjects_('メンバーマスタ'); }
function getAccounts_() { return sheetObjects_('ユーザーアカウントマスタ'); }

/** 「有効」列の値をゆるく真偽判定（チェックボックス / TRUE / ○ / 1 などを許容） */
function truthy_(v) {
  if (v === true || v === 1) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === '有効' || s === '○' || s === 'yes' || s === 'y';
}


/* ============================================================
 *  認証
 * ============================================================ */
function login(userId, password) {
  userId = String(userId || '').trim();
  const m = getMembers_().find(function (x) {
    return String(x.userId).trim() === userId && truthy_(x.有効);
  });
  if (!m || String(m.passwordHash) !== hash_(password)) {
    return { ok: false, message: 'ユーザーIDまたはパスワードが違います。' };
  }
  const token = Utilities.getUuid() + Utilities.getUuid();
  saveToken_(token, userId);
  return { ok: true, token: token, user: publicUser_(m) };
}

function autoLogin(token) {
  if (!token) return { ok: false };
  const tSh = ss_().getSheetByName('ログイントークン');
  if (!tSh) return { ok: false };
  const rows = sheetObjects_('ログイントークン');
  const th = hash_(token);
  const rec = rows.find(function (r) { return r.tokenHash === th && truthy_(r.有効); });
  if (!rec) return { ok: false };

  if (new Date(rec.有効期限).getTime() < Date.now()) {
    tSh.getRange(rec._row, 6).setValue(false);   // 期限切れ → 無効化
    return { ok: false };
  }
  const m = getMembers_().find(function (x) {
    return String(x.userId).trim() === String(rec.userId).trim() && truthy_(x.有効);
  });
  if (!m) return { ok: false };

  // 最終アクセスから6時間以上なら、期限を今日から30日先へ延長
  const lastAccess = new Date(rec.最終アクセス日時).getTime();
  if (isNaN(lastAccess) || Date.now() - lastAccess > 6 * 3600 * 1000) {
    const exp = new Date(); exp.setDate(exp.getDate() + TOKEN_DAYS);
    tSh.getRange(rec._row, 3, 1, 2).setValues([[new Date(), exp]]);
  }
  return { ok: true, user: publicUser_(m) };
}

function logout(token) {
  if (!token) return;
  const tSh = ss_().getSheetByName('ログイントークン');
  const th = hash_(token);
  sheetObjects_('ログイントークン')
    .filter(function (r) { return r.tokenHash === th; })
    .forEach(function (r) { tSh.getRange(r._row, 6).setValue(false); });
}

function saveToken_(token, userId) {
  const tSh = ss_().getSheetByName('ログイントークン');
  const exp = new Date(); exp.setDate(exp.getDate() + TOKEN_DAYS);
  tSh.appendRow([hash_(token), userId, new Date(), exp, '', true]);
}

function publicUser_(m) {
  return {
    userId: m.userId, 氏名: m.氏名, 表示担当者名: m.表示担当者名,
    役割: m.役割, チーム: m.チーム, 権限: m.権限
  };
}

/** すべてのサーバー関数の先頭で呼ぶ。トークンを検証してユーザーを返す */
function requireUser_(token) {
  const r = autoLogin(token);
  if (!r.ok) throw new Error('セッションが切れました。ページを再読み込みしてログインし直してください。');
  return r.user;
}


/* ============================================================
 *  ログイン後の初期データ
 * ============================================================ */
function getBootstrap(token) {
  const user = requireUser_(token);
  const myAccounts = getAccounts_()
    .filter(function (a) { return String(a.userId).trim() === user.userId && truthy_(a.有効); })
    .map(function (a) { return a.アカウント名; });
  const owners = getMembers_()
    .map(function (m) { return String(m.表示担当者名 || '').trim(); })
    .filter(function (v, i, arr) { return v && arr.indexOf(v) === i; });
  return {
    user: user,
    myAccounts: myAccounts,
    owners: owners,
    options: { media: OPT_MEDIA, project: OPT_PROJECT, status: OPT_STATUS }
  };
}


/* ============================================================
 *  顧客一覧（全チーム表示）／未対応リスト（自分の担当のみ）
 *
 *  顧客リストは過去データが複数の場所に分かれて入っている（行の位置が
 *  新しい／古いを意味しない）ため、並び順は行の位置ではなく応募日・NO列で判定する。
 *   - mode='recent'（既定）：NOが大きい順（新しい順）に limit 件だけ返す
 *   - mode='search'         ：全行を走査し、条件（月・媒体・案件名・担当者・
 *                             ステータス・重複ありのみ等）に合うものを、
 *                             応募日が古い順（同日はNO昇順）で最大 limit 件返す
 *   - mode='mine'           ：ログインユーザーが担当した顧客のうち、重複と
 *                             完了ステータス（面談済み／辞退・ブロック）を
 *                             除いたものを、応募日が古い順で返す（未対応リスト用）
 *
 *  戻り値： { rows:[...], total, scanned, truncated, mode }
 * ============================================================ */
function getCustomers(token, opts) {
  const user = requireUser_(token);
  opts = opts || {};
  const limit = Math.min(Math.max(Number(opts.limit) || 500, 1), 2000);
  const mode = (opts.mode === 'search') ? 'search' : (opts.mode === 'mine') ? 'mine' : 'recent';
  const q = opts.query || {};

  const sh = ss_().getSheetByName(CUSTOMER_SHEET);
  const last = sh.getLastRow();
  if (last < CUSTOMER_START_ROW) return { rows: [], total: 0, scanned: 0, truncated: false, mode: mode };
  const tz = Session.getScriptTimeZone();

  const values = sh.getRange(CUSTOMER_START_ROW, 1, last - CUSTOMER_START_ROW + 1, COL_LAST).getValues();

  const word = String(q.word || '').trim().toLowerCase();
  const matched = [];
  values.forEach(function (r) {
    const name = String(r[COL.名前 - 1] || '').trim();
    const id = String(r[COL.ID - 1] || '').trim();
    if (!name && !id) return;   // 空行はスキップ
    const o = mapCustomerRow_(r, tz);

    if (mode === 'search') {
      if (word) {
        const hay = ((o.名前 || '') + ' ' + (o.ID || '') + ' ' + (o.運用アカウント || '')).toLowerCase();
        if (hay.indexOf(word) < 0) return;
      }
      if (q.media && o.媒体 !== q.media) return;
      if (q.project && o.案件名 !== q.project) return;
      if (q.owner && o.担当者 !== q.owner) return;
      if (q.status && o.ステータス !== q.status) return;
      if (q.dupOnly && String(o.重複 || '').indexOf('重複') < 0) return;
      if (q.month && String(o.応募日 || '').indexOf(q.month) !== 0) return;
    } else if (mode === 'mine') {
      // 「未対応リスト」：自分が担当した顧客のうち、重複と完了ステータス（面談済み／辞退・ブロック／対応不可／垢バン・否認）を除いたもの
      if (String(o.担当者 || '').trim() !== String(user.表示担当者名 || '').trim()) return;
      if (String(o.重複 || '').indexOf('重複') >= 0) return;
      if (MINE_DONE_STATUS.indexOf(o.ステータス) >= 0) return;
      if (word) {
        const hay = ((o.名前 || '') + ' ' + (o.ID || '')).toLowerCase();
        if (hay.indexOf(word) < 0) return;
      }
    }
    matched.push(o);
  });

  if (mode === 'recent') {
    matched.sort(function (a, b) { return (Number(b.no) || 0) - (Number(a.no) || 0); });   // NOが大きい＝新しい順
  } else {
    // 検索結果／未対応リストは、応募日が古い順（同日はNOが小さい順）
    matched.sort(function (a, b) {
      const da = a.応募日 || '', db = b.応募日 || '';
      if (da !== db) return da < db ? -1 : 1;
      return (Number(a.no) || 0) - (Number(b.no) || 0);
    });
  }

  return {
    rows: matched.slice(0, limit),
    total: matched.length,
    scanned: values.length,
    truncated: matched.length > limit,
    mode: mode
  };
}

/* ============================================================
 *  顧客一覧の「対象月」プルダウン用：応募日(K列)に存在する年月の一覧を返す
 * ============================================================ */
function getMonthOptions(token) {
  requireUser_(token);
  const sh = ss_().getSheetByName(CUSTOMER_SHEET);
  const last = sh.getLastRow();
  if (last < CUSTOMER_START_ROW) return [];
  const tz = Session.getScriptTimeZone();
  const vals = sh.getRange(CUSTOMER_START_ROW, COL.応募日, last - CUSTOMER_START_ROW + 1, 1).getValues();
  const set = {};
  vals.forEach(function (r) {
    const v = r[0];
    if (!v) return;
    const ym = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM') : String(v).slice(0, 7);
    if (/^\d{4}-\d{2}$/.test(ym)) set[ym] = true;
  });
  return Object.keys(set).sort();   // 昇順（古い月→新しい月）
}

/* ============================================================
 *  ステータス更新（「未対応リスト」から入力）
 *   自分が担当した顧客だけステータスを更新できる（他人の顧客は不可）
 * ============================================================ */
function updateCustomerStatus(token, no, status) {
  const user = requireUser_(token);
  const st = String(status || '').trim();
  if (OPT_STATUS.indexOf(st) < 0) throw new Error('ステータスが選択肢と一致しません：「' + st + '」');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const sh = ss_().getSheetByName(CUSTOMER_SHEET);
    const last = sh.getLastRow();
    if (last < CUSTOMER_START_ROW) throw new Error('対象の顧客が見つかりません。');
    const nos = sh.getRange(CUSTOMER_START_ROW, COL.NO, last - CUSTOMER_START_ROW + 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < nos.length; i++) {
      if (Number(nos[i][0]) === Number(no)) { targetRow = CUSTOMER_START_ROW + i; break; }
    }
    if (targetRow < 0) throw new Error('NO:' + no + ' の顧客が見つかりません。');

    const owner = sh.getRange(targetRow, COL.担当者).getValue();
    if (String(owner || '').trim() !== String(user.表示担当者名 || '').trim()) {
      throw new Error('自分が担当した顧客のみステータスを更新できます。');
    }
    sh.getRange(targetRow, COL.ステータス).setValue(st);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function mapCustomerRow_(r, tz) {
  return {
    no: r[COL.NO - 1],
    外部申請: r[COL.外部申請 - 1] === true,
    Line誘導: r[COL.Line誘導 - 1] === true,
    名前: r[COL.名前 - 1],
    ID: r[COL.ID - 1],
    年齢: r[COL.年齢 - 1],
    地域: r[COL.地域 - 1],
    応募金額: r[COL.応募金額 - 1],
    応募日: fmtDate_(r[COL.応募日 - 1], tz),
    案件名: r[COL.案件名 - 1],
    担当者: r[COL.担当者 - 1],
    運用アカウント: r[COL.運用アカウント - 1],
    媒体: r[COL.媒体 - 1],
    LINE追加: r[COL.LINE追加 - 1] === true,
    ステータス: r[COL.ステータス - 1],
    備考: r[COL.備考 - 1],
    重複: r[COL.重複 - 1],
    重複元NO: r[COL.重複元NO - 1],
    次工程: r[COL.次工程 - 1]
  };
}

function fmtDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === 0 || v ? String(v) : '';
}

/* ライブ重複チェック（複数ID版）：{ id: [該当NO...] } を返す */
function checkDuplicates(token, ids) {
  requireUser_(token);
  const want = {};
  (ids || []).forEach(function (x) {
    const v = String(x || '').trim();
    if (v) want[v] = [];
  });
  if (!Object.keys(want).length) return want;
  const sh = ss_().getSheetByName(CUSTOMER_SHEET);
  const last = sh.getLastRow();
  if (last >= CUSTOMER_START_ROW) {
    const grid = sh.getRange(CUSTOMER_START_ROW, 1, last - CUSTOMER_START_ROW + 1, COL.ID).getValues();
    grid.forEach(function (r) {
      const rid = String(r[COL.ID - 1] || '').trim();
      if (rid && Object.prototype.hasOwnProperty.call(want, rid)) want[rid].push(r[COL.NO - 1]);
    });
  }
  return want;
}

/* 1件版（後方互換） */
function checkDuplicate(token, id) {
  const map = checkDuplicates(token, [id]);
  const nos = map[String(id || '').trim()] || [];
  return { dup: nos.length > 0, nos: nos };
}


/* ============================================================
 *  新規登録（まとめ登録）
 *   - LockService で同時登録を1件ずつ順番に処理
 *   - 顧客リストを1回だけ読み直して ID重複チェック／採番
 *   - 書き込み位置：APPEND_SEARCH_START_ROW から下へ向かって、
 *     「名前もIDも空」の行を探して詰めて書き込む。
 *     （行挿入は他シートも巻き込むセル数上限エラーの原因になったため
 *       使わない。sh.getLastRow() は他列の残骸に影響されるため
 *       使わない。→ 指定した行から下の空白行を実際に探すのが最も確実）
 *   - 重複IDでも登録する（応募数カウントのため）。重複時は
 *     V=重複あり／W=重複元NO／X=不可 を記録し、行を赤くする。
 *   - 担当者(M列) ＝ ログインユーザーの表示担当者名で固定
 *   - 応募日(K列) ＝ 登録日（今日）で固定／NOは自動採番
 *     （NOは行の位置に関係なく常に既存最大＋1。顧客一覧の「新しい順」
 *      表示はこのNOの大小で判定している＝どこに書き込まれても新しい順に見える）
 *
 *   rows: [{名前, ID, 年齢, 地域, 応募金額, 案件名, 媒体, 運用アカウント, LINE追加, 備考}, ...]
 *   戻り値: { results:[{index, ok, no, dup, dupNos, message}] }
 * ============================================================ */
function registerBatch(token, rows) {
  const user = requireUser_(token);
  if (!Array.isArray(rows) || !rows.length) throw new Error('登録データがありません。');
  if (rows.length > 100) throw new Error('一度に登録できるのは100件までです。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    throw new Error('現在ほかの登録処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const sh = ss_().getSheetByName(CUSTOMER_SHEET);
    const lastRaw = sh.getLastRow();   // 重複チェック・採番はシート全体を対象にする

    let grid = [];
    if (lastRaw >= CUSTOMER_START_ROW) {
      grid = sh.getRange(CUSTOMER_START_ROW, 1, lastRaw - CUSTOMER_START_ROW + 1, COL.ID).getValues();
    }

    // 既存ID索引・既存の最大NO（シート全体、行の位置は問わない）
    const idIndex = {};
    let maxNo = 0;
    grid.forEach(function (r) {
      const n = Number(r[COL.NO - 1]);
      if (!isNaN(n) && isFinite(n) && n > maxNo) maxNo = n;
      const rid = String(r[COL.ID - 1] || '').trim();
      if (rid) (idIndex[rid] = idIndex[rid] || []).push(r[COL.NO - 1]);
    });

    // 書き込み先の行：APPEND_SEARCH_START_ROW から下の空白行を必要数ぶん確保
    const targetRows = findBlankRows_(sh, APPEND_SEARCH_START_ROW, rows.length);

    const today = new Date();
    const results = [];

    rows.forEach(function (payload, k) {
      try {
        const name    = String(payload.名前 || '').trim();
        const id      = String(payload.ID || '').trim();
        const project = String(payload.案件名 || '');
        const media   = String(payload.媒体 || '');
        const account = String(payload.運用アカウント || '');
        if (!name || !id || !project || !media || !account) {
          throw new Error('必須項目（名前・ID・案件名・媒体・運用アカウント）が未入力です。');
        }
        if (OPT_PROJECT.indexOf(project) < 0) throw new Error('案件名が選択肢と一致しません：「' + project + '」');
        if (OPT_MEDIA.indexOf(media) < 0)     throw new Error('媒体が選択肢と一致しません：「' + media + '」');

        // 重複していても登録する（応募数として数え、重複あり／不可 として記録）
        const dupNos = (idIndex[id] || []).slice();   // 既存＋このバッチで先に確定したID
        const isDup = dupNos.length > 0;

        maxNo += 1;
        const newNo = maxNo;
        const targetRow = targetRows[k];

        writeCustomerRow_(sh, targetRow, {
          no: newNo, 名前: name, ID: id,
          年齢: payload.年齢, 地域: payload.地域, 応募金額: payload.応募金額,
          応募日: today, 案件名: project, 担当者: user.表示担当者名,
          運用アカウント: account, 媒体: media,
          LINE追加: payload.LINE追加 === true, 備考: payload.備考,
          dup: isDup, dupNos: dupNos
        });

        if (payload.LINE追加 === true) addJisseki_(newNo, user, 'LINE追加', 1);

        (idIndex[id] = idIndex[id] || []).push(newNo);   // 後続行の重複判定用
        results.push({ index: k, ok: true, no: newNo, dup: isDup, dupNos: dupNos });
      } catch (e) {
        results.push({ index: k, ok: false, message: e.message });
      }
    });

    SpreadsheetApp.flush();
    return { results: results };
  } finally {
    lock.releaseLock();
  }
}

/* 顧客リストの1行に書き込む（列がとびとびなのでセル単位） */
function writeCustomerRow_(sh, targetRow, d) {
  const put = function (col, val) { sh.getRange(targetRow, col).setValue(val); };
  put(COL.NO, d.no);
  put(COL.外部申請, false);              // 登録後、顧客詳細の操作で切り替え（Phase 2）
  put(COL.Line誘導, false);              // 同上
  put(COL.名前, d.名前);
  put(COL.ID, d.ID);
  put(COL.年齢, d.年齢 || '');
  put(COL.地域, d.地域 || '');
  if (d.応募金額 !== '' && d.応募金額 != null && !isNaN(Number(d.応募金額))) put(COL.応募金額, Number(d.応募金額));
  put(COL.応募日, d.応募日);             // K列 ＝ 今日
  put(COL.案件名, d.案件名);
  put(COL.担当者, d.担当者);             // M列 ＝ ログインユーザー固定
  put(COL.運用アカウント, d.運用アカウント);
  put(COL.媒体, d.媒体);
  put(COL.LINE追加, d.LINE追加 === true);
  if (d.備考) put(COL.備考, String(d.備考));
  if (d.dup) {
    put(COL.重複, '重複あり');
    put(COL.重複元NO, d.dupNos.join(','));
    put(COL.次工程, '不可');
    sh.getRange(targetRow, 1, 1, COL_LAST).setBackground('#f8c9c4');  // 行を赤く
  }
}

/* 1件版（後方互換） */
function registerCustomer(token, payload) {
  const r = registerBatch(token, [payload]);
  const res = r.results[0];
  if (!res.ok) throw new Error(res.message);
  return { ok: true, no: res.no, dup: res.dup, dupNos: res.dupNos };
}

function addJisseki_(no, user, type, value) {
  const sh = ss_().getSheetByName('システム実績');
  if (!sh) return;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([today, no, user.表示担当者名, user.チーム, type, value, '', user.表示担当者名, new Date()]);
}


/* ============================================================
 *  診断：新規登録がどこに入るか確認する
 * ============================================================ */
function peekCustomers() {
  const ss = ss_();
  const sh = ss.getSheetByName(CUSTOMER_SHEET);
  const out = [];
  out.push('スプレッドシート名: ' + ss.getName());
  out.push('スプレッドシートID : ' + ss.getId());
  out.push('シート名           : ' + (sh ? sh.getName() : '★見つかりません（CUSTOMER_SHEET の値を確認）'));
  if (!sh) { Logger.log(out.join('\n')); return out.join('\n'); }
  out.push('sh.getLastRow()         : ' + sh.getLastRow() + '（シート全体で何か値がある最後の行。参考情報）');
  out.push('APPEND_SEARCH_START_ROW : ' + APPEND_SEARCH_START_ROW + '（この行から下の空白行を探して書き込みます）');

  const targets = findBlankRows_(sh, APPEND_SEARCH_START_ROW, 5);
  out.push('次に書き込まれる行（先頭5件）: ' + targets.join(', '));

  out.push('----- APPEND_SEARCH_START_ROW の前後（確認用） -----');
  const from = Math.max(CUSTOMER_START_ROW, APPEND_SEARCH_START_ROW - 5);
  const to = APPEND_SEARCH_START_ROW + 9;
  const vals = sh.getRange(from, 1, to - from + 1, COL_LAST).getValues();
  vals.forEach(function (r, i) {
    out.push('行' + (from + i)
      + '｜NO=' + r[COL.NO - 1]
      + '｜名前=' + r[COL.名前 - 1]
      + '｜ID=' + r[COL.ID - 1]
      + '｜担当者=' + r[COL.担当者 - 1]
      + '｜案件名=' + r[COL.案件名 - 1]
      + '｜応募日=' + r[COL.応募日 - 1]
      + '｜重複=' + r[COL.重複 - 1]);
  });
  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

/* ============================================================
 *  診断：サーバー側だけで新規登録を1件試す（Webアプリを経由しない）
 * ============================================================ */
function testRegister() {
  const lg = login('ajima', '1234');
  if (!lg.ok) { Logger.log('login失敗: ' + JSON.stringify(lg)); return JSON.stringify(lg); }
  const res = registerCustomer(lg.token, {
    名前: 'サーバーテスト太郎',
    ID: 'srvtest_' + new Date().getTime(),
    案件名: OPT_PROJECT[0],
    媒体: 'CW',
    運用アカウント: 'gdcc',
    年齢: '30代',
    地域: '東京',
    LINE追加: false,
    備考: 'testRegister による書き込み'
  });
  Logger.log('registerCustomer 結果: ' + JSON.stringify(res));
  Logger.log(peekCustomers());
  return JSON.stringify(res);
}


/* ============================================================
 *  動作確認用（任意）：エディタから実行するとログインの流れを試せる
 * ============================================================ */
function _selftest() {
  const r = login('ajima', '1234');
  Logger.log(JSON.stringify(r));
  if (r.ok) {
    Logger.log(JSON.stringify(getBootstrap(r.token)));
    var cs = getCustomers(r.token, { mode: 'recent', limit: 50 });
    Logger.log('customers: 取得' + cs.rows.length + '件 / シート全' + cs.scanned + '行');
  }
}
