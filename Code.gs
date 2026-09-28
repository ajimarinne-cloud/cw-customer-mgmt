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
  ステータス: 19, 備考: 21, 重複: 22, 重複元NO: 23, 次工程: 24,
  // Q列：面談日を入れる既存の日付列（2026-09-20 ユーザー指定。Y列は人名プルダウン、O列は媒体のため使えない）
  面談予定日: 17
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

/* ========== 数字管理（KPI）関連（2026-09-15 追加） ==========
 * YES数・採用数はステータスに追加せず、面談カレンダーで日別に入力した数を
 * システム実績シートに「種別=YES／採用」で記録する。
 * 目標値は採用者ごとに用意してもらう「別スプシ」内の「KPI目標」タブ
 * （列：対象月／指標／月次目標）から読み込む。そのスプシIDはメンバーマスタの
 * 「目標スプシID」列に保存する（未設定なら目標なしで実績のみ表示）。
 * ※このKPI目標タブは各採用者の“別スプシ”側に作ってもらうものなので、
 *   このシステム本体のスプシ（MGMT_SHEETS）には同名のシートは作らない。
 * 集計は「応募日が属する月」を基準にしたコホート方式（旧スプシと同じ考え方）。
 */
const KPI_RECRUITER_METRICS = ['応募数', 'LINE追加数', 'アポ日程切数', 'アポ予定数（予約）', 'アポ済数（面談完了）', 'YES数', '採用数'];
const KPI_OPERATOR_METRICS  = ['応募数', 'LINE追加数', 'アポ日程切数'];
const KPI_TARGET_SHEET_NAME = 'KPI目標';

/* ========== 管理シートの見出し定義 ========== */
const MGMT_SHEETS = {
  'メンバーマスタ':          ['userId', 'passwordHash', '氏名', '表示担当者名', '役割', 'チーム', '権限', '有効', '目標スプシID'],
  'ユーザーアカウントマスタ': ['userId', 'アカウント名', '有効'],
  'ログイントークン':        ['tokenHash', 'userId', '最終アクセス日時', '有効期限', 'userAgent', '有効'],
  'システム実績':            ['日付', '顧客NO', '担当者', 'チーム', '種別', '数値', 'メモ', '更新者', '更新日時'],
  '操作履歴':                ['日時', 'userId', '内容', 'データ']
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
 *   ※ システム実績は消しません（見出しだけ用意）
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

  // システム実績の見出しだけ用意
  ['システム実績'].forEach(function (name) {
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
  let targetSheetId = '';
  if (user.役割 === '採用者') {
    const rec = getMembers_().find(function (m) { return String(m.userId).trim() === user.userId; });
    targetSheetId = rec ? String(rec.目標スプシID || '').trim() : '';
  }
  return {
    user: user,
    myAccounts: myAccounts,
    owners: owners,
    options: { media: OPT_MEDIA, project: OPT_PROJECT, status: OPT_STATUS },
    targetSheetId: targetSheetId
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

/**
 * 顧客リストからNOで行を探し、自分が担当した顧客であることを確認して行番号を返す。
 * 権限がない／見つからない場合は例外を投げる（更新系の関数で共通して使う）。
 */
function findOwnRow_(sh, no, user, errLabel) {
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
    throw new Error('自分が担当した顧客のみ' + (errLabel || '更新') + 'できます。');
  }
  return targetRow;
}

/** システム実績に、その顧客NOでまだ記録されていない種別だけ追記する（二重記録を防ぐ） */
function logJissekiOnce_(no, user, type) {
  const already = sheetObjects_('システム実績').some(function (r) {
    return Number(r.顧客NO) === Number(no) && r.種別 === type;
  });
  if (already) return false;
  addJisseki_(no, user, type, 1);
  return true;
}

/* ============================================================
 *  元に戻す（操作履歴）（2026-09-22 追加）
 *   顧客の更新（LINE追加／ステータス／面談予定日／面談実施チェック）と、
 *   カレンダーのYES数・採用数の保存を、操作の直前の状態に戻せる。
 *   - 戻せるのは自分の操作だけ。直近 UNDO_KEEP 件・UNDO_DAYS 日以内。
 *   - 操作したあとに別の変更（他の人を含む）が入っていたら、戻さずにエラーにする。
 *   - 履歴は「操作履歴」シートに、操作ごとに1行（内容＋JSON）で残す。
 * ============================================================ */
const UNDO_SHEET = '操作履歴';
const UNDO_KEEP = 10;
const UNDO_DAYS = 7;
const CUSTOMER_JIS_TYPES = ['LINE追加', 'アポ日程切', '面談完了'];
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

function undoSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(UNDO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(UNDO_SHEET);
    const h = MGMT_SHEETS[UNDO_SHEET];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 'yyyy-mm-dd' → 'm/d'（年を出さない表示用） */
function mdShort_(ymd) {
  const m = String(ymd || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? Number(m[2]) + '/' + Number(m[3]) : String(ymd || '');
}

function recordUndo_(user, label, payload) {
  const sh = undoSheet_();
  sh.appendRow([new Date(), user.userId, label, JSON.stringify(payload)]);
  pruneUndo_(sh, user.userId);
}

/** 古い履歴（UNDO_DAYS日より前）と、そのユーザーの UNDO_KEEP 件を超えた分を消す */
function pruneUndo_(sh, userId) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  const limit = new Date().getTime() - UNDO_DAYS * 24 * 60 * 60 * 1000;
  const del = [], mine = [];
  vals.forEach(function (r, i) {
    const row = i + 2;
    const t = (r[0] instanceof Date) ? r[0].getTime() : 0;
    if (t < limit) { del.push(row); return; }
    if (String(r[1]).trim() === userId) mine.push(row);
  });
  if (mine.length > UNDO_KEEP) {
    mine.slice(0, mine.length - UNDO_KEEP).forEach(function (row) { del.push(row); });
  }
  del.sort(function (a, b) { return b - a; }).forEach(function (row) { sh.deleteRow(row); });
}

/** そのユーザーの有効な履歴（古い→新しい順） */
function userUndoEntries_(user) {
  const limit = new Date().getTime() - UNDO_DAYS * 24 * 60 * 60 * 1000;
  return sheetObjects_(UNDO_SHEET).filter(function (r) {
    const t = (r.日時 instanceof Date) ? r.日時.getTime() : 0;
    return String(r.userId).trim() === user.userId && t >= limit;
  });
}

/** 「元に戻す」ボタン用：戻せる件数と、直前の操作の内容 */
function getUndoInfo(token) {
  const user = requireUser_(token);
  const es = userUndoEntries_(user);
  return { count: es.length, label: es.length ? String(es[es.length - 1].内容) : '' };
}

/** 顧客1行の状態（LINE追加／ステータス／面談予定日、withJis=trueなら数字の記録も）を取る */
function customerSnapshot_(sh, row, no, withJis) {
  const tz = Session.getScriptTimeZone();
  const v = sh.getRange(row, 1, 1, COL_LAST).getValues()[0];
  const snap = {
    name: String(v[COL.名前 - 1] || ''),
    line: v[COL.LINE追加 - 1] === true,
    status: String(v[COL.ステータス - 1] || '').trim(),
    md: fmtDate_(v[COL.面談予定日 - 1], tz)
  };
  if (withJis) {
    snap.jis = sheetObjects_('システム実績').filter(function (r) {
      return Number(r.顧客NO) === Number(no) && CUSTOMER_JIS_TYPES.indexOf(r.種別) >= 0;
    }).map(function (r) {
      return [fmtDate_(r.日付, tz), r.種別, (r.数値 === '' || r.数値 == null) ? 1 : Number(r.数値)];
    });
  }
  return snap;
}

function jisKey_(arr) {
  return (arr || []).map(function (j) { return JSON.stringify(j); }).sort().join('|');
}

/** 面談予定日セルに書く（空なら消す）。入力規則があっても書けるようにする */
function setMeetingCell_(sh, row, md) {
  const cell = sh.getRange(row, COL.面談予定日);
  if (!md) { cell.clearContent(); return; }
  if (!DATE_RE.test(md)) { cell.setValue(md); return; }
  const p = md.split('-');
  cell.clearDataValidations();
  cell.setValue(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
}

function writeCustomerCells_(sh, row, snap) {
  sh.getRange(row, COL.LINE追加).setValue(snap.line === true);
  const sc = sh.getRange(row, COL.ステータス);
  if (snap.status) sc.setValue(snap.status); else sc.clearContent();
  setMeetingCell_(sh, row, snap.md);
}

function describeChange_(before, after) {
  const parts = [];
  if (before.line !== after.line) parts.push('LINE追加を' + (after.line ? 'オン' : 'オフ'));
  if (before.status !== after.status) parts.push('ステータス「' + (before.status || '未設定') + '」→「' + (after.status || '未設定') + '」');
  if (before.md !== after.md) parts.push('面談日 ' + (mdShort_(before.md) || 'なし') + '→' + (mdShort_(after.md) || 'なし'));
  return parts.join('、');
}

/**
 * 顧客1件への変更を実行し、前後の状態を比べて変わっていれば履歴に残す。
 * mutate(before) の中で、顧客リストとシステム実績への書き込みを行う。
 * describe(before, after) は履歴に出す説明文を返す。
 */
function applyCustomerChange_(user, sh, row, no, describe, mutate) {
  const before = customerSnapshot_(sh, row, no, true);
  mutate(before);
  SpreadsheetApp.flush();
  const afterFull = customerSnapshot_(sh, row, no, true);
  const changed = before.line !== afterFull.line || before.status !== afterFull.status || before.md !== afterFull.md
    || jisKey_(before.jis) !== jisKey_(afterFull.jis);
  if (!changed) return;
  const desc = (describe ? describe(before, afterFull) : '') || describeChange_(before, afterFull) || '数字の記録を更新';
  recordUndo_(user, 'NO:' + no + ' ' + before.name + '：' + desc, {
    kind: 'customer', no: Number(no), before: before,
    after: { line: afterFull.line, status: afterFull.status, md: afterFull.md }
  });
}

/**
 * 顧客の LINE追加／ステータス／面談予定日 をまとめて更新する共通処理（自分の担当顧客のみ）。
 * patch: { line: true|false, status: '面談予約' など（'' は未設定に戻す）, md: 'yyyy-mm-dd' }
 * 指定した項目だけ変える。数字（システム実績）は次のルールで合わせる。
 *  - LINE追加オン ：種別「LINE追加」を1回だけ記録／オフ：その記録を消す
 *  - ステータスが「面談予約」「面談済み」 ：種別「アポ日程切」を1回だけ記録
 *  - ステータスを上記以外へ変更   ：「アポ日程切」の記録を消す
 *  - ステータスを「面談済み」以外へ変更 ：「面談完了」（アポ済み）の記録も消す
 *  - 「面談予約」にするときは面談予定日が必須
 */
function editCustomer_(token, no, patch, errLabel) {
  const user = requireUser_(token);
  patch = patch || {};
  let st = null, md = null;
  if (patch.status !== undefined && patch.status !== null) {
    st = String(patch.status).trim();
    if (st && OPT_STATUS.indexOf(st) < 0) throw new Error('ステータスが選択肢と一致しません：「' + st + '」');
  }
  if (patch.md !== undefined && patch.md !== null) {
    md = String(patch.md).trim();
    if (md && !DATE_RE.test(md)) throw new Error('面談予定日が正しくありません。');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const sh = ss_().getSheetByName(CUSTOMER_SHEET);
    const row = findOwnRow_(sh, no, user, errLabel || '更新');
    applyCustomerChange_(user, sh, row, no, null, function (before) {
      const newStatus = (st === null) ? before.status : st;
      const newMd = (md === null) ? before.md : md;
      if (newStatus === '面談予約' && !DATE_RE.test(newMd)) {
        throw new Error('「面談予約」にするときは、面談予定日を入力してください。');
      }

      if (patch.line !== undefined && patch.line !== null) {
        const on = (patch.line === true || patch.line === 'true');
        sh.getRange(row, COL.LINE追加).setValue(on);
        if (on) logJissekiOnce_(no, user, 'LINE追加');
        else deleteJissekiRows_(function (r) { return Number(r.顧客NO) === Number(no) && r.種別 === 'LINE追加'; });
      }

      if (st !== null) {
        const sc = sh.getRange(row, COL.ステータス);
        if (st) sc.setValue(st); else sc.clearContent();
        if (st === '面談予約' || st === '面談済み') {
          logJissekiOnce_(no, user, 'アポ日程切');
        } else if (st !== before.status) {
          deleteJissekiRows_(function (r) { return Number(r.顧客NO) === Number(no) && r.種別 === 'アポ日程切'; });
        }
        if (st !== before.status && st !== '面談済み') {
          deleteJissekiRows_(function (r) { return Number(r.顧客NO) === Number(no) && r.種別 === '面談完了'; });
        }
      }

      if (md !== null && md !== before.md) setMeetingCell_(sh, row, md);
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 *  ステータス更新（「未対応リスト」から入力）
 *   自分が担当した顧客だけ更新できる。「面談予約」にするときは面談予定日
 *   （meetingDate: 'yyyy-mm-dd'）が必須。数字の合わせ方は editCustomer_ を参照。
 *   ※アポ済数（面談完了）は、数字管理タブの「面談カレンダー」で
 *     「面談実施」にチェックした日を数える（setMeetingDone）。
 * ============================================================ */
function updateCustomerStatus(token, no, status, meetingDate) {
  const st = String(status || '').trim();
  if (OPT_STATUS.indexOf(st) < 0) throw new Error('ステータスが選択肢と一致しません：「' + st + '」');
  const patch = { status: st };
  if (st === '面談予約') patch.md = String(meetingDate || '').trim();
  return editCustomer_(token, no, patch, 'ステータスを更新');
}

/** LINE追加にする（「未対応リスト」のボタン）。自分が担当した顧客のみ */
function markLineAdded(token, no) {
  return editCustomer_(token, no, { line: true }, 'LINE追加に');
}

/** 顧客一覧の編集（自分の担当顧客のみ）：patch = { line, status, md } の指定した項目だけ更新 */
function updateCustomerFields(token, no, patch) {
  return editCustomer_(token, no, patch, '更新');
}

/* ============================================================
 *  面談実施のチェック（数字管理タブの「面談カレンダー」から）
 *   done=true ：ステータスを「面談済み」に変え、その日（dateStr='yyyy-mm-dd'）の
 *               アポ済み（種別=面談完了）としてシステム実績に記録する。
 *   done=false：アポ済みの記録を消し、ステータスが「面談済み」なら「面談予約」に戻す。
 *   自分が担当した顧客だけ操作できる。
 * ============================================================ */
function setMeetingDone(token, no, done, dateStr) {
  const user = requireUser_(token);
  const day = String(dateStr || '').trim();
  if (done && !DATE_RE.test(day)) throw new Error('日付が正しくありません。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const sh = ss_().getSheetByName(CUSTOMER_SHEET);
    const targetRow = findOwnRow_(sh, no, user, '面談を更新');
    const describe = function () {
      return done ? '面談実施にチェック（' + mdShort_(day) + '）' : '面談実施のチェックを外した';
    };
    applyCustomerChange_(user, sh, targetRow, no, describe, function (before) {
      const statusCell = sh.getRange(targetRow, COL.ステータス);
      // 同じ顧客の既存のアポ済み記録は一度消して、チェックした日で入れ直す
      deleteJissekiRows_(function (r) { return Number(r.顧客NO) === Number(no) && r.種別 === '面談完了'; });
      if (done) {
        statusCell.setValue('面談済み');
        logJissekiOnce_(no, user, 'アポ日程切');
        addJisseki_(no, user, '面談完了', 1, day);
      } else if (before.status === '面談済み') {
        statusCell.setValue('面談予約');
      }
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** その日に顧客に紐づけず入力されたYES数・採用数の合計（担当者ごと） */
function manualDayValues_(user, day) {
  const out = { yes: 0, hire: 0 };
  sheetObjects_('システム実績').forEach(function (r) {
    if (!(r.顧客NO === '' || r.顧客NO == null)) return;
    if (String(r.担当者 || '').trim() !== user.表示担当者名 || fmtDate_(r.日付) !== day) return;
    const n = (r.数値 === '' || r.数値 == null) ? 1 : (Number(r.数値) || 0);
    if (r.種別 === 'YES') out.yes += n;
    else if (r.種別 === '採用') out.hire += n;
  });
  return out;
}

/** その日のYES・採用の入力行を消して入れ直す。nums = { 'YES': 数|'' , '採用': 数|'' }（''・0は記録なし） */
function writeDayCounts_(user, day, nums) {
  Object.keys(nums).forEach(function (type) {
    deleteJissekiRows_(function (r) {
      return (r.顧客NO === '' || r.顧客NO == null) && r.種別 === type
        && String(r.担当者 || '').trim() === user.表示担当者名 && fmtDate_(r.日付) === day;
    });
    if (nums[type] !== '' && nums[type] > 0) addJisseki_('', user, type, nums[type], day);
  });
}

/* ============================================================
 *  日別のYES数・採用数の入力（数字管理タブの「面談カレンダー」から）
 *   その日（dateStr）の数字を上書き保存する。空欄は「入力なし」（記録を消す）。
 *   顧客には紐づけず、システム実績に 種別=YES／採用・数値=入力した数 で1日1行持つ。
 * ============================================================ */
function saveDayCounts(token, dateStr, yes, hire) {
  const user = requireUser_(token);
  const day = String(dateStr || '').trim();
  if (!DATE_RE.test(day)) throw new Error('日付が正しくありません。');

  const entries = { 'YES': yes, '採用': hire };
  const nums = {};
  Object.keys(entries).forEach(function (type) {
    const raw = entries[type];
    if (raw === '' || raw == null) { nums[type] = ''; return; }
    const n = Number(raw);
    if (isNaN(n) || n < 0 || Math.floor(n) !== n) throw new Error('数字は0以上の整数で入力してください。');
    nums[type] = n;
  });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const before = manualDayValues_(user, day);
    writeDayCounts_(user, day, nums);
    SpreadsheetApp.flush();
    const after = manualDayValues_(user, day);
    if (before.yes !== after.yes || before.hire !== after.hire) {
      recordUndo_(user,
        mdShort_(day) + ' のYES・採用を保存（YES ' + before.yes + '→' + after.yes + '、採用 ' + before.hire + '→' + after.hire + '）',
        { kind: 'day', day: day, before: before, after: after });
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 *  元に戻す（直前の自分の操作を1つ取り消す）
 * ============================================================ */
function undoLast(token) {
  const user = requireUser_(token);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    const es = userUndoEntries_(user);
    if (!es.length) throw new Error('元に戻せる操作がありません。');
    const e = es[es.length - 1];
    const hist = undoSheet_();
    let data;
    try { data = JSON.parse(e.データ); } catch (x) {
      hist.deleteRow(e._row);
      throw new Error('この操作は元に戻せません。');
    }
    const conflict = function () {
      hist.deleteRow(e._row);
      throw new Error('その後に別の変更があったため、「' + e.内容 + '」は元に戻せません。');
    };

    if (data.kind === 'customer') {
      const cs = ss_().getSheetByName(CUSTOMER_SHEET);
      const row = findOwnRow_(cs, data.no, user, '元に戻');
      const cur = customerSnapshot_(cs, row, data.no, false);
      if (cur.line !== data.after.line || cur.status !== data.after.status || cur.md !== data.after.md) conflict();
      writeCustomerCells_(cs, row, data.before);
      deleteJissekiRows_(function (r) {
        return Number(r.顧客NO) === Number(data.no) && CUSTOMER_JIS_TYPES.indexOf(r.種別) >= 0;
      });
      (data.before.jis || []).forEach(function (j) { addJisseki_(data.no, user, j[1], j[2], j[0]); });
    } else if (data.kind === 'day') {
      const cur = manualDayValues_(user, data.day);
      if (cur.yes !== data.after.yes || cur.hire !== data.after.hire) conflict();
      writeDayCounts_(user, data.day, { 'YES': data.before.yes || '', '採用': data.before.hire || '' });
    } else {
      hist.deleteRow(e._row);
      throw new Error('この操作は元に戻せません。');
    }
    SpreadsheetApp.flush();
    hist.deleteRow(e._row);
    return { ok: true, label: String(e.内容) };
  } finally {
    lock.releaseLock();
  }
}

/** システム実績から、条件に合う行をすべて削除する（下の行から消す） */
function deleteJissekiRows_(predicate) {
  const sh = ss_().getSheetByName('システム実績');
  if (!sh) return;
  sheetObjects_('システム実績')
    .filter(predicate)
    .map(function (r) { return r._row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (row) { sh.deleteRow(row); });
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
    次工程: r[COL.次工程 - 1],
    面談予定日: fmtDate_(r[COL.面談予定日 - 1], tz)
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

/** システム実績に1行追記する。dateStr（'yyyy-mm-dd'）を省略すると今日の日付で記録する */
function addJisseki_(no, user, type, value, dateStr) {
  const sh = ss_().getSheetByName('システム実績');
  if (!sh) return;
  const day = dateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([day, no, user.表示担当者名, user.チーム, type, value, '', user.表示担当者名, new Date()]);
}


/* ============================================================
 *  数字管理（KPI）（2026-09-15 日付ベースに再設計）
 *
 *  考え方：
 *   - 応募数だけは「応募日が属する月」のコホートで数える（応募日＝その人が発生した日そのもの）。
 *   - アポ日程切数は、ステータスを初めて「面談予約」（または「面談済み」）にした日を数える。
 *   - LINE追加数は「LINE追加にする」を押した日、アポ済数（面談完了）は面談カレンダーで
 *     「面談実施」にチェックした日（面談予定日）、YES数・採用数は面談カレンダーで日別に
 *     入力した数を、システム実績へ日付つきで記録し、対象月の分を担当者別に足し合わせる。
 *     「今の状態」のスナップショットではなく実際に起きた日で正確に月集計できる。
 *   - アポ予定数（予約）は、ステータスの現在値ではなく「面談予定日」（Q列の面談日）が
 *     対象月に入っている件数で数える。数字管理タブの「面談カレンダー」では、この面談予定日
 *     ごとに名前を確認し、面談実施のチェック・YES数／採用数の入力ができる。
 *   - 採用者：応募数／LINE追加数／アポ日程切数／アポ予定数（予約）／アポ済数（面談完了）／
 *     YES数／採用数　を自分の担当ぶんで集計。あわせて同じチームの運用者の数字も見える。
 *   - 達成率は転換率（ひとつ前の段階に対する割合）：LINE追加÷応募、アポ日程切÷LINE追加、
 *     アポ予定÷アポ日程切、アポ済÷アポ予定、YES÷アポ済、採用÷YES。応募数だけは 応募数÷月目標。
 *     ギャップ＝月目標−現状。「今日までの目標」＝月次目標×（経過日数÷その月の日数）で、
 *     現状がそれに届いていなければ画面で赤字にする（判定は画面側）。
 *   - 運用者：応募数／LINE追加数／アポ日程切数　を自分の運用アカウントぶんで集計（目標なし）。
 *     ※運用者側は現時点では「今の状態」のスナップショット集計のまま（要望があれば同様に
 *       日付ベースへ拡張できる）。
 *   - 目標値は採用者ごとに用意してもらう別スプシの「KPI目標」タブ（対象月／指標／月次目標）
 *     から読み込む。未設定なら「目標なし」として実績のみ表示する。
 * ============================================================ */
function pad2_(n) { return ('0' + n).slice(-2); }

function kpiBlankCounts_() {
  return { 応募数: 0, LINE追加数: 0, アポ日程切数: 0, 'アポ予定数（予約）': 0, 'アポ済数（面談完了）': 0, YES数: 0, 採用数: 0 };
}

/** opts: {mode:'month', year, month} または {mode:'quarter', year, quarter(1-4)} → 対象月(YYYY-MM)の配列 */
function monthsForRequest_(opts) {
  opts = opts || {};
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const year = Number(opts.year) || Number(Utilities.formatDate(now, tz, 'yyyy'));
  if (opts.mode === 'quarter') {
    const q = Math.min(Math.max(Number(opts.quarter) || 1, 1), 4);
    const startMonth = (q - 1) * 3 + 1;
    const out = [];
    for (let i = 0; i < 3; i++) out.push(year + '-' + pad2_(startMonth + i));
    return out;
  }
  const month = Math.min(Math.max(Number(opts.month) || Number(Utilities.formatDate(now, tz, 'M')), 1), 12);
  return [year + '-' + pad2_(month)];
}

/**
 * 顧客リストを1回だけ読み、次の3つを集計する。
 *  - byOwnerApplications: 応募数（応募日が対象月コホートに入る件数、担当者別）
 *  - byOwnerMeetings    : アポ予定数（予約）（面談予定日が対象月に入る件数、担当者別）
 *                         ※2026-09-15〜：ステータスの現在値ではなく実際の面談予定日で数える
 *  - byAccount          : 運用者・チーム表示用（現在値のスナップショット、運用アカウント別）
 *                         ※LINE追加は「今の状態」を見ているだけで、後から日付を
 *                           正確に遡れるわけではない。運用者側は今回まだそこまで求められて
 *                           いないため、従来どおりのスナップショット集計のまま残している。
 */
function computeFunnelCounts_(months) {
  const sh = ss_().getSheetByName(CUSTOMER_SHEET);
  const last = sh.getLastRow();
  const byOwnerApplications = {}, byOwnerMeetings = {}, byAccount = {};
  const monthSet = {};
  (months || []).forEach(function (m) { monthSet[m] = true; });
  if (last < CUSTOMER_START_ROW) return { byOwnerApplications: byOwnerApplications, byOwnerMeetings: byOwnerMeetings, byAccount: byAccount };

  const tz = Session.getScriptTimeZone();
  const values = sh.getRange(CUSTOMER_START_ROW, 1, last - CUSTOMER_START_ROW + 1, COL_LAST).getValues();

  values.forEach(function (r) {
    const name = String(r[COL.名前 - 1] || '').trim();
    const id = String(r[COL.ID - 1] || '').trim();
    if (!name && !id) return;

    const owner = String(r[COL.担当者 - 1] || '').trim();
    const account = String(r[COL.運用アカウント - 1] || '').trim();
    const appliedYm = fmtDate_(r[COL.応募日 - 1], tz).slice(0, 7);
    const meetingYm = fmtDate_(r[COL.面談予定日 - 1], tz).slice(0, 7);
    const status = r[COL.ステータス - 1];
    const lineAdded = r[COL.LINE追加 - 1] === true;

    if (owner && monthSet[appliedYm]) {
      byOwnerApplications[owner] = (byOwnerApplications[owner] || 0) + 1;
    }
    if (owner && monthSet[meetingYm]) {
      byOwnerMeetings[owner] = (byOwnerMeetings[owner] || 0) + 1;
    }
    if (account && monthSet[appliedYm]) {
      if (!byAccount[account]) byAccount[account] = { 応募数: 0, LINE追加数: 0, アポ日程切数: 0 };
      byAccount[account].応募数++;
      if (lineAdded) byAccount[account].LINE追加数++;
      if (status === '面談予約' || status === '面談済み') byAccount[account].アポ日程切数++;
    }
  });

  return { byOwnerApplications: byOwnerApplications, byOwnerMeetings: byOwnerMeetings, byAccount: byAccount };
}

/**
 * システム実績を1回だけ読み、対象月に発生した「その日の記録」を担当者別に集計する。
 * LINE追加は「LINE追加にする」を押した日、面談完了はカレンダーでチェックした日、
 * YES・採用はカレンダーで入力した日に、システム実績へ日付つきで記録されているため、
 * 「応募日のコホート」ではなく「実際に起きた日」で正確に月集計できる。
 * 各行の「数値」を足し合わせる（1件の記録は数値1、日別入力は入力した数）。
 * ※日付セルがスプシで日付型に変換されていても読めるよう fmtDate_ で文字列に揃える。
 */
function computeEventCounts_(months) {
  const monthSet = {};
  (months || []).forEach(function (m) { monthSet[m] = true; });
  const byOwner = {};
  sheetObjects_('システム実績').forEach(function (r) {
    const ym = fmtDate_(r.日付).slice(0, 7);
    if (!monthSet[ym]) return;
    const owner = String(r.担当者 || '').trim();
    if (!owner) return;
    const n = (r.数値 === '' || r.数値 == null) ? 1 : (Number(r.数値) || 0);
    if (!byOwner[owner]) byOwner[owner] = { LINE追加数: 0, アポ日程切数: 0, 'アポ済数（面談完了）': 0, YES数: 0, 採用数: 0 };
    if (r.種別 === 'LINE追加') byOwner[owner].LINE追加数 += n;
    else if (r.種別 === 'アポ日程切') byOwner[owner].アポ日程切数 += n;
    else if (r.種別 === '面談完了') byOwner[owner]['アポ済数（面談完了）'] += n;
    else if (r.種別 === 'YES') byOwner[owner].YES数 += n;
    else if (r.種別 === '採用') byOwner[owner].採用数 += n;
  });
  return { byOwner: byOwner };
}

/**
 * 面談カレンダー用：ログインユーザーが担当した顧客の、指定した年月ぶんの情報を返す。
 *  days    : 面談予定日（Q列）が入っている日ごとの顧客 [{no,名前,ステータス,done}]
 *            done＝その顧客のアポ済み（面談完了）が記録されているか
 *  totals  : 日ごとの数字 { 'YYYY-MM-DD': {done, yes, hire} }（システム実績の合計）
 *  manual  : 日別入力の値 { 'YYYY-MM-DD': {yes, hire} }（顧客に紐づかない入力ぶん）
 * 戻り値： { year, month, days, totals, manual }
 */
function getMeetingCalendar(token, opts) {
  const user = requireUser_(token);
  opts = opts || {};
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const year = Number(opts.year) || Number(Utilities.formatDate(now, tz, 'yyyy'));
  const month = Math.min(Math.max(Number(opts.month) || Number(Utilities.formatDate(now, tz, 'M')), 1), 12);
  const ym = year + '-' + pad2_(month);

  // システム実績：面談完了済みの顧客NO、日ごとの合計、日別入力ぶん
  const doneNos = {}, totals = {}, manual = {};
  sheetObjects_('システム実績').forEach(function (r) {
    if (String(r.担当者 || '').trim() !== user.表示担当者名) return;
    const day = fmtDate_(r.日付, tz);
    const n = (r.数値 === '' || r.数値 == null) ? 1 : (Number(r.数値) || 0);
    if (r.種別 === '面談完了') doneNos[Number(r.顧客NO)] = true;
    if (day.indexOf(ym) !== 0) return;
    if (!totals[day]) totals[day] = { done: 0, yes: 0, hire: 0 };
    if (r.種別 === '面談完了') totals[day].done += n;
    else if (r.種別 === 'YES') totals[day].yes += n;
    else if (r.種別 === '採用') totals[day].hire += n;
    if ((r.種別 === 'YES' || r.種別 === '採用') && (r.顧客NO === '' || r.顧客NO == null)) {
      if (!manual[day]) manual[day] = { yes: 0, hire: 0 };
      if (r.種別 === 'YES') manual[day].yes += n; else manual[day].hire += n;
    }
  });

  const sh = ss_().getSheetByName(CUSTOMER_SHEET);
  const last = sh.getLastRow();
  const days = {};
  if (last >= CUSTOMER_START_ROW) {
    const values = sh.getRange(CUSTOMER_START_ROW, 1, last - CUSTOMER_START_ROW + 1, COL_LAST).getValues();
    values.forEach(function (r) {
      const name = String(r[COL.名前 - 1] || '').trim();
      const id = String(r[COL.ID - 1] || '').trim();
      if (!name && !id) return;
      const owner = String(r[COL.担当者 - 1] || '').trim();
      if (owner !== user.表示担当者名) return;
      const md = fmtDate_(r[COL.面談予定日 - 1], tz);
      if (!md || md.indexOf(ym) !== 0) return;
      if (!days[md]) days[md] = [];
      const no = r[COL.NO - 1];
      days[md].push({ no: no, 名前: name, ステータス: r[COL.ステータス - 1], done: doneNos[Number(no)] === true });
    });
  }
  return { year: year, month: month, days: days, totals: totals, manual: manual };
}

/** 見出し行から列名の位置(0始まり)を探す。前後や途中の空白（全角含む）は無視する。無ければ -1 */
function hdrIdx_(headers, name) {
  const norm = function (s) { return String(s == null ? '' : s).replace(/[\s　]/g, ''); };
  const want = norm(name);
  for (let i = 0; i < headers.length; i++) if (norm(headers[i]) === want) return i;
  return -1;
}

/** 「対象月」セルの値を 'YYYY-MM' に揃える（文字列 '2026-10' / '2026/10' / 日付型のどれでも可） */
function normYm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  const m = String(v || '').trim().match(/^(\d{4})[-\/](\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) : '';
}

function ownTargetSheetId_(user) {
  const rec = getMembers_().find(function (m) { return String(m.userId).trim() === user.userId; });
  return rec ? String(rec.目標スプシID || '').trim() : '';
}

/** 目標スプシの「KPI目標」タブから、対象月ごとの目標を返す（{ 'YYYY-MM': { 指標: 数値 } }） */
function readKpiTargetsByMonth_(spreadsheetId, months) {
  const out = {};
  (months || []).forEach(function (m) { out[m] = {}; });
  const total = readKpiTargetRows_(spreadsheetId, months);
  total.forEach(function (t) { out[t.ym][t.metric] = (out[t.ym][t.metric] || 0) + t.val; });
  return out;
}

/** 対象月ぶんの目標を指標ごとに合計して返す */
function readKpiTargets_(spreadsheetId, months) {
  const out = {};
  readKpiTargetRows_(spreadsheetId, months).forEach(function (t) {
    out[t.metric] = (out[t.metric] || 0) + t.val;
  });
  return out;
}

/**
 * 「今日までの目標」を返す：月次目標 × 経過日数 ÷ その月の日数 を対象月ぶん合計する。
 * 過ぎた月は月全体の目標、これから来る月は0（達成率を月の日数で出すため）。
 */
function kpiTargetsToDate_(byMonth, months) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const curYm = Utilities.formatDate(now, tz, 'yyyy-MM');
  const day = Number(Utilities.formatDate(now, tz, 'd'));
  const out = {};
  (months || []).forEach(function (ym) {
    const p = ym.split('-');
    const daysInMonth = new Date(Number(p[0]), Number(p[1]), 0).getDate();
    const frac = ym < curYm ? 1 : (ym === curYm ? day / daysInMonth : 0);
    const t = byMonth[ym] || {};
    Object.keys(t).forEach(function (metric) {
      out[metric] = (out[metric] || 0) + t[metric] * frac;
    });
  });
  Object.keys(out).forEach(function (k) { out[k] = Math.round(out[k] * 10) / 10; });
  return out;
}

/** 目標スプシのKPI目標タブを1行ずつ読み、対象月に入る行を [{ym, metric, val}] で返す（空欄は除く） */
function readKpiTargetRows_(spreadsheetId, months) {
  const rows = [];
  const id = String(spreadsheetId || '').trim();
  if (!id) return rows;
  try {
    const targetSs = SpreadsheetApp.openById(id);
    const sh = targetSs.getSheetByName(KPI_TARGET_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return rows;
    const values = sh.getDataRange().getValues();
    const headers = values.shift();
    const idx = { month: hdrIdx_(headers, '対象月'), metric: hdrIdx_(headers, '指標'), target: hdrIdx_(headers, '月次目標') };
    if (idx.month < 0 || idx.metric < 0 || idx.target < 0) return rows;
    const monthSet = {};
    (months || []).forEach(function (m) { monthSet[m] = true; });
    values.forEach(function (row) {
      const ym = normYm_(row[idx.month]);
      if (!monthSet[ym]) return;
      if (row[idx.target] === '' || row[idx.target] == null) return;   // 空欄は「目標なし」
      rows.push({ ym: ym, metric: String(row[idx.metric] || '').trim(), val: Number(row[idx.target]) || 0 });
    });
  } catch (e) {
    // スプシが開けない・権限がない等の場合は「目標なし」として扱う
  }
  return rows;
}

/** サイト上の「目標を編集」用：指定した月の目標を返す（{hasTargetSheet, values:{指標:数値}}） */
function getMyKpiTargets(token, ym) {
  const user = requireUser_(token);
  if (user.役割 !== '採用者') throw new Error('採用者のみ利用できます。');
  const id = ownTargetSheetId_(user);
  if (!id) return { hasTargetSheet: false, values: {} };
  return { hasTargetSheet: true, values: readKpiTargets_(id, [normYm_(ym)]) };
}

/**
 * サイト上の「目標を編集」から、指定した月の目標を目標スプシの「KPI目標」タブへ書き込む。
 * values: { 指標名: 数値または空文字 }（空文字は「目標なし」にする）。
 * すでに同じ月・指標の行があれば上書き、なければ行を追加する。
 * ※目標スプシに「編集権限」が必要。タブが無ければ自動で作る。
 */
function saveMyKpiTargets(token, ym, values) {
  const user = requireUser_(token);
  if (user.役割 !== '採用者') throw new Error('採用者のみ設定できます。');
  const month = normYm_(ym);
  if (!month) throw new Error('対象月が正しくありません。');
  const id = ownTargetSheetId_(user);
  if (!id) throw new Error('先に「目標スプシ」のIDを登録してください。');

  const updates = {};
  KPI_RECRUITER_METRICS.forEach(function (metric) {
    if (!values || !Object.prototype.hasOwnProperty.call(values, metric)) return;
    const raw = values[metric];
    if (raw === '' || raw == null) { updates[metric] = ''; return; }
    const n = Number(raw);
    if (isNaN(n) || n < 0) throw new Error('「' + metric + '」は0以上の数字で入力してください。');
    updates[metric] = n;
  });

  let targetSs;
  try { targetSs = SpreadsheetApp.openById(id); } catch (e) {
    throw new Error('目標スプシを開けません。IDと共有設定（編集権限）を確認してください。');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('現在ほかの処理が実行中です。数秒待ってからもう一度お試しください。');
  }
  try {
    let sh = targetSs.getSheetByName(KPI_TARGET_SHEET_NAME);
    if (!sh) sh = targetSs.insertSheet(KPI_TARGET_SHEET_NAME);
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, 3).setValues([['対象月', '指標', '月次目標']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const lastCol = Math.max(sh.getLastColumn(), 3);
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const idx = { month: hdrIdx_(headers, '対象月') + 1, metric: hdrIdx_(headers, '指標') + 1, target: hdrIdx_(headers, '月次目標') + 1 };
    if (idx.month < 1 || idx.metric < 1 || idx.target < 1) {
      throw new Error('目標スプシの「KPI目標」タブの見出しが「対象月／指標／月次目標」になっていません。');
    }

    const lastRow = sh.getLastRow();
    const rows = lastRow >= 2 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    Object.keys(updates).forEach(function (metric) {
      const val = updates[metric];
      const hits = [];
      rows.forEach(function (r, i) {
        if (normYm_(r[idx.month - 1]) === month && String(r[idx.metric - 1] || '').trim() === metric) hits.push(i + 2);
      });
      if (hits.length) {
        sh.getRange(hits[0], idx.target).setValue(val);
        for (let k = 1; k < hits.length; k++) sh.getRange(hits[k], idx.target).setValue('');   // 重複行は空にして二重計上を防ぐ
      } else if (val !== '') {
        const row = sh.getLastRow() + 1;
        sh.getRange(row, idx.month).setNumberFormat('@').setValue(month);
        sh.getRange(row, idx.metric).setValue(metric);
        sh.getRange(row, idx.target).setValue(val);
      }
    });
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 *  目標スプシの「KPI実績」タブへ、月ごとの現状の数字と転換率を書き出す（2026-09-22 追加）
 *   転換率＝ひとつ前の段階の現状に対する割合
 *   （LINE追加÷応募、アポ日程切÷LINE追加、アポ予定÷アポ日程切、アポ済÷アポ予定、YES÷アポ済、採用÷YES）
 *   サイト → スプシの一方通行。スプシ側で直してもサイトには戻らない。
 * ============================================================ */
const KPI_ACTUAL_SHEET_NAME = 'KPI実績';
const KPI_PREV = {
  'LINE追加数': '応募数',
  'アポ日程切数': 'LINE追加数',
  'アポ予定数（予約）': 'アポ日程切数',
  'アポ済数（面談完了）': 'アポ予定数（予約）',
  'YES数': 'アポ済数（面談完了）',
  '採用数': 'YES数'
};
const KPI_RATE_HEADERS = ['LINE追加率', 'アポ日程切率', 'アポ予定率', 'アポ済率', 'YES率', '採用率'];

/** 転換率（%・小数第1位）。ひとつ前の段階が0または無いときは null */
function convRate_(counts, metric) {
  const prev = KPI_PREV[metric];
  if (!prev) return null;
  const d = Number(counts[prev]) || 0;
  if (!d) return null;
  return Math.round((Number(counts[metric]) || 0) / d * 1000) / 10;
}

function writeKpiActuals_(spreadsheetId, month, counts) {
  const targetSs = SpreadsheetApp.openById(spreadsheetId);
  let sh = targetSs.getSheetByName(KPI_ACTUAL_SHEET_NAME);
  const headers = ['対象月'].concat(KPI_RECRUITER_METRICS, KPI_RATE_HEADERS, ['更新日時']);
  const allZero = KPI_RECRUITER_METRICS.every(function (m) { return !Number(counts[m]); });

  let row = -1;
  if (sh && sh.getLastRow() >= 2) {
    const months = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < months.length; i++) {
      if (normYm_(months[i][0]) === month) { row = i + 2; break; }
    }
  }
  if (row < 0 && allZero) return;   // まだ何も無い月は行を作らない

  if (!sh) sh = targetSs.insertSheet(KPI_ACTUAL_SHEET_NAME);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (row < 0) row = sh.getLastRow() + 1;

  const rates = KPI_RECRUITER_METRICS.filter(function (m) { return !!KPI_PREV[m]; })
    .map(function (m) { const r = convRate_(counts, m); return r == null ? '' : r; });
  const vals = [month]
    .concat(KPI_RECRUITER_METRICS.map(function (m) { return Number(counts[m]) || 0; }))
    .concat(rates)
    .concat([new Date()]);
  sh.getRange(row, 1).setNumberFormat('@');
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  sh.getRange(row, 1 + KPI_RECRUITER_METRICS.length + 1, 1, rates.length).setNumberFormat('0.0"%"');
  sh.getRange(row, vals.length).setNumberFormat('yyyy-mm-dd hh:mm');
}

/**
 * 数字管理ダッシュボードのメイン関数。
 * opts: {mode:'month'|'quarter', year, month, quarter}
 * 戻り値：
 *   採用者 → { months, role, own:{label,counts}, targets, hasTargetSheet, metrics, team:[{氏名,counts}] }
 *   運用者 → { months, role, own:{label,counts}, metrics }
 */
function getKpiDashboard(token, opts) {
  const user = requireUser_(token);
  const months = monthsForRequest_(opts);

  const funnel = computeFunnelCounts_(months);
  const events = computeEventCounts_(months);

  function operatorCounts_(accountNames) {
    const counts = { 応募数: 0, LINE追加数: 0, アポ日程切数: 0 };
    accountNames.forEach(function (acct) {
      const c = funnel.byAccount[acct];
      if (!c) return;
      counts.応募数 += c.応募数;
      counts.LINE追加数 += c.LINE追加数;
      counts.アポ日程切数 += c.アポ日程切数;
    });
    return counts;
  }

  const result = { months: months, role: user.役割 };

  if (user.役割 === '採用者') {
    const ownCounts = kpiBlankCounts_();
    ownCounts.応募数 = funnel.byOwnerApplications[user.表示担当者名] || 0;
    ownCounts['アポ予定数（予約）'] = funnel.byOwnerMeetings[user.表示担当者名] || 0;
    const ev = events.byOwner[user.表示担当者名] || {};
    ownCounts.LINE追加数 = ev.LINE追加数 || 0;
    ownCounts.アポ日程切数 = ev.アポ日程切数 || 0;
    ownCounts['アポ済数（面談完了）'] = ev['アポ済数（面談完了）'] || 0;
    ownCounts.YES数 = ev.YES数 || 0;
    ownCounts.採用数 = ev.採用数 || 0;
    result.own = { label: user.表示担当者名, counts: ownCounts };

    const targetSheetId = ownTargetSheetId_(user);
    const byMonth = readKpiTargetsByMonth_(targetSheetId, months);
    result.targets = readKpiTargets_(targetSheetId, months);
    result.targetsToDate = kpiTargetsToDate_(byMonth, months);   // 今日までの目標（月の日数で日割り）
    result.hasTargetSheet = !!targetSheetId;
    result.metrics = KPI_RECRUITER_METRICS;

    // 月次表示のときは、目標スプシの「KPI実績」タブへ現状の数字と転換率を書き出す（四半期表示では書かない）
    result.actualsSync = null;
    if (targetSheetId && opts && opts.mode !== 'quarter') {
      try {
        writeKpiActuals_(targetSheetId, months[0], ownCounts);
        result.actualsSync = { ok: true };
      } catch (e) {
        result.actualsSync = { ok: false, message: '目標スプシの「' + KPI_ACTUAL_SHEET_NAME + '」タブに書き込めませんでした（編集権限を確認してください）。' };
      }
    }

    const accounts = getAccounts_();
    result.team = getMembers_()
      .filter(function (m) {
        return String(m.チーム).trim() === String(user.チーム).trim()
          && String(m.役割).trim() === '運用者' && truthy_(m.有効);
      })
      .map(function (m) {
        const myAccts = accounts
          .filter(function (a) { return String(a.userId).trim() === String(m.userId).trim() && truthy_(a.有効); })
          .map(function (a) { return String(a.アカウント名).trim(); });
        return { 氏名: m.氏名, counts: operatorCounts_(myAccts) };
      });
  } else {
    const myAccts = getAccounts_()
      .filter(function (a) { return String(a.userId).trim() === user.userId && truthy_(a.有効); })
      .map(function (a) { return String(a.アカウント名).trim(); });
    result.own = { label: user.表示担当者名, counts: operatorCounts_(myAccts) };
    result.metrics = KPI_OPERATOR_METRICS;
  }

  return result;
}

/** 採用者が自分の目標スプシIDを登録・変更する（未設定なら列ごと自動作成） */
function setMyTargetSheetId(token, spreadsheetId) {
  const user = requireUser_(token);
  if (user.役割 !== '採用者') throw new Error('採用者のみ設定できます。');
  // URLをそのまま貼っても、https://docs.google.com/spreadsheets/d/【ID】/edit の【ID】部分だけを取り出す
  const raw = String(spreadsheetId || '').trim();
  const urlHit = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const id = urlHit ? urlHit[1] : raw;
  if (id) {
    try { SpreadsheetApp.openById(id); } catch (e) {
      throw new Error('スプレッドシートを開けません。IDが正しいか、このシステムのGoogleアカウントに共有されているか（編集権限）を確認してください。');
    }
  }
  const mSh = ss_().getSheetByName('メンバーマスタ');
  const members = sheetObjects_('メンバーマスタ');
  const rec = members.find(function (m) { return String(m.userId).trim() === user.userId; });
  if (!rec) throw new Error('メンバー情報が見つかりません。');

  const headers = mSh.getRange(1, 1, 1, mSh.getLastColumn()).getValues()[0];
  let col = headers.indexOf('目標スプシID') + 1;
  if (col <= 0) {
    col = mSh.getLastColumn() + 1;
    mSh.getRange(1, col).setValue('目標スプシID');
  }
  mSh.getRange(rec._row, col).setValue(id);
  return { ok: true, id: id };
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
