/**
 * 棚ページ バックエンド (GAS)
 * Slice 1〜5: シート初期化 / getShelf / report / 棚キー認証+承認フロー / 在庫CRUD+メッセージ編集
 * +初回セットアップ動線: createShelf（shelf_config が空のときだけ棚を作れる）
 *
 * 設計原則:
 * - 購入者の個人情報（IP, UA, Cookie 等）は一切記録しない
 * - reward_message は getShelf では返さない（report のレスポンスでのみ返す）
 * - 報告は在庫を直接減らさず reports キューに積み、棚主の承認で sold になる
 * - 棚キーは平文で保存しない（SHA-256ハッシュのみ。§4.4）
 */

const SHEETS = {
  BOOKS: 'books',
  REPORTS: 'reports',
  CONFIG: 'shelf_config',
};

/* ============================================================
 * 初期セットアップ
 * スプレッドシートに紐づくApps Scriptとして作成し、
 * setup() をエディタから一度だけ実行する（シート3枚をヘッダのみで生成）
 *
 * 棚の作り方は2通り:
 * - 間借り一人型/個人型: admin.html を開くと「棚をつくる」画面になる
 *   （shelf_config が空のあいだだけ。棚ID・棚名・棚キーを決めれば完了）
 * - 多テナント型: 運営者が shelf_config に行を追加する
 *   （shelf_id=既存の棚番号・記号、shelf_key_hash は空のまま →
 *     棚主が admin.html?s=棚ID を開いて棚キーを設定）
 *
 * 見本データで動きを眺めたいときは setupDemo() を実行（shelf-2f が生える）
 * ============================================================ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheet_(ss, SHEETS.BOOKS, [
    'book_id','shelf_id','title','author','price',
    'owner_comment','reward_message','status','added_at','sold_at'
  ], []);

  createSheet_(ss, SHEETS.REPORTS, [
    'report_id','shelf_id','book_id','buyer_comment','reported_at','status'
  ], []);

  createSheet_(ss, SHEETS.CONFIG, [
    'shelf_id','shelf_name','owner_intro','owner_message','default_rewards',
    'sns_note','sns_instagram','sns_x','theme','palette','bg_image','shelf_key_hash'
  ], []);
}

// 見本の棚（shelf-2f）と3冊を追加する。試したあとは行ごと消してよい
function setupDemo() {
  setup();
  if (findRow_(SHEETS.CONFIG, r => r.shelf_id === 'shelf-2f')) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(SHEETS.CONFIG).appendRow(
    ['shelf-2f','ことづて棚',
     '手放したくないのに手放す本ばかり置いています。どれも二度読みました。',
     '先週『銀河鉄道の夜』を買ってくださった方へ。ジョバンニの切符の話、私も同じところで泣きました。よい旅を。',
     'この本を選んでくださって、ありがとう。よい読書の時間を。\nその本、棚のいちばん端で、ずっと誰かを待っていました。\n読み終わったら、次の誰かに手渡してあげてください。',
     'https://note.com/','','','bunko','sukkiri','','']);

  const books = ss.getSheetByName(SHEETS.BOOKS);
  [
    ['b001','shelf-2f','銀河鉄道の夜','宮沢賢治',600,
     '何度目かの読み返しで、カムパネルラの気持ちがやっとわかった気がしました。',
     'ほんとうのさいわいを、あなたも探しに行くのですね。','on_shelf',new Date(),''],
    ['b002','shelf-2f','檸檬','梶井基次郎',550,
     '丸善に檸檬を置いて帰る勇気を、私はまだ持っていません。',
     '','on_shelf',new Date(),''],
    ['b003','shelf-2f','方丈記','鴨長明',600,
     'ゆく河の流れは絶えずして。ミニマリストの元祖。',
     '','on_shelf',new Date(),''],
  ].forEach(r => books.appendRow(r));
}

function createSheet_(ss, name, headers, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    rows.forEach(r => sh.appendRow(r));
    sh.setFrozenRows(1);
  }
}

/* ============================================================
 * ルーティング
 * GET  : ?action=getShelf&shelf=shelf-2f
 * POST : Content-Type text/plain で JSON 文字列を送る（CORS回避）
 *        { action:'report', shelfId, bookId, comment }
 * ============================================================ */
function doGet(e) {
  try {
    const action = (e.parameter.action || '');
    if (action === 'getShelf') return json_(getShelf_(e.parameter.shelf));
    return json_({ ok:false, error:'unknown action' });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'report')        return json_(report_(req));
    if (req.action === 'createShelf')   return json_(createShelf_(req));
    if (req.action === 'setKey')        return json_(setKey_(req));
    if (req.action === 'getAdmin')      return json_(getAdmin_(req));
    if (req.action === 'approveReport') return json_(approveReport_(req));
    if (req.action === 'rejectReport')  return json_(rejectReport_(req));
    if (req.action === 'upsertBook')    return json_(upsertBook_(req));
    if (req.action === 'removeBook')    return json_(removeBook_(req));
    if (req.action === 'updateMessage') return json_(updateMessage_(req));
    return json_({ ok:false, error:'unknown action' });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 * getShelf: 棚情報＋陳列中の本
 * 注意: reward_message は絶対に含めない
 * ============================================================ */
function getShelf_(shelfId) {
  if (!shelfId) return { ok:false, error:'shelf required' };

  const cfg = findRow_(SHEETS.CONFIG, r => r.shelf_id === shelfId);
  if (!cfg) return { ok:false, error:'shelf not found' };

  const books = readRows_(SHEETS.BOOKS)
    .filter(r => r.shelf_id === shelfId && r.status === 'on_shelf')
    .map(r => ({
      id: r.book_id,
      title: r.title,
      author: r.author,
      price: Number(r.price) || 0,
      note: r.owner_comment,
    }));

  const sns = [];
  if (cfg.sns_note)      sns.push({ label:'note',      url:cfg.sns_note });
  if (cfg.sns_instagram) sns.push({ label:'Instagram', url:cfg.sns_instagram });
  if (cfg.sns_x)         sns.push({ label:'X',         url:cfg.sns_x });

  return { ok:true, data:{
    shelf: {
      name: cfg.shelf_name,
      intro: cfg.owner_intro,
      message: cfg.owner_message,
      theme: cfg.theme || 'bunko',
      palette: cfg.palette || 'sukkiri',
      bg: cfg.bg_image || '',
      sns: sns,
    },
    books: books,
  }};
}

/* ============================================================
 * report: 匿名の購入報告 → キューに積み、返礼を返す
 * ============================================================ */
function report_(req) {
  const shelfId = String(req.shelfId || '');
  const bookId  = String(req.bookId || '');
  let comment   = String(req.comment || '').slice(0, 200);

  const book = findRow_(SHEETS.BOOKS,
    r => r.book_id === bookId && r.shelf_id === shelfId && r.status === 'on_shelf');
  if (!book) return { ok:false, error:'book not found' };

  // URLを含む一言は保留フラグ（棚主承認画面で確認）
  const status = /https?:\/\//i.test(comment) ? 'held' : 'pending';

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    // 同一 book_id の pending/held 重複は積まない（返礼は返す）
    const dup = findRow_(SHEETS.REPORTS,
      r => r.book_id === bookId && (r.status === 'pending' || r.status === 'held'));
    if (!dup) {
      SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEETS.REPORTS)
        .appendRow([Utilities.getUuid(), shelfId, bookId, comment, new Date(), status]);
    }
  } finally {
    lock.releaseLock();
  }

  // 返礼: 本ごとの reward_message → なければ棚共通からランダム
  let reward = String(book.reward_message || '').trim();
  if (!reward) {
    const cfg = findRow_(SHEETS.CONFIG, r => r.shelf_id === shelfId);
    const pool = String(cfg && cfg.default_rewards || '')
      .split('\n').map(s => s.trim()).filter(Boolean);
    reward = pool.length
      ? pool[Math.floor(Math.random() * pool.length)]
      : 'ご報告ありがとうございました。よい読書の時間を。';
  }
  return { ok:true, data:{ reward: reward } };
}

/* ============================================================
 * Slice 3: 棚キー認証
 * - shelf_key_hash には SHA-256(shelf_id + ':' + 棚キー) の16進を保持
 * - 初期状態は未設定（空文字）。setKey は未設定のときだけ受理（§4.4）
 * - リセットは運営者がシート上で shelf_key_hash を空にする → 再設定
 * ============================================================ */
function hashKey_(shelfId, key) {
  return Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      shelfId + ':' + key,
      Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0'))
    .join('');
}

// 認証エラーの error 値: 'shelf not found' / 'no_key'（未設定→初回設定へ誘導） / 'auth'（キー不一致）
function auth_(shelfId, key) {
  const cfg = findRow_(SHEETS.CONFIG, r => r.shelf_id === shelfId);
  if (!cfg) return { ok:false, error:'shelf not found' };
  const stored = String(cfg.shelf_key_hash || '').trim();
  if (!stored) return { ok:false, error:'no_key' };
  if (hashKey_(shelfId, String(key || '')) !== stored) return { ok:false, error:'auth' };
  return { ok:true, cfg };
}

/* ============================================================
 * createShelf: 初回セットアップ（間借り一人型・個人型の動線）
 * - shelf_config に1行も棚がないときだけ受理する
 *   （多テナント型は運営者がシートで行を割り当てるので、ここは通らない。
 *     公開URLを知る第三者が棚を勝手に作ることもできない）
 * - 2つ目以降の棚（ZINE棚など）は運営者としてシートに行を追加する
 * - 棚IDはURLの ?s= に入るため半角英数字・ハイフン・アンダースコアに限定
 * ============================================================ */
function createShelf_(req) {
  const shelfId = String(req.shelfId || '').trim();
  const name    = String(req.shelfName || '').trim().slice(0, 50);
  const intro   = String(req.intro || '').trim().slice(0, 300);
  const newKey  = String(req.newKey || '');

  if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,31}$/.test(shelfId))
    return { ok:false, error:'bad shelf id' };
  if (!name) return { ok:false, error:'name required' };
  if (newKey.length < 8) return { ok:false, error:'key too short' };

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    if (readRows_(SHEETS.CONFIG).length)
      return { ok:false, error:'create not allowed' };
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG)
      .appendRow([shelfId, name, intro, '', '', '', '', '',
                  'bunko', 'sukkiri', '', hashKey_(shelfId, newKey)]);
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

function setKey_(req) {
  const shelfId = String(req.shelfId || '');
  const newKey  = String(req.newKey || '');
  if (newKey.length < 8) return { ok:false, error:'key too short' };

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const cfg = findRow_(SHEETS.CONFIG, r => r.shelf_id === shelfId);
    if (!cfg) return { ok:false, error:'shelf not found' };
    if (String(cfg.shelf_key_hash || '').trim()) return { ok:false, error:'key already set' };
    updateRow_(SHEETS.CONFIG, r => r.shelf_id === shelfId,
      { shelf_key_hash: hashKey_(shelfId, newKey) });
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * Slice 3: 管理データ取得＋承認フロー
 * ============================================================ */
function getAdmin_(req) {
  const a = auth_(req.shelfId, req.key);
  if (!a.ok) {
    // 棚が1つもない＝初回セットアップ前。admin.html を「棚をつくる」へ誘導する
    if (a.error === 'shelf not found' && !readRows_(SHEETS.CONFIG).length)
      return { ok:false, error:'shelf not found', canCreate:true };
    return a;
  }

  const books = readRows_(SHEETS.BOOKS)
    .filter(r => r.shelf_id === req.shelfId)
    .map(r => ({
      id: r.book_id,
      title: r.title,
      author: r.author,
      price: Number(r.price) || 0,
      note: r.owner_comment,
      reward: r.reward_message,
      status: r.status,
      soldAt: r.sold_at || '',
    }));

  const reports = readRows_(SHEETS.REPORTS)
    .filter(r => r.shelf_id === req.shelfId &&
                 (r.status === 'pending' || r.status === 'held'))
    .map(r => ({
      id: r.report_id,
      bookId: r.book_id,
      comment: r.buyer_comment,
      at: r.reported_at,
      status: r.status,
    }));

  return { ok:true, data:{
    books: books,
    reports: reports,
    config: {
      message: String(a.cfg.owner_message || ''),
      rewards: String(a.cfg.default_rewards || ''),
    },
  }};
}

// 承認: report を approved に、該当の本を sold に（sold_at 記録）
function approveReport_(req) {
  const a = auth_(req.shelfId, req.key);
  if (!a.ok) return a;

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const rep = updateRow_(SHEETS.REPORTS,
      r => r.report_id === req.reportId && r.shelf_id === req.shelfId &&
           (r.status === 'pending' || r.status === 'held'),
      { status: 'approved' });
    if (!rep) return { ok:false, error:'report not found' };

    updateRow_(SHEETS.BOOKS,
      r => r.book_id === rep.book_id && r.shelf_id === req.shelfId &&
           r.status === 'on_shelf',
      { status: 'sold', sold_at: new Date() });

    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

function rejectReport_(req) {
  const a = auth_(req.shelfId, req.key);
  if (!a.ok) return a;

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const rep = updateRow_(SHEETS.REPORTS,
      r => r.report_id === req.reportId && r.shelf_id === req.shelfId &&
           (r.status === 'pending' || r.status === 'held'),
      { status: 'rejected' });
    if (!rep) return { ok:false, error:'report not found' };
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * Slice 4: 在庫CRUD＋「棚主より」編集
 * ============================================================ */
// 追加（book.id なし）or 編集（book.id あり）。status:'on_shelf' を渡すと
// 販売済の本を棚に戻せる（誤承認の訂正用。sold_at はクリア）
function upsertBook_(req) {
  const a = auth_(req.shelfId, req.key);
  if (!a.ok) return a;

  const b = req.book || {};
  const title = String(b.title || '').trim();
  if (!title) return { ok:false, error:'title required' };
  const fields = {
    title: title,
    author: String(b.author || '').trim().slice(0, 100),
    price: Number(b.price) || 0,
    owner_comment: String(b.note || '').slice(0, 500),
    reward_message: String(b.reward || '').slice(0, 500),
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    if (b.id) {
      const updates = Object.assign({}, fields);
      if (b.status === 'on_shelf') {
        updates.status = 'on_shelf';
        updates.sold_at = '';
      }
      const row = updateRow_(SHEETS.BOOKS,
        r => r.book_id === b.id && r.shelf_id === req.shelfId && r.status !== 'removed',
        updates);
      if (!row) return { ok:false, error:'book not found' };
      return { ok:true, data:{ id: b.id } };
    }
    const id = 'b' + Date.now().toString(36);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.BOOKS)
      .appendRow([id, req.shelfId, fields.title, fields.author, fields.price,
                  fields.owner_comment, fields.reward_message, 'on_shelf', new Date(), '']);
    return { ok:true, data:{ id: id } };
  } finally {
    lock.releaseLock();
  }
}

function removeBook_(req) {
  const a = auth_(req.shelfId, req.key);
  if (!a.ok) return a;

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const row = updateRow_(SHEETS.BOOKS,
      r => r.book_id === String(req.bookId || '') && r.shelf_id === req.shelfId &&
           r.status !== 'removed',
      { status: 'removed' });
    if (!row) return { ok:false, error:'book not found' };
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

// 「棚主より」(owner_message) と共通返礼 (default_rewards) の更新
function updateMessage_(req) {
  const a = auth_(req.shelfId, req.key);
  if (!a.ok) return a;

  const updates = {};
  if ('message' in req) updates.owner_message = String(req.message || '').slice(0, 1000);
  if ('rewards' in req) updates.default_rewards = String(req.rewards || '').slice(0, 2000);
  if (!Object.keys(updates).length) return { ok:false, error:'nothing to update' };

  updateRow_(SHEETS.CONFIG, r => r.shelf_id === req.shelfId, updates);
  return { ok:true };
}

/* ============================================================
 * シート読み取りユーティリティ（ヘッダ行→オブジェクト配列）
 * ============================================================ */
function readRows_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  return values.map(row => {
    const o = {};
    head.forEach((h, i) => o[h] = row[i]);
    return o;
  });
}

function findRow_(sheetName, pred) {
  return readRows_(sheetName).find(pred) || null;
}

// pred に合う最初の行の指定列だけ更新し、更新前の行オブジェクトを返す
function updateRow_(sheetName, pred, updates) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const head = values[0];
  for (let i = 1; i < values.length; i++) {
    const o = {};
    head.forEach((h, j) => o[h] = values[i][j]);
    if (pred(o)) {
      Object.keys(updates).forEach(k => {
        const col = head.indexOf(k);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(updates[k]);
      });
      return o;
    }
  }
  return null;
}
