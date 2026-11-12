/**
 * OnePercent - may chu ban quyen (Google Apps Script, gan voi mot Google Sheet).
 *
 * Add-in gui POST JSON { action, ... } toi URL web app (/exec) va nhan JSON tra ve.
 * Dich vu bao bien dong so du (SePay) goi cung URL do voi ?hook=sepay&key=<WEBHOOK_KEY>.
 *
 * Bi mat nam trong Script Properties (Project Settings > Script properties), KHONG nam trong Sheet:
 *   PRIVATE_KEY  - noi dung file private_key.pem (dan nguyen van, xuong dong hay khong deu duoc)
 *   WEBHOOK_KEY  - chuoi bi mat cho webhook; ham setup() tu tao neu chua co
 *
 * Cac sheet: Config, Plans, Users, Machines, Orders, Codes, BankLog, Usage, ToolStats - ham setup() tao san.
 * Usage = so luot dung tung tool cua tung tai khoan (add-in gui len); ToolStats = tong theo tool, tu tinh lai.
 * Cach cai dat: xem HUONG_DAN.md.
 */

// ---------------------------------------------------------------- Sheets ----

var T_CONFIG = { name: 'Config', head: ['Key', 'Value', 'Ghi chú'] };
var T_PLANS = { name: 'Plans', head: ['Id', 'Tên (VI)', 'Name (EN)', 'Months', 'Price', 'Active'] };
var T_USERS = { name: 'Users', head: ['Email', 'Name', 'Plan', 'Expires', 'Created', 'LastSeen', 'Blocked', 'Note'] };
var T_MACHINES = { name: 'Machines', head: ['Email', 'MachineId', 'MachineName', 'Revit', 'AppVersion', 'SessionHash', 'FirstSeen', 'LastSeen'] };
var T_ORDERS = { name: 'Orders', head: ['Code', 'Email', 'PlanId', 'Amount', 'Status', 'Created', 'PaidAt', 'TxId', 'Note'] };
var T_CODES = { name: 'Codes', head: ['Code', 'PlanId', 'Created', 'UsedBy', 'UsedAt'] };
var T_BANK = { name: 'BankLog', head: ['TxId', 'Time', 'Amount', 'Content', 'Order', 'Result'] };
var T_USAGE = { name: 'Usage', head: ['Email', 'Tool', 'Count', 'FirstUsed', 'LastUsed'] };
var T_STATS = { name: 'ToolStats', head: ['Tool', 'Total', 'Users', 'LastUsed'] };

// Vi tri cot (0-based) trong tung sheet.
var G = { email: 0, tool: 1, count: 2, first: 3, last: 4 };
var MAX_USAGE_TOOLS = 400;          // so tool toi da trong mot lan gui
var MAX_USAGE_COUNT = 100000;       // so luot toi da cho mot tool trong mot lan gui
var U = { email: 0, name: 1, plan: 2, expires: 3, created: 4, lastSeen: 5, blocked: 6, note: 7 };
var M = { email: 0, machine: 1, name: 2, revit: 3, app: 4, session: 5, first: 6, last: 7 };
var O = { code: 0, email: 1, plan: 2, amount: 3, status: 4, created: 5, paidAt: 6, tx: 7, note: 8 };
var C = { code: 0, plan: 1, created: 2, usedBy: 3, usedAt: 4 };
var P = { id: 0, name: 1, nameEn: 2, months: 3, price: 4, active: 5 };

var LIFETIME = 'LIFETIME';
var LIFETIME_SECONDS = 4102444800;                    // 2100-01-01, chi de token co mot con so
var ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';    // bo I, O, 0, 1 cho de doc
var ORDER_CODE_LENGTH = 6;

var DEFAULT_CONFIG = [
  ['ProductName', 'OnePercent', 'Tên sản phẩm'],
  ['GoogleClientId', '', 'Client ID của OAuth client loại Desktop (giống trong LicenseConfig.cs)'],
  ['TrialDays', 30, 'Số ngày dùng thử cho tài khoản mới'],
  ['MaxMachines', 2, 'Số máy tối đa cho mỗi tài khoản'],
  ['OfflineDays', 7, 'Số ngày dùng được khi không có mạng'],
  ['SessionDays', 90, 'Máy không kết nối quá số ngày này thì phải đăng nhập lại'],
  ['OrderPrefix', 'OP', 'Tiền tố nội dung chuyển khoản (chỉ chữ in hoa, không dấu)'],
  ['BankBin', '', 'Mã BIN ngân hàng cho VietQR, vd Techcombank 970407, Vietcombank 970436'],
  ['BankName', '', 'Tên ngân hàng hiển thị cho khách'],
  ['AccountNo', '', 'Số tài khoản nhận tiền'],
  ['AccountName', '', 'Tên chủ tài khoản (in hoa, không dấu)'],
  ['SupportZalo', '', 'Số Zalo hoặc link zalo.me'],
  ['SupportFacebook', '', 'Link Facebook (https://...)'],
  ['LatestVersion', '', 'Phiên bản mới nhất, vd 1.1.0 (để trống nếu không cần báo cập nhật)'],
  ['DownloadUrl', '', 'Link tải bản mới (https://...)']
];

var DEFAULT_PLANS = [
  ['M1', 'Gói 1 tháng', '1 month', 1, 50000, true],
  ['M6', 'Gói 6 tháng', '6 months', 6, 250000, true],
  ['M12', 'Gói 12 tháng', '12 months', 12, 450000, true],
  ['LIFE', 'Gói vĩnh viễn', 'Lifetime', 0, 1000000, true]
];

// ------------------------------------------------------------ Entry points ----

function doGet() {
  return ContentService.createTextOutput('OnePercent license server: OK');
}

function doPost(e) {
  var out;
  try {
    var query = (e && e.parameter) || {};
    if (query.hook === 'sepay') {
      out = handleBankHook_(e, query);
    } else {
      out = route_(JSON.parse(e.postData.contents));
    }
  } catch (err) {
    out = { ok: false, code: 'server', error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  if (!req || typeof req !== 'object') return fail_('server', 'Bad request');

  var action = String(req.action || '');
  if (action === 'info') return publicInfo_({ ok: true });

  // Moi thu con lai deu doc-roi-ghi, nen chay lan luot tung yeu cau mot.
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    switch (action) {
      case 'login': return login_(req);
      case 'refresh': return refresh_(req);
      case 'order': return order_(req);
      case 'redeem': return redeem_(req);
      case 'logout': return logout_(req);
      case 'usage': return usage_(req);
      default: return fail_('server', 'Unknown action: ' + action);
    }
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------------------- Actions ----

/** Dang nhap: xac nhan ID token voi Google, tao tai khoan neu moi, gan may, cap phien. */
function login_(req) {
  var cfg = config_();
  if (!cfg.GoogleClientId) return fail_('config', 'GoogleClientId is empty');

  var who = verifyGoogle_(String(req.idToken || ''), String(cfg.GoogleClientId));
  if (!who) return fail_('google', 'Google token rejected');

  var machineId = cleanId_(req.machine);
  if (!machineId) return fail_('server', 'Missing machine id');

  var now = new Date();
  var users = table_(T_USERS);
  var machines = table_(T_MACHINES);

  var u = findRow_(users, U.email, who.email);
  if (u < 0) {
    // May nay da tung dung thu bang tai khoan khac thi tai khoan moi khong co them ngay thu.
    var machineKnown = machines.rows.some(function (r) { return String(r[M.machine]) === machineId; });
    var trialDays = machineKnown ? 0 : Number(cfg.TrialDays) || 0;
    var row = [];
    row[U.email] = who.email;
    row[U.name] = who.name;
    row[U.plan] = 'TRIAL';
    row[U.expires] = new Date(now.getTime() + trialDays * 86400000);
    row[U.created] = now;
    row[U.lastSeen] = now;
    row[U.blocked] = false;
    row[U.note] = machineKnown ? 'Máy đã dùng thử bằng tài khoản khác' : '';
    u = appendRow_(users, row);
  } else {
    if (isTrue_(users.rows[u][U.blocked])) return fail_('blocked', 'Account blocked');
    users.rows[u][U.name] = who.name || users.rows[u][U.name];
    users.rows[u][U.lastSeen] = now;
    writeRow_(users, u);
  }

  var session = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var sessionHash = sha256Hex_(session);

  var m = -1, count = 0;
  for (var i = 0; i < machines.rows.length; i++) {
    if (sameEmail_(machines.rows[i][M.email], who.email)) {
      count++;
      if (String(machines.rows[i][M.machine]) === machineId) m = i;
    }
  }

  if (m < 0) {
    var max = Number(cfg.MaxMachines) || 1;
    if (count >= max) return fail_('machines', 'Machine limit reached (' + max + ')');
    var mrow = [];
    mrow[M.email] = who.email;
    mrow[M.machine] = machineId;
    mrow[M.name] = short_(req.machineName, 60);
    mrow[M.revit] = short_(req.revit, 60);
    mrow[M.app] = short_(req.app, 20);
    mrow[M.session] = sessionHash;
    mrow[M.first] = now;
    mrow[M.last] = now;
    m = appendRow_(machines, mrow);
    count++;
  } else {
    machines.rows[m][M.name] = short_(req.machineName, 60);
    machines.rows[m][M.revit] = short_(req.revit, 60);
    machines.rows[m][M.app] = short_(req.app, 20);
    machines.rows[m][M.session] = sessionHash;
    machines.rows[m][M.last] = now;
    writeRow_(machines, m);
  }

  var out = licenceReply_(cfg, users.rows[u], machineId, count);
  out.session = session;
  return out;
}

/** Gia han giay phep; kem ma don thi tra them trang thai don do. */
function refresh_(req) {
  var cfg = config_();
  var ctx = session_(req);
  if (ctx.error) return ctx.error;

  var now = new Date();
  ctx.users.rows[ctx.u][U.lastSeen] = now;
  writeRow_(ctx.users, ctx.u);
  ctx.machines.rows[ctx.m][M.last] = now;
  ctx.machines.rows[ctx.m][M.revit] = short_(req.revit, 60);
  ctx.machines.rows[ctx.m][M.app] = short_(req.app, 20);
  writeRow_(ctx.machines, ctx.m);

  var out = licenceReply_(cfg, ctx.users.rows[ctx.u], ctx.machineId, ctx.machineCount);

  var code = String(req.order || '').toUpperCase().trim();
  if (code) {
    var orders = table_(T_ORDERS);
    var o = findRow_(orders, O.code, code);
    if (o >= 0 && sameEmail_(orders.rows[o][O.email], ctx.email)) {
      out.order = orderReply_(cfg, orders.rows[o]);
    }
  }
  return out;
}

/** Mo don mua mot goi: tra ve so tien, noi dung chuyen khoan va link anh QR. */
function order_(req) {
  var cfg = config_();
  var ctx = session_(req);
  if (ctx.error) return ctx.error;

  if (!cfg.AccountNo || !cfg.BankBin) return fail_('config', 'Bank account is not configured');

  var plan = findPlan_(String(req.plan || ''));
  if (!plan) return fail_('plan', 'Plan not available');

  var now = new Date();
  var orders = table_(T_ORDERS);

  // Don dang cho cua cung goi, trong 24 gio, voi dung gia hien tai thi dung lai - bam Mua nhieu lan khong de ra rac.
  var o = -1;
  for (var i = orders.rows.length - 1; i >= 0; i--) {
    var r = orders.rows[i];
    if (sameEmail_(r[O.email], ctx.email) && String(r[O.plan]) === plan.id && String(r[O.status]) === 'pending'
        && Number(r[O.amount]) === plan.price && now.getTime() - toMs_(r[O.created]) < 86400000) {
      o = i;
      break;
    }
  }

  if (o < 0) {
    var prefix = orderPrefix_(cfg);
    var code;
    do {
      code = prefix + randomCode_(ORDER_CODE_LENGTH);
    } while (findRow_(orders, O.code, code) >= 0);

    var row = [];
    row[O.code] = code;
    row[O.email] = ctx.email;
    row[O.plan] = plan.id;
    row[O.amount] = plan.price;
    row[O.status] = 'pending';
    row[O.created] = now;
    row[O.paidAt] = '';
    row[O.tx] = '';
    row[O.note] = '';
    o = appendRow_(orders, row);
  }

  var out = licenceReply_(cfg, ctx.users.rows[ctx.u], ctx.machineId, ctx.machineCount);
  out.order = orderReply_(cfg, orders.rows[o]);
  return out;
}

/** Dung mot ma kich hoat. */
function redeem_(req) {
  var cfg = config_();
  var ctx = session_(req);
  if (ctx.error) return ctx.error;

  var code = short_(req.code, 64).toUpperCase().replace(/[^A-Z0-9]/g, '');   // khach go co gach hay cach deu duoc
  var codes = table_(T_CODES);
  var c = code ? findRow_(codes, C.code, code) : -1;
  if (c < 0) return fail_('code_invalid', 'Unknown code');
  if (String(codes.rows[c][C.usedBy] || '') !== '') return fail_('code_used', 'Code already used');

  var plan = findPlan_(String(codes.rows[c][C.plan]), true);
  if (!plan) return fail_('code_invalid', 'Code points to a missing plan');

  applyPlan_(ctx.users, ctx.u, plan);
  codes.rows[c][C.usedBy] = ctx.email;
  codes.rows[c][C.usedAt] = new Date();
  writeRow_(codes, c);

  return licenceReply_(cfg, ctx.users.rows[ctx.u], ctx.machineId, ctx.machineCount);
}

/** Dang xuat: xoa phien cua may nay (may van giu cho trong tai khoan). */
function logout_(req) {
  var ctx = session_(req);
  if (ctx.error) return { ok: true };
  ctx.machines.rows[ctx.m][M.session] = '';
  writeRow_(ctx.machines, ctx.m);
  return { ok: true };
}

// ---------------------------------------------------------- Luot dung tool ----

/**
 * Add-in gui { tools: { "Ten tool": so luot, ... } } = so luot dung ke tu lan gui truoc.
 * Cong don vao sheet Usage (moi tai khoan + tool mot dong), roi tinh lai sheet ToolStats.
 */
function usage_(req) {
  var ctx = session_(req);
  if (ctx.error) return ctx.error;

  var tools = req.tools;
  if (!tools || typeof tools !== 'object' || Array.isArray(tools)) return fail_('server', 'Missing tools');

  ensureSheet_(T_USAGE);
  var usage = table_(T_USAGE);
  var now = new Date();

  var index = {};
  for (var i = 0; i < usage.rows.length; i++) index[usageKey_(usage.rows[i][G.email], usage.rows[i][G.tool])] = i;

  var oldCount = usage.rows.length;
  var first = -1, last = -1;                       // khoang dong cu bi sua
  var names = Object.keys(tools).slice(0, MAX_USAGE_TOOLS);
  for (var n = 0; n < names.length; n++) {
    var tool = short_(names[n], 60).replace(/\s+/g, ' ').trim();
    var count = Math.floor(Number(tools[names[n]]));
    if (!tool || !(count >= 1)) continue;
    count = Math.min(count, MAX_USAGE_COUNT);

    var key = usageKey_(ctx.email, tool);
    var r = index[key];
    if (r === undefined) {
      var row = [];
      row[G.email] = ctx.email;
      row[G.tool] = tool;
      row[G.count] = count;
      row[G.first] = now;
      row[G.last] = now;
      usage.rows.push(pad_(row, usage.width));
      index[key] = usage.rows.length - 1;
    } else {
      usage.rows[r][G.count] = (Number(usage.rows[r][G.count]) || 0) + count;
      usage.rows[r][G.last] = now;
      if (r < oldCount) {
        if (first < 0 || r < first) first = r;
        if (r > last) last = r;
      }
    }
  }

  // Ghi hai lan la cung: mot khoi cho cac dong cu bi sua, mot khoi cho cac dong moi.
  if (first >= 0) writeRows_(usage, first, last);
  if (usage.rows.length > oldCount) writeRows_(usage, oldCount, usage.rows.length - 1);

  rebuildToolStats_(usage);
  return { ok: true };
}

function usageKey_(email, tool) {
  return String(email).toLowerCase().trim() + '\n' + String(tool).toLowerCase().trim();
}

/** Tong hop theo tool: tong luot, so tai khoan da dung, lan dung gan nhat; nhieu luot nhat o tren. */
function rebuildToolStats_(usage) {
  var byTool = {};
  usage.rows.forEach(function (r) {
    var name = String(r[G.tool]).trim();
    if (!name) return;
    var key = name.toLowerCase();
    var s = byTool[key] || (byTool[key] = { name: name, total: 0, users: 0, last: 0 });
    s.total += Number(r[G.count]) || 0;
    s.users++;
    s.last = Math.max(s.last, toMs_(r[G.last]));
  });

  var rows = Object.keys(byTool).map(function (key) {
    var s = byTool[key];
    return pad_([s.name, s.total, s.users, s.last ? new Date(s.last) : ''], T_STATS.head.length);
  }).sort(function (a, b) { return b[1] - a[1] || String(a[0]).localeCompare(String(b[0])); });

  var sheet = ensureSheet_(T_STATS);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, T_STATS.head.length).clearContent();
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, T_STATS.head.length).setValues(rows);
}

/** Menu: tinh lai ToolStats sau khi sua / xoa tay trong sheet Usage. */
function refreshToolStats() {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    ensureSheet_(T_USAGE);
    rebuildToolStats_(table_(T_USAGE));
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.getUi().alert('Đã tính lại sheet ToolStats.');
}

// ------------------------------------------------------ Thanh toan tu dong ----

/**
 * SePay goi vao day moi khi tai khoan co tien vao. Tim ma don trong noi dung chuyen khoan,
 * du tien thi danh dau da tra va cong han dung. Goi lai nhieu lan cho cung giao dich khong sao.
 */
function handleBankHook_(e, query) {
  var secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY');
  if (!secret || String(query.key || '') !== secret) return { success: false, error: 'forbidden' };

  var tx = JSON.parse(e.postData.contents);
  if (String(tx.transferType || 'in') !== 'in') return { success: true, ignored: 'not incoming' };

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var cfg = config_();
    var txId = String(tx.id || tx.referenceCode || '');
    var amount = Number(tx.transferAmount) || 0;
    var content = String(tx.content || '') + ' ' + String(tx.code || '') + ' ' + String(tx.description || '');

    var bank = table_(T_BANK);
    if (txId && findRow_(bank, 0, txId) >= 0) return { success: true, duplicate: true };

    var result = settlePayment_(cfg, content, amount, txId);
    appendRow_(bank, [txId, tx.transactionDate || new Date(), amount, short_(content, 200), result.order, result.status]);
    return { success: true, result: result.status };
  } finally {
    lock.releaseLock();
  }
}

/** Tim don dang cho khop voi noi dung chuyen khoan va xu ly. Tra ve { order, status }. */
function settlePayment_(cfg, content, amount, txId) {
  var prefix = orderPrefix_(cfg);
  var text = String(content || '').toUpperCase();
  var orders = table_(T_ORDERS);
  var length = prefix.length + ORDER_CODE_LENGTH;

  // Quet moi vi tri co tien to, vi ngan hang hay chen them chu truoc/sau noi dung.
  for (var at = text.indexOf(prefix); at >= 0; at = text.indexOf(prefix, at + 1)) {
    var code = text.substr(at, length);
    if (code.length < length) break;

    var o = findRow_(orders, O.code, code);
    if (o < 0) continue;

    var row = orders.rows[o];
    if (String(row[O.status]) === 'paid') return { order: code, status: 'already_paid' };

    if (amount < Number(row[O.amount])) {
      row[O.status] = 'underpaid';
      row[O.tx] = txId;
      row[O.note] = 'Nhận ' + amount + ', thiếu ' + (Number(row[O.amount]) - amount);
      writeRow_(orders, o);
      return { order: code, status: 'underpaid' };
    }

    return { order: code, status: markPaid_(orders, o, txId) };
  }
  return { order: '', status: 'no_order' };
}

/** Danh dau don da tra va cong han dung cho nguoi mua. Tra ve 'paid' hoac ly do khong lam duoc. */
function markPaid_(orders, o, txId) {
  var row = orders.rows[o];
  var plan = findPlan_(String(row[O.plan]), true);
  if (!plan) return 'plan_missing';

  var users = table_(T_USERS);
  var u = findRow_(users, U.email, String(row[O.email]));
  if (u < 0) return 'user_missing';

  applyPlan_(users, u, plan);
  row[O.status] = 'paid';
  row[O.paidAt] = new Date();
  row[O.tx] = txId || row[O.tx] || 'manual';
  writeRow_(orders, o);
  return 'paid';
}

/** Cong goi vao han dung: tinh tu ngay het han hien tai neu con han, khong thi tu hom nay. */
function applyPlan_(users, u, plan) {
  var row = users.rows[u];
  if (String(row[U.expires]) === LIFETIME) {
    // Da vinh vien thi khong con gi de cong: giu nguyen.
  } else if (plan.months <= 0) {
    row[U.expires] = LIFETIME;
    row[U.plan] = plan.id;
  } else {
    var base = new Date(Math.max(Date.now(), toMs_(row[U.expires])));
    base.setMonth(base.getMonth() + plan.months);
    row[U.expires] = base;
    row[U.plan] = plan.id;
  }
  writeRow_(users, u);
}

// ----------------------------------------------------------------- Replies ----

function fail_(code, error) {
  return { ok: false, code: code, error: error };
}

/** Phan cong khai: bang gia, lien he, ban moi. */
function publicInfo_(out, cfg) {
  cfg = cfg || config_();
  out.trialDays = Number(cfg.TrialDays) || 0;
  out.plans = table_(T_PLANS).rows
    .filter(function (r) { return String(r[P.id]) !== '' && isTrue_(r[P.active]); })
    .map(function (r) {
      return { id: String(r[P.id]), name: String(r[P.name]), nameEn: String(r[P.nameEn]),
               months: Number(r[P.months]) || 0, price: Number(r[P.price]) || 0 };
    });
  out.support = { zalo: String(cfg.SupportZalo || ''), facebook: String(cfg.SupportFacebook || '') };
  out.update = { version: String(cfg.LatestVersion || ''), url: String(cfg.DownloadUrl || '') };
  return out;
}

/** Giay phep da ky cho mot tai khoan tren mot may, kem phan cong khai. */
function licenceReply_(cfg, userRow, machineId, machineCount) {
  var now = Math.floor(Date.now() / 1000);
  var lifetime = String(userRow[U.expires]) === LIFETIME;
  var plan = String(userRow[U.plan] || 'TRIAL');

  var payload = {
    v: 1,
    email: String(userRow[U.email]),
    name: String(userRow[U.name] || ''),
    machine: machineId,
    plan: plan,
    exp: lifetime ? LIFETIME_SECONDS : Math.floor(toMs_(userRow[U.expires]) / 1000),
    iat: now,
    offline: now + (Number(cfg.OfflineDays) || 7) * 86400,
    trial: plan === 'TRIAL' && !lifetime,
    life: lifetime
  };

  var out = publicInfo_({ ok: true }, cfg);
  out.token = sign_(payload);
  out.user = {
    email: payload.email,
    name: payload.name,
    machines: machineCount,
    maxMachines: Number(cfg.MaxMachines) || 1
  };
  return out;
}

function orderReply_(cfg, row) {
  var code = String(row[O.code]);
  var amount = Number(row[O.amount]) || 0;
  // Ban da ky cua thong tin chuyen khoan: add-in chi hien thi nhung gi nam trong nay, nen
  // khong ai chen giua duong truyen ma doi duoc so tai khoan, so tien hay ma don.
  var signed = sign_({
    t: 'order',
    code: code,
    planId: String(row[O.plan]),
    amount: amount,
    email: String(row[O.email]).toLowerCase().trim(),
    bankBin: String(cfg.BankBin || ''),
    bankName: String(cfg.BankName || ''),
    accountNo: String(cfg.AccountNo || ''),
    accountName: String(cfg.AccountName || '')
  });

  return {
    signed: signed,
    code: code,
    planId: String(row[O.plan]),
    amount: amount,
    status: String(row[O.status]),
    bankBin: String(cfg.BankBin || ''),     // add-in tu ve ma QR tu BIN + so tai khoan + so tien + ma don
    bankName: String(cfg.BankName || ''),
    accountNo: String(cfg.AccountNo || ''),
    accountName: String(cfg.AccountName || ''),
    qr: 'https://img.vietqr.io/image/' + encodeURIComponent(String(cfg.BankBin)) + '-'
      + encodeURIComponent(String(cfg.AccountNo)) + '-compact2.png?amount=' + amount
      + '&addInfo=' + encodeURIComponent(code)
      + '&accountName=' + encodeURIComponent(String(cfg.AccountName || ''))
  };
}

// ------------------------------------------------------- Phien va chu ky ----

/** Tim tai khoan + may tu phien gui len. Tra ve { error } neu phien khong con dung. */
function session_(req) {
  var session = String(req.session || '');
  var machineId = cleanId_(req.machine);
  if (!session || !machineId) return { error: fail_('session', 'No session') };

  var hash = sha256Hex_(session);
  var machines = table_(T_MACHINES);
  var m = -1;
  for (var i = 0; i < machines.rows.length; i++) {
    if (String(machines.rows[i][M.session]) === hash && String(machines.rows[i][M.machine]) === machineId) {
      m = i;
      break;
    }
  }
  if (m < 0) return { error: fail_('session', 'Session not found') };

  // Phien khong song mai: may im lang qua SessionDays thi xoa phien, khach dang nhap lai.
  // (Mot phien bi lo, hoac dang xuat luc mat mang, vi the cung tu het.)
  var sessionDays = Number(config_().SessionDays) || 90;
  if (Date.now() - toMs_(machines.rows[m][M.last]) > sessionDays * 86400000) {
    machines.rows[m][M.session] = '';
    writeRow_(machines, m);
    return { error: fail_('session', 'Session expired') };
  }

  var email = String(machines.rows[m][M.email]);
  var users = table_(T_USERS);
  var u = findRow_(users, U.email, email);
  if (u < 0) return { error: fail_('session', 'Account not found') };
  if (isTrue_(users.rows[u][U.blocked])) return { error: fail_('blocked', 'Account blocked') };

  var count = machines.rows.filter(function (r) { return sameEmail_(r[M.email], email); }).length;
  return { email: email, machineId: machineId, users: users, u: u, machines: machines, m: m, machineCount: count };
}

/** Hoi Google xem ID token co that khong va co dung la cap cho ung dung nay khong. */
function verifyGoogle_(idToken, clientId) {
  if (!idToken) return null;
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;

  var info = JSON.parse(res.getContentText());
  if (String(info.aud) !== clientId) return null;
  if (String(info.email_verified) !== 'true') return null;
  if (Number(info.exp) * 1000 < Date.now()) return null;
  if (!info.email) return null;

  return { email: String(info.email).toLowerCase().trim(), name: String(info.name || '') };
}

/** token = base64url(JSON) + "." + base64url(chu ky RSA-SHA256 cua phan dau). */
function sign_(payload) {
  var body = b64url_(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  var signature = Utilities.computeRsaSha256Signature(body, privateKey_());
  return body + '.' + b64url_(signature);
}

function privateKey_() {
  var raw = PropertiesService.getScriptProperties().getProperty('PRIVATE_KEY');
  if (!raw) throw new Error('Script property PRIVATE_KEY is missing');

  // O Script Properties hay bi mat dau xuong dong, nen dung lai dung dang PEM.
  var body = raw.replace(/-----[A-Z ]+-----/g, '').replace(/\s+/g, '');
  var lines = body.match(/.{1,64}/g) || [];
  return '-----BEGIN PRIVATE KEY-----\n' + lines.join('\n') + '\n-----END PRIVATE KEY-----\n';
}

function b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] & 0xff;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

// ------------------------------------------------------------ Sheet tools ----

function book_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Doc ca sheet (bo dong tieu de) vao bo nho: { sheet, width, rows }. */
function table_(def) {
  var sheet = book_().getSheetByName(def.name);
  if (!sheet) throw new Error('Sheet "' + def.name + '" is missing - run setup()');
  var last = sheet.getLastRow();
  var rows = last > 1 ? sheet.getRange(2, 1, last - 1, def.head.length).getValues() : [];
  return { sheet: sheet, width: def.head.length, rows: rows };
}

function writeRow_(table, index) {
  table.sheet.getRange(index + 2, 1, 1, table.width).setValues([pad_(table.rows[index], table.width)]);
}

/** Ghi cac dong tu `from` den `to` (0-based, tinh ca hai dau) trong mot lan goi. */
function writeRows_(table, from, to) {
  var block = [];
  for (var i = from; i <= to; i++) block.push(pad_(table.rows[i], table.width));
  table.sheet.getRange(from + 2, 1, block.length, table.width).setValues(block);
}

/** Tao sheet neu chua co, dat dong tieu de. Sheet da co thi giu nguyen du lieu. */
function ensureSheet_(def) {
  var book = book_();
  var sheet = book.getSheetByName(def.name);
  if (sheet) return sheet;
  sheet = book.insertSheet(def.name);
  sheet.getRange(1, 1, 1, def.head.length).setValues([def.head]).setFontWeight('bold').setBackground('#D6E6F7');
  sheet.setFrozenRows(1);
  return sheet;
}

function appendRow_(table, row) {
  var full = pad_(row, table.width);
  table.sheet.appendRow(full);
  table.rows.push(full);
  return table.rows.length - 1;
}

function pad_(row, width) {
  var out = [];
  for (var i = 0; i < width; i++) out.push(row[i] === undefined || row[i] === null ? '' : safeCell_(row[i]));
  return out;
}

/**
 * Chu do nguoi ngoai gui len (ten may, ten Google, noi dung chuyen khoan) ma bat dau bang
 * = + - @ thi Sheet se hieu la cong thuc. Them dau nhay don phia truoc de no chi la chu.
 * Doc lai bang getValues() khong thay dau nhay, nen ghi lai nhieu lan van on.
 */
function safeCell_(value) {
  return typeof value === 'string' && /^[=+\-@]/.test(value) ? "'" + value : value;
}

/** Dong dau tien co cot `col` bang `value` (khong phan biet hoa thuong), -1 neu khong co. */
function findRow_(table, col, value) {
  var want = String(value).toLowerCase().trim();
  for (var i = 0; i < table.rows.length; i++) {
    if (String(table.rows[i][col]).toLowerCase().trim() === want) return i;
  }
  return -1;
}

function config_() {
  var cfg = {};
  table_(T_CONFIG).rows.forEach(function (r) {
    var key = String(r[0]).trim();
    // Gia tri dan vao o hay mang theo dau tab / dau cach thua: cat di.
    if (key) cfg[key] = typeof r[1] === 'string' ? r[1].trim() : r[1];
  });
  return cfg;
}

function findPlan_(id, includeInactive) {
  var plans = table_(T_PLANS);
  var i = id ? findRow_(plans, P.id, id) : -1;
  if (i < 0) return null;
  var r = plans.rows[i];
  if (!includeInactive && !isTrue_(r[P.active])) return null;
  return { id: String(r[P.id]), months: Number(r[P.months]) || 0, price: Number(r[P.price]) || 0 };
}

function orderPrefix_(cfg) {
  var prefix = String(cfg.OrderPrefix || 'OP').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return prefix || 'OP';
}

function randomCode_(length) {
  var out = '';
  for (var i = 0; i < length; i++) out += ALPHABET.charAt(Math.floor(Math.random() * ALPHABET.length));
  return out;
}

/** O ngay trong Sheet co the la Date, so, chu ISO hoac LIFETIME. Khong doc duoc = 0 (da het han). */
function toMs_(value) {
  if (value instanceof Date) return value.getTime();
  if (typeof value === 'number') return value;
  if (String(value) === LIFETIME) return LIFETIME_SECONDS * 1000;
  var parsed = Date.parse(String(value));
  return isNaN(parsed) ? 0 : parsed;
}

function isTrue_(value) {
  return value === true || /^(true|yes|1|x)$/i.test(String(value).trim());
}

function sameEmail_(a, b) {
  return String(a).toLowerCase().trim() === String(b).toLowerCase().trim();
}

function cleanId_(value) {
  var id = String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  return id.length >= 8 && id.length <= 64 ? id : '';
}

function short_(value, max) {
  return String(value === undefined || value === null ? '' : value).substring(0, max);
}

// --------------------------------------------- Menu cho nguoi ban (trong Sheet) ----

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bản quyền')
    .addItem('1. Tạo các sheet (setup)', 'setup')
    .addItem('2. Kiểm tra cấu hình', 'checkSetup')
    .addSeparator()
    .addItem('Duyệt đơn ở dòng đang chọn (đã nhận tiền)', 'approveSelectedOrder')
    .addItem('Cấp gói cho tài khoản ở dòng đang chọn…', 'grantPlanToSelectedUser')
    .addItem('Tạo mã kích hoạt…', 'createCodes')
    .addItem('Xem link webhook SePay', 'showWebhookUrl')
    .addItem('Tính lại thống kê tool (ToolStats)', 'refreshToolStats')
    .addToUi();
}

/** Tao cac sheet con thieu, dien cau hinh va bang gia mau. Chay lai khong mat du lieu. */
function setup() {
  var book = book_();
  [T_CONFIG, T_PLANS, T_USERS, T_MACHINES, T_ORDERS, T_CODES, T_BANK, T_USAGE, T_STATS].forEach(function (def) {
    var sheet = book.getSheetByName(def.name) || book.insertSheet(def.name);
    sheet.getRange(1, 1, 1, def.head.length).setValues([def.head]).setFontWeight('bold').setBackground('#D6E6F7');
    sheet.setFrozenRows(1);
  });

  var config = table_(T_CONFIG);
  DEFAULT_CONFIG.forEach(function (row) {
    if (findRow_(config, 0, row[0]) < 0) appendRow_(config, row);
  });

  var plans = table_(T_PLANS);
  if (plans.rows.length === 0) DEFAULT_PLANS.forEach(function (row) { appendRow_(plans, row); });

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('WEBHOOK_KEY')) {
    props.setProperty('WEBHOOK_KEY', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  }

  SpreadsheetApp.getUi().alert('Đã tạo xong các sheet.\n\nTiếp theo: điền sheet Config và Plans, '
    + 'dán PRIVATE_KEY vào Script Properties, rồi chạy "Kiểm tra cấu hình".');
}

/** Bao tung muc cau hinh da on hay chua. */
function checkSetup() {
  var cfg = config_();
  var lines = [];
  function line(ok, text) { lines.push((ok ? '✔ ' : '✘ ') + text); }

  var signed = false, why = '';
  try {
    signed = sign_({ v: 1, test: true }).split('.').length === 2;
  } catch (err) {
    why = ' (' + err.message + ')';
  }
  line(signed, 'PRIVATE_KEY ký được giấy phép' + why);
  line(!!cfg.GoogleClientId, 'Config.GoogleClientId');
  line(!!cfg.BankBin && !!cfg.AccountNo && !!cfg.AccountName, 'Config.BankBin / AccountNo / AccountName');
  line(table_(T_PLANS).rows.some(function (r) { return isTrue_(r[P.active]); }), 'Plans có ít nhất một gói đang bán');
  line(!!PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY'), 'WEBHOOK_KEY (cho SePay)');

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/** Duyet tay: chon mot dong trong sheet Orders roi chay - dung khi khong dung SePay. */
function approveSelectedOrder() {
  var ui = SpreadsheetApp.getUi();
  var sheet = book_().getActiveSheet();
  if (sheet.getName() !== T_ORDERS.name) { ui.alert('Hãy mở sheet Orders và chọn dòng của đơn cần duyệt.'); return; }

  var index = sheet.getActiveRange().getRow() - 2;
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var orders = table_(T_ORDERS);
    if (index < 0 || index >= orders.rows.length) { ui.alert('Hãy chọn một dòng có đơn.'); return; }

    var row = orders.rows[index];
    if (String(row[O.status]) === 'paid') { ui.alert('Đơn ' + row[O.code] + ' đã được duyệt rồi.'); return; }

    var answer = ui.alert('Duyệt đơn ' + row[O.code] + '?',
      row[O.email] + '\nGói ' + row[O.plan] + ' - ' + row[O.amount] + ' đ', ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return;

    var status = markPaid_(orders, index, 'manual');
    ui.alert(status === 'paid' ? 'Đã duyệt. Khách bấm "Kiểm tra lại" trong add-in là dùng được.'
                               : 'Không duyệt được: ' + status);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cap mot goi cho tai khoan: chon mot dong trong sheet Users roi chay. Cong han dung dung nhu
 * khi khach mua goi do (tinh tiep tu ngay het han hien tai neu con han, khong thi tu hom nay).
 * Go tay vao o Plan KHONG lam viec nay - o Plan chi la ten goi, han dung nam o o Expires.
 */
function grantPlanToSelectedUser() {
  var ui = SpreadsheetApp.getUi();
  var sheet = book_().getActiveSheet();
  if (sheet.getName() !== T_USERS.name) { ui.alert('Hãy mở sheet Users và chọn dòng của tài khoản cần cấp gói.'); return; }

  var index = sheet.getActiveRange().getRow() - 2;
  var answer = ui.prompt('Cấp gói', 'Id của gói (cột Id trong sheet Plans), vd M6 hoặc LIFE:', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  var plan = findPlan_(answer.getResponseText().trim(), true);
  if (!plan) { ui.alert('Không có gói đó trong sheet Plans.'); return; }

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var users = table_(T_USERS);
    if (index < 0 || index >= users.rows.length || String(users.rows[index][U.email]) === '') {
      ui.alert('Hãy chọn một dòng có tài khoản.');
      return;
    }

    applyPlan_(users, index, plan);

    var expires = users.rows[index][U.expires];
    var shown = String(expires) === LIFETIME ? 'vĩnh viễn'
      : expires.getDate() + '/' + (expires.getMonth() + 1) + '/' + expires.getFullYear();
    ui.alert('Đã cấp gói ' + plan.id + ' cho ' + users.rows[index][U.email] + '.\nHạn dùng mới: ' + shown
      + '\n\nKhách bấm "Kiểm tra lại" trong add-in là thấy.');
  } finally {
    lock.releaseLock();
  }
}

/** Nhac khi ai do go tay vao o Plan cua sheet Users: viec do khong doi han dung. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== T_USERS.name) return;
    if (e.range.getRow() < 2 || e.range.getColumn() !== U.plan + 1) return;
    e.source.toast('Ô Plan chỉ là tên gói; hạn dùng nằm ở ô Expires. Muốn cấp gói kèm hạn dùng: '
      + 'menu Bản quyền → Cấp gói cho tài khoản ở dòng đang chọn.', 'Lưu ý', 15);
  } catch (err) {
    // Chi la loi nhac; khong duoc lam hong viec sua o.
  }
}

/** Tao mot loat ma kich hoat cho mot goi; ma hien trong sheet Codes. */
function createCodes() {
  var ui = SpreadsheetApp.getUi();
  var planAnswer = ui.prompt('Tạo mã kích hoạt', 'Id của gói (cột Id trong sheet Plans), vd M12:', ui.ButtonSet.OK_CANCEL);
  if (planAnswer.getSelectedButton() !== ui.Button.OK) return;

  var plan = findPlan_(planAnswer.getResponseText().trim(), true);
  if (!plan) { ui.alert('Không có gói đó trong sheet Plans.'); return; }

  var countAnswer = ui.prompt('Tạo mã kích hoạt', 'Số mã cần tạo (1-200):', ui.ButtonSet.OK_CANCEL);
  if (countAnswer.getSelectedButton() !== ui.Button.OK) return;

  var count = Math.floor(Number(countAnswer.getResponseText()));
  if (!(count >= 1 && count <= 200)) { ui.alert('Số mã phải từ 1 đến 200.'); return; }

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var codes = table_(T_CODES);
    var now = new Date();
    for (var i = 0; i < count; i++) {
      var code;
      do {
        code = randomCode_(12);
      } while (findRow_(codes, C.code, code) >= 0);
      appendRow_(codes, [code, plan.id, now, '', '']);
    }
  } finally {
    lock.releaseLock();
  }
  ui.alert('Đã tạo ' + count + ' mã cho gói ' + plan.id + ' - xem sheet Codes.');
}

function showWebhookUrl() {
  // Khong tu doc URL web app o day: ScriptApp tra ve link /dev khi chay tu trinh soan thao.
  var key = PropertiesService.getScriptProperties().getProperty('WEBHOOK_KEY');
  SpreadsheetApp.getUi().alert(key
    ? 'Link webhook cho SePay = URL Web app (kết thúc bằng /exec) + đoạn dưới đây. Giữ bí mật.\n\n'
      + '?hook=sepay&key=' + key
    : 'Chưa chạy setup.');
}
