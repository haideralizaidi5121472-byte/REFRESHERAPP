const { DateTime } = require("luxon");
const APP_ZONE = process.env.APP_ZONE || "Asia/Karachi";
const BUSINESS_DAY_START_HOUR = Number(process.env.BUSINESS_DAY_START_HOUR || 4) || 4;
/* IMPORTANT: used in computeMonthly SQLite timezone shifting */
function getAppTzOffsetHours() {
  return Math.trunc(DateTime.now().setZone(APP_ZONE).offset / 60);
}console.log("APP_ZONE", APP_ZONE);
console.log(
  "NOW (Local)",
  DateTime.now().setZone(APP_ZONE).toFormat("ccc LLL dd yyyy HH:mm:ss 'GMT'ZZ")
);
console.log("NOW ISO (UTC)", new Date().toISOString());
console.log("APP_TZ_OFFSET_HOURS", getAppTzOffsetHours());
const path = require("path");
const fs = require("fs");
const express = require("express");
const session = require("express-session");
const SQLiteStore = require("connect-sqlite3")(session);
const bodyParser = require("body-parser");
const multer = require("multer");
const XLSX = require("xlsx");
const bcrypt = require("bcryptjs");

const config = require("./config");
const { initDb, nextCustomerId, getOutstanding } = require("./db");
const { requireRole, safeInt, todayYmd, ymdOffset, normalizePkPhone } = require("./utils");

const app = express();

/* =========================
   NEW SETTINGS
   ========================= */
const DEFAULT_RIDER_PASSWORD = String(process.env.DEFAULT_RIDER_PASSWORD || config.defaultRiderPassword || "5121").trim() || "5121";
const DEFAULT_CUSTOMER_PASSWORD = String(process.env.DEFAULT_CUSTOMER_PASSWORD || config.defaultCustomerPassword || "5121").trim() || "5121";
const SALES_LOGIN_WINDOW_MINUTES = Number(process.env.SALES_LOGIN_WINDOW_MINUTES || 10) || 10;
const SALES_ADMIN_CREDENTIALS = config.salesAdminCredentials || { username: "sales", password: "sales123" };
/* ========================= */

const uploadsDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

const db = initDb(path.join(__dirname, "data.sqlite"));

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static(path.join(__dirname, "public")));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.use(
  session({
    store: new SQLiteStore({ db: "sessions.sqlite", dir: __dirname }),
secret: (process.env.SESSION_SECRET || "refresher_super_secret_key"),
    resave: false,
    saveUninitialized: false,
    cookie: { maxAge: 1000 * 60 * 60 * 24 * 30 },
  })
);

const upload = multer({ dest: uploadsDir });

function nowIso() {
  return new Date().toISOString();
}

/* =========================
   BUSINESS DAY 4 AM HELPERS
   ========================= */
console.log("BD", businessDayInfo());

function pad2(n) {
  const x = Number(n || 0);
  return x < 10 ? "0" + x : String(x);
}
function inactiveCutoffIso(days = 14) {
  try {
    const d = DateTime.now().setZone(APP_ZONE).minus({ days: Number(days || 14) });
    return d.toUTC().toISO();
  } catch (e) {
    const x = new Date();
    x.setDate(x.getDate() - Number(days || 14));
    return x.toISOString();
  }
}

function ymdFromLocalDate(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function businessDayInfo(date = new Date()) {
  try {
    const now = DateTime.fromJSDate(new Date(date)).setZone(APP_ZONE);

    let base = now.startOf("day");
    if (now.hour < BUSINESS_DAY_START_HOUR) base = base.minus({ days: 1 });

    const startLocal = base.set({
      hour: BUSINESS_DAY_START_HOUR,
      minute: 0,
      second: 0,
      millisecond: 0,
    });
    const endLocal = startLocal.plus({ days: 1 });

    return {
      ymd: startLocal.toISODate(),
      startIso: startLocal.toUTC().toISO(),
      endIso: endLocal.toUTC().toISO(),
      startLocal: startLocal.toJSDate(),
      endLocal: endLocal.toJSDate(),
    };
  } catch (e) {
    const fallback = new Date();
    const y = fallback.getFullYear();
    const m = fallback.getMonth();
    const d = fallback.getDate();
    const startUtc = new Date(Date.UTC(y, m, d, 0, 0, 0, 0));
    const endUtc = new Date(startUtc.getTime() + 24 * 60 * 60 * 1000);
    return {
      ymd: `${y}-${pad2(m + 1)}-${pad2(d)}`,
      startIso: startUtc.toISOString(),
      endIso: endUtc.toISOString(),
      startLocal: startUtc,
      endLocal: endUtc,
    };
  }
}

function businessDayInfoForYmd(ymd) {
  try {
    const base = DateTime.fromISO(String(ymd), { zone: APP_ZONE });

    const startLocal = base.startOf("day").set({
      hour: BUSINESS_DAY_START_HOUR,
      minute: 0,
      second: 0,
      millisecond: 0,
    });
    const endLocal = startLocal.plus({ days: 1 });

    return {
      ymd: startLocal.toISODate(),
      startIso: startLocal.toUTC().toISO(),
      endIso: endLocal.toUTC().toISO(),
      startLocal: startLocal.toJSDate(),
      endLocal: endLocal.toJSDate(),
    };
  } catch (e) {
    return businessDayInfo();
  }
}
/* ========================= */

function tableCols(tableName) {
  try {
    return db
      .prepare(`PRAGMA table_info(${tableName})`)
      .all()
      .map((r) => r.name);
  } catch (e) {
    return [];
  }
}

function ensureColumn(tableName, colName, ddl) {
  try {
    const cols = tableCols(tableName);
    if (!cols.includes(colName)) db.exec(ddl);
  } catch (e) {}
}

/* seed water types safely */
function ensureWaterTypesSeed() {
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS water_types (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        list_price INTEGER NOT NULL DEFAULT 0
      )
    `);
  } catch (e) {}

  try {
    ensureColumn("water_types", "list_price", "ALTER TABLE water_types ADD COLUMN list_price INTEGER NOT NULL DEFAULT 0");
  } catch (e) {}

  try {
    const name = "Ozonated Mineral Boosted";
    const ex = db.prepare("SELECT id FROM water_types WHERE LOWER(name)=LOWER(?)").get(name);
    if (!ex) {
      db.prepare("INSERT INTO water_types (name, list_price) VALUES (?, ?)").run(name, 160);
    } else {
      const row = db.prepare("SELECT list_price FROM water_types WHERE id=?").get(ex.id);
      const price = Number(row && row.list_price ? row.list_price : 0);
      if (!price || price === 0) db.prepare("UPDATE water_types SET list_price=? WHERE id=?").run(160, ex.id);
    }
  } catch (e) {}
}

function ensureSchemaCompat() {
try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS expenses (

        id INTEGER PRIMARY KEY AUTOINCREMENT,
        exp_date TEXT NOT NULL,
        category TEXT NOT NULL,
        amount INTEGER NOT NULL,
        method TEXT NOT NULL,
        staff_id TEXT,
        note TEXT,
        created_at TEXT NOT NULL
      )
    `);
  } catch (e) {}
  /* SPOT FILLING (walk-in sale) */
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS spot_fillings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        spot_date TEXT NOT NULL,          /* business ymd e.g. 2026-02-13 */
        qty INTEGER NOT NULL DEFAULT 0,
        price INTEGER NOT NULL DEFAULT 0, /* unit price */
        total INTEGER NOT NULL DEFAULT 0, /* qty*price */
        created_by TEXT,
        created_at TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_spot_fillings_date ON spot_fillings(spot_date);
    `);
  } catch (e) {}
  ensureColumn("orders", "payment_received", "ALTER TABLE orders ADD COLUMN payment_received INTEGER NOT NULL DEFAULT 0");

  /* extra compat for older dbs */
  ensureColumn("orders", "delivered_qty", "ALTER TABLE orders ADD COLUMN delivered_qty INTEGER NOT NULL DEFAULT 0");
  ensureColumn("orders", "empty_returned_qty", "ALTER TABLE orders ADD COLUMN empty_returned_qty INTEGER NOT NULL DEFAULT 0");
  ensureColumn("orders", "payment_amount", "ALTER TABLE orders ADD COLUMN payment_amount INTEGER NOT NULL DEFAULT 0");
  ensureColumn("orders", "payment_method", "ALTER TABLE orders ADD COLUMN payment_method TEXT");
  ensureColumn("orders", "delivered_at", "ALTER TABLE orders ADD COLUMN delivered_at TEXT");
  ensureColumn("orders", "list_price", "ALTER TABLE orders ADD COLUMN list_price INTEGER NOT NULL DEFAULT 0");
// NEW: who created order (customer/admin)
ensureColumn("orders", "created_source", "ALTER TABLE orders ADD COLUMN created_source TEXT");

// Backfill old rows as customer (safe)
try {
  db.prepare("UPDATE orders SET created_source='customer' WHERE created_source IS NULL OR TRIM(created_source)=''").run();
} catch (e) {}

  /* customer bottles balance compat */
  ensureColumn("customers", "opening_balance", "ALTER TABLE customers ADD COLUMN opening_balance INTEGER NOT NULL DEFAULT 0");
  ensureColumn("customers", "opening_bottle", "ALTER TABLE customers ADD COLUMN opening_bottle INTEGER NOT NULL DEFAULT 0");
  ensureColumn("customers", "bottles_balance", "ALTER TABLE customers ADD COLUMN bottles_balance INTEGER NOT NULL DEFAULT 0");
  ensureColumn("customers", "password_plain", "ALTER TABLE customers ADD COLUMN password_plain TEXT");

  /* NEW rider password support */
  ensureColumn("riders", "password_plain", "ALTER TABLE riders ADD COLUMN password_plain TEXT");

  /* ledger compat */
  ensureColumn("ledger_entries", "method", "ALTER TABLE ledger_entries ADD COLUMN method TEXT");
  ensureColumn("ledger_entries", "note", "ALTER TABLE ledger_entries ADD COLUMN note TEXT");
  /* SALES ADMIN ONLY finance tables */
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS sa_riders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS sa_rider_tx (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rider_id INTEGER NOT NULL,
        tx_date TEXT NOT NULL,
        tx_type TEXT NOT NULL, /* advance or return */
        amount INTEGER NOT NULL,
        method TEXT NOT NULL DEFAULT 'cash',
        note TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(rider_id) REFERENCES sa_riders(id)
      );

      CREATE INDEX IF NOT EXISTS idx_sa_rider_tx_date ON sa_rider_tx(tx_date);
      CREATE INDEX IF NOT EXISTS idx_sa_rider_tx_rider ON sa_rider_tx(rider_id);

      CREATE TABLE IF NOT EXISTS sa_employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        monthly_salary INTEGER NOT NULL DEFAULT 0,
        salary_day INTEGER NOT NULL DEFAULT 1,
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS sa_employee_adv (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        adv_date TEXT NOT NULL,
        amount INTEGER NOT NULL,
        method TEXT NOT NULL DEFAULT 'cash',
        note TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(employee_id) REFERENCES sa_employees(id)
      );

      CREATE INDEX IF NOT EXISTS idx_sa_emp_adv_date ON sa_employee_adv(adv_date);
      CREATE INDEX IF NOT EXISTS idx_sa_emp_adv_emp ON sa_employee_adv(employee_id);
      CREATE TABLE IF NOT EXISTS sa_employee_salary_tx (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        salary_date TEXT NOT NULL,
        amount INTEGER NOT NULL,
        method TEXT NOT NULL DEFAULT 'cash',
        note TEXT,
        created_at TEXT NOT NULL,
        FOREIGN KEY(employee_id) REFERENCES sa_employees(id)
      );

      CREATE INDEX IF NOT EXISTS idx_sa_emp_salary_date ON sa_employee_salary_tx(salary_date);
      CREATE INDEX IF NOT EXISTS idx_sa_emp_salary_emp ON sa_employee_salary_tx(employee_id);
    `);
  } catch (e) {}
  /* INACTIVE CUSTOMERS status table */
  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS inactive_customer_status (
        customer_id TEXT PRIMARY KEY,
        status TEXT NOT NULL DEFAULT 'inqueue',  /* inqueue or leftforever */
        updated_at TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_inactive_customer_status_status ON inactive_customer_status(status);
    `);
  } catch (e) {}
  ensureWaterTypesSeed();
}

ensureSchemaCompat();

/* set default password for existing riders if empty */
function ensureRidersDefaultPasswords() {
  try {
    db.prepare("UPDATE riders SET password_plain=? WHERE (password_plain IS NULL OR TRIM(password_plain)='')").run(DEFAULT_RIDER_PASSWORD);
  } catch (e) {}
}
ensureRidersDefaultPasswords();

/* set default password for existing customers if empty */
function ensureCustomersDefaultPasswords() {
  try {
    db.prepare("UPDATE customers SET password_plain=? WHERE (password_plain IS NULL OR TRIM(password_plain)='')").run(DEFAULT_CUSTOMER_PASSWORD);
  } catch (e) {}
}
ensureCustomersDefaultPasswords();

/* =========================
   SALES ADMIN GATE AND AUTH
   ========================= */
function setSalesGate(req) {
  try {
    req.session.sales_login_until = Date.now() + SALES_LOGIN_WINDOW_MINUTES * 60 * 1000;
  } catch (e) {}
}

function salesGateOk(req) {
  try {
    const until = Number(req.session && req.session.sales_login_until ? req.session.sales_login_until : 0);
    if (!until) return false;
    return Date.now() <= until;
  } catch (e) {
    return false;
  }
}

function requireSales(req, res, next) {
  if (req.session && req.session.sales_user) return next();
  return res.redirect("/sales/login");
}

/* BLOCK 1
   Is function ko apne code me requireSalesAdmin wali jagah replace kar do
*/
function requireSalesAdmin(req, res, next) {
  if (req.session && req.session.sales_admin) return next();
  try { req.session.sales_admin_redirect = req.originalUrl || "/sales_admin/report"; } catch (e) {}
  return res.redirect("/sales_admin/login");
}
/* owner like session */
function isOwnerLike(req) {
  try {
    if (req.session && req.session.user && req.session.user.role === "owner") return true;
  } catch (e) {}
  try {
    if (req.session && req.session.sales_admin) return true;
  } catch (e) {}
  return false;
}

function isManagerAdmin(req) {
  try {
    return !!(req.session && req.session.user && req.session.user.role === "admin");
  } catch (e) {
    return false;
  }
}

function getEffectiveUser(req, overrideUser) {
  if (overrideUser) return overrideUser;

  try {
    if (req.session && req.session.user) return req.session.user;
  } catch (e) {}

  try {
    if (req.session && req.session.sales_admin) {
      return { role: "owner", id: "0", name: req.session.sales_admin.name || "owner" };
    }
  } catch (e) {}

  try {
    if (req.session && req.session.sales_user) return req.session.sales_user;
  } catch (e) {}

  return null;
}

/* staff middleware allows admin or owner or sales_admin */
function requireStaff(req, res, next) {
  const u = getEffectiveUser(req);
  if (u && (u.role === "admin" || u.role === "owner")) return next();
  if (req.session && req.session.sales_admin) return next();
  return res.redirect("/login/admin");
}

/* allow admin or sales admin to access compare */
function requireAdminOrSalesAdmin(req, res, next) {
  if (req.session && req.session.user && (req.session.user.role === "admin" || req.session.user.role === "owner")) return next();
  if (req.session && req.session.sales_admin) return next();
  try {
    req.session.sales_admin_redirect = req.originalUrl || "/sales_admin/report";
  } catch (e) {}
  return res.redirect("/sales_admin/login");
}
/* ========================= */

function pickBody(req, keys, fallback = "") {
  for (const k of keys) {
    const v = req.body && req.body[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return fallback;
}

function normStatus(s) {
  let v = (s || "").toString().trim().toLowerCase();
  v = v.replace(/\s+/g, "");
  if (v === "on_the_way" || v === "ontheway") v = "onway";
  if (v === "on_way") v = "onway";
  return v;
}

const ALLOWED_STATUSES = new Set(["pending", "assigned", "accepted", "onway", "delivered", "cancelled"]);

function waNumber(raw) {
  try {
    const n = normalizePkPhone ? normalizePkPhone(raw || "") : raw || "";
    return String(n || "").replace(/[^\d]/g, "");
  } catch (e) {
    return String(raw || "").replace(/[^\d]/g, "");
  }
}

function waLink(phone, text) {
  const p = waNumber(phone);
  const t = encodeURIComponent(text || "");
  return `https://wa.me/${p}?text=${t}`;
}

function wantsJsonReq(req) {
  try {
    if (!req) return false;
    if (req.xhr) return true;
    const accept = String((req.headers && req.headers.accept) ? req.headers.accept : "");
    if (accept.includes("application/json")) return true;
    const p = String(req.path || "");
    if (p.startsWith("/api/")) return true;
    return false;
  } catch (e) {
    return false;
  }
}

function renderAny(res, viewNames, data, fallbackJson) {
  const names = Array.isArray(viewNames) ? viewNames : [viewNames];
  const req = res && res.req ? res.req : null;

  const tryNext = (i) => {
    if (i >= names.length) {
      if (fallbackJson && wantsJsonReq(req)) return res.json(fallbackJson);
      return res.status(500).send("View render error");
    }
    return res.render(names[i], data, (err, html) => {
      if (err) {
        try {
          console.error("EJS render error on view:", names[i], err && err.message ? err.message : err);
        } catch (e) {}
        return tryNext(i + 1);
      }
      return res.send(html);
    });
  };
  return tryNext(0);
}

/* helpers */
function billQtyFromOrder(o, deliveredQty) {
  const dq = Number(deliveredQty || 0);
  if (dq > 0) return dq;
  const existing = Number(o && o.delivered_qty ? o.delivered_qty : 0);
  if (existing > 0) return existing;
  return Number(o && o.quantity_requested ? o.quantity_requested : 0);
}

function totalFromOrder(o, billQty) {
  const uq = Number(o && o.unit_price ? o.unit_price : 0);
  return uq * Number(billQty || 0);
}

function safeRedirectBack(req, res, fallback) {
  const ref = (req.get && req.get("referer")) || "";
  if (ref && typeof ref === "string") return res.redirect(ref);
  return res.redirect(fallback || "/admin");
}

function ensureSaleLedgerOnce(customerId, oid, totalBill, createdAt) {
  try {
    const ex = db
      .prepare("SELECT id FROM ledger_entries WHERE entry_type='sale' AND ref_type='order' AND ref_id=? LIMIT 1")
      .get(String(oid));
    if (ex) return;
  } catch (e) {}

  try {
    db.prepare(
      "INSERT INTO ledger_entries (customer_id, entry_type, ref_type, ref_id, amount, created_at) VALUES (?, 'sale', 'order', ?, ?, ?)"
    ).run(customerId, String(oid), totalBill, createdAt);
  } catch (e) {}
}

/* payment ledger only when payment_received is 1 */
function ensurePaymentLedgerOnce(customerId, oid, amountReceived, method, createdAt, paymentReceived = 1) {
  const amt = Number(amountReceived || 0);
  if (!amt || amt <= 0) return;

  const received = Number(paymentReceived) ? 1 : 0;
  if (!received) return;

  try {
    const ex = db
      .prepare(
        "SELECT id FROM ledger_entries WHERE entry_type='payment' AND ref_type='order' AND ref_id=? AND amount=? LIMIT 1"
      )
      .get(String(oid), -amt);
    if (ex) return;
  } catch (e) {}

  try {
    db.prepare(
      "INSERT INTO ledger_entries (customer_id, entry_type, ref_type, ref_id, amount, method, created_at) VALUES (?, 'payment', 'order', ?, ?, ?, ?)"
    ).run(customerId, String(oid), -amt, method, createdAt);
  } catch (e) {}
}

/* manual udhaar receive */
function receiveUdhaar(customerId, amount, method, note, refTypeIn, refIdIn) {
  const cid = String(customerId || "").trim();
  const amt = Number(amount || 0);
  if (!cid || !amt || amt <= 0) return { ok: false, msg: "Invalid receive data" };

  const mRaw = String(method || "cash").trim().toLowerCase();
  const m = mRaw === "jazzcash" ? "jazzcash" : "cash";
  const n = String(note || "").trim();

  const refType = String(refTypeIn || "udhaar").trim();
  const refId = String(refIdIn || "manual").trim() || "manual";

  try {
    db.prepare(
      "INSERT INTO ledger_entries (customer_id, entry_type, ref_type, ref_id, amount, method, note, created_at) VALUES (?, 'payment', ?, ?, ?, ?, ?, ?)"
    ).run(cid, refType, refId, -amt, m, n, nowIso());
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: "Receive failed" };
  }
}
/* bottles balance update */
function updateCustomerBottlesBalance(customerId, deliveredQty, emptyReturnedQty) {
  try {
    const cid = String(customerId || "").trim();
    if (!cid) return;

    const row = db.prepare("SELECT bottles_balance FROM customers WHERE CAST(id AS TEXT)=?").get(cid);
    const prev = Number(row && row.bottles_balance !== undefined ? row.bottles_balance : 0);

    const d = Number(deliveredQty || 0);
    const r = Number(emptyReturnedQty || 0);
    const nextRaw = prev + (d - r);
    const next = nextRaw < 0 ? 0 : nextRaw;

    db.prepare("UPDATE customers SET bottles_balance=? WHERE CAST(id AS TEXT)=?").run(next, cid);
  } catch (e) {}
}

/* total outstanding across all customers */
function totalOutstandingAll() {
  try {
    const row = db
      .prepare(
        `
        SELECT COALESCE(SUM(CASE WHEN bal > 0 THEN bal ELSE 0 END),0) as total
        FROM (
          SELECT customer_id, COALESCE(SUM(amount),0) as bal
          FROM ledger_entries
          GROUP BY customer_id
        )
      `
      )
      .get();
    return Number(row && row.total ? row.total : 0);
  } catch (e) {
    return 0;
  }
}
/* Spot filling totals for a given business ymd */
function spotTotalsForYmd(ymd) {
  const out = { qty: 0, total: 0, avgPrice: 0 };
  try {
    const row = db.prepare(`
      SELECT 
        COALESCE(SUM(qty),0) as q,
        COALESCE(SUM(total),0) as t
      FROM spot_fillings
      WHERE spot_date=?
    `).get(String(ymd));

    out.qty = Number(row && row.q ? row.q : 0);
    out.total = Number(row && row.t ? row.t : 0);
    out.avgPrice = out.qty > 0 ? Math.round(out.total / out.qty) : 0;
  } catch (e) {}
  return out;
}
/* admin delivered core */
function markDeliveredAdminCore(oid, overrides = {}) {
  const o = db.prepare("SELECT * FROM orders WHERE id=?").get(oid);
  if (!o) return { ok: false, msg: "Order not found" };

  const now = nowIso();

  const deliveredQty = Number(overrides.delivered_qty || overrides.deliveredQty || 0) || 0;
  const billQty = billQtyFromOrder(o, deliveredQty);

  const methodRaw = (overrides.payment_method || overrides.method || o.payment_method || "cash").toString().trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  const amountReceived = Number(overrides.payment_amount || overrides.amount_received || o.payment_amount || 0) || 0;
  const emptyReturned = Number(overrides.empty_returned_qty || overrides.emptyReturned || o.empty_returned_qty || 0) || 0;

  const totalBill = totalFromOrder(o, billQty);
const paymentReceivedFlag =
  overrides.payment_received !== undefined && overrides.payment_received !== null
    ? Number(overrides.payment_received) ? 1 : 0
    : method === "jazzcash"
      ? 0
      : (Number(amountReceived || 0) > 0 ? 1 : 0);
  db.transaction(() => {
    db.prepare(
      `
      UPDATE orders
      SET status='delivered',
          delivered_qty=?,
          empty_returned_qty=?,
          payment_amount=?,
          payment_method=?,
          delivered_at=?,
          payment_received=?
      WHERE id=?
    `
    ).run(billQty, emptyReturned, amountReceived, method, now, paymentReceivedFlag, oid);

    ensureSaleLedgerOnce(o.customer_id, oid, totalBill, now);
    ensurePaymentLedgerOnce(o.customer_id, oid, amountReceived, method, now, paymentReceivedFlag);

    updateCustomerBottlesBalance(o.customer_id, billQty, emptyReturned);
  })();

  return { ok: true };
}

/* verify jazzcash core */
function verifyJazzCashCore(oid, overrideAmount) {
  const o = db.prepare("SELECT * FROM orders WHERE id=?").get(oid);
  if (!o) return { ok: false, msg: "Order not found" };

  if (String(o.payment_method || "").toLowerCase() !== "jazzcash") {
    return { ok: false, msg: "Not a JazzCash order" };
  }

  const now = nowIso();
  const qty = Number(o.delivered_qty || o.quantity_requested || 0);
  const totalBill = Number(o.unit_price || 0) * qty;

  const amt =
    overrideAmount !== undefined && overrideAmount !== null && Number(overrideAmount) > 0
      ? Number(overrideAmount)
      : Number(o.payment_amount || 0) > 0
        ? Number(o.payment_amount)
        : totalBill;

  db.transaction(() => {
    db.prepare("UPDATE orders SET payment_received=1, payment_amount=? WHERE id=?").run(amt, oid);

    ensureSaleLedgerOnce(o.customer_id, String(oid), totalBill, now);
    ensurePaymentLedgerOnce(o.customer_id, String(oid), amt, "jazzcash", now, 1);
  })();

  return { ok: true };
}

/* JazzCash whatsapp message */
function buildJazzCashPendingMsg(orderRow) {
  const qty = Number(orderRow.delivered_qty || orderRow.quantity_requested || 0);
  const total = Number(orderRow.unit_price || 0) * qty;
  const paid = Number(orderRow.payment_amount || 0);
  const due = Math.max(0, total - paid);

  const msg =
    `${config.appName}\n` +
    `JazzCash pending\n` +
    `Order #${orderRow.id}\n` +
    `Customer: ${orderRow.customer_name}\n` +
    `Bottles: ${qty}\n` +
    `Total bill: ${total}\n` +
    `Paid: ${paid}\n` +
    `Pending: ${due}\n` +
    `Kindly JazzCash transfer kar dein. Shukriya`;

  return msg;
}

function getGlobalStats() {
  const bd = businessDayInfo();
  try {
    const totalClients = db.prepare("SELECT COUNT(*) as c FROM customers").get().c || 0;

    const todayOrders =
      db
        .prepare("SELECT COUNT(*) as c FROM orders WHERE created_at >= ? AND created_at < ?")
        .get(bd.startIso, bd.endIso).c || 0;

    let todayBottles = 0;
    try {
      const row = db
        .prepare(
          `
        SELECT COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as b
        FROM orders
        WHERE status='delivered' AND delivered_at >= ? AND delivered_at < ?
      `
        )
        .get(bd.startIso, bd.endIso);
      todayBottles = Number(row && row.b ? row.b : 0);
    } catch (e) {}

    let totalDeliveredBottles = 0;
    try {
      const row = db
        .prepare(
          `
        SELECT COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as b
        FROM orders
        WHERE status='delivered'
      `
        )
        .get();
      totalDeliveredBottles = Number(row && row.b ? row.b : 0);
    } catch (e) {}

    return { totalClients, todayOrders, todayBottles, totalDeliveredBottles };
  } catch (e) {
    return { totalClients: 0, todayOrders: 0, todayBottles: 0, totalDeliveredBottles: 0 };
  }
}

function viewData(req, override = {}) {
  const bd = businessDayInfo();
  const mainUser = req.session.user || null;
  const salesUser = req.session.sales_user || null;
  const salesAdmin = req.session.sales_admin || null;

  const effective = getEffectiveUser(req, override.user);

  const role = effective ? effective.role : null;
  const isOwner = role === "owner";

  return {
    appName: config.appName,
    logoUrl: config.logoUrl,
waLink,
waNumber,
    config,
    DEFAULT_CUSTOMER_PASSWORD,
    DEFAULT_RIDER_PASSWORD,
    user: effective,
    mainUser,
    salesUser,
    salesAdmin,
    isOwner,
    currency: config.baseCurrency,
    globalStats: getGlobalStats(),
    businessDay: bd,
    dashboardUrl: (effective && (effective.role === "admin" || effective.role === "owner")) ? "/admin" : (effective && effective.role === "rider") ? "/rider" : (effective && effective.role === "customer") ? "/customer" : "/",
  };
}

/* UPDATED admin stats with expenses split and net */
function adminStatsForDay(dayOrInfo) {
  const info =
    typeof dayOrInfo === "string"
      ? businessDayInfoForYmd(dayOrInfo)
      : dayOrInfo && dayOrInfo.startIso && dayOrInfo.endIso
        ? dayOrInfo
        : businessDayInfo();

  const todayYmd = String(info.ymd);

  // 1) Sales only delivered in today business window
  const salesRow = db.prepare(`
    SELECT
      COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales
    FROM orders
    WHERE status='delivered'
      AND delivered_at >= ?
      AND delivered_at < ?
  `).get(info.startIso, info.endIso);

  let salesToday = Number(salesRow && salesRow.sales ? salesRow.sales : 0);

  // Spot filling counts as sale and cash today
  const spot = spotTotalsForYmd(info.ymd);

  // 2) Expenses today ymd
  const expRow = db.prepare(`
    SELECT 
      COALESCE(SUM(CASE WHEN LOWER(method)='cash' THEN amount ELSE 0 END),0) as expCash,
      COALESCE(SUM(CASE WHEN LOWER(method)='jazzcash' THEN amount ELSE 0 END),0) as expJazz,
      COALESCE(SUM(amount),0) as expTotal
    FROM expenses
    WHERE exp_date = ?
  `).get(info.ymd);

  const expCashToday = Number(expRow && expRow.expCash ? expRow.expCash : 0);
  const expJazzToday = Number(expRow && expRow.expJazz ? expRow.expJazz : 0);
  const totalExpenses = Number(expRow && expRow.expTotal ? expRow.expTotal : 0);

  // 3) Today order payments only for orders delivered today business ymd
  const todayOrderPay = db.prepare(`
    SELECT
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(le.method,''))='cash' THEN -le.amount ELSE 0 END),0) as cash,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(le.method,''))='jazzcash' THEN -le.amount ELSE 0 END),0) as jazz
    FROM ledger_entries le
    JOIN orders o ON CAST(o.id AS TEXT)=CAST(le.ref_id AS TEXT)
    WHERE le.entry_type='payment'
      AND le.amount < 0
      AND le.ref_type='order'
      AND le.created_at >= ?
      AND le.created_at < ?
      AND (
        CASE
          WHEN o.delivered_at IS NULL OR o.delivered_at=''
            THEN ''
          ELSE (
            CASE
              WHEN CAST(strftime('%H', o.delivered_at) AS INTEGER) < ${BUSINESS_DAY_START_HOUR}
                THEN date(o.delivered_at, '-1 day')
              ELSE date(o.delivered_at)
            END
          )
        END
      ) = ?
  `).get(info.startIso, info.endIso, todayYmd);

  let todayCashCollected = Number(todayOrderPay && todayOrderPay.cash ? todayOrderPay.cash : 0);
  let todayJazzCollected = Number(todayOrderPay && todayOrderPay.jazz ? todayOrderPay.jazz : 0);

  // 4) Today udhaar received only ref_id == todayYmd
  const todayUdhaarPay = db.prepare(`
    SELECT
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='cash' THEN -amount ELSE 0 END),0) as cash,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='jazzcash' THEN -amount ELSE 0 END),0) as jazz
    FROM ledger_entries
    WHERE entry_type='payment'
      AND amount < 0
      AND ref_type IN ('udhaar_day','udhaar')
      AND ref_id = ?
      AND created_at >= ?
      AND created_at < ?
  `).get(todayYmd, info.startIso, info.endIso);

  todayCashCollected += Number(todayUdhaarPay && todayUdhaarPay.cash ? todayUdhaarPay.cash : 0);
  todayJazzCollected += Number(todayUdhaarPay && todayUdhaarPay.jazz ? todayUdhaarPay.jazz : 0);

  // 5) Previous JazzCash collected today, but order delivered on older business day
  const prevJazzRow = db.prepare(`
    SELECT COALESCE(SUM(-le.amount),0) as jazz
    FROM ledger_entries le
    JOIN orders o ON CAST(o.id AS TEXT)=CAST(le.ref_id AS TEXT)
    WHERE le.entry_type='payment'
      AND le.amount < 0
      AND le.ref_type='order'
      AND LOWER(COALESCE(le.method,''))='jazzcash'
      AND le.created_at >= ?
      AND le.created_at < ?
      AND (
        CASE
          WHEN o.delivered_at IS NULL OR o.delivered_at=''
            THEN ''
          ELSE (
            CASE
              WHEN CAST(strftime('%H', o.delivered_at) AS INTEGER) < ${BUSINESS_DAY_START_HOUR}
                THEN date(o.delivered_at, '-1 day')
              ELSE date(o.delivered_at)
            END
          )
        END
      ) <> ?
  `).get(info.startIso, info.endIso, todayYmd);

  const previousJazzCollected = Number(prevJazzRow && prevJazzRow.jazz ? prevJazzRow.jazz : 0);

  // 6) Previous udhaar collected today, ref_id not equal todayYmd
  const prevUdhaarRow = db.prepare(`
    SELECT COALESCE(SUM(-amount),0) as amt
    FROM ledger_entries
    WHERE entry_type='payment'
      AND amount < 0
      AND ref_type IN ('udhaar_day','udhaar')
      AND ref_id <> ?
      AND created_at >= ?
      AND created_at < ?
  `).get(todayYmd, info.startIso, info.endIso);

  const previousUdhaarCollected = Number(prevUdhaarRow && prevUdhaarRow.amt ? prevUdhaarRow.amt : 0);

  // Spot filling adds to sale and cash today
  if (spot && Number(spot.total || 0) > 0) {
    salesToday += Number(spot.total || 0);
    todayCashCollected += Number(spot.total || 0);
  }

  // 7) Pending JazzCash all time, never resets until verified
  const pendingJazzRow = db.prepare(`
    SELECT COALESCE(SUM(
      CASE
        WHEN COALESCE(o.payment_amount,0) > 0 THEN COALESCE(o.payment_amount,0)
        ELSE (COALESCE(o.unit_price,0) * CASE WHEN COALESCE(o.delivered_qty,0) > 0 THEN COALESCE(o.delivered_qty,0) ELSE COALESCE(o.quantity_requested,0) END)
      END
    ),0) as pending
    FROM orders o
    WHERE o.status='delivered'
      AND LOWER(COALESCE(o.payment_method,''))='jazzcash'
      AND COALESCE(o.payment_received,0)=0
  `).get();

  const pendingJazzAll = Number(pendingJazzRow && pendingJazzRow.pending ? pendingJazzRow.pending : 0);

  // 8) Total outstanding udhaar all time
  const udhaarAll = totalOutstandingAll();

  // 9) Total net
  const totalNetToday = (todayCashCollected + todayJazzCollected) - totalExpenses;
  const totalNetWithPrevious = totalNetToday + previousJazzCollected + previousUdhaarCollected;

  return {
    businessYmd: info.ymd,
    businessStartIso: info.startIso,
    businessEndIso: info.endIso,

    salesToday,

    cashInHand: todayCashCollected,
    totalJazz: todayJazzCollected,

    previousJazzCollected,
    previousUdhaarCollected,

    pendingJazzAll,
    totalUdhaarAll: udhaarAll,

    spotQty: Number(spot.qty || 0),
    spotTotal: Number(spot.total || 0),
    spotAvgPrice: Number(spot.avgPrice || 0),

    totalExpenses,
    expCashToday,
    expJazzToday,

    cashNet: todayCashCollected - expCashToday,
    jazzNet: todayJazzCollected - expJazzToday,

    totalNetToday,
    totalNetWithPrevious,
  };
}
function riderTodayStats(rid) {
  const bd = businessDayInfo();
  const row = db
    .prepare(
      `
    SELECT 
      COUNT(*) as deliveries,
      COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
      COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned,
      COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales,
      COALESCE(SUM(CASE WHEN payment_method='cash' AND COALESCE(payment_received,1)=1 THEN payment_amount ELSE 0 END),0) as cash,
      COALESCE(SUM(CASE WHEN payment_method='jazzcash' AND COALESCE(payment_received,0)=1 THEN payment_amount ELSE 0 END),0) as jazz
    FROM orders
    WHERE rider_id=? AND status='delivered' AND delivered_at >= ? AND delivered_at < ?
  `
    )
    .get(rid, bd.startIso, bd.endIso);

  const deliveries = Number(row.deliveries || 0);
  const bottles = Number(row.bottles || 0);
  const returned = Number(row.returned || 0);
  const sales = Number(row.sales || 0);
  const cash = Number(row.cash || 0);
  const jazz = Number(row.jazz || 0);
  const due = Math.max(0, sales - (cash + jazz));

  return { deliveries, bottles, returned, sales, cash, jazz, due };
}

/* rider history last N days */
function riderHistoryStats(rid, daysBack = 14) {
  const days = Math.max(1, Number(daysBack || 14) || 14);

const from = new Date();
from.setHours(0, 0, 0, 0);
from.setDate(from.getDate() - days);
const fromIso = from.toISOString();
  try {
    const rows = db
      .prepare(
        `
      SELECT 
        substr(delivered_at,1,10) as day,
        COUNT(*) as deliveries,
        COUNT(DISTINCT customer_id) as clients,
        COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
        COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned,
        COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales,
        COALESCE(SUM(CASE WHEN payment_method='cash' AND COALESCE(payment_received,1)=1 THEN payment_amount ELSE 0 END),0) as cash,
        COALESCE(SUM(CASE WHEN payment_method='jazzcash' AND COALESCE(payment_received,0)=1 THEN payment_amount ELSE 0 END),0) as jazz
      FROM orders
      WHERE rider_id=? AND status='delivered' AND delivered_at >= ?
      GROUP BY substr(delivered_at,1,10)
      ORDER BY day DESC
      LIMIT 60
    `
      )
      .all(rid, fromIso);

    return rows.map((r) => {
      const sales = Number(r.sales || 0);
      const cash = Number(r.cash || 0);
      const jazz = Number(r.jazz || 0);
      return {
        day: r.day,
        deliveries: Number(r.deliveries || 0),
        clients: Number(r.clients || 0),
        bottles: Number(r.bottles || 0),
        returned: Number(r.returned || 0),
        sales,
        cash,
        jazz,
        due: Math.max(0, sales - (cash + jazz)),
      };
    });
  } catch (e) {
    return [];
  }
}

/* AUTH */
app.get("/", (req, res) => {
  if (req.session && req.session.sales_admin) return res.redirect("/sales_admin/report");
  if (req.session.user) return res.redirect("/" + (req.session.user.role === "owner" ? "admin" : req.session.user.role));
  return res.redirect("/login/customer");
});

app.get("/login", (req, res) => renderAny(res, ["login"], { ...viewData(req), msg: null }, { ok: false }));

/* ONLY allow these roles on public login route */
app.get("/login/:role", (req, res) => {
  const role = String(req.params.role || "").toLowerCase().trim();
  if (!["customer", "admin", "rider"].includes(role)) return res.redirect("/login");
  return renderAny(res, ["login_" + role], { ...viewData(req), msg: null }, { ok: false, role });
});

/* customer login now supports optional password, default password works without typing */
app.post("/login/customer", (req, res) => {
  const id = (req.body.id || "").toString().trim();
  const passInRaw = pickBody(req, ["password", "pass", "pin"], "");
  const passIn = String(passInRaw || "").trim();

  const c = db.prepare("SELECT * FROM customers WHERE CAST(id AS TEXT) = ? AND is_active=1").get(id);
  if (!c) return renderAny(res, ["login_customer"], { ...viewData(req), msg: "ID not found" }, { ok: false });

  const expected = String((c.password_plain || "").trim() || DEFAULT_CUSTOMER_PASSWORD);

  if (!passIn) {
    if (expected !== DEFAULT_CUSTOMER_PASSWORD) {
      return renderAny(res, ["login_customer"], { ...viewData(req), msg: "Password required" }, { ok: false });
    }
  } else {
    if (passIn !== expected) {
      return renderAny(res, ["login_customer"], { ...viewData(req), msg: "Invalid password" }, { ok: false });
    }
  }

  req.session.user = { role: "customer", id: String(c.id), name: c.name };
  return res.redirect("/customer");
});

/* rider login requires password (admin-set password) */
app.post("/login/rider", (req, res) => {
  const id = (req.body.id || "").toString().trim();
  const passInRaw = pickBody(req, ["password", "pass", "pin"], "");
  const passIn = String(passInRaw || "").trim();

  const r = db.prepare("SELECT * FROM riders WHERE id = ? AND is_active = 1").get(id);
  if (!r) return renderAny(res, ["login_rider"], { ...viewData(req), msg: "Rider ID not found" }, { ok: false });

  const expected = String((r.password_plain || "").trim() || DEFAULT_RIDER_PASSWORD);

  // ✅ If admin has set a custom password, rider must enter it
  if (!passIn) {
    return renderAny(res, ["login_rider"], { ...viewData(req), msg: "Password required" }, { ok: false });
  }

  if (passIn !== expected) {
    return renderAny(res, ["login_rider"], { ...viewData(req), msg: "Invalid password" }, { ok: false });
  }

  req.session.user = { role: "rider", id: String(r.id), name: r.name };
  return res.redirect("/rider");
});
app.post("/login/admin", (req, res) => {
  const username = (req.body.username || "").toString().trim();
  const password = (req.body.password || "").toString();

  let ok = false;

  if (config.adminCredentials && username === config.adminCredentials.username && password === config.adminCredentials.password) ok = true;
  if ((username === "ali" && password === "51214") || (username === "admin" && password === "admin123")) ok = true;

  if (!ok) {
    try {
      const u = db.prepare("SELECT * FROM users WHERE username=?").get(username);
      if (u && bcrypt.compareSync(password, u.password_hash)) ok = true;
    } catch (e) {}
  }

  if (!ok) return renderAny(res, ["login_admin"], { ...viewData(req), msg: "Invalid credentials" }, { ok: false });

  req.session.user = { role: "admin", id: "0", name: username || "admin" };
  return res.redirect("/admin");
});

app.get("/logout", (req, res) => req.session.destroy(() => res.redirect("/")));

/* =========================
   SALES ADMIN FLOW
   ========================= */

/* admin clicks sales report, now goes to sales admin report */
app.get("/admin/sales_report", requireStaff, (req, res) => {
  try {
    req.session.sales_admin_redirect = "/sales_admin/report";
  } catch (e) {}
  return res.redirect("/sales_admin/report");
});
app.get("/admin/sales-report", requireStaff, (req, res) => res.redirect("/admin/sales_report"));
app.get("/admin/salesreport", requireStaff, (req, res) => res.redirect("/admin/sales_report"));

/* old gate based sales login routes stay for compatibility */
app.get("/sales/login", (req, res) => {
  if (req.session.sales_user) return res.redirect("/sales");

  if (!salesGateOk(req)) {
    return res.redirect("/admin");
  }

  return renderAny(
    res,
    ["login_sales_admin", "login_sales", "login_admin"],
    { ...viewData(req), msg: null },
    { ok: true }
  );
});

app.post("/sales/login", (req, res) => {
  if (!salesGateOk(req)) return res.redirect("/admin");

  const username = String(req.body.username || req.body.user || "").trim();
  const password = String(req.body.password || "").toString();

  let ok = false;

  if (SALES_ADMIN_CREDENTIALS && username === SALES_ADMIN_CREDENTIALS.username && password === SALES_ADMIN_CREDENTIALS.password) ok = true;

  if (!ok) {
    try {
      const u = db.prepare("SELECT * FROM users WHERE username=?").get(username);
      if (u && (u.role === "sales" || u.role === "sales_admin") && bcrypt.compareSync(password, u.password_hash)) ok = true;
    } catch (e) {}
  }

  if (!ok) {
    return renderAny(
      res,
      ["login_sales_admin", "login_sales", "login_admin"],
      { ...viewData(req), msg: "Invalid credentials" },
      { ok: false }
    );
  }

  req.session.sales_user = { role: "sales", id: "0", name: username || "sales" };

  try {
    req.session.sales_login_until = 0;
  } catch (e) {}

  return res.redirect("/sales");
});

app.get("/sales/logout", (req, res) => {
  try {
    req.session.sales_user = null;
  } catch (e) {}
  return res.redirect("/admin");
});


/* NEW sales admin direct login and report */
app.get("/sales_admin", requireSalesAdmin, (req, res) => res.redirect("/sales_admin/report"));

app.get("/sales_admin/login", (req, res) => {
  return renderAny(res, ["login_sales_admin", "login_sales", "login_admin"], { ...viewData(req), msg: null }, { ok: false });
});
app.post("/sales_admin/login", (req, res) => {
  const username = (req.body.username || "").toString().trim();
  const password = (req.body.password || "").toString();

  let ok = false;

  if (SALES_ADMIN_CREDENTIALS) {
    if (username === SALES_ADMIN_CREDENTIALS.username && password === SALES_ADMIN_CREDENTIALS.password) ok = true;
  }

  if (!ok) {
    if ((username === "sales" && password === "sales123") || (username === "salesadmin" && password === "sales123")) ok = true;
  }

  if (!ok) {
    return renderAny(res, ["login_sales_admin"], { ...viewData(req), msg: "Invalid credentials" }, { ok: false });
  }

  const redirectTo = (req.session && req.session.sales_admin_redirect) ? req.session.sales_admin_redirect : "/sales_admin/report";
  try { req.session.sales_admin_redirect = null; } catch (e) {}

  req.session.regenerate((err) => {
    if (err) {
      try { req.session.user = null; } catch (e) {}
      req.session.sales_admin = { role: "sales_admin", id: "0", name: username || "salesadmin" };
      req.session.sales_user = { role: "sales", id: "0", name: username || "salesadmin" };
      return res.redirect(redirectTo || "/sales_admin/report");
    }

    try { req.session.user = null; } catch (e) {}
    req.session.sales_admin = { role: "sales_admin", id: "0", name: username || "salesadmin" };
    req.session.sales_user = { role: "sales", id: "0", name: username || "salesadmin" };
    return res.redirect(redirectTo || "/sales_admin/report");
  });
});
app.get("/sales_admin/logout", (req, res) => {
  try { req.session.sales_admin = null; } catch (e) {}
  try { req.session.sales_user = null; } catch (e) {}
try { req.session.user = null; } catch (e) {}
  return res.redirect("/admin");
});


/* ========================= */
/* RIDER */
app.get("/rider", requireRole("rider"), (req, res) => {
  const rid = req.session.user.id;
  const bd = businessDayInfo();

  const orders = db
    .prepare(
      `
      SELECT o.*, c.name as customer_name, c.phone as customer_phone, c.address, c.map_url, c.bottles_balance
      FROM orders o JOIN customers c ON c.id = o.customer_id 
      WHERE o.rider_id = ? AND o.status != 'delivered' AND o.status != 'cancelled'
      ORDER BY o.id DESC
    `
    )
    .all(rid);

  const cash =
    db
      .prepare(
        "SELECT COALESCE(SUM(payment_amount),0) as s FROM orders WHERE rider_id=? AND payment_method='cash' AND status='delivered' AND delivered_at >= ? AND delivered_at < ? AND COALESCE(payment_received,1)=1"
      )
      .get(rid, bd.startIso, bd.endIso).s || 0;

  const jazz =
    db
      .prepare(
        "SELECT COALESCE(SUM(payment_amount),0) as s FROM orders WHERE rider_id=? AND payment_method='jazzcash' AND status='delivered' AND delivered_at >= ? AND delivered_at < ? AND COALESCE(payment_received,0)=1"
      )
      .get(rid, bd.startIso, bd.endIso).s || 0;

  const today = riderTodayStats(rid);

  const totals = {
    deliveries: Number(today.deliveries || 0),
    bottles: Number(today.bottles || 0),
  };

  return renderAny(
    res,
    ["rider_dashboard"],
    { ...viewData(req), orders, summary: { cash, jazz }, totals },
    { orders, summary: { cash, jazz }, totals }
  );
});
/* rider can update status accepted or onway */
app.get("/rider/daily", requireRole("rider"), (req, res) => {
  const rid = String(req.session.user.id || "").trim();
  const info = businessDayInfo();

  const rider = db
    .prepare("SELECT id, name, phone FROM riders WHERE id=?")
    .get(rid);

  const deliveredOrders = db
    .prepare(
      `
      SELECT 
        o.*,
        c.name as customer_name,
        c.phone as customer_phone,
        c.area,
        c.address,
        c.map_url
      FROM orders o
      JOIN customers c ON c.id = o.customer_id
      WHERE o.rider_id=? 
        AND o.status='delivered'
        AND o.delivered_at >= ?
        AND o.delivered_at < ?
      ORDER BY o.delivered_at DESC, o.id DESC
    `
    )
    .all(rid, info.startIso, info.endIso);

  const summaryRow = db
    .prepare(
      `
      SELECT 
        COUNT(*) as deliveries,
        COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
        COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned
      FROM orders
      WHERE rider_id=? 
        AND status='delivered'
        AND delivered_at >= ?
        AND delivered_at < ?
    `
    )
    .get(rid, info.startIso, info.endIso);

  const summary = {
    deliveries: Number(summaryRow && summaryRow.deliveries ? summaryRow.deliveries : 0),
    bottles: Number(summaryRow && summaryRow.bottles ? summaryRow.bottles : 0),
    returned: Number(summaryRow && summaryRow.returned ? summaryRow.returned : 0),
  };

  return renderAny(
    res,
    ["rider_daily"],
    {
      ...viewData(req),
      rider,
      ymd: info.ymd,
      businessDay: info,
      deliveredOrders,
      summary,
    },
    {
      ok: true,
      rider,
      ymd: info.ymd,
      deliveredOrders,
      summary,
    }
  );
});
app.post("/rider/orders/:id/status", requireRole("rider"), (req, res) => {
  const oid = String(req.params.id || "").trim();
  const rid = String(req.session.user.id || "").trim();
  const st = normStatus(req.body.status || "");

  if (!["accepted", "onway"].includes(st)) return res.redirect("/rider");

  try {
    const o = db.prepare("SELECT * FROM orders WHERE id=?").get(oid);
    if (!o) return res.redirect("/rider");
    if (String(o.rider_id || "") !== rid) return res.redirect("/rider");
    if (String(o.status || "") === "delivered") return res.redirect("/rider");

    db.prepare("UPDATE orders SET status=? WHERE id=?").run(st, oid);
  } catch (e) {}

  return res.redirect("/rider");
});

app.post("/complete-order/:id", requireRole("rider"), (req, res) => {
  const oid = req.params.id;

  const deliveredQty = Number(pickBody(req, ["delivered_qty", "deliveredQty", "qty_delivered", "qty"], 0)) || 0;
  const emptyReturned = Number(pickBody(req, ["empty_returned_qty", "emptyReturned", "empty_qty"], 0)) || 0;
  const amountReceived = Number(pickBody(req, ["amount_received", "payment_amount", "amount", "received"], 0)) || 0;

  const methodRaw = (pickBody(req, ["method", "payment_method", "pay_method"], "cash") || "cash").toString().trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  const now = nowIso();

  const o = db.prepare("SELECT * FROM orders WHERE id=?").get(oid);
  if (!o) return res.redirect("/rider");

  const billQty = deliveredQty > 0 ? deliveredQty : Number(o.quantity_requested || 0);
  const totalBill = Number(o.unit_price || 0) * billQty;

  const paymentReceivedFlag = method === "jazzcash" ? 0 : (Number(amountReceived || 0) > 0 ? 1 : 0);

  db.transaction(() => {
    db.prepare(
      `
      UPDATE orders
      SET status='delivered',
          delivered_qty=?,
          empty_returned_qty=?,
          payment_amount=?,
          payment_method=?,
          delivered_at=?,
          payment_received=?
      WHERE id=?
    `
    ).run(billQty, emptyReturned, amountReceived, method, now, paymentReceivedFlag, oid);

    ensureSaleLedgerOnce(o.customer_id, String(oid), totalBill, now);
    ensurePaymentLedgerOnce(o.customer_id, String(oid), amountReceived, method, now, paymentReceivedFlag);

    updateCustomerBottlesBalance(o.customer_id, billQty, emptyReturned);
  })();

  return res.redirect("/rider");
});

/* ADMIN DASHBOARD */
app.get("/admin", requireStaff, (req, res) => {
  const bd = businessDayInfo();
  const ownerLike = isOwnerLike(req);

  const ledgerBalJoin = `
    LEFT JOIN (
      SELECT customer_id, COALESCE(SUM(amount),0) as bal
      FROM ledger_entries
      GROUP BY customer_id
    ) lb ON CAST(lb.customer_id AS TEXT)=CAST(c.id AS TEXT)
  `;

  const openOrders = db
    .prepare(
      `
      SELECT o.*, c.name as customer_name, c.area, c.phone as customer_phone, c.bottles_balance,
             COALESCE(lb.bal,0) as outstanding
      FROM orders o
      JOIN customers c ON o.customer_id = c.id
      ${ledgerBalJoin}
      WHERE o.status != 'delivered'
      ORDER BY 
        CASE 
          WHEN o.status = 'pending' THEN 0
          WHEN o.status = 'assigned' THEN 1
          WHEN o.status = 'accepted' THEN 2
          WHEN o.status = 'onway' THEN 3
          ELSE 4
        END,
        o.id DESC
    `
    )
    .all();

  /* delivered but jazzcash not verified yet */
  const pendingPaymentOrders = db
    .prepare(
      `
      SELECT o.*, c.name as customer_name, c.area, c.phone as customer_phone, r.name as rider_name, c.bottles_balance,
             COALESCE(lb.bal,0) as outstanding
      FROM orders o
      JOIN customers c ON o.customer_id = c.id
      LEFT JOIN riders r ON r.id = o.rider_id
      ${ledgerBalJoin}
      WHERE o.status = 'delivered'
        AND o.payment_method = 'jazzcash'
        AND COALESCE(o.payment_received,0) = 0
      ORDER BY o.delivered_at DESC, o.id DESC
      LIMIT 200
    `
    );

const pendingPaymentOrdersRows = pendingPaymentOrders.all();
  /* completed means delivered and payment settled */
  const completedOrders = db
    .prepare(
      `
      SELECT o.*, c.name as customer_name, c.area, c.phone as customer_phone, r.name as rider_name, c.bottles_balance,
             COALESCE(lb.bal,0) as outstanding
      FROM orders o
      JOIN customers c ON o.customer_id = c.id
      LEFT JOIN riders r ON r.id = o.rider_id
      ${ledgerBalJoin}
      WHERE o.status = 'delivered'
        AND (o.payment_method != 'jazzcash' OR COALESCE(o.payment_received,0) = 1)
        ${(!ownerLike && isManagerAdmin(req)) ? "AND o.delivered_at >= ? AND o.delivered_at < ?" : ""}
      ORDER BY o.delivered_at DESC, o.id DESC
      LIMIT 200
    `
    );

  const completedOrdersRows = (!ownerLike && isManagerAdmin(req))
    ? completedOrders.all(bd.startIso, bd.endIso)
    : completedOrders.all();

  const riders = db.prepare("SELECT id, name, phone FROM riders WHERE is_active=1 ORDER BY id").all();
  const stats = adminStatsForDay(bd);

  const pendingJazzRawStmt = db
    .prepare(
      `
    SELECT o.id, o.payment_amount, o.delivered_at, c.name as cust_name, c.phone as cust_phone
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    WHERE o.status='delivered'
      AND o.payment_method='jazzcash'
      AND COALESCE(o.payment_received,0)=0
    ORDER BY o.delivered_at DESC
    LIMIT 500
  `
    );

const pendingJazzRaw = pendingJazzRawStmt.all();
  const pendingJazz = pendingJazzRaw.map((p) => {
    const msg =
      `${config.appName}\n` +
      `JazzCash pending\n` +
      `Order #${p.id}\n` +
      `Customer: ${p.cust_name}\n` +
      `Pending amount: ${Number(p.payment_amount || 0)}\n` +
      `Kindly JazzCash transfer kar dein. Shukriya`;
    return { ...p, wa_url: waLink(p.cust_phone || "", msg) };
  });

  const riderStats = riders.map((r) => ({
    id: r.id,
    name: r.name,
    phone: r.phone || "",
    today: riderTodayStats(r.id),
    history: (ownerLike ? riderHistoryStats(r.id, 14) : []),
  }));

  return renderAny(
    res,
    ["admin_dashboard"],
    {
      ...viewData(req, { user: getEffectiveUser(req) }),
      orders: openOrders,
      pendingPaymentOrders: pendingPaymentOrdersRows,
      completedOrders: completedOrdersRows,
      riders,
      stats,
      pendingJazz,
      riderStats,
    },
    {
      orders: openOrders,
      pendingPaymentOrders: pendingPaymentOrdersRows,
      completedOrders: completedOrdersRows,
      riders,
      stats,
      pendingJazz,
      riderStats,
    }
  );
});

app.get("/admin/dashboard", requireStaff, (req, res) => res.redirect("/admin"));
/* INACTIVE CUSTOMERS tab data */
app.get("/admin/inactive_customers", requireStaff, (req, res) => {
  const cutoffIso = inactiveCutoffIso(14);

  // Only customers who have at least one order
  // and last order is older than 14 days
  const rows = db.prepare(`
    SELECT
      c.id,
      c.name,
      c.phone,
      MAX(o.created_at) as last_order_at,
      COALESCE(s.status, 'inqueue') as status
    FROM customers c
    JOIN orders o ON o.customer_id = c.id
    LEFT JOIN inactive_customer_status s
      ON CAST(s.customer_id AS TEXT) = CAST(c.id AS TEXT)
    GROUP BY c.id
    HAVING MAX(o.created_at) < ?
    ORDER BY last_order_at DESC
  `).all(cutoffIso);

  const inactive = [];
  const leftForever = [];

  for (const r of rows) {
    const st = String(r.status || "inqueue").toLowerCase().trim();
    if (st === "leftforever") leftForever.push(r);
    else inactive.push(r);
  }

if (wantsJsonReq(req)) {
  return res.json({ ok: true, cutoffIso, inactive, leftForever });
}

return renderAny(
  res,
  ["admin_inactive_customers"],
  { ...viewData(req), cutoffIso, cutoffDays: 14, inactive, leftForever },
  { ok: true, cutoffIso, inactive, leftForever }
);
});

/* Update inactive customer status */
app.post("/admin/inactive_customers/:id/status", requireStaff, (req, res) => {
  const cid = String(req.params.id || "").trim();
  const raw = String(req.body.status || "").toLowerCase().trim();

  const status = (raw === "leftforever") ? "leftforever" : "inqueue";
  if (!cid) return safeRedirectBack(req, res, "/admin/inactive_customers");

  try {
    db.prepare(`
      INSERT INTO inactive_customer_status (customer_id, status, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(customer_id) DO UPDATE SET
        status=excluded.status,
        updated_at=excluded.updated_at
    `).run(cid, status, nowIso());
  } catch (e) {}

  if (wantsJsonReq(req)) {
  return res.json({ ok: true, customer_id: cid, status });
}
return safeRedirectBack(req, res, "/admin/inactive_customers");
});
function spotAddHandler(req, res) {
  const bd = businessDayInfo();
  const spotDate = String(pickBody(req, ["date", "spot_date"], bd.ymd)).trim() || bd.ymd;

  let qty = Number(pickBody(req, ["qty", "quantity"], 0)) || 0;
  let price = Number(pickBody(req, ["price", "unit_price"], 0)) || 0;

  if (qty < 0) qty = 0;
  if (price < 0) price = 0;

  if (qty > 10000) qty = 10000;
  if (price > 1000000) price = 1000000;

  const total = qty * price;
  if (!qty || !price || total <= 0) {
    if (wantsJsonReq(req)) return res.json({ ok: false, msg: "Invalid qty price" });
    return safeRedirectBack(req, res, "/admin");
  }

  const u = getEffectiveUser(req);
  const createdBy = u ? (u.name || u.id || u.role) : "";

  try {
    db.prepare(`
      INSERT INTO spot_fillings (spot_date, qty, price, total, created_by, created_at)
      VALUES (?, ?, ?, ?, ?, ?)
    `).run(spotDate, qty, price, total, createdBy, nowIso());
  } catch (e) {
    if (wantsJsonReq(req)) return res.json({ ok: false, msg: "Spot save failed" });
    return safeRedirectBack(req, res, "/admin");
  }

  if (wantsJsonReq(req)) {
    return res.json({ ok: true, spot_date: spotDate, qty, price, total });
  }
  return safeRedirectBack(req, res, "/admin");
}

app.post("/admin/spot/add", requireStaff, spotAddHandler);
app.post("/sales_admin/spot/add", requireSalesAdmin, spotAddHandler);
/* =========================
   ADMIN: GENERATE ORDER (API)
   ========================= */

// 1) Customer search (name/id/phone)
app.get("/api/admin/customers/search", requireStaff, (req, res) => {
  const q = String(req.query.q || "").trim();
  if (!q) return res.json({ ok: true, customers: [] });

  // safe LIKE
  const like = `%${q.replace(/[%_]/g, "")}%`;

  try {
    const rows = db.prepare(`
      SELECT
        c.id, c.name, c.phone, c.area,
        c.water_type_id, c.unit_price,
        w.name as water_name,
        w.list_price
      FROM customers c
      LEFT JOIN water_types w ON w.id = c.water_type_id
      WHERE c.is_active=1
        AND (
          CAST(c.id AS TEXT) LIKE ?
          OR LOWER(c.name) LIKE LOWER(?)
          OR REPLACE(COALESCE(c.phone,''),' ','') LIKE REPLACE(?,' ','')
        )
      ORDER BY CAST(c.id AS INTEGER) ASC
      LIMIT 20
    `).all(like, like, like);

    return res.json({ ok: true, customers: rows });
  } catch (e) {
    return res.json({ ok: false, msg: "Search failed", customers: [] });
  }
});

// 2) Create order by admin (same as customer logic)
app.post("/api/admin/orders/generate", requireStaff, (req, res) => {
  const customerId = String(pickBody(req, ["customer_id", "cid", "id"], "")).trim();
  let qty = safeInt(pickBody(req, ["quantity", "qty", "quantity_requested"], 1), 1);
  if (qty < 1) qty = 1;
  if (qty > 100) qty = 100;

  const riderId = String(pickBody(req, ["rider_id", "rider"], "")).trim();

  if (!customerId) return res.status(400).json({ ok: false, msg: "Customer required" });

  const customer = db.prepare(`
    SELECT c.id, c.water_type_id, c.unit_price, w.list_price
    FROM customers c
    LEFT JOIN water_types w ON w.id = c.water_type_id
    WHERE CAST(c.id AS TEXT)=? AND c.is_active=1
  `).get(customerId);

  if (!customer) return res.status(404).json({ ok: false, msg: "Customer not found" });

  // rider optional
  let rid = null;
  let status = "pending";
  if (riderId) {
    const r = db.prepare("SELECT id FROM riders WHERE id=? AND is_active=1").get(riderId);
    if (!r) return res.status(400).json({ ok: false, msg: "Invalid rider" });
    rid = riderId;
    status = "assigned";
  }

  const waterTypeId = Number(customer.water_type_id || 1) || 1;
  const unitPrice = Number(customer.unit_price || 0) || 0;
  const listPrice = Number(customer.list_price || 0) || 0;

  try {
    const ins = db.prepare(`
      INSERT INTO orders
      (customer_id, water_type_id, unit_price, list_price, quantity_requested, status, rider_id, created_at, created_source)
      VALUES
      (?, ?, ?, ?, ?, ?, ?, ?, 'admin')
    `).run(customerId, waterTypeId, unitPrice, listPrice, qty, status, rid, nowIso());
try {
  db.prepare(
    "DELETE FROM inactive_customer_status WHERE CAST(customer_id AS TEXT)=CAST(? AS TEXT)"
  ).run(String(customerId));
} catch (e) {}
    return res.json({ ok: true, order_id: ins.lastInsertRowid });
  } catch (e) {
    return res.status(500).json({ ok: false, msg: "Order create failed" });
  }
});

/* ADMIN RIDERS */
app.get("/admin/riders", requireStaff, (req, res) => {
  const ownerLike = isOwnerLike(req);
  const riders = db.prepare("SELECT * FROM riders ORDER BY is_active DESC, id DESC").all();
  const rows = riders.map((r) => ({ ...r, today: riderTodayStats(r.id), history: (ownerLike ? riderHistoryStats(r.id, 30) : []) }));
  return renderAny(res, ["admin_riders"], { ...viewData(req), riders: rows, msg: null }, { riders: rows });
});

app.get("/admin/manage_riders", requireStaff, (req, res) => res.redirect("/admin/riders"));
app.get("/admin/rider", requireStaff, (req, res) => res.redirect("/admin/riders"));

/* rider stats api for dropdown */
app.get("/admin/riders/:id/stats", requireStaff, (req, res) => {
  const rid = String(req.params.id || "").trim();
  const days = Number(req.query.days || 14) || 14;
  const rider = db.prepare("SELECT * FROM riders WHERE id=?").get(rid);
  if (!rider) return res.status(404).json({ ok: false, msg: "Rider not found" });

  const ownerLike = isOwnerLike(req);

  const payload = {
    ok: true,
    rider: { id: rider.id, name: rider.name, phone: rider.phone || "", is_active: rider.is_active },
    today: riderTodayStats(rid),
    history: ownerLike ? riderHistoryStats(rid, days) : [],
  };

  return res.json(payload);
});
/* sales admin rider stats api for dropdown */
app.get("/sales_admin/riders/:id/stats", requireSalesAdmin, (req, res) => {
  const rid = String(req.params.id || "").trim();
  const days = Number(req.query.days || 14) || 14;

  const rider = db.prepare("SELECT * FROM riders WHERE id=?").get(rid);
  if (!rider) return res.status(404).json({ ok: false, msg: "Rider not found" });

  const ownerLike = isOwnerLike(req);

  const payload = {
    ok: true,
    rider: { id: rider.id, name: rider.name, phone: rider.phone || "", is_active: rider.is_active },
    today: riderTodayStats(rid),
    history: ownerLike ? riderHistoryStats(rid, days) : [],
  };

  return res.json(payload);
});

/* rider daily report */
app.get("/admin/riders/:id/daily", requireAdminOrSalesAdmin, (req, res) => {
  const rid = String(req.params.id || "").trim();

  const ownerLike = isOwnerLike(req);
  const ymdIncoming = String(req.query.ymd || businessDayInfo().ymd).trim();
  const ymd = (!ownerLike && isManagerAdmin(req)) ? businessDayInfo().ymd : ymdIncoming;

  const info = businessDayInfoForYmd(ymd);

  const rider = db.prepare("SELECT id, name, phone FROM riders WHERE id=?").get(rid);
  if (!rider) return res.status(404).send("Rider not found");

  const orders = db
    .prepare(
      `
      SELECT o.*, c.name as customer_name, c.area
      FROM orders o
      JOIN customers c ON c.id = o.customer_id
      WHERE o.rider_id=? AND o.status='delivered' AND o.delivered_at >= ? AND o.delivered_at < ?
      ORDER BY o.delivered_at DESC, o.id DESC
    `
    )
    .all(rid, info.startIso, info.endIso);

  const summaryRow = db
    .prepare(
      `
    SELECT 
      COUNT(*) as deliveries,
      COUNT(DISTINCT customer_id) as clients,
      COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
      COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned,
      COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales,
      COALESCE(SUM(CASE WHEN payment_method='cash' AND COALESCE(payment_received,1)=1 THEN payment_amount ELSE 0 END),0) as cash,
      COALESCE(SUM(CASE WHEN payment_method='jazzcash' AND COALESCE(payment_received,0)=1 THEN payment_amount ELSE 0 END),0) as jazz
    FROM orders
    WHERE rider_id=? AND status='delivered' AND delivered_at >= ? AND delivered_at < ?
  `
    )
    .get(rid, info.startIso, info.endIso);

  const summary = {
    deliveries: Number(summaryRow.deliveries || 0),
    clients: Number(summaryRow.clients || 0),
    bottles: Number(summaryRow.bottles || 0),
    returned: Number(summaryRow.returned || 0),
    sales: Number(summaryRow.sales || 0),
    cash: Number(summaryRow.cash || 0),
    jazz: Number(summaryRow.jazz || 0),
    due: Math.max(0, Number(summaryRow.sales || 0) - (Number(summaryRow.cash || 0) + Number(summaryRow.jazz || 0))),
  };

  return renderAny(
    res,
    ["admin_rider_daily"],
    { ...viewData(req), rider, ymd: info.ymd, orders, summary, businessDay: info },
    { ok: true, rider, ymd: info.ymd, orders, summary }
  );
});
app.get("/sales_admin/riders/:id/daily", requireSalesAdmin, (req, res) => {
  const id = String(req.params.id || "").trim();
  const qsIndex = req.originalUrl.indexOf("?");
  const qs = qsIndex >= 0 ? req.originalUrl.slice(qsIndex) : "";
  return res.redirect("/admin/riders/" + id + "/daily" + qs);
});

/* rider date range report for monthly and custom ranges */
app.get("/admin/riders/:id/report", requireStaff, (req, res) => {
  const rid = String(req.params.id || "").trim();

  const ownerLike = isOwnerLike(req);
  let from = String(req.query.from || "").trim();
  let to = String(req.query.to || "").trim();

  if (!ownerLike && isManagerAdmin(req)) {
    const cur = businessDayInfo().ymd;
    from = cur;
    to = cur;
  }

  const rider = db.prepare("SELECT * FROM riders WHERE id=?").get(rid);
  if (!rider) return res.status(404).json({ ok: false, msg: "Rider not found" });

  let where = "WHERE rider_id=? AND status='delivered'";
  const args = [rid];

  if (from) {
    where += " AND substr(delivered_at,1,10) >= ?";
    args.push(from);
  }
  if (to) {
    where += " AND substr(delivered_at,1,10) <= ?";
    args.push(to);
  }

  const rows = db
    .prepare(
      `
      SELECT
        substr(delivered_at,1,10) as day,
        COUNT(*) as deliveries,
        COUNT(DISTINCT customer_id) as clients,
        COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
        COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned,
        COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales,
        COALESCE(SUM(CASE WHEN payment_method='cash' AND COALESCE(payment_received,1)=1 THEN payment_amount ELSE 0 END),0) as cash,
        COALESCE(SUM(CASE WHEN payment_method='jazzcash' AND COALESCE(payment_received,0)=1 THEN payment_amount ELSE 0 END),0) as jazz
      FROM orders
      ${where}
      GROUP BY substr(delivered_at,1,10)
      ORDER BY day DESC
      LIMIT 2000
    `
    )
    .all(...args);

  return res.json({
    ok: true,
    rider: { id: rider.id, name: rider.name, phone: rider.phone || "" },
    from: from || null,
    to: to || null,
    rows: rows.map((r) => {
      const sales = Number(r.sales || 0);
      const cash = Number(r.cash || 0);
      const jazz = Number(r.jazz || 0);
      return {
        day: r.day,
        deliveries: Number(r.deliveries || 0),
        clients: Number(r.clients || 0),
        bottles: Number(r.bottles || 0),
        returned: Number(r.returned || 0),
        sales,
        cash,
        jazz,
        due: Math.max(0, sales - (cash + jazz)),
      };
    }),
  });
});

app.post("/admin/riders/new", requireStaff, (req, res) => {
  const id = (req.body.id || "").toString().trim();
  const name = (req.body.name || "").toString().trim();
  const phone = (req.body.phone || "").toString().trim();
  const pass = String(pickBody(req, ["password_plain", "password", "pass"], DEFAULT_RIDER_PASSWORD)).trim() || DEFAULT_RIDER_PASSWORD;

  if (!id || !name) return res.redirect("/admin/riders");

  const existing = db.prepare("SELECT id FROM riders WHERE id=?").get(id);
  if (existing) return res.redirect("/admin/riders");

  try {
    db.prepare("INSERT INTO riders (id, name, phone, password_plain, created_at, is_active) VALUES (?, ?, ?, ?, ?, 1)").run(
      id,
      name,
      phone,
      pass,
      nowIso()
    );
  } catch (e) {}

  return res.redirect("/admin/riders");
});

/* admin rider update with password edit */
app.post("/admin/riders/:id/update", requireStaff, (req, res) => {
  const rid = String(req.params.id || "").trim();
  const name = String(req.body.name || "").trim();
  const phone = String(req.body.phone || "").trim();
  const isActive = Number(req.body.is_active || 1) ? 1 : 0;
  const pass = String(pickBody(req, ["password_plain", "password", "pass"], "")).trim();

  try {
    if (pass) {
      db.prepare("UPDATE riders SET name=?, phone=?, is_active=?, password_plain=? WHERE id=?").run(name, phone, isActive, pass, rid);
    } else {
      db.prepare("UPDATE riders SET name=?, phone=?, is_active=? WHERE id=?").run(name, phone, isActive, rid);
    }
  } catch (e) {}

  return res.redirect("/admin/riders");
});

/* admin rider delete */
app.post("/admin/riders/:id/delete", requireStaff, (req, res) => {
  const rid = String(req.params.id || "").trim();
  try {
    db.prepare("DELETE FROM riders WHERE id=?").run(rid);
  } catch (e) {}
  return res.redirect("/admin/riders");
});

/* ADMIN ORDERS */
app.get("/admin/orders", requireStaff, (req, res) => {
  const bd = businessDayInfo();
  const ownerLike = isOwnerLike(req);

  const ledgerBalJoin = `
    LEFT JOIN (
      SELECT customer_id, COALESCE(SUM(amount),0) as bal
      FROM ledger_entries
      GROUP BY customer_id
    ) lb ON CAST(lb.customer_id AS TEXT)=CAST(c.id AS TEXT)
  `;

  const stmt = db
    .prepare(
      `
    SELECT 
      o.*,
      c.name as customer_name,
      c.phone as customer_phone,
      c.bottles_balance,
      w.name as water_name,
      r.name as rider_name,
      COALESCE(lb.bal,0) as outstanding
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    LEFT JOIN water_types w ON w.id = o.water_type_id
    LEFT JOIN riders r ON r.id = o.rider_id
    ${ledgerBalJoin}
    WHERE 1=1
      ${(!ownerLike && isManagerAdmin(req)) ? "AND (o.status != 'delivered' OR (o.delivered_at >= ? AND o.delivered_at < ?))" : ""}
    ORDER BY o.id DESC
    LIMIT 500
  `
    );

  const orders = (!ownerLike && isManagerAdmin(req))
    ? stmt.all(bd.startIso, bd.endIso)
    : stmt.all();

  const riders = db.prepare("SELECT id, name FROM riders WHERE is_active=1 ORDER BY id").all();

  return renderAny(res, ["admin_orders"], { ...viewData(req), orders, riders }, { orders, riders });
});

/* admin set payment method */
app.post("/admin/orders/:id/payment_method", requireStaff, (req, res) => {
  const oid = String(req.params.id || "").trim();
  const methodRaw = String(req.body.payment_method || "").trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  try {
    db.prepare("UPDATE orders SET payment_method=? WHERE id=?").run(method, oid);
  } catch (e) {}

  return safeRedirectBack(req, res, "/admin/orders");
});

/* assign alias */
function assignOrderCore(oid, riderId) {
  const rid = String(riderId || "").trim();
  if (!rid) return;

  const rider = db.prepare("SELECT id FROM riders WHERE id=? AND is_active=1").get(rid);
  if (!rider) return;

  const o = db.prepare("SELECT id, status FROM orders WHERE id=?").get(oid);
  if (!o) return;

  const cur = String(o.status || "").trim().toLowerCase();
  if (cur === "delivered" || cur === "cancelled") return;

  const nextStatus = (cur === "pending" || !cur) ? "assigned" : cur;
  db.prepare("UPDATE orders SET rider_id=?, status=? WHERE id=?").run(rid, nextStatus, oid);
}

app.get("/admin/orders/:id/assign", requireStaff, (req, res) => {
  const oid = req.params.id;
  const riderId = (req.query.rider_id || "").toString().trim();
  if (!riderId) return safeRedirectBack(req, res, "/admin/orders");

  try {
    assignOrderCore(oid, riderId);
  } catch (e) {}

  return safeRedirectBack(req, res, "/admin/orders");
});

app.post("/admin/orders/:id/assign", requireStaff, (req, res) => {
  const oid = req.params.id;
  const riderId = (req.body.rider_id || "").toString().trim();
  if (!riderId) return safeRedirectBack(req, res, "/admin/orders");

  try {
    assignOrderCore(oid, riderId);
  } catch (e) {}

  return safeRedirectBack(req, res, "/admin/orders");
});

/* extra alias for dashboard assign */
app.get("/admin/dashboard/orders/:id/assign", requireStaff, (req, res) => {
  const oid = req.params.id;
  const riderId = (req.query.rider_id || "").toString().trim();
  if (!riderId) return safeRedirectBack(req, res, "/admin");
  try {
    assignOrderCore(oid, riderId);
  } catch (e) {}
  return safeRedirectBack(req, res, "/admin");
});
app.post("/admin/dashboard/orders/:id/assign", requireStaff, (req, res) => {
  const oid = req.params.id;
  const riderId = (req.body.rider_id || "").toString().trim();
  if (!riderId) return safeRedirectBack(req, res, "/admin");
  try {
    assignOrderCore(oid, riderId);
  } catch (e) {}
  return safeRedirectBack(req, res, "/admin");
});

/* status via GET */
app.get("/admin/orders/:id/status", requireStaff, (req, res) => {
  const oid = req.params.id;
  const st = normStatus(req.query.status || "");
  if (!ALLOWED_STATUSES.has(st)) return safeRedirectBack(req, res, "/admin/orders");

  try {
    if (st === "delivered") {
      markDeliveredAdminCore(oid, req.query || {});
    } else {
      db.prepare("UPDATE orders SET status=? WHERE id=?").run(st, oid);
    }
  } catch (e) {}

  return safeRedirectBack(req, res, "/admin/orders");
});

/* mark delivered aliases */
app.get("/admin/orders/:id/mark_delivered", requireStaff, (req, res) => {
  const oid = req.params.id;
  try {
    markDeliveredAdminCore(oid, req.query || {});
  } catch (e) {}
  return safeRedirectBack(req, res, "/admin/orders");
});

app.get("/admin/orders/:id/mark-delivered", requireStaff, (req, res) => {
  const oid = req.params.id;
  try {
    markDeliveredAdminCore(oid, req.query || {});
  } catch (e) {}
  return safeRedirectBack(req, res, "/admin/orders");
});

app.post("/admin/orders/:id/status", requireStaff, (req, res) => {
  const oid = req.params.id;
  const st = normStatus(req.body.status || "");
  if (!ALLOWED_STATUSES.has(st)) return safeRedirectBack(req, res, "/admin/orders");

  try {
    if (st === "delivered") {
      markDeliveredAdminCore(oid, req.body || {});
    } else {
      db.prepare("UPDATE orders SET status=? WHERE id=?").run(st, oid);
    }
  } catch (e) {}

  return safeRedirectBack(req, res, "/admin/orders");
});

/* invoice whatsapp */
app.get("/admin/orders/:id/whatsapp", requireStaff, (req, res) => {
  const oid = req.params.id;

  const o = db
    .prepare(
      `
    SELECT 
      o.*,
      c.name as customer_name,
      c.phone as customer_phone
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    WHERE o.id=?
  `
    )
    .get(oid);

  if (!o) return safeRedirectBack(req, res, "/admin/orders");

  const qty = Number(o.delivered_qty || o.quantity_requested || 0);
  const total = Number(o.unit_price || 0) * qty;
  const paid = Number(o.payment_amount || 0);
  const due = Math.max(0, total - paid);

  const msg =
    `${config.appName}\n` +
    `Invoice for order #${o.id}\n` +
    `Customer: ${o.customer_name}\n` +
    `Qty: ${qty}\n` +
    `Total: ${total}\n` +
    `Paid: ${paid}\n` +
    `Due: ${due}\n` +
    `Thank you`;

  return res.redirect(waLink(o.customer_phone, msg));
});

/* JazzCash pending whatsapp route */
app.get("/admin/orders/:id/jazzcash_whatsapp", requireStaff, (req, res) => {
  const oid = req.params.id;

  const o = db
    .prepare(
      `
    SELECT 
      o.*,
      c.name as customer_name,
      c.phone as customer_phone
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    WHERE o.id=?
  `
    )
    .get(oid);

  if (!o) return safeRedirectBack(req, res, "/admin/pending_jazzcash");

  const msg = buildJazzCashPendingMsg(o);
  return res.redirect(waLink(o.customer_phone, msg));
});

app.get("/admin/orders/:id/jazzcash-whatsapp", requireStaff, (req, res) => {
  const oid = req.params.id;

  const o = db
    .prepare(
      `
    SELECT 
      o.*,
      c.name as customer_name,
      c.phone as customer_phone
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    WHERE o.id=?
  `
    )
    .get(oid);

  if (!o) return safeRedirectBack(req, res, "/admin/pending_jazzcash");

  const msg = buildJazzCashPendingMsg(o);
  return res.redirect(waLink(o.customer_phone, msg));
});

/* PENDING JAZZCASH */
app.get("/admin/pending_jazzcash", requireStaff, (req, res) => {
  const bd = businessDayInfo();
  const ownerLike = isOwnerLike(req);

  const pendingJazzRawStmt = db
    .prepare(
      `
    SELECT 
      o.id,
      o.payment_amount,
      o.delivered_at,
      c.name as cust_name,
      c.phone as cust_phone
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    WHERE o.status='delivered'
      AND o.payment_method='jazzcash'
      AND COALESCE(o.payment_received,0)=0
      ${(!ownerLike && isManagerAdmin(req)) ? "AND o.delivered_at >= ? AND o.delivered_at < ?" : ""}
    ORDER BY o.delivered_at DESC
    LIMIT 500
  `
    );

  const pendingJazzRaw = (!ownerLike && isManagerAdmin(req))
    ? pendingJazzRawStmt.all(bd.startIso, bd.endIso)
    : pendingJazzRawStmt.all();

  const pendingJazz = pendingJazzRaw.map((p) => {
    const msg =
      `${config.appName}\n` +
      `JazzCash pending\n` +
      `Order #${p.id}\n` +
      `Customer: ${p.cust_name}\n` +
      `Pending amount: ${Number(p.payment_amount || 0)}\n` +
      `Kindly JazzCash transfer kar dein. Shukriya`;
    return { ...p, wa_url: waLink(p.cust_phone || "", msg), wa_route: `/admin/orders/${p.id}/jazzcash_whatsapp` };
  });

  return renderAny(res, ["admin_pending_jazzcash"], { ...viewData(req), pendingJazz }, { pendingJazz });
});

app.get("/admin/pending-jazzcash", requireStaff, (req, res) => res.redirect("/admin/pending_jazzcash"));

app.post("/admin/orders/:id/verify_jazzcash", requireStaff, (req, res) => {
  const oid = req.params.id;
  try {
    const overrideAmount = pickBody(req, ["amount", "payment_amount", "paid"], null);
    verifyJazzCashCore(oid, overrideAmount);
  } catch (e) {}
  return safeRedirectBack(req, res, "/admin/pending_jazzcash");
});

/* EXPENSES */
app.get("/admin/expenses", requireStaff, (req, res) => {
  const bd = businessDayInfo();
  const today = bd.ymd;

  const ownerLike = isOwnerLike(req);

  let expenses = [];
  try {
    if (!ownerLike && isManagerAdmin(req)) {
      expenses = db.prepare("SELECT * FROM expenses WHERE exp_date=? ORDER BY id DESC LIMIT 5000").all(today);
    } else {
      expenses = db.prepare("SELECT * FROM expenses ORDER BY id DESC LIMIT 5000").all();
    }
  } catch (e) {
    expenses = [];
  }

  const totalExpensesToday = db
    .prepare("SELECT COALESCE(SUM(amount),0) as s FROM expenses WHERE exp_date=?")
    .get(today).s;

  const totalExpensesAll = db.prepare("SELECT COALESCE(SUM(amount),0) as s FROM expenses").get().s;

  const stats = adminStatsForDay(bd);

  const payload = {
    ...viewData(req),
    expenses,
    today,
    stats,
    totalExpenses: Number(totalExpensesAll || 0),
    totalExpensesAll: Number(totalExpensesAll || 0),
    totalExpensesToday: Number(totalExpensesToday || 0),
    total: Number(totalExpensesAll || 0),
    todayTotal: Number(totalExpensesToday || 0),
    msg: null,
  };

  return renderAny(res, ["admin_expenses", "daily_expense", "expenses"], payload, {
    expenses,
    today,
    stats,
    totalExpensesAll: Number(totalExpensesAll || 0),
    totalExpensesToday: Number(totalExpensesToday || 0),
  });
});

app.get("/admin/daily_expense", requireStaff, (req, res) => res.redirect("/admin/expenses"));
app.get("/admin/daily-expense", requireStaff, (req, res) => res.redirect("/admin/expenses"));
app.get("/admin/expense", requireStaff, (req, res) => res.redirect("/admin/expenses"));

function addExpenseHandler(req, res) {
  const bd = businessDayInfo();
  const expDate = (pickBody(req, ["exp_date", "date"], bd.ymd) || bd.ymd).toString().trim();
  const category = pickBody(req, ["category", "type"], "").toString().trim();
  const amount = Number(pickBody(req, ["amount", "expense_amount", "cost"], 0)) || 0;

  const methodRaw = (pickBody(req, ["method", "payment_method"], "cash") || "cash").toString().trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  const staffId = pickBody(req, ["staff_id", "staff"], "").toString().trim();
  const note = pickBody(req, ["note", "details", "desc", "description"], "").toString().trim();

  if (!category || !amount) return res.redirect("/admin/expenses");

  try {
    db.prepare(
      "INSERT INTO expenses (exp_date, category, amount, method, staff_id, note, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)"
    ).run(expDate, category, amount, method, staffId, note, nowIso());
  } catch (e) {}

  return res.redirect("/admin/expenses");
}

app.post("/admin/expenses/add", requireStaff, addExpenseHandler);
function addExpenseHandlerSales(req, res) {
  const bd = businessDayInfo();
  const expDate = (pickBody(req, ["exp_date", "date"], bd.ymd) || bd.ymd).toString().trim();
  const category = pickBody(req, ["category", "type"], "").toString().trim();
  const amount = Number(pickBody(req, ["amount", "expense_amount", "cost"], 0)) || 0;

  const methodRaw = (pickBody(req, ["method", "payment_method"], "cash") || "cash").toString().trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  const staffId = pickBody(req, ["staff_id", "staff"], "").toString().trim();
  const note = pickBody(req, ["note", "details", "desc", "description"], "").toString().trim();

  if (!category || !amount) return res.redirect("/sales_admin/expenses?date=" + encodeURIComponent(expDate));

  try {
    db.prepare(
      "INSERT INTO expenses (exp_date, category, amount, method, staff_id, note, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)"
    ).run(expDate, category, amount, method, staffId, note, nowIso());
  } catch (e) {}

  return res.redirect("/sales_admin/expenses?date=" + encodeURIComponent(expDate));
}

app.post("/sales_admin/expenses/add", requireSalesAdmin, addExpenseHandlerSales);

app.post("/sales_admin/expenses/:id/update", requireSalesAdmin, (req, res) => {
  const id = Number(req.params.id || 0);
  const category = String(req.body.category || "").trim();
  const amount = Number(req.body.amount || 0);

  const methodRaw = String(req.body.method || "cash").trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  const note = String(req.body.note || "").trim();

  const backDate = String(req.body.exp_date || req.body.date || req.query.date || "").trim();

  if (!id || !category || amount <= 0) {
    return res.redirect("/sales_admin/expenses" + (backDate ? "?date=" + encodeURIComponent(backDate) : ""));
  }

  try {
    db.prepare("UPDATE expenses SET category=?, amount=?, method=?, note=? WHERE id=?")
      .run(category, amount, method, note, id);
  } catch (e) {}

  return res.redirect("/sales_admin/expenses" + (backDate ? "?date=" + encodeURIComponent(backDate) : ""));
});
app.post("/admin/expenses/:id/update", requireStaff, (req, res) => {
  const id = Number(req.params.id || 0);
  const category = String(req.body.category || "").trim();
  const amount = Number(req.body.amount || 0);

  const methodRaw = String(req.body.method || "cash").trim().toLowerCase();
  const method = methodRaw === "jazzcash" ? "jazzcash" : "cash";

  const note = String(req.body.note || "").trim();

  if (!id || !category || amount <= 0) return safeRedirectBack(req, res, "/admin/expenses");

  try {
    db.prepare("UPDATE expenses SET category=?, amount=?, method=?, note=? WHERE id=?")
      .run(category, amount, method, note, id);
  } catch (e) {}

  return safeRedirectBack(req, res, "/admin/expenses");
});
app.post("/admin/expenses/new", requireStaff, addExpenseHandler);
app.post("/admin/expenses/create", requireStaff, addExpenseHandler);
app.post("/admin/daily_expense/add", requireStaff, addExpenseHandler);

function monthLabelFromYmd(ymd) {
  const parts = String(ymd || "").split("-");
  const y = Number(parts[0] || 0);
  const m = Number(parts[1] || 0);
  const names = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  if (!y || !m || m < 1 || m > 12) return "";
  return `${names[m - 1]} ${y}`;
}

function businessMonthInfoForYmd(ymd) {
  const cur = businessDayInfo();
  const safe = String(ymd || cur.ymd);

  const base = DateTime.fromISO(safe, { zone: APP_ZONE });

  const startLocal = base.startOf("month").set({
    hour: BUSINESS_DAY_START_HOUR,
    minute: 0,
    second: 0,
    millisecond: 0,
  });

  const endLocal = startLocal.plus({ months: 1 });

  const monthStartYmd = startLocal.toISODate().slice(0, 7) + "-01";
  const monthEndYmd = endLocal.toISODate().slice(0, 7) + "-01";

  return {
    monthLabel: monthLabelFromYmd(monthStartYmd),
    startIso: startLocal.toUTC().toISO(),
    endIso: endLocal.toUTC().toISO(),
    monthStartYmd,
    monthEndYmd,
  };
}
function billQtyForRow(o) {
  const dq = Number(o && o.delivered_qty ? o.delivered_qty : 0);
  if (dq > 0) return dq;
  return Number(o && o.quantity_requested ? o.quantity_requested : 0);
}

function buildOrderViewRow(o) {
  const qty = billQtyForRow(o);
  const unit = Number(o && o.unit_price ? o.unit_price : 0);
  const billTotal = unit * qty;

  const method = String(o && o.payment_method ? o.payment_method : "cash").toLowerCase();
  const receivedFlag = Number(o && o.payment_received ? o.payment_received : 0) === 1;

  const payAmt = Number(o && o.payment_amount ? o.payment_amount : 0);

  const receivedTotal = receivedFlag ? payAmt : 0;

  let dueTotal = 0;
  if (receivedFlag) {
    const diff = billTotal - payAmt;
    dueTotal = diff > 0 ? diff : 0;
  } else {
    if (method === "jazzcash") {
      dueTotal = payAmt > 0 ? payAmt : billTotal;
    } else {
      const diff = billTotal - payAmt;
      dueTotal = diff > 0 ? diff : 0;
    }
  }

return {
  id: o.id,
  customer_name: o.customer_name || "",
  delivered_qty: Number(qty || 0),
  empty_returned_qty: Number(o.empty_returned_qty || 0),
  payment_method: method,
  payment_received: receivedFlag ? 1 : 0,

  bill_total: billTotal,
  received_total: receivedTotal,
  due_total: dueTotal,
};
}

function manualUdhaarReceivedForDay(ymd) {
  const out = { cash: 0, jazz: 0, total: 0 };
  try {
    const rows = db.prepare(
      `
      SELECT LOWER(COALESCE(method,'')) as m, COALESCE(SUM(-amount),0) as amt
      FROM ledger_entries
      WHERE entry_type='payment'
        AND amount < 0
        AND ref_type IN ('udhaar_day','udhaar')
        AND ref_id = ?
      GROUP BY LOWER(COALESCE(method,''))
    `
    ).all(String(ymd));

    for (const r of rows) {
      const m = String(r.m || "").toLowerCase();
      const amt = Number(r.amt || 0);
      if (m === "jazzcash") out.jazz += amt;
      else out.cash += amt;
      out.total += amt;
    }
  } catch (e) {}
  return out;
}

function staffMoneyForDay(ymd) {
const out = { riderAdv: 0, riderRet: 0, empAdv: 0, empSalary: 0 };  try {
    const r = db.prepare(
      `
      SELECT
        COALESCE(SUM(CASE WHEN tx_type='advance' THEN amount ELSE 0 END),0) as adv,
        COALESCE(SUM(CASE WHEN tx_type='return' THEN amount ELSE 0 END),0) as ret
      FROM sa_rider_tx
      WHERE tx_date=?
    `
    ).get(String(ymd));
    out.riderAdv = Number(r && r.adv ? r.adv : 0);
    out.riderRet = Number(r && r.ret ? r.ret : 0);
  } catch (e) {}

  try {
    const e = db.prepare(
      `
      SELECT COALESCE(SUM(amount),0) as adv
      FROM sa_employee_adv
      WHERE adv_date=?
    `
    ).get(String(ymd));
    out.empAdv = Number(e && e.adv ? e.adv : 0);
  } catch (e) {}
  try {
    const s = db.prepare(
      `
      SELECT COALESCE(SUM(amount),0) as sal
      FROM sa_employee_salary_tx
      WHERE salary_date=?
    `
    ).get(String(ymd));
    out.empSalary = Number(s && s.sal ? s.sal : 0);
  } catch (e) {}
  return out;
}

function loadSalesAdminRiders() {
  try {
    const riders = db.prepare("SELECT id, name FROM sa_riders WHERE is_active=1 ORDER BY name ASC").all();
    return riders.map((r) => {
      const row = db.prepare(
        `
        SELECT
          COALESCE(SUM(CASE WHEN tx_type='advance' THEN amount ELSE 0 END),0) as adv,
          COALESCE(SUM(CASE WHEN tx_type='return' THEN amount ELSE 0 END),0) as ret
        FROM sa_rider_tx
        WHERE rider_id=?
      `
      ).get(r.id);
      const adv = Number(row && row.adv ? row.adv : 0);
      const ret = Number(row && row.ret ? row.ret : 0);
      return { id: r.id, name: r.name, advanceTotal: adv, returnTotal: ret, balance: Math.max(0, adv - ret) };
    });
  } catch (e) {
    return [];
  }
}

function loadSalesAdminEmployees(monthInfo) {
  const ym = String(monthInfo && monthInfo.monthStartYmd ? monthInfo.monthStartYmd : "").slice(0, 7);
  try {
    const emps = db.prepare("SELECT id, name, monthly_salary, salary_day FROM sa_employees WHERE is_active=1 ORDER BY name ASC").all();

    return emps.map((e) => {
      const advRow = db.prepare(
        `
        SELECT COALESCE(SUM(amount),0) as adv
        FROM sa_employee_adv
        WHERE employee_id=?
          AND substr(adv_date,1,7)=?
        `
      ).get(e.id, ym);

      const salRow = db.prepare(
        `
        SELECT COALESCE(SUM(amount),0) as sal
        FROM sa_employee_salary_tx
        WHERE employee_id=?
          AND substr(salary_date,1,7)=?
        `
      ).get(e.id, ym);

      const salaryPaidThisMonth = Number(salRow && salRow.sal ? salRow.sal : 0);
      const salary = Number(e.monthly_salary || 0);
      const adv = Number(advRow && advRow.adv ? advRow.adv : 0);

      const totalPaidThisMonth = adv + salaryPaidThisMonth;
      const remaining = Math.max(0, salary - totalPaidThisMonth);

      const sd = Number(e.salary_day || 1);
      const payDate = ym ? (ym + "-" + pad2(sd)) : "";

      return {
        id: e.id,
        name: e.name,
        salary,
        salaryDay: sd,
        payDate,
        advancesThisMonth: adv,
        salaryPaidThisMonth,
        totalPaidThisMonth,
        remainingSalary: remaining,
      };
    });
  } catch (e) {
    return [];
  }
}
function computeDayReport(info) {  const rows = db.prepare(
    `
    SELECT 
      o.*,
      c.name as customer_name
    FROM orders o
    JOIN customers c ON c.id = o.customer_id
    WHERE o.status='delivered'
      AND o.delivered_at >= ?
      AND o.delivered_at < ?
    ORDER BY o.delivered_at DESC, o.id DESC
  `
  ).all(info.startIso, info.endIso);

  const orders = rows.map(buildOrderViewRow);

  const expRow = db.prepare(
    `
    SELECT
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='cash' THEN amount ELSE 0 END),0) as expCash,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='jazzcash' THEN amount ELSE 0 END),0) as expJazz,
      COALESCE(SUM(amount),0) as expTotal
    FROM expenses
    WHERE exp_date=?
  `
  ).get(info.ymd);

  const expensesCash = Number(expRow && expRow.expCash ? expRow.expCash : 0);
  const expensesJazz = Number(expRow && expRow.expJazz ? expRow.expJazz : 0);
  const expenses = Number(expRow && expRow.expTotal ? expRow.expTotal : 0);
  const spot = spotTotalsForYmd(info.ymd);

  let sales = 0;
  let cash = 0;
  let jazz = 0;
  let udhaar = 0;
  let pendingJazz = 0;

  let bottlesDelivered = 0;
  let emptyReturned = 0;

  for (const o of orders) {
    sales += Number(o.bill_total || 0);
    udhaar += Number(o.due_total || 0);

    bottlesDelivered += Number(o.delivered_qty || 0);
    emptyReturned += Number(o.empty_returned_qty || 0);

    const pm = String(o.payment_method || "").toLowerCase();
    const rec = Number(o.received_total || 0);

    if (pm === "jazzcash") jazz += rec;
    else cash += rec;

    if (pm === "jazzcash" && Number(o.payment_received || 0) !== 1) {
      pendingJazz += Number(o.due_total || 0);
    }
  }
  if (spot && Number(spot.total || 0) > 0) {
    sales += Number(spot.total || 0);
    cash += Number(spot.total || 0);
  }
  const manualUdhaar = (info && info.ymd) ? manualUdhaarReceivedForDay(info.ymd) : { cash: 0, jazz: 0, total: 0 };

  /* NOTE
     Ye adjustments hum sirf sales admin route se enable karain ge
     is liye yahan safe fields add kar rahe hain
  */
  let cashAdj = cash;
  let jazzAdj = jazz;
  let udhaarAdj = udhaar;

  if (info && info.applyManualUdhaar === true) {
    cashAdj += Number(manualUdhaar.cash || 0);
    jazzAdj += Number(manualUdhaar.jazz || 0);
    udhaarAdj = Math.max(0, udhaarAdj - Number(manualUdhaar.total || 0));
  }

  const staff = (info && info.applyStaffAdjust === true)
    ? staffMoneyForDay(info.ymd)
    : { riderAdv: 0, riderRet: 0, empAdv: 0, empSalary: 0 };
  const staffNetAdjust = (info && info.applyStaffAdjust === true)
    ? (0 - Number(staff.riderAdv || 0) - Number(staff.empAdv || 0) - Number(staff.empSalary || 0) + Number(staff.riderRet || 0))
    : 0;
  const profitAccrualBase = sales - expenses;
  const profitCashBase = (cashAdj + jazzAdj) - expenses;

  let profitAccrualAdj = profitAccrualBase;
  let profitCashAdj = profitCashBase;

  if (info && info.applyStaffAdjust === true) {
    profitAccrualAdj =
      profitAccrualAdj
      - Number(staff.riderAdv || 0)
      - Number(staff.empAdv || 0)
      - Number(staff.empSalary || 0)
      + Number(staff.riderRet || 0);

    profitCashAdj =
      profitCashAdj
      - Number(staff.riderAdv || 0)
      - Number(staff.empAdv || 0)
      - Number(staff.empSalary || 0)
      + Number(staff.riderRet || 0);
  }
  const cashNet = (cashAdj - expensesCash) + staffNetAdjust;
  const jazzNet = (jazzAdj - expensesJazz);

  // Net after expenses should follow same adjusted logic as profitCash
  const netAfterExpenses = profitCashAdj;       // adjusted (includes staff)
  const netAfterExpensesBase = profitCashBase;  // without staff

  return {
    sales,
    cash: cashAdj,
    jazz: jazzAdj,
    expenses,
    expensesCash,
    expensesJazz,
    cashNet,
    jazzNet,

    // NEW: net after expenses (adjusted + base)
    netAfterExpenses,
    netAfterExpensesBase,

    udhaar: udhaarAdj,
    pendingJazz,
    bottlesDelivered,
    emptyReturned,

    profitAccrual: profitAccrualAdj,
    profitCash: profitCashAdj,

    profitAccrualBase,
    profitCashBase,

    manualUdhaarReceived: manualUdhaar,

    riderAdvances: Number(staff.riderAdv || 0),
    riderReturns: Number(staff.riderRet || 0),
    employeeAdvances: Number(staff.empAdv || 0),
    employeeSalaries: Number(staff.empSalary || 0),
    spotQty: Number(spot.qty || 0),
    spotTotal: Number(spot.total || 0),
    spotAvgPrice: Number(spot.avgPrice || 0),
    orders,
  };
}
function businessYmdFromUtcIso(iso) {
  const dt = DateTime.fromISO(String(iso), { zone: "utc" }).setZone(APP_ZONE);
  const base = dt.hour < BUSINESS_DAY_START_HOUR ? dt.minus({ days: 1 }) : dt;
  return base.toISODate();
}
function computeMonthly(infoDay) {
  const monthInfo = businessMonthInfoForYmd(infoDay.ymd);

  const ordAgg = db.prepare(
  `
  WITH local AS (
  SELECT
    o.*,
    datetime(
      replace(replace(o.delivered_at,'T',' '),'Z',''),
      printf('%+d hours', ?)
    ) AS dt_local
  FROM orders o
  WHERE o.status='delivered'
    AND o.delivered_at >= ?
    AND o.delivered_at < ?
)  SELECT
    CASE
      WHEN CAST(strftime('%H', dt_local) AS INTEGER) < ?
        THEN date(dt_local, '-1 day')
      ELSE date(dt_local)
    END AS day,
    COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales,
    COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
    COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned,
    COALESCE(SUM(CASE WHEN payment_method='cash' AND COALESCE(payment_received,1)=1 THEN payment_amount ELSE 0 END),0) as cash,
    COALESCE(SUM(CASE WHEN payment_method='jazzcash' AND COALESCE(payment_received,0)=1 THEN payment_amount ELSE 0 END),0) as jazz,
    COALESCE(SUM(
      CASE 
        WHEN COALESCE(payment_received,0)=1 THEN
          CASE
            WHEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0)) > 0
            THEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0))
            ELSE 0
          END
        ELSE
          CASE
            WHEN LOWER(COALESCE(payment_method,''))='jazzcash' THEN
              CASE
                WHEN COALESCE(payment_amount,0) > 0 THEN COALESCE(payment_amount,0)
                ELSE (unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END)
              END
            ELSE
              CASE
                WHEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0)) > 0
                THEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0))
                ELSE 0
              END
          END
      END
    ),0) as udhaar
  FROM local
  GROUP BY day
  ORDER BY day DESC
`
).all(getAppTzOffsetHours(), monthInfo.startIso, monthInfo.endIso, BUSINESS_DAY_START_HOUR);  const expAgg = db.prepare(
    `
    SELECT
      exp_date as day,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='cash' THEN amount ELSE 0 END),0) as expCash,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='jazzcash' THEN amount ELSE 0 END),0) as expJazz,
      COALESCE(SUM(amount),0) as expenses
    FROM expenses
    WHERE exp_date >= ?
      AND exp_date < ?
    GROUP BY exp_date
  `
  ).all(monthInfo.monthStartYmd, monthInfo.monthEndYmd);

  const expMap = {};
  for (const e of expAgg) {
    expMap[String(e.day)] = {
      expenses: Number(e.expenses || 0),
      expCash: Number(e.expCash || 0),
      expJazz: Number(e.expJazz || 0),
    };
  }

  const manualMap = {};
  if (infoDay && infoDay.applyManualUdhaar === true) {
    try {
      const mRows = db.prepare(
        `
        SELECT ref_id as day, LOWER(COALESCE(method,'')) as m, COALESCE(SUM(-amount),0) as amt
        FROM ledger_entries
        WHERE entry_type='payment'
          AND amount < 0
          AND ref_type IN ('udhaar_day','udhaar')
          AND ref_id >= ?
          AND ref_id < ?
        GROUP BY ref_id, LOWER(COALESCE(method,''))
      `
      ).all(monthInfo.monthStartYmd, monthInfo.monthEndYmd);

      for (const r of mRows) {
        const day = String(r.day || "");
        if (!manualMap[day]) manualMap[day] = { cash: 0, jazz: 0, total: 0 };
        const amt = Number(r.amt || 0);
        const m = String(r.m || "").toLowerCase();
        if (m === "jazzcash") manualMap[day].jazz += amt;
        else manualMap[day].cash += amt;
        manualMap[day].total += amt;
      }
    } catch (e) {}
  }

  const staffMap = {};
  if (infoDay && infoDay.applyStaffAdjust === true) {
    try {
      const rRows = db.prepare(
        `
        SELECT tx_date as day,
          COALESCE(SUM(CASE WHEN tx_type='advance' THEN amount ELSE 0 END),0) as adv,
          COALESCE(SUM(CASE WHEN tx_type='return' THEN amount ELSE 0 END),0) as ret
        FROM sa_rider_tx
        WHERE tx_date >= ?
          AND tx_date < ?
        GROUP BY tx_date
      `
      ).all(monthInfo.monthStartYmd, monthInfo.monthEndYmd);

      for (const r of rRows) {
        staffMap[String(r.day)] = staffMap[String(r.day)] || { riderAdv: 0, riderRet: 0, empAdv: 0 };
        staffMap[String(r.day)].riderAdv = Number(r.adv || 0);
        staffMap[String(r.day)].riderRet = Number(r.ret || 0);
      }
    } catch (e) {}
    try {
      const sRows = db.prepare(
        `
        SELECT salary_date as day, COALESCE(SUM(amount),0) as sal
        FROM sa_employee_salary_tx
        WHERE salary_date >= ?
          AND salary_date < ?
        GROUP BY salary_date
      `
      ).all(monthInfo.monthStartYmd, monthInfo.monthEndYmd);

      for (const r of sRows) {
        staffMap[String(r.day)] = staffMap[String(r.day)] || { riderAdv: 0, riderRet: 0, empAdv: 0, empSalary: 0 };
        staffMap[String(r.day)].empSalary = Number(r.sal || 0);
      }
    } catch (e) {}
    try {
      const eRows = db.prepare(
        `
        SELECT adv_date as day, COALESCE(SUM(amount),0) as adv
        FROM sa_employee_adv
        WHERE adv_date >= ?
          AND adv_date < ?
        GROUP BY adv_date
      `
      ).all(monthInfo.monthStartYmd, monthInfo.monthEndYmd);

      for (const r of eRows) {
        staffMap[String(r.day)] = staffMap[String(r.day)] || { riderAdv: 0, riderRet: 0, empAdv: 0 };
        staffMap[String(r.day)].empAdv = Number(r.adv || 0);
      }
    } catch (e) {}
  }

 const monthlyRows = ordAgg.map((r) => {
  const day = String(r.day);
  const ex = expMap[day] || { expenses: 0, expCash: 0, expJazz: 0 };

  const row = {
    day,
    sales: Number(r.sales || 0),
    bottles: Number(r.bottles || 0),
    returned: Number(r.returned || 0),
    cash: Number(r.cash || 0),
    jazz: Number(r.jazz || 0),
    udhaar: Number(r.udhaar || 0),

    expenses: Number(ex.expenses || 0),
    expensesCash: Number(ex.expCash || 0),
    expensesJazz: Number(ex.expJazz || 0),

    manualUdhaarReceived: 0,

    riderAdvances: 0,
    riderReturns: 0,
    employeeAdvances: 0,
    employeeSalaries: 0,

    netAfterExpensesBase: 0,
    netAfterExpenses: 0,
  };

  // 1) Pehle manual adjust lagao
  if (infoDay && infoDay.applyManualUdhaar === true) {
    const mm = manualMap[day] || { cash: 0, jazz: 0, total: 0 };
    row.cash = Number(row.cash || 0) + Number(mm.cash || 0);
    row.jazz = Number(row.jazz || 0) + Number(mm.jazz || 0);
    row.udhaar = Math.max(0, Number(row.udhaar || 0) - Number(mm.total || 0));
    row.manualUdhaarReceived = Number(mm.total || 0);
  }

  // 2) Phir base netAfterExpensesBase dubara calculate karo
  row.netAfterExpensesBase =
    (Number(row.cash || 0) + Number(row.jazz || 0)) - Number(row.expenses || 0);

  // 3) Default adjusted equals base
  row.netAfterExpenses = Number(row.netAfterExpensesBase || 0);

  // 4) Staff adjust on ho to netAfterExpenses calculate karo
  if (infoDay && infoDay.applyStaffAdjust === true) {
    const st = staffMap[day] || { riderAdv: 0, riderRet: 0, empAdv: 0, empSalary: 0 };

    row.riderAdvances = Number(st.riderAdv || 0);
    row.riderReturns = Number(st.riderRet || 0);
    row.employeeAdvances = Number(st.empAdv || 0);
    row.employeeSalaries = Number(st.empSalary || 0);

    row.netAfterExpenses =
      Number(row.netAfterExpensesBase || 0)
      - Number(row.riderAdvances || 0)
      - Number(row.employeeAdvances || 0)
      - Number(row.employeeSalaries || 0)
      + Number(row.riderReturns || 0);
  }

  // 5) Har haal me end pe return row
  return row;
});
  let mSales = 0, mCash = 0, mJazz = 0, mUdhaar = 0, mExp = 0, mExpCash = 0, mExpJazz = 0, mBottles = 0, mReturned = 0;
  for (const r of monthlyRows) {
    mSales += Number(r.sales || 0);
    mCash += Number(r.cash || 0);
    mJazz += Number(r.jazz || 0);
    mUdhaar += Number(r.udhaar || 0);
    mExp += Number(r.expenses || 0);
    mExpCash += Number(r.expensesCash || 0);
    mExpJazz += Number(r.expensesJazz || 0);
    mBottles += Number(r.bottles || 0);
    mReturned += Number(r.returned || 0);
  }

  const mProfitCashBase = (mCash + mJazz) - mExp;

    let mRiderAdv = 0, mRiderRet = 0, mEmpAdv = 0, mEmpSalary = 0;
  for (const r of monthlyRows) {
    mRiderAdv += Number(r.riderAdvances || 0);
    mRiderRet += Number(r.riderReturns || 0);
    mEmpAdv += Number(r.employeeAdvances || 0);
    mEmpSalary += Number(r.employeeSalaries || 0);
  }
    const mProfitCash = (infoDay && infoDay.applyStaffAdjust === true)
    ? (mProfitCashBase - mRiderAdv - mEmpAdv - mEmpSalary + mRiderRet)
    : mProfitCashBase;
  const monthly = {
    monthLabel: monthInfo.monthLabel,
    sales: mSales,
    bottlesDelivered: mBottles,
    emptyReturned: mReturned,
    cash: mCash,
    jazz: mJazz,
    expenses: mExp,
    expensesCash: mExpCash,
    expensesJazz: mExpJazz,
    udhaar: mUdhaar,
    profitCash: mProfitCash,
    profitCashBase: mProfitCashBase,
    netAfterExpenses: mProfitCash,
    netAfterExpensesBase: mProfitCashBase,
    riderAdvances: mRiderAdv,
    riderReturns: mRiderRet,
    employeeAdvances: mEmpAdv,
employeeSalaries: mEmpSalary,
    cashNet: mCash - mExpCash,
    jazzNet: mJazz - mExpJazz,
  };

  return { monthly, monthlyRows };
}
function computeLifetime(infoDay) {
  const ordRow = db.prepare(
    `
    SELECT
      COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as sales,
      COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as bottles,
      COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as returned,
      COALESCE(SUM(CASE WHEN payment_method='cash' AND COALESCE(payment_received,1)=1 THEN payment_amount ELSE 0 END),0) as cash,
      COALESCE(SUM(CASE WHEN payment_method='jazzcash' AND COALESCE(payment_received,0)=1 THEN payment_amount ELSE 0 END),0) as jazz,
      COALESCE(SUM(
        CASE 
          WHEN COALESCE(payment_received,0)=1 THEN
            CASE
              WHEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0)) > 0
              THEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0))
              ELSE 0
            END
          ELSE
            CASE
              WHEN LOWER(COALESCE(payment_method,''))='jazzcash' THEN
                CASE
                  WHEN COALESCE(payment_amount,0) > 0 THEN COALESCE(payment_amount,0)
                  ELSE (unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END)
                END
              ELSE
                CASE
                  WHEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0)) > 0
                  THEN ((unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END) - COALESCE(payment_amount,0))
                  ELSE 0
                END
            END
        END
      ),0) as udhaar
    FROM orders
    WHERE status='delivered'
    `
  ).get();

  const expRow = db.prepare(
    `
    SELECT
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='cash' THEN amount ELSE 0 END),0) as expCash,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='jazzcash' THEN amount ELSE 0 END),0) as expJazz,
      COALESCE(SUM(amount),0) as expenses
    FROM expenses
    `
  ).get();

  const manual = { cash: 0, jazz: 0, total: 0 };
  try {
    const rows = db.prepare(
      `
      SELECT LOWER(COALESCE(method,'')) as m, COALESCE(SUM(-amount),0) as amt
      FROM ledger_entries
      WHERE entry_type='payment'
        AND amount < 0
        AND ref_type IN ('udhaar_day','udhaar')
      GROUP BY LOWER(COALESCE(method,''))
      `
    ).all();
    for (const r of rows) {
      const m = String(r.m || "").toLowerCase();
      const amt = Number(r.amt || 0);
      if (m === "jazzcash") manual.jazz += amt;
      else manual.cash += amt;
      manual.total += amt;
    }
  } catch (e) {}

  const staffTotals = { riderAdv: 0, riderRet: 0, empAdv: 0, empSalary: 0 };
  try {
    const r = db.prepare(
      `
      SELECT
        COALESCE(SUM(CASE WHEN tx_type='advance' THEN amount ELSE 0 END),0) as adv,
        COALESCE(SUM(CASE WHEN tx_type='return' THEN amount ELSE 0 END),0) as ret
      FROM sa_rider_tx
      `
    ).get();
    staffTotals.riderAdv = Number(r && r.adv ? r.adv : 0);
    staffTotals.riderRet = Number(r && r.ret ? r.ret : 0);
  } catch (e) {}

  try {
    const e = db.prepare(
      `
      SELECT COALESCE(SUM(amount),0) as adv
      FROM sa_employee_adv
      `
    ).get();
    staffTotals.empAdv = Number(e && e.adv ? e.adv : 0);
  } catch (e) {}

  try {
    const s = db.prepare(
      `
      SELECT COALESCE(SUM(amount),0) as sal
      FROM sa_employee_salary_tx
      `
    ).get();
    staffTotals.empSalary = Number(s && s.sal ? s.sal : 0);
  } catch (e) {}

  const sales = Number(ordRow && ordRow.sales ? ordRow.sales : 0);
  const cashBase = Number(ordRow && ordRow.cash ? ordRow.cash : 0);
  const jazzBase = Number(ordRow && ordRow.jazz ? ordRow.jazz : 0);
  let udhaarBase = Number(ordRow && ordRow.udhaar ? ordRow.udhaar : 0);

  const expenses = Number(expRow && expRow.expenses ? expRow.expenses : 0);
  const expensesCash = Number(expRow && expRow.expCash ? expRow.expCash : 0);
  const expensesJazz = Number(expRow && expRow.expJazz ? expRow.expJazz : 0);

  const cash = cashBase + Number(manual.cash || 0);
  const jazz = jazzBase + Number(manual.jazz || 0);
  udhaarBase = Math.max(0, udhaarBase - Number(manual.total || 0));

  const netAfterExpensesBase = (cash + jazz) - expenses;

  const netAfterExpenses =
    netAfterExpensesBase
    - Number(staffTotals.riderAdv || 0)
    - Number(staffTotals.empAdv || 0)
    - Number(staffTotals.empSalary || 0)
    + Number(staffTotals.riderRet || 0);

  return {
    sales,
    cash,
    jazz,
    expenses,
    expensesCash,
    expensesJazz,
    udhaar: udhaarBase,

    netAfterExpenses,
    netAfterExpensesBase,

    riderAdvances: Number(staffTotals.riderAdv || 0),
    riderReturns: Number(staffTotals.riderRet || 0),
    employeeAdvances: Number(staffTotals.empAdv || 0),
    employeeSalaries: Number(staffTotals.empSalary || 0),

    bottlesDelivered: Number(ordRow && ordRow.bottles ? ordRow.bottles : 0),
    emptyReturned: Number(ordRow && ordRow.returned ? ordRow.returned : 0),
    manualUdhaarReceived: manual,
  };
}
function salesReportPage(req, res) {
  const ownerLike = isOwnerLike(req);
  const managerAdmin = (!ownerLike && isManagerAdmin(req));

  const incomingDate = String(req.query.date || "").trim();
  const incomingCompare = String(req.query.compare || "").trim();

  const todayInfo = businessDayInfo();

  const selected = managerAdmin ? todayInfo.ymd : (incomingDate ? incomingDate : todayInfo.ymd);
  const selectedInfo = businessDayInfoForYmd(selected);

const isSalesPortal = !!(req.session && (req.session.sales_admin || req.session.sales_user));
const isSalesOnly = isSalesPortal;
selectedInfo.applyManualUdhaar = isSalesPortal;
selectedInfo.applyStaffAdjust = isSalesPortal;
  const report = computeDayReport(selectedInfo);
  let compareDate = managerAdmin ? "" : incomingCompare;
  let compareReport = null;

  if (compareDate) {
    if (compareDate === selectedInfo.ymd) {
      compareDate = "";
    } else {
      const cmpInfo = businessDayInfoForYmd(compareDate);
      compareDate = cmpInfo.ymd;
      compareReport = computeDayReport(cmpInfo);
    }
  }

  const { monthly, monthlyRows } = managerAdmin ? { monthly: null, monthlyRows: [] } : computeMonthly(selectedInfo);

  const effectiveUser = getEffectiveUser(req);
  // Riders list for dropdown (same as admin dashboard)
  const riders = db.prepare("SELECT id, name FROM riders WHERE is_active=1 ORDER BY id").all();
const salesRiders = isSalesPortal ? loadSalesAdminRiders() : [];
const salesEmployees = isSalesPortal ? loadSalesAdminEmployees(businessMonthInfoForYmd(selectedInfo.ymd)) : [];
const lifeTime = isSalesPortal ? computeLifetime(selectedInfo) : null;
  const payload = {
    ...viewData(req, { user: effectiveUser }),
    selectedDate: selectedInfo.ymd,
    compareDate: compareDate || "",
    report,
    compareReport,
    monthly,
    monthlyRows,
    salesRiders,
isSalesPortal,
    riders,
    salesEmployees,
lifeTime,
    msg: null,
  };

  const views = ["admin_compare", "admin_sales_compare", "sales_compare"];

  return renderAny(
    res,
    views,
    payload,
    {
      ok: true,
      selectedDate: payload.selectedDate,
      compareDate: payload.compareDate,
      report,
      compareReport,
      monthly,
      monthlyRows,
    }
  );
} // yahan function close ho gaya

/* Sales admin links jo EJS me aate hain unke liye redirects */
function salesAdminExpensesPage(req, res) {
  const bd = businessDayInfo();

  const incoming = String(req.query.date || req.query.ymd || "").trim();
  const selectedYmd = incoming ? businessDayInfoForYmd(incoming).ymd : bd.ymd;

  let expenses = [];
  try {
    expenses = db
      .prepare("SELECT * FROM expenses WHERE exp_date=? ORDER BY id DESC LIMIT 5000")
      .all(selectedYmd);
  } catch (e) {
    expenses = [];
  }

  const totalRow = db.prepare(`
    SELECT
      COALESCE(SUM(amount),0) as total,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='cash' THEN amount ELSE 0 END),0) as cash,
      COALESCE(SUM(CASE WHEN LOWER(COALESCE(method,''))='jazzcash' THEN amount ELSE 0 END),0) as jazz
    FROM expenses
    WHERE exp_date=?
  `).get(selectedYmd);

  const totalAll = db.prepare("SELECT COALESCE(SUM(amount),0) as s FROM expenses").get().s || 0;

  const payload = {
    ...viewData(req),
    expenses,
    today: selectedYmd,
    selectedDate: selectedYmd,
    isSalesPortal: true,

    totalExpensesToday: Number(totalRow && totalRow.total ? totalRow.total : 0),
    expCashToday: Number(totalRow && totalRow.cash ? totalRow.cash : 0),
    expJazzToday: Number(totalRow && totalRow.jazz ? totalRow.jazz : 0),

    totalExpensesAll: Number(totalAll || 0),
    totalExpenses: Number(totalAll || 0),

    msg: null,
  };

  return renderAny(
    res,
    ["sales_admin_expenses", "admin_expenses", "daily_expense", "expenses"],
    payload,
    {
      ok: true,
      selectedDate: selectedYmd,
      totalExpensesToday: payload.totalExpensesToday,
      expenses,
    }
  );
}

app.get("/sales_admin/expenses", requireSalesAdmin, salesAdminExpensesPage);app.get("/sales_admin/udhaar", requireSalesAdmin, (req, res) => res.redirect("/admin/udhaar"));
app.get("/sales_admin/pending_jazzcash", requireSalesAdmin, (req, res) => res.redirect("/admin/pending_jazzcash"));

/* Sales report main routes */
app.get("/sales_admin/report", requireSalesAdmin, salesReportPage);

app.post("/sales_admin/riders/create", requireSalesAdmin, (req, res) => {
  const name = String(pickBody(req, ["name", "rider_name"], "")).trim();
  if (!name) return safeRedirectBack(req, res, "/sales_admin/report");
  try {
    db.prepare("INSERT OR IGNORE INTO sa_riders (name, created_at, is_active) VALUES (?, ?, 1)").run(name, nowIso());
  } catch (e) {}
  return safeRedirectBack(req, res, "/sales_admin/report");
});

app.post("/sales_admin/riders/advance", requireSalesAdmin, (req, res) => {
  const riderId = Number(pickBody(req, ["rider_id"], 0)) || 0;
  const amount = Number(pickBody(req, ["amount"], 0)) || 0;
  const date = String(pickBody(req, ["date", "tx_date"], businessDayInfo().ymd)).trim();
  const method = String(pickBody(req, ["method"], "cash")).trim().toLowerCase();
  const note = String(pickBody(req, ["note"], "")).trim();
  if (!riderId || amount <= 0) return safeRedirectBack(req, res, "/sales_admin/report");
  try {
    db.prepare("INSERT INTO sa_rider_tx (rider_id, tx_date, tx_type, amount, method, note, created_at) VALUES (?, ?, 'advance', ?, ?, ?, ?)")
      .run(riderId, date, amount, (method === "jazzcash" ? "jazzcash" : "cash"), note, nowIso());
  } catch (e) {}
  return safeRedirectBack(req, res, "/sales_admin/report?date=" + encodeURIComponent(date));
});

app.post("/sales_admin/riders/return", requireSalesAdmin, (req, res) => {
  const riderId = Number(pickBody(req, ["rider_id"], 0)) || 0;
  const amount = Number(pickBody(req, ["amount"], 0)) || 0;
  const date = String(pickBody(req, ["date", "tx_date"], businessDayInfo().ymd)).trim();
  const method = String(pickBody(req, ["method"], "cash")).trim().toLowerCase();
  const note = String(pickBody(req, ["note"], "")).trim();
  if (!riderId || amount <= 0) return safeRedirectBack(req, res, "/sales_admin/report");
  try {
    db.prepare("INSERT INTO sa_rider_tx (rider_id, tx_date, tx_type, amount, method, note, created_at) VALUES (?, ?, 'return', ?, ?, ?, ?)")
      .run(riderId, date, amount, (method === "jazzcash" ? "jazzcash" : "cash"), note, nowIso());
  } catch (e) {}
  return safeRedirectBack(req, res, "/sales_admin/report?date=" + encodeURIComponent(date));
});

app.post("/sales_admin/employees/create", requireSalesAdmin, (req, res) => {
  const name = String(pickBody(req, ["name", "employee_name"], "")).trim();
  const salary = Number(pickBody(req, ["monthly_salary", "salary"], 0)) || 0;
  const salaryDay = Number(pickBody(req, ["salary_day"], 1)) || 1;
  if (!name) return safeRedirectBack(req, res, "/sales_admin/report");
  try {
    db.prepare("INSERT OR IGNORE INTO sa_employees (name, monthly_salary, salary_day, created_at, is_active) VALUES (?, ?, ?, ?, 1)")
      .run(name, salary, salaryDay, nowIso());
  } catch (e) {}
  return safeRedirectBack(req, res, "/sales_admin/report");
});

app.post("/sales_admin/employees/advance", requireSalesAdmin, (req, res) => {
  const employeeId = Number(pickBody(req, ["employee_id"], 0)) || 0;
  const amount = Number(pickBody(req, ["amount"], 0)) || 0;
  const date = String(pickBody(req, ["date", "adv_date"], businessDayInfo().ymd)).trim();
  const method = String(pickBody(req, ["method"], "cash")).trim().toLowerCase();
  const note = String(pickBody(req, ["note"], "")).trim();
  if (!employeeId || amount <= 0) return safeRedirectBack(req, res, "/sales_admin/report");
  try {
    db.prepare("INSERT INTO sa_employee_adv (employee_id, adv_date, amount, method, note, created_at) VALUES (?, ?, ?, ?, ?, ?)")
      .run(employeeId, date, amount, (method === "jazzcash" ? "jazzcash" : "cash"), note, nowIso());
  } catch (e) {}
  return safeRedirectBack(req, res, "/sales_admin/report?date=" + encodeURIComponent(date));
});

app.post("/sales_admin/employees/pay_salary", requireSalesAdmin, (req, res) => {
  const employeeId = Number(pickBody(req, ["employee_id"], 0)) || 0;
  const amount = Number(pickBody(req, ["amount"], 0)) || 0;
  const date = String(pickBody(req, ["date", "salary_date"], businessDayInfo().ymd)).trim();
  const method = String(pickBody(req, ["method"], "cash")).trim().toLowerCase();
  const note = String(pickBody(req, ["note"], "")).trim();

  if (!employeeId || amount <= 0) return safeRedirectBack(req, res, "/sales_admin/report");

  try {
    db.prepare(
      "INSERT INTO sa_employee_salary_tx (employee_id, salary_date, amount, method, note, created_at) VALUES (?, ?, ?, ?, ?, ?)"
    ).run(employeeId, date, amount, (method === "jazzcash" ? "jazzcash" : "cash"), note, nowIso());
  } catch (e) {}

  return safeRedirectBack(req, res, "/sales_admin/report?date=" + encodeURIComponent(date));
});

app.get("/sales_admin/compare", requireSalesAdmin, (req, res) => res.redirect("/sales_admin/report"));
app.get("/sales_admin/sales_compare", requireSalesAdmin, (req, res) => res.redirect("/sales_admin/report"));
app.get("/sales_admin/sales-compare", requireSalesAdmin, (req, res) => res.redirect("/sales_admin/report"));

app.post("/sales_admin/riders/:id/delete", requireSalesAdmin, (req, res) => {
  const id = Number(req.params.id || 0);
  if (!id) return res.redirect("/sales_admin/report");

  try {
    db.transaction(() => {
      db.prepare("DELETE FROM sa_rider_tx WHERE rider_id=?").run(id);
      db.prepare("DELETE FROM sa_riders WHERE id=?").run(id);
    })();
  } catch (e) {}

  return res.redirect("/sales_admin/report");
});

app.post("/sales_admin/employees/:id/delete", requireSalesAdmin, (req, res) => {
  const id = Number(req.params.id || 0);
  if (!id) return res.redirect("/sales_admin/report");

  try {
    db.transaction(() => {
      db.prepare("DELETE FROM sa_employee_adv WHERE employee_id=?").run(id);
      db.prepare("DELETE FROM sa_employee_salary_tx WHERE employee_id=?").run(id);
      db.prepare("DELETE FROM sa_employees WHERE id=?").run(id);
    })();
  } catch (e) {}

  return res.redirect("/sales_admin/report");
});

/* Old sales route compatibility */
app.get("/sales", requireSales, (req, res) => res.redirect("/sales_admin/report"));

/* Admin compare routes updated */
app.get("/admin/compare", requireAdminOrSalesAdmin, salesReportPage);
app.get("/admin/sales_compare", requireAdminOrSalesAdmin, salesReportPage);
app.get("/admin/sales-compare", requireAdminOrSalesAdmin, salesReportPage);
app.get("/admin/salescompare", requireAdminOrSalesAdmin, salesReportPage);
app.get("/admin/compare_sales", requireAdminOrSalesAdmin, salesReportPage);
app.get("/admin/compare-sales", requireAdminOrSalesAdmin, salesReportPage);
/* IMPORT */
app.get("/admin/import", requireStaff, (req, res) => {
  const expectedHeaders = ["CodeNo", "Name", "MobileNo", "Address", "location", "Balance", "OpeningBalance", "OpeningBottle", "BottlesBalance", "Area", "Password", "WaterTypeId", "WaterType", "UnitPrice"];

  let waterTypes = [];
  try {
    waterTypes = db.prepare("SELECT * FROM water_types ORDER BY id").all();
  } catch (e) {
    waterTypes = [];
  }

  return renderAny(
    res,
    ["admin_import"],
    { ...viewData(req), msg: null, expectedHeaders, customers: [], waterTypes, imported: [] },
    { ok: true }
  );
});
app.get("/admin/upload", requireStaff, (req, res) => res.redirect("/admin/import"));

function parseTableFile(filePath) {
  const buf = fs.readFileSync(filePath);
  const ext = (path.extname(filePath) || "").toLowerCase();

  if (ext === ".xlsx" || ext === ".xls") {
    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  }

  const text = buf.toString("latin1");
  const lines = text.split(/\r?\n/).filter((x) => x.trim() !== "");
  if (!lines.length) return [];

  const head = lines[0];
  let sep = ",";
  if (head.includes("\t")) sep = "\t";
  else if (head.includes(";") && !head.includes(",")) sep = ";";

  const headers = lines[0].split(sep).map((h) => h.trim());
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const parts = lines[i].split(sep);
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = (parts[j] || "").trim();
    rows.push(obj);
  }
  return rows;
}

function buildCustomerIdAllocator(dbRef) {
  let used = new Set();
  try {
    const rows = dbRef.prepare("SELECT CAST(id AS INTEGER) as id FROM customers WHERE CAST(id AS INTEGER) > 0").all();
    for (const r of rows) {
      const n = Number(r && r.id ? r.id : 0);
      if (n > 0) used.add(n);
    }
  } catch (e) {}

  let cursor = 1;
  while (used.has(cursor)) cursor++;

  return {
    take: () => {
      while (used.has(cursor)) cursor++;
      const id = cursor;
      used.add(id);
      cursor++;
      return String(id);
    },
  };
}

function nextReusableCustomerIdSingle(dbRef) {
  try {
    const rows = dbRef
      .prepare("SELECT CAST(id AS INTEGER) as id FROM customers WHERE CAST(id AS INTEGER) > 0 ORDER BY CAST(id AS INTEGER) ASC")
      .all();
    let expected = 1;
    for (const r of rows) {
      const n = Number(r && r.id ? r.id : 0);
      if (!n || n <= 0) continue;
      if (n < expected) continue;
      if (n === expected) {
        expected++;
        continue;
      }
      if (n > expected) break;
    }
    return expected;
  } catch (e) {
    return 1;
  }
}

function safeStr(v) {
  if (v === undefined || v === null) return "";
  return String(v);
}

app.post("/admin/import/customers", requireStaff, upload.any(), (req, res) => {
  const expectedHeaders = ["CodeNo", "Name", "MobileNo", "Address", "location", "Balance", "OpeningBalance", "OpeningBottle", "BottlesBalance", "Area", "Password", "WaterTypeId", "WaterType", "UnitPrice"];
  let waterTypes = [];

  const f = req.file || (req.files && req.files.length ? req.files[0] : null);
  if (!f) return res.redirect("/admin/import");
  try {
    const rows = parseTableFile(f.path);

    waterTypes = db.prepare("SELECT * FROM water_types ORDER BY id").all();
    const defaultWaterTypeId = waterTypes.length ? waterTypes[0].id : 1;
    const defaultUnitPrice = waterTypes.length ? Number(waterTypes[0].list_price || 0) : 0;

    let okCount = 0;
    let skipCount = 0;

    const alloc = buildCustomerIdAllocator(db);
    const seenInFile = new Set();

    db.transaction(() => {
      for (const r of rows) {
        const name = safeStr(r.Name || r.name || r.customer_name || "").trim();
        if (!name) {
          skipCount++;
          continue;
        }

        let rawId = safeStr(r.CodeNo || r.codeno || r.code || r.ID || r.id || "").trim();
        if (rawId && /^\d+$/.test(rawId)) rawId = String(Number(rawId));

        let id = rawId ? String(rawId) : alloc.take();

        if (seenInFile.has(String(id))) {
          skipCount++;
          continue;
        }
        seenInFile.add(String(id));

        const existingRow = db.prepare("SELECT * FROM customers WHERE CAST(id AS TEXT)=?").get(String(id));

        const phone = safeStr(r.MobileNo || r.phone || r.mobile || r.Mobile || "").trim();
        const address = safeStr(r.Address || r.address || "").trim();
        const mapUrl = safeStr(r.location || r.map_url || r.mapUrl || "").trim();
        const area = safeStr(r.Area || r.area || "").trim();

        const openingBalance = Number(r.OpeningBalance || r.Balance || 0) || 0;
        const openingBottle = Number(r.OpeningBottle || 0) || 0;
        const bottlesBalance = Number(r.BottlesBalance || 0) || 0;

        const passCell = safeStr(r.Password || r.password || r.pass || r.pin || "").trim();
        const pass = passCell ? passCell : String((existingRow && existingRow.password_plain) ? existingRow.password_plain : DEFAULT_CUSTOMER_PASSWORD);

        const waterTypeIdCell = safeStr(r.WaterTypeId || r.water_type_id || r.waterTypeId || "").trim();
        const waterNameCell = safeStr(r.WaterType || r.water_type || r.water_name || r.WaterName || r.water || "").trim();

        let waterTypeId = waterTypeIdCell !== "" ? (Number(waterTypeIdCell) || 0) : (existingRow ? Number(existingRow.water_type_id || 0) : 0);
        if (!waterTypeId && waterNameCell) {
          const found = waterTypes.find((w) => String(w.name || "").trim().toLowerCase() === String(waterNameCell).toLowerCase());
          if (found) waterTypeId = found.id;
        }
        if (!waterTypeId) waterTypeId = defaultWaterTypeId;

        const unitPriceCell = safeStr(r.UnitPrice || r.unit_price || r.price || r.rate || "").trim();
        let unitPrice = unitPriceCell !== "" ? (Number(unitPriceCell) || 0) : (existingRow ? Number(existingRow.unit_price || 0) : 0);
        if (!unitPrice) {
          const wrow = waterTypes.find((w) => Number(w.id) === Number(waterTypeId));
          unitPrice = wrow ? Number(wrow.list_price || 0) : 0;
        }
        if (!unitPrice) unitPrice = defaultUnitPrice;

        if (!existingRow) {
          db.prepare(
            `
            INSERT INTO customers
            (id, name, phone, address, map_url, area, water_type_id, unit_price, opening_balance, opening_bottle, bottles_balance, password_plain, created_at, is_active)
            VALUES
            (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
          `
          ).run(
            String(id),
            name,
            phone,
            address,
            mapUrl,
            area,
            waterTypeId,
            unitPrice,
            openingBalance,
            openingBottle,
            bottlesBalance,
            pass,
            nowIso()
          );

          if (openingBalance && openingBalance !== 0) {
            db.prepare(
              "INSERT INTO ledger_entries (customer_id, entry_type, ref_type, ref_id, amount, note, created_at) VALUES (?, 'opening', 'import', ?, ?, ?, ?)"
            ).run(String(id), "opening_balance", openingBalance, "Opening balance import", nowIso());
          }

          okCount++;
        } else {
          db.prepare(
            `
            UPDATE customers
            SET name=?,
                phone=?,
                address=?,
                map_url=?,
                area=?,
                water_type_id=?,
                unit_price=?,
                opening_balance=?,
                opening_bottle=?,
                bottles_balance=?,
                password_plain=?
            WHERE CAST(id AS TEXT)=?
          `
          ).run(
            name || String(existingRow.name || ""),
            phone || String(existingRow.phone || ""),
            address || String(existingRow.address || ""),
            mapUrl || String(existingRow.map_url || ""),
            area || String(existingRow.area || ""),
            waterTypeId || Number(existingRow.water_type_id || defaultWaterTypeId),
            unitPrice || Number(existingRow.unit_price || defaultUnitPrice),
            Number(openingBalance || 0),
            Number(openingBottle || 0),
            Number(bottlesBalance || 0),
            pass,
            String(id)
          );

          okCount++;
        }
      }
    })();

    try {
      fs.unlinkSync(f.path);
    } catch (e) {}

    return renderAny(
      res,
      ["admin_import"],
      { ...viewData(req), msg: `Imported ${okCount}, skipped ${skipCount}`, expectedHeaders, waterTypes, imported: [], customers: [] },
      { okCount, skipCount }
    );
  } catch (e) {
    try {
      fs.unlinkSync(f.path);
    } catch (x) {}

    return renderAny(
      res,
      ["admin_import"],
      { ...viewData(req), msg: "Import failed", expectedHeaders, waterTypes, imported: [], customers: [] },
      { ok: false }
    );
  }
});

/* CUSTOMERS with inactive list */
app.get("/admin/customers", requireStaff, (req, res) => {
  let customersAll = [];
  let waterTypes = [];

  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  const cutoffIso = sevenDaysAgo.toISOString();

  try {
    customersAll = db
      .prepare(
        `
        SELECT 
          c.*, 
          w.name as water_name, 
          w.list_price,
          (SELECT MAX(o.created_at) FROM orders o WHERE o.customer_id = c.id) as last_order_at
        FROM customers c
        LEFT JOIN water_types w ON w.id = c.water_type_id
        ORDER BY CAST(c.id AS INTEGER) DESC, c.id DESC
      `
      )
      .all();
  } catch (e) {
    customersAll = [];
  }

  try {
    waterTypes = db.prepare("SELECT * FROM water_types ORDER BY id").all();
  } catch (e) {
    waterTypes = [];
  }

  const activeCustomers = [];
  const inactiveCustomers = [];

  for (const c of customersAll) {
    const last = (c.last_order_at || "").toString();
    const isInactive = !last || last < cutoffIso;
    if (isInactive) inactiveCustomers.push(c);
    else activeCustomers.push(c);
  }

  const added = String(req.query.added || "").trim();

  return renderAny(
    res,
    ["admin_customers"],
    { ...viewData(req), customers: customersAll, activeCustomers, inactiveCustomers, waterTypes, msg: null, cutoffIso, added },
    { customers: customersAll, activeCustomers, inactiveCustomers, waterTypes, cutoffIso, added }
  );
});

app.get("/admin/customer", requireStaff, (req, res) => res.redirect("/admin/customers"));

/* customers export download */
app.get("/admin/customers/download", requireStaff, (req, res) => {
  try {
    const rows = db
      .prepare(
        `
        SELECT 
          c.*,
          w.name as water_name,
          w.list_price
        FROM customers c
        LEFT JOIN water_types w ON w.id = c.water_type_id
        ORDER BY CAST(c.id AS INTEGER) ASC, c.id ASC
      `
      )
      .all();

    const out = rows.map((c) => {
      const cid = String(c.id);
      const balance = Number(getOutstanding(db, cid) || 0);
      const pw = String((c.password_plain || "").trim() || DEFAULT_CUSTOMER_PASSWORD);

      return {
        CodeNo: cid,
        Name: c.name || "",
        MobileNo: c.phone || "",
        Address: c.address || "",
        location: c.map_url || "",
        Balance: balance,
        OpeningBalance: Number(c.opening_balance || 0),
        OpeningBottle: Number(c.opening_bottle || 0),
        BottlesBalance: Number(c.bottles_balance || 0),
        Area: c.area || "",
        Password: pw,
        WaterTypeId: Number(c.water_type_id || 0),
        WaterType: c.water_name || "",
        UnitPrice: Number(c.unit_price || 0),
      };
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(out, { skipHeader: false });
    XLSX.utils.book_append_sheet(wb, ws, "Customers");

    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=customers.xlsx");
    return res.send(buf);
  } catch (e) {
    return res.status(500).send("Export failed");
  }
});
app.get("/admin/customers/export", requireStaff, (req, res) => res.redirect("/admin/customers/download"));
app.get("/admin/customers/excel", requireStaff, (req, res) => res.redirect("/admin/customers/download"));

/* باقي سارا code bilkul same rahega jaisa tumhara tha */
app.get("/admin/customers/new", requireStaff, (req, res) => {
  const waterTypes = db.prepare("SELECT * FROM water_types ORDER BY id").all();
  return renderAny(res, ["admin_customer_new"], { ...viewData(req), waterTypes, msg: null }, { waterTypes });
});

app.get("/admin/customer/new", requireStaff, (req, res) => res.redirect("/admin/customers/new"));
app.get("/admin/add_customer", requireStaff, (req, res) => res.redirect("/admin/customers/new"));
app.get("/admin/new_customer", requireStaff, (req, res) => res.redirect("/admin/customers/new"));

function createCustomerFromBody(req) {
  let id = pickBody(req, ["id", "customer_id", "cid", "CodeNo"], "").toString().trim();
  if (!id) {
    try {
      const numeric = nextReusableCustomerIdSingle(db);
      id = String(numeric);
    } catch (e) {
      id = nextCustomerId(db, config.customerIdPrefix, config.customerIdPad);
    }
  }

  const name = pickBody(req, ["name", "customer_name", "Name"], "").toString().trim();
  const phone = pickBody(req, ["phone", "mobile", "MobileNo", "Mobile", "mobile_no", "mobileNo"], "").toString().trim();
  const area = pickBody(req, ["area", "Area"], "").toString().trim();
  const address = pickBody(req, ["address", "Address"], "").toString().trim();
  const mapUrl = pickBody(req, ["map_url", "mapUrl", "location", "map"], "").toString().trim();

  const waterTypeId = Number(pickBody(req, ["water_type_id", "waterTypeId", "water", "water_type"], 1)) || 1;
  const unitPrice = Number(pickBody(req, ["unit_price", "price", "rate", "unitPrice", "UnitPrice"], 0)) || 0;

  const openingBalance = Number(pickBody(req, ["opening_balance", "OpeningBalance", "Balance"], 0)) || 0;
  const openingBottle = Number(pickBody(req, ["opening_bottle", "OpeningBottle"], 0)) || 0;
  const bottlesBalance = Number(pickBody(req, ["bottles_balance", "BottlesBalance"], 0)) || 0;

  if (!name) return { ok: false, msg: "Name required" };

  db.prepare(
    `
    INSERT INTO customers
    (id, name, phone, address, map_url, area, water_type_id, unit_price, opening_balance, opening_bottle, bottles_balance, password_plain, created_at, is_active)
    VALUES
    (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
  `
  ).run(
    String(id),
    name,
    phone,
    address,
    mapUrl,
    area,
    waterTypeId,
    unitPrice,
    openingBalance,
    openingBottle,
    bottlesBalance,
    DEFAULT_CUSTOMER_PASSWORD,
    nowIso()
  );

  if (openingBalance && openingBalance !== 0) {
    db.prepare(
      "INSERT INTO ledger_entries (customer_id, entry_type, ref_type, ref_id, amount, note, created_at) VALUES (?, 'opening', 'import', ?, ?, ?, ?)"
    ).run(String(id), "opening_balance", openingBalance, "Opening balance import", nowIso());
  }

  return { ok: true, id: String(id) };
}

function createCustomerHandler(req, res) {
  try {
    const out = createCustomerFromBody(req);
    if (!out.ok) return res.status(400).send(out.msg);
    return res.redirect(303, "/admin/customers?added=" + encodeURIComponent(out.id));
  } catch (e) {
    return res.status(400).send("Customer create failed");
  }
}

app.post("/admin/customers/new", requireStaff, createCustomerHandler);
app.post("/admin/customers/add", requireStaff, createCustomerHandler);
app.post("/admin/customers/create", requireStaff, createCustomerHandler);
app.post("/admin/add_customer", requireStaff, createCustomerHandler);
app.post("/admin/customer/new", requireStaff, createCustomerHandler);

app.get("/admin/customers/:id/edit", requireStaff, (req, res) => {
  const cid = String(req.params.id);

  const customer = db
    .prepare("SELECT * FROM customers WHERE CAST(id AS TEXT)=?")
    .get(cid);

  if (!customer) return res.redirect("/admin/customers");

  const waterTypes = db.prepare("SELECT * FROM water_types ORDER BY id").all();

  const msg = String(req.query.msg || "").trim() || null;

  return renderAny(
    res,
    ["admin_customer_edit"],
    { ...viewData(req), customer, waterTypes, msg },
    { customer, waterTypes }
  );
});
app.get("/admin/customer/:id/edit", requireStaff, (req, res) => res.redirect("/admin/customers/" + req.params.id + "/edit"));

app.post("/admin/customers/:id/update", requireStaff, (req, res) => {
  const cid = String(req.params.id);

  const existing = db.prepare("SELECT * FROM customers WHERE CAST(id AS TEXT)=?").get(cid);
  if (!existing) return res.redirect("/admin/customers");

  const nameIn = pickBody(req, ["name", "customer_name", "Name"], "").toString().trim();
  const phoneIn = pickBody(req, ["phone", "mobile", "MobileNo", "Mobile", "mobile_no", "mobileNo"], "").toString().trim();
  const areaIn = pickBody(req, ["area", "Area"], "").toString().trim();
  const addressIn = pickBody(req, ["address", "Address"], "").toString().trim();
  const mapUrlIn = pickBody(req, ["map_url", "mapUrl", "location", "map"], "").toString().trim();
  const waterTypeIdIn = Number(pickBody(req, ["water_type_id", "waterTypeId", "water_type", "water"], "")) || 0;
  const unitPriceIn = Number(pickBody(req, ["unit_price", "price", "rate", "unitPrice", "UnitPrice"], "")) || 0;
  const isActiveIn = pickBody(req, ["is_active", "active"], null);
  const passIn = String(pickBody(req, ["password_plain", "password", "pass", "pin"], "")).trim();

  const name = nameIn || String(existing.name || "");
  const phone = phoneIn || String(existing.phone || "");
  const area = areaIn || String(existing.area || "");
  const address = addressIn || String(existing.address || "");
  const mapUrl = mapUrlIn || String(existing.map_url || "");
  const waterTypeId = waterTypeIdIn ? waterTypeIdIn : Number(existing.water_type_id || 1) || 1;
const unitPriceRaw = pickBody(req, ["unit_price", "price", "rate", "unitPrice", "UnitPrice"], null);

const unitPrice =
  unitPriceRaw !== null && String(unitPriceRaw).trim() !== ""
    ? Number(unitPriceIn || 0)
    : Number(existing.unit_price || 0);
  const isActive = (isActiveIn === null || isActiveIn === undefined) ? Number(existing.is_active || 1) : (Number(isActiveIn || 0) ? 1 : 0);

  try {
    if (passIn) {
      db.prepare(
        `
        UPDATE customers
        SET name=?,
            phone=?,
            address=?,
            map_url=?,
            area=?,
            water_type_id=?,
            unit_price=?,
            is_active=?,
            password_plain=?
        WHERE CAST(id AS TEXT)=?
      `
      ).run(name, phone, address, mapUrl, area, waterTypeId, unitPrice, isActive, passIn, cid);
    } else {
      db.prepare(
        `
        UPDATE customers
        SET name=?,
            phone=?,
            address=?,
            map_url=?,
            area=?,
            water_type_id=?,
            unit_price=?,
            is_active=?
        WHERE CAST(id AS TEXT)=?
      `
      ).run(name, phone, address, mapUrl, area, waterTypeId, unitPrice, isActive, cid);
    }
} catch (e) {
  console.error("Customer update failed:", e);
}
  return res.redirect("/admin/customers/" + cid + "/edit?msg=" + encodeURIComponent("Saved successfully"));
});

/* ... tumhara baqi routes same ... */
app.post("/admin/customers/:id/delete", requireStaff, (req, res) => {
  const cid = String(req.params.id);
  try {
    db.transaction(() => {
      db.prepare("DELETE FROM orders WHERE CAST(customer_id AS TEXT)=?").run(cid);
      db.prepare("DELETE FROM ledger_entries WHERE customer_id=?").run(cid);
      db.prepare("DELETE FROM customers WHERE CAST(id AS TEXT)=?").run(cid);
    })();
  } catch (e) {}
  return res.redirect("/admin/customers");
});

app.get("/admin/customers/:id", requireStaff, (req, res) => {
  const cid = String(req.params.id);

  const customer = db.prepare("SELECT * FROM customers WHERE CAST(id AS TEXT)=?").get(cid);
  if (!customer) return res.status(404).send("Customer not found");

  const outstanding = Number(getOutstanding(db, cid) || 0);

  let ledger = [];
  try {
    ledger = db.prepare("SELECT * FROM ledger_entries WHERE customer_id=? ORDER BY id DESC LIMIT 500").all(cid);
  } catch (e) {
    ledger = [];
  }

  let waterTypes = [];
  try {
    waterTypes = db.prepare("SELECT * FROM water_types ORDER BY id").all();
  } catch (e) {
    waterTypes = [];
  }

  return renderAny(res, ["admin_customer_view"], { ...viewData(req), customer, outstanding, ledger, waterTypes }, { customer, outstanding, ledger });
});

app.get("/admin/customer/:id", requireStaff, (req, res) => res.redirect("/admin/customers/" + req.params.id));

/* UDHAAR */
app.get("/admin/udhaar", requireStaff, (req, res) => {
  const modeIn = String(req.query.mode || "day").trim().toLowerCase();
  const mode = modeIn === "all" ? "all" : "day";

  const ymdIn = String(req.query.date || "").trim();
  const info = ymdIn ? businessDayInfoForYmd(ymdIn) : businessDayInfo();

  const selectedDate = info.ymd;

  const baseRows = db
    .prepare(
      `
      SELECT
        c.id,
        c.name,
        c.area,
        c.phone,
        COALESCE(lb.bal,0) as bal
      FROM customers c
      LEFT JOIN (
        SELECT customer_id, COALESCE(SUM(amount),0) as bal
        FROM ledger_entries
        GROUP BY customer_id
      ) lb ON CAST(lb.customer_id AS TEXT)=CAST(c.id AS TEXT)
      ORDER BY CAST(c.id AS INTEGER) ASC, c.id ASC
    `
    )
    .all();

  const dayOrders = db
    .prepare(
      `
      SELECT id, customer_id, unit_price, delivered_qty, quantity_requested
      FROM orders
      WHERE status='delivered'
        AND delivered_at >= ?
        AND delivered_at < ?
    `
    )
    .all(info.startIso, info.endIso);

  const orderIds = dayOrders.map((o) => String(o.id || "")).filter((x) => x);

  const saleByOrder = {};
  const payAllByOrder = {};
  const payDayByOrder = {};

  if (orderIds.length) {
    const ph = orderIds.map(() => "?").join(",");

    try {
      const saleRows = db
        .prepare(
          `
          SELECT ref_id as oid, COALESCE(SUM(amount),0) as s
          FROM ledger_entries
          WHERE entry_type='sale'
            AND ref_type='order'
            AND ref_id IN (${ph})
          GROUP BY ref_id
        `
        )
        .all(...orderIds);

      for (const r of saleRows) saleByOrder[String(r.oid)] = Number(r.s || 0);
    } catch (e) {}

    try {
      const payAllRows = db
        .prepare(
          `
          SELECT ref_id as oid, COALESCE(SUM(amount),0) as s
          FROM ledger_entries
          WHERE entry_type='payment'
            AND ref_type='order'
            AND ref_id IN (${ph})
          GROUP BY ref_id
        `
        )
        .all(...orderIds);

      for (const r of payAllRows) payAllByOrder[String(r.oid)] = Number(r.s || 0);
    } catch (e) {}

    try {
      const payDayRows = db
        .prepare(
          `
          SELECT ref_id as oid, COALESCE(SUM(amount),0) as s
          FROM ledger_entries
          WHERE entry_type='payment'
            AND ref_type='order'
            AND ref_id IN (${ph})
            AND created_at >= ?
            AND created_at < ?
          GROUP BY ref_id
        `
        )
        .all(...orderIds, info.startIso, info.endIso);

      for (const r of payDayRows) payDayByOrder[String(r.oid)] = Number(r.s || 0);
    } catch (e) {}
  }

  const dayTotalByCustomer = {};
  const dayRemainByCustomer = {};

  for (const o of dayOrders) {
    const oid = String(o.id || "");
    const cid = String(o.customer_id || "");

    const qty =
      Number(o.delivered_qty || 0) > 0
        ? Number(o.delivered_qty || 0)
        : Number(o.quantity_requested || 0);

    const fallbackSale = Number(o.unit_price || 0) * Number(qty || 0);
    const sale = saleByOrder[oid] !== undefined ? Number(saleByOrder[oid] || 0) : fallbackSale;

    const payAllAbs = Math.max(0, -Number(payAllByOrder[oid] || 0));
    const payDayAbs = Math.max(0, -Number(payDayByOrder[oid] || 0));

    const initialDue = Math.max(0, sale - payDayAbs);
    const currentDue = Math.max(0, sale - payAllAbs);

    if (initialDue > 0) dayTotalByCustomer[cid] = (dayTotalByCustomer[cid] || 0) + initialDue;
    if (currentDue > 0) dayRemainByCustomer[cid] = (dayRemainByCustomer[cid] || 0) + currentDue;
  }

  const manualPaidAbsByCustomer = {};
  try {
    const manualRows = db
      .prepare(
        `
        SELECT customer_id, COALESCE(SUM(amount),0) as s
        FROM ledger_entries
        WHERE entry_type='payment'
          AND amount < 0
          AND ref_type IN ('udhaar_day','udhaar')
          AND ref_id = ?
        GROUP BY customer_id
      `
      )
      .all(selectedDate);

    for (const r of manualRows) {
      manualPaidAbsByCustomer[String(r.customer_id)] = Math.max(0, -Number(r.s || 0));
    }
  } catch (e) {}

  const rows = baseRows.map((r) => {
    const cid = String(r.id || "");
    const totalOutstanding = Math.max(0, Number(r.bal || 0));

    const dayTotal = Math.max(0, Number(dayTotalByCustomer[cid] || 0));

    let dayRemain = Math.max(0, Number(dayRemainByCustomer[cid] || 0));
    const manualPaid = Math.max(0, Number(manualPaidAbsByCustomer[cid] || 0));

    if (manualPaid > 0) dayRemain = Math.max(0, dayRemain - manualPaid);

    if (totalOutstanding <= 0) dayRemain = 0;
    else dayRemain = Math.min(dayRemain, totalOutstanding);

    const prev = Math.max(0, totalOutstanding - dayRemain);

    const msg =
      `${config.appName}\n` +
      `Payment reminder\n` +
      `Total pending: ${totalOutstanding}\n` +
      `Selected day pending: ${dayRemain}\n` +
      `Previous pending: ${prev}\n` +
      `Please clear pending amount. Thank you`;

    return {
      id: r.id,
      name: r.name || "",
      area: r.area || "",
      phone: r.phone || "",
      outstanding_total: totalOutstanding,
      outstanding_day: dayRemain,
      day_total: dayTotal,
      outstanding: totalOutstanding,
      wa_url: waLink(r.phone || "", msg),
    };
  });

  const customersDay = rows.filter((c) => Number(c.outstanding_day || 0) > 0);
  const customersAll = rows.filter((c) => Number(c.outstanding_total || 0) > 0);

  const totalUdhaarDayTotal = rows.reduce((a, c) => a + Number(c.day_total || 0), 0);
  const totalUdhaarDayRemaining = rows.reduce((a, c) => a + Number(c.outstanding_day || 0), 0);
  const totalUdhaarDayCleared = Math.max(0, totalUdhaarDayTotal - totalUdhaarDayRemaining);

  const totalUdhaarAll = customersAll.reduce((a, c) => a + Number(c.outstanding_total || 0), 0);

  return renderAny(
    res,
    ["admin_udhaar"],
    {
      ...viewData(req),
      selectedDate,
      mode,
      totalUdhaarDay: totalUdhaarDayTotal,
      totalUdhaarDayTotal,
      totalUdhaarDayRemaining,
      totalUdhaarDayCleared,
      totalUdhaarAll,
      customersDay,
      customersAll,
      customers: customersAll,
    },
    {
      selectedDate,
      mode,
      totalUdhaarDay: totalUdhaarDayTotal,
      totalUdhaarDayTotal,
      totalUdhaarDayRemaining,
      totalUdhaarDayCleared,
      totalUdhaarAll,
      customersDay,
      customersAll,
    }
  );
});
app.get("/admin/udhaar_list", requireStaff, (req, res) => res.redirect("/admin/udhaar"));
app.get("/admin/udhaar-list", requireStaff, (req, res) => res.redirect("/admin/udhaar"));

function redirectUdhaarBack(res, mode, date) {
  const m = String(mode || "day");
  const d = String(date || "");
  const q = [];
  if (m) q.push("mode=" + encodeURIComponent(m));
  if (d) q.push("date=" + encodeURIComponent(d));
  const url = "/admin/udhaar" + (q.length ? "?" + q.join("&") : "");
  return res.redirect(url);
}

/* NEW receive routes for your new page */
app.post("/admin/udhaar/receive_full", requireStaff, (req, res) => {
  const customerId = String(pickBody(req, ["customer_id", "cid", "id"], "")).trim();
  const method = String(pickBody(req, ["method", "payment_method"], "cash")).trim().toLowerCase();
  const mode = String(pickBody(req, ["mode"], "day")).trim().toLowerCase();
  const date = String(pickBody(req, ["date"], "")).trim();

  const amtIn = Number(pickBody(req, ["amount"], 0)) || 0;
  if (!customerId || amtIn <= 0) return redirectUdhaarBack(res, mode, date);

  const curOutstanding = Number(getOutstanding(db, String(customerId)) || 0);
  const apply = Math.min(curOutstanding > 0 ? curOutstanding : 0, amtIn);

  if (apply <= 0) return redirectUdhaarBack(res, mode, date);

const dateRef = date || businessDayInfo().ymd;
receiveUdhaar(customerId, apply, method, "Udhaar received full", "udhaar_day", dateRef);
  return redirectUdhaarBack(res, mode, date);
});

app.post("/admin/udhaar/receive_partial", requireStaff, (req, res) => {
  const customerId = String(pickBody(req, ["customer_id", "cid", "id"], "")).trim();
  const method = String(pickBody(req, ["method", "payment_method"], "cash")).trim().toLowerCase();
  const mode = String(pickBody(req, ["mode"], "day")).trim().toLowerCase();
  const date = String(pickBody(req, ["date"], "")).trim();

  const amtIn = Number(pickBody(req, ["amount"], 0)) || 0;
  if (!customerId || amtIn <= 0) return redirectUdhaarBack(res, mode, date);

  const curOutstanding = Number(getOutstanding(db, String(customerId)) || 0);
  const apply = Math.min(curOutstanding > 0 ? curOutstanding : 0, amtIn);

  if (apply <= 0) return redirectUdhaarBack(res, mode, date);

const dateRef = date || businessDayInfo().ymd;
receiveUdhaar(customerId, apply, method, "Udhaar received partial", "udhaar_day", dateRef);  return redirectUdhaarBack(res, mode, date);
});

/* old route keep for compatibility */
app.post("/admin/udhaar/receive", requireStaff, (req, res) => {
  const customerId = pickBody(req, ["customer_id", "cid", "id"], "").toString().trim();
  const amount = Number(pickBody(req, ["amount", "received", "pay", "payment_amount"], 0)) || 0;
  const method = pickBody(req, ["method", "payment_method"], "cash").toString().trim();
  const note = pickBody(req, ["note", "details"], "").toString().trim();

const date = String(pickBody(req, ["date"], businessDayInfo().ymd)).trim();
const dateRef = date || businessDayInfo().ymd;
const out = receiveUdhaar(customerId, amount, method, note, "udhaar_day", dateRef);  if (wantsJsonReq(req)) {
    return res.json({ ok: !!out.ok, msg: out.msg || null, outstanding: Number(getOutstanding(db, String(customerId)) || 0) });
  }
  return safeRedirectBack(req, res, "/admin/udhaar");
});
app.post("/admin/udhaar_receive", requireStaff, (req, res) => res.redirect(307, "/admin/udhaar/receive"));
/* CUSTOMER */
app.post("/customer/order", requireRole("customer"), (req, res) => {
  const cid = String(req.session.user.id);
  let qty = safeInt(pickBody(req, ["quantity", "qty", "bottles"], 0), 0);
  if (!qty || qty < 1) qty = 1;
  if (qty > 100) qty = 100;

  const data = db
    .prepare(
      `
      SELECT c.unit_price, c.water_type_id, w.list_price
      FROM customers c
      JOIN water_types w ON w.id = c.water_type_id
      WHERE CAST(c.id AS TEXT) = ?
    `
    )
    .get(cid);

  if (!data) return res.redirect("/customer");

  db.prepare(
    `
    INSERT INTO orders
    (customer_id, water_type_id, unit_price, list_price, quantity_requested, status, rider_id, created_at, created_source)
    VALUES
    (?, ?, ?, ?, ?, 'pending', NULL, ?, 'customer')  `
).run(cid, data.water_type_id, data.unit_price, data.list_price, qty, nowIso());

try {
  db.prepare(
    "DELETE FROM inactive_customer_status WHERE CAST(customer_id AS TEXT)=CAST(? AS TEXT)"
  ).run(String(cid));
} catch (e) {}

return res.redirect("/customer"); });
app.get("/customer", requireRole("customer"), (req, res) => {
  const cid = String(req.session.user.id);

  const customer = db
    .prepare(
      `
      SELECT c.*, w.name as water_name, w.list_price
      FROM customers c
      JOIN water_types w ON w.id = c.water_type_id
      WHERE CAST(c.id AS TEXT) = ?
    `
    )
    .get(cid);

  let recentOrders = [];
  try {
    recentOrders = db.prepare("SELECT * FROM orders WHERE customer_id=? ORDER BY id DESC LIMIT 50").all(cid);
  } catch (e) {
    recentOrders = [];
  }

  let ledger = [];
  try {
    ledger = db.prepare("SELECT * FROM ledger_entries WHERE customer_id=? ORDER BY id DESC LIMIT 200").all(cid);
  } catch (e) {
    ledger = [];
  }

  const delivered_total = db
    .prepare(
      `
    SELECT COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as b
    FROM orders
    WHERE customer_id=? AND status='delivered'
  `
    )
    .get(cid).b;

  const returned_total = db
    .prepare(
      `
    SELECT COALESCE(SUM(COALESCE(empty_returned_qty,0)),0) as r
    FROM orders
    WHERE customer_id=? AND status='delivered'
  `
    )
    .get(cid).r;

  const spent_total = db
    .prepare(
      `
    SELECT COALESCE(SUM(unit_price * CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as s
    FROM orders
    WHERE customer_id=? AND status='delivered'
  `
    )
    .get(cid).s;

  const monthlyConsumption = db
    .prepare(
      `
    SELECT substr(delivered_at,1,7) as ym,
           COALESCE(SUM(CASE WHEN COALESCE(delivered_qty,0) > 0 THEN delivered_qty ELSE quantity_requested END),0) as qty
    FROM orders
    WHERE customer_id=? AND status='delivered' AND delivered_at IS NOT NULL AND delivered_at != ''
    GROUP BY substr(delivered_at,1,7)
    ORDER BY ym DESC
    LIMIT 24
  `
    )
    .all(cid);

  return renderAny(
    res,
    ["customer_dashboard"],
    {
      ...viewData(req),
      customer,
      stats: {
        delivered_total: Number(delivered_total || 0),
        returned_total: Number(returned_total || 0),
        spent_total: Number(spent_total || 0),
      },
      recentOrders,
      outstanding: getOutstanding(db, cid),
      discountPer: null,
      ledger,
      monthlyConsumption,
    },
    { customer, recentOrders, ledger }
  );
});

/* 404 */
app.use((req, res) => {
  res.status(404).send("Route not found");
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server running on http://localhost:${port}`));     
