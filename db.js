const Database = require("better-sqlite3");
const bcrypt = require("bcryptjs");

function nowIso() {
  return new Date().toISOString();
}

function tableCols(db, tableName) {
  try {
    return db
      .prepare(`PRAGMA table_info(${tableName})`)
      .all()
      .map((r) => r.name);
  } catch (e) {
    return [];
  }
}

function ensureColumn(db, tableName, colName, ddl) {
  try {
    const cols = tableCols(db, tableName);
    if (!cols.includes(colName)) db.exec(ddl);
  } catch (e) {}
}

function ensureWaterTypesSeed(db) {
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
    ensureColumn(db, "water_types", "list_price", "ALTER TABLE water_types ADD COLUMN list_price INTEGER NOT NULL DEFAULT 0");
  } catch (e) {}

  const upsert = (name, price) => {
    try {
      const ex = db.prepare("SELECT id, list_price FROM water_types WHERE LOWER(name)=LOWER(?)").get(name);
      if (!ex) {
        db.prepare("INSERT INTO water_types (name, list_price) VALUES (?, ?)").run(name, Number(price || 0));
        return;
      }
      const cur = Number(ex.list_price || 0) || 0;
      if (!cur || cur === 0) db.prepare("UPDATE water_types SET list_price=? WHERE id=?").run(Number(price || 0), ex.id);
    } catch (e) {}
  };

  upsert("Classic Purified", 110);
  upsert("Mineral Boosted", 130);
  upsert("Ozonated Mineral Boosted", 160);
}

/*
  NEW
  Smallest available numeric customer id
  If 1 2 3 5 exist then it returns 4
*/
function nextAvailableNumericCustomerId(db) {
  try {
    const rows = db
      .prepare(
        `
        SELECT CAST(id AS INTEGER) as n
        FROM customers
        WHERE TRIM(COALESCE(id,'')) != ''
          AND id GLOB '[0-9]*'
        ORDER BY n ASC
      `
      )
      .all();

    let expected = 1;
    for (const r of rows) {
      const n = Number(r && r.n ? r.n : 0);
      if (!n || n < 1) continue;
      if (n === expected) {
        expected += 1;
        continue;
      }
      if (n > expected) break;
    }
    return String(expected);
  } catch (e) {
    return "1";
  }
}

function initDb(dbFile) {
  const path = require("path");
  const fs = require("fs");

  const dir = path.dirname(dbFile);
  if (dir && dir !== "." && !fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  const db = new Database(dbFile);
  db.pragma("journal_mode = WAL");

  db.exec(`
    CREATE TABLE IF NOT EXISTS meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS customers (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      phone TEXT,
      address TEXT,
      map_url TEXT,
      area TEXT,
      water_type_id INTEGER NOT NULL DEFAULT 1,
      unit_price INTEGER NOT NULL DEFAULT 0,
      opening_balance INTEGER NOT NULL DEFAULT 0,
      opening_bottle INTEGER NOT NULL DEFAULT 0,
      bottles_balance INTEGER NOT NULL DEFAULT 0,
      pin_hash TEXT,
      password_plain TEXT,
      created_at TEXT NOT NULL,
      is_active INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS riders (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      phone TEXT,
      password_plain TEXT,
      created_at TEXT NOT NULL,
      is_active INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS orders (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      customer_id TEXT NOT NULL,
      water_type_id INTEGER NOT NULL,
      unit_price INTEGER NOT NULL,
      list_price INTEGER NOT NULL,
      quantity_requested INTEGER NOT NULL,
      status TEXT NOT NULL,
      rider_id TEXT,
      created_at TEXT NOT NULL,

      delivered_at TEXT,
      delivered_qty INTEGER,
      empty_returned_qty INTEGER,

      payment_method TEXT,
      payment_amount INTEGER,
      payment_received INTEGER NOT NULL DEFAULT 0,
      payment_received_at TEXT,

      note TEXT
    );

    CREATE TABLE IF NOT EXISTS ledger_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      customer_id TEXT NOT NULL,
      entry_type TEXT NOT NULL,
      ref_type TEXT,
      ref_id TEXT,
      amount INTEGER NOT NULL,
      method TEXT,
      note TEXT,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS expenses (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      exp_date TEXT NOT NULL,
      category TEXT NOT NULL,
      amount INTEGER NOT NULL,
      method TEXT NOT NULL,
      staff_type TEXT,
      staff_id TEXT,
      is_advance INTEGER NOT NULL DEFAULT 0,
      note TEXT,
      created_at TEXT NOT NULL
    );
  `);

  try {
    ensureColumn(db, "customers", "map_url", "ALTER TABLE customers ADD COLUMN map_url TEXT");
    ensureColumn(db, "customers", "area", "ALTER TABLE customers ADD COLUMN area TEXT");
    ensureColumn(db, "customers", "opening_balance", "ALTER TABLE customers ADD COLUMN opening_balance INTEGER NOT NULL DEFAULT 0");
    ensureColumn(db, "customers", "opening_bottle", "ALTER TABLE customers ADD COLUMN opening_bottle INTEGER NOT NULL DEFAULT 0");
    ensureColumn(db, "customers", "bottles_balance", "ALTER TABLE customers ADD COLUMN bottles_balance INTEGER NOT NULL DEFAULT 0");
    ensureColumn(db, "customers", "pin_hash", "ALTER TABLE customers ADD COLUMN pin_hash TEXT");
    ensureColumn(db, "customers", "password_plain", "ALTER TABLE customers ADD COLUMN password_plain TEXT");
    ensureColumn(db, "customers", "is_active", "ALTER TABLE customers ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1");
    ensureColumn(db, "customers", "water_type_id", "ALTER TABLE customers ADD COLUMN water_type_id INTEGER NOT NULL DEFAULT 1");
    ensureColumn(db, "customers", "unit_price", "ALTER TABLE customers ADD COLUMN unit_price INTEGER NOT NULL DEFAULT 0");
  } catch (e) {}

  try {
    ensureColumn(db, "riders", "phone", "ALTER TABLE riders ADD COLUMN phone TEXT");
    ensureColumn(db, "riders", "password_plain", "ALTER TABLE riders ADD COLUMN password_plain TEXT");
    ensureColumn(db, "riders", "is_active", "ALTER TABLE riders ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1");
  } catch (e) {}

  try {
    ensureColumn(db, "orders", "delivered_at", "ALTER TABLE orders ADD COLUMN delivered_at TEXT");
    ensureColumn(db, "orders", "delivered_qty", "ALTER TABLE orders ADD COLUMN delivered_qty INTEGER");
    ensureColumn(db, "orders", "empty_returned_qty", "ALTER TABLE orders ADD COLUMN empty_returned_qty INTEGER");
    ensureColumn(db, "orders", "payment_method", "ALTER TABLE orders ADD COLUMN payment_method TEXT");
    ensureColumn(db, "orders", "payment_amount", "ALTER TABLE orders ADD COLUMN payment_amount INTEGER");
    ensureColumn(db, "orders", "payment_received", "ALTER TABLE orders ADD COLUMN payment_received INTEGER NOT NULL DEFAULT 0");
    ensureColumn(db, "orders", "payment_received_at", "ALTER TABLE orders ADD COLUMN payment_received_at TEXT");
    ensureColumn(db, "orders", "note", "ALTER TABLE orders ADD COLUMN note TEXT");
  } catch (e) {}

  try {
    ensureColumn(db, "ledger_entries", "method", "ALTER TABLE ledger_entries ADD COLUMN method TEXT");
    ensureColumn(db, "ledger_entries", "note", "ALTER TABLE ledger_entries ADD COLUMN note TEXT");
  } catch (e) {}

  try {
    ensureColumn(db, "expenses", "staff_type", "ALTER TABLE expenses ADD COLUMN staff_type TEXT");
    ensureColumn(db, "expenses", "staff_id", "ALTER TABLE expenses ADD COLUMN staff_id TEXT");
    ensureColumn(db, "expenses", "is_advance", "ALTER TABLE expenses ADD COLUMN is_advance INTEGER NOT NULL DEFAULT 0");
    ensureColumn(db, "expenses", "note", "ALTER TABLE expenses ADD COLUMN note TEXT");
  } catch (e) {}

  ensureWaterTypesSeed(db);

  try {
    const adminCount = db.prepare("SELECT COUNT(*) AS c FROM users").get().c || 0;
    if (adminCount === 0) {
      const hash = bcrypt.hashSync("admin123", 10);
      db.prepare("INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)").run("admin", hash, nowIso());
    }
  } catch (e) {}

  try {
    const seq = db.prepare("SELECT value FROM meta WHERE key='customer_seq'").get();
    if (!seq) db.prepare("INSERT INTO meta (key, value) VALUES ('customer_seq','1')").run();
  } catch (e) {}

  return db;
}

function nextCustomerId(db, prefix, pad) {
  const row = db.prepare("SELECT value FROM meta WHERE key='customer_seq'").get();
  const n = parseInt(String(row && row.value ? row.value : "1"), 10) || 1;
  const id = String(prefix || "") + String(n).padStart(Number(pad || 4), "0");
  db.prepare("UPDATE meta SET value=? WHERE key='customer_seq'").run(String(n + 1));
  return id;
}

function getOutstanding(db, customerId) {
  const row = db
    .prepare(
      `
      SELECT COALESCE(SUM(amount),0) AS bal
      FROM ledger_entries
      WHERE customer_id = ?
    `
    )
    .get(String(customerId));
  return Number(row && row.bal ? row.bal : 0) || 0;
}

module.exports = {
  initDb,
  nextCustomerId,
  nextAvailableNumericCustomerId,
  getOutstanding
};
