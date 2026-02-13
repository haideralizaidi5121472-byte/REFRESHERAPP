const crypto = require("crypto");
const config = require("./config");

function money(n) {
  return Number(Number(n || 0).toFixed(2));
}

function safeInt(v, def = 0) {
  const n = parseInt(v, 10);
  return Number.isNaN(n) ? def : n;
}

function getBusinessDate(dateObj = new Date()) {
  const adjusted = new Date(dateObj.getTime() - 2 * 60 * 60 * 1000);
  return adjusted;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatYmd(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function parseYmd(ymd) {
  const s = String(ymd || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  const parts = s.split("-").map((x) => parseInt(x, 10));
  const y = parts[0];
  const m = parts[1];
  const da = parts[2];
  if (!y || !m || !da) return null;

  const d = new Date(y, m - 1, da, 12, 0, 0, 0);
  if (Number.isNaN(d.getTime())) return null;
  return d;
}

function todayYmd(dateObj) {
  const d = getBusinessDate(dateObj || new Date());
  return formatYmd(d);
}

function ymdOffset(a, b) {
  if (typeof b === "number") {
    const base = parseYmd(a);
    const d = base ? new Date(base.getFullYear(), base.getMonth(), base.getDate(), 12, 0, 0, 0) : getBusinessDate();
    d.setDate(d.getDate() + b);
    return formatYmd(d);
  }

  const days = Number(a || 0);
  const d = getBusinessDate();
  d.setDate(d.getDate() + days);
  return formatYmd(d);
}

function normalizePkPhone(p) {
  if (!p) return "";
  let s = String(p).replace(/\D/g, "");
  if (s.startsWith("92")) return s;
  if (s.startsWith("03")) return "92" + s.substring(1);
  if (s.startsWith("3")) return "92" + s;
  return s;
}

function requireRole(role) {
  return (req, res, next) => {
    if (!req.session.user) return res.redirect("/login/customer");

    if (role === "admin" && req.session.user.role === "owner") return next();

    if (req.session.user.role !== role) return res.redirect("/login/" + role);

    next();
  };
}

module.exports = {
  money,
  safeInt,
  todayYmd,
  ymdOffset,
  normalizePkPhone,
  requireRole,
  getBusinessDate,
};
