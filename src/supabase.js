// ═══════════════════════════════════════════════════════════════
// SUPABASE CLIENT — TALA COSTING
// ═══════════════════════════════════════════════════════════════
const SUPABASE_URL = "https://jdeasptreourrqkjwgdv.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkZWFzcHRyZW91cnJxa2p3Z2R2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE1MzIyOTEsImV4cCI6MjA5NzEwODI5MX0._628XK35-jwGUBbH1XVTzOp4CjHNPBdu_TsYvaRSVsE";

const headers = {
  "Content-Type": "application/json",
  "apikey": SUPABASE_KEY,
  "Authorization": `Bearer ${SUPABASE_KEY}`,
  "Prefer": "return=representation"
};

// ── Generic fetch helpers ──────────────────────────────────────
const sbGet = async (table) => {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?select=*`, { headers });
  if (!res.ok) throw new Error(`GET ${table} failed: ${res.status}`);
  return res.json();
};

const sbUpsert = async (table, data) => {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
    method: "POST",
    headers: { ...headers, "Prefer": "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify(data)
  });
  if (!res.ok) throw new Error(`UPSERT ${table} failed: ${res.status}`);
};

// ── Load all data from Supabase ────────────────────────────────
export const loadFromCloud = async () => {
  const [rawRow, prepRow, prodsRow, clsRow, usersRow, modsRow, salesRow, mpRow] =
    await Promise.all([
      sbGet("tc_raw"),
      sbGet("tc_prep"),
      sbGet("tc_products"),
      sbGet("tc_classes"),
      sbGet("tc_users"),
      sbGet("tc_modifiers"),
      sbGet("tc_sales"),
      sbGet("tc_monthly_prices"),
    ]);

  const pick = (rows, def) => rows.length > 0 ? rows[0].data : def;

  return {
    raw:          pick(rawRow,   []),
    prep:         pick(prepRow,  []),
    products:     pick(prodsRow, []),
    classes:      pick(clsRow,   null),
    users:        pick(usersRow, null),
    modifiers:    pick(modsRow,  []),
    sales:        pick(salesRow, []),
    monthlyPrices:pick(mpRow,    {}),
  };
};

// ── Save one key to Supabase ───────────────────────────────────
export const saveToCloud = async (key, data) => {
  const tableMap = {
    raw:           "tc_raw",
    prep:          "tc_prep",
    products:      "tc_products",
    classes:       "tc_classes",
    users:         "tc_users",
    modifiers:     "tc_modifiers",
    sales:         "tc_sales",
    monthlyPrices: "tc_monthly_prices",
  };
  const table = tableMap[key];
  if (!table) return;
  await sbUpsert(table, { id: 1, data });
};

// updated Tue Jun 16 23:01:02 UTC 2026
