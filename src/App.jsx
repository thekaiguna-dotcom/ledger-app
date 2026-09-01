import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

const CATS = ["FOOD", "TRAVEL", "SNACKS", "SHOP", "BILLS", "MISC"];
const CAT_LABEL = { FOOD: "Food", TRAVEL: "Travel", SNACKS: "Snacks", SHOP: "Shop", BILLS: "Bills", MISC: "Misc" };
const PRESETS = [10, 20, 50, 100, 200];
const LS_ENTRIES = "ledger.entries";
const LS_SETTINGS = "ledger.settings";
const LS_OWES = "ledger.owes";

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
function startOfDay(ts) {
  const d = new Date(ts);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}
function startOfWeek(ts) {
  const d = new Date(ts);
  const day = d.getDay();
  const diff = day === 0 ? 6 : day - 1;
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - diff);
  return d.getTime();
}
function startOfMonth(ts) {
  const d = new Date(ts);
  d.setHours(0, 0, 0, 0);
  d.setDate(1);
  return d.getTime();
}
function fmtTime(ts) {
  return new Date(ts).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", hour12: true });
}
function fmtDateShort(ts) {
  return new Date(ts).toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
}
function money(n) {
  return "\u20B9" + n.toLocaleString("en-IN", { maximumFractionDigits: 0 });
}

function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

export default function App() {
  const [tab, setTab] = useState("ledger"); // ledger | owe

  // ---- Ledger state ----
  const [entries, setEntries] = useState(() => loadJSON(LS_ENTRIES, []));
  const [budget, setBudget] = useState(() => loadJSON(LS_SETTINGS, {}).dailyBudget ?? null);
  const [amountInput, setAmountInput] = useState("");
  const [selectedCat, setSelectedCat] = useState(null);
  const [note, setNote] = useState("");
  const [range, setRange] = useState("today");
  const [showSettings, setShowSettings] = useState(false);
  const [budgetInput, setBudgetInput] = useState("");
  const [flash, setFlash] = useState(false);
  const [error, setError] = useState("");
  const amountRef = useRef(null);

  // ---- Owe state ----
  const [owes, setOwes] = useState(() => loadJSON(LS_OWES, []));
  const [owePersonInput, setOwePersonInput] = useState("");
  const [oweAmountInput, setOweAmountInput] = useState("");
  const [oweNoteInput, setOweNoteInput] = useState("");
  const [oweError, setOweError] = useState("");
  const owePersonRef = useRef(null);

  useEffect(() => { localStorage.setItem(LS_ENTRIES, JSON.stringify(entries)); }, [entries]);
  useEffect(() => { localStorage.setItem(LS_SETTINGS, JSON.stringify({ dailyBudget: budget })); }, [budget]);
  useEffect(() => { localStorage.setItem(LS_OWES, JSON.stringify(owes)); }, [owes]);

  function addEntry() {
    const amt = parseFloat(amountInput);
    if (!amt || amt <= 0) {
      setError("Enter an amount first.");
      setTimeout(() => setError(""), 1800);
      return;
    }
    if (!selectedCat) {
      setError("Pick a category.");
      setTimeout(() => setError(""), 1800);
      return;
    }
    const entry = { id: uid(), amount: amt, category: selectedCat, note: note.trim(), ts: Date.now() };
    setEntries((prev) => [entry, ...prev]);
    setAmountInput("");
    setSelectedCat(null);
    setNote("");
    setFlash(true);
    setTimeout(() => setFlash(false), 500);
    if (amountRef.current) amountRef.current.focus();
  }

  function deleteEntry(id) {
    setEntries((prev) => prev.filter((e) => e.id !== id));
  }

  function saveBudget() {
    const v = parseFloat(budgetInput);
    setBudget(v > 0 ? v : null);
    setShowSettings(false);
  }

  function addOwe() {
    const amt = parseFloat(oweAmountInput);
    if (!owePersonInput.trim()) {
      setOweError("Enter a name.");
      setTimeout(() => setOweError(""), 1800);
      return;
    }
    if (!amt || amt <= 0) {
      setOweError("Enter an amount.");
      setTimeout(() => setOweError(""), 1800);
      return;
    }
    const owe = { id: uid(), person: owePersonInput.trim(), amount: amt, note: oweNoteInput.trim(), ts: Date.now(), settled: false };
    setOwes((prev) => [owe, ...prev]);
    setOwePersonInput("");
    setOweAmountInput("");
    setOweNoteInput("");
    if (owePersonRef.current) owePersonRef.current.focus();
  }

  function toggleSettled(id) {
    setOwes((prev) => prev.map((o) => (o.id === id ? { ...o, settled: !o.settled } : o)));
  }

  function deleteOwe(id) {
    setOwes((prev) => prev.filter((o) => o.id !== id));
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function exportExcel() {
    const sorted = entries.slice().sort((a, b) => a.ts - b.ts);
    const entryRows = sorted.map((e) => ({
      Date: new Date(e.ts).toLocaleDateString("en-IN"),
      Time: fmtTime(e.ts),
      Category: e.category,
      Amount: e.amount,
      Note: e.note || "",
    }));
    const ws1 = XLSX.utils.json_to_sheet(entryRows);
    ws1["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 30 }];

    const oweRows = owes.slice().sort((a, b) => b.ts - a.ts).map((o) => ({
      Date: new Date(o.ts).toLocaleDateString("en-IN"),
      Person: o.person,
      Amount: o.amount,
      Note: o.note || "",
      Status: o.settled ? "Paid" : "Pending",
    }));
    const ws2 = XLSX.utils.json_to_sheet(oweRows);
    ws2["!cols"] = [{ wch: 12 }, { wch: 16 }, { wch: 10 }, { wch: 24 }, { wch: 10 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, "Expenses");
    XLSX.utils.book_append_sheet(wb, ws2, "Owe");
    XLSX.writeFile(wb, `ledger-export-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  const now = Date.now();
  const todayStart = startOfDay(now);
  const weekStart = startOfWeek(now);
  const monthStart = startOfMonth(now);

  const totals = useMemo(() => {
    let today = 0, week = 0, month = 0;
    for (const e of entries) {
      if (e.ts >= todayStart) today += e.amount;
      if (e.ts >= weekStart) week += e.amount;
      if (e.ts >= monthStart) month += e.amount;
    }
    return { today, week, month };
  }, [entries, todayStart, weekStart, monthStart]);

  const rangeStart = range === "today" ? todayStart : range === "week" ? weekStart : monthStart;
  const rangeTotal = range === "today" ? totals.today : range === "week" ? totals.week : totals.month;
  const filtered = useMemo(() => entries.filter((e) => e.ts >= rangeStart), [entries, rangeStart]);

  const byCat = useMemo(() => {
    const m = {};
    for (const c of CATS) m[c] = 0;
    for (const e of filtered) m[e.category] = (m[e.category] || 0) + e.amount;
    return m;
  }, [filtered]);
  const maxCat = Math.max(1, ...Object.values(byCat));

  const budgetPct = budget != null ? Math.min(100, Math.round((totals.today / budget) * 100)) : 0;
  const overBudget = budget != null && totals.today > budget;

  const activeOwes = useMemo(() => owes.filter((o) => !o.settled).sort((a, b) => b.ts - a.ts), [owes]);
  const settledOwes = useMemo(() => owes.filter((o) => o.settled).sort((a, b) => b.ts - a.ts), [owes]);
  const totalOwed = useMemo(() => activeOwes.reduce((s, o) => s + o.amount, 0), [activeOwes]);

  return (
    <div style={styles.page}>
      <style>{globalCSS}</style>

      <div style={styles.container}>
        <div style={styles.topBar}>
          <div style={styles.pageTitle}>Ledger</div>
          <button style={styles.iconBtn} onClick={exportExcel} title="Export Excel">⇩ Export</button>
        </div>

        <div style={styles.tabBar}>
          <button className="tap" style={{ ...styles.tabBtn, ...(tab === "ledger" ? styles.tabBtnActive : {}) }} onClick={() => setTab("ledger")}>Spending</button>
          <button className="tap" style={{ ...styles.tabBtn, ...(tab === "owe" ? styles.tabBtnActive : {}) }} onClick={() => setTab("owe")}>Owe me</button>
        </div>

        {tab === "ledger" ? (
          <>
            <div style={styles.heroCard}>
              <div style={styles.heroLabel}>Spent today</div>
              <div className={flash ? "flash-total" : ""} style={styles.heroAmount}>{money(totals.today)}</div>
              {budget != null && (
                <div style={styles.budgetRow}>
                  <div style={styles.progressTrack}>
                    <div style={{ ...styles.progressFill, width: `${budgetPct}%`, background: overBudget ? "#FF3B30" : "#34C759" }} />
                  </div>
                  <div style={{ ...styles.budgetText, color: overBudget ? "#FF3B30" : "#6E6E73" }}>
                    {overBudget ? "Over" : `${money(Math.max(0, budget - totals.today))} left`} · budget {money(budget)}/day
                  </div>
                </div>
              )}
              <button style={styles.linkBtn} onClick={() => { setBudgetInput(budget ? String(budget) : ""); setShowSettings(true); }}>
                {budget != null ? "Edit budget" : "Set a daily budget"}
              </button>
            </div>

            <div style={styles.card}>
              <div style={styles.cardLabel}>Add expense</div>
              <div style={styles.presetRow}>
                {PRESETS.map((p) => (
                  <button key={p} className="tap"
                    style={{ ...styles.pill, ...(amountInput === String(p) ? styles.pillActive : {}) }}
                    onClick={() => setAmountInput(String(p))}>
                    {money(p)}
                  </button>
                ))}
              </div>
              <input ref={amountRef} style={{ ...styles.input, marginTop: 12 }} type="number" inputMode="decimal"
                placeholder="Custom amount" value={amountInput} onChange={(e) => setAmountInput(e.target.value)} />
              <div style={styles.catGrid}>
                {CATS.map((c) => (
                  <button key={c} className="tap"
                    style={{ ...styles.catPill, ...(selectedCat === c ? styles.catPillActive : {}) }}
                    onClick={() => setSelectedCat(c)}>
                    {CAT_LABEL[c]}
                  </button>
                ))}
              </div>
              <input style={{ ...styles.input, marginTop: 10 }} type="text" placeholder="Note (optional)" value={note}
                onChange={(e) => setNote(e.target.value)} maxLength={40} />
              {error && <div style={styles.errorText}>{error}</div>}
              <button className="tap" style={{ ...styles.primaryBtn, marginTop: 14 }} onClick={addEntry}>Add expense</button>
            </div>

            <div style={styles.card}>
              <div style={styles.segment}>
                {["today", "week", "month"].map((r) => (
                  <button key={r} className="tap" style={{ ...styles.segmentItem, ...(range === r ? styles.segmentItemActive : {}) }} onClick={() => setRange(r)}>
                    {r[0].toUpperCase() + r.slice(1)}
                  </button>
                ))}
              </div>
              <div style={styles.rangeTotalRow}>
                <span style={styles.rangeTotalLabel}>Total</span>
                <span style={styles.rangeTotalAmt}>{money(rangeTotal)}</span>
              </div>
              {rangeTotal > 0 && (
                <div style={{ marginTop: 14 }}>
                  {CATS.filter((c) => byCat[c] > 0).map((c) => (
                    <div key={c} style={styles.barRow}>
                      <div style={styles.barLabel}>{CAT_LABEL[c]}</div>
                      <div style={styles.barTrack}><div style={{ ...styles.barFill, width: `${(byCat[c] / maxCat) * 100}%` }} /></div>
                      <div style={styles.barAmt}>{money(byCat[c])}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div style={styles.card}>
              {filtered.length === 0 ? (
                <div style={styles.emptyState}>Nothing logged {range === "today" ? "today" : `this ${range}`} yet.</div>
              ) : (
                filtered.slice().sort((a, b) => b.ts - a.ts).map((e, i, arr) => (
                  <div key={e.id} className="row tap" style={{ ...styles.entryRow, borderBottom: i === arr.length - 1 ? "none" : "1px solid #E5E5EA" }} onClick={() => deleteEntry(e.id)} title="Tap to delete">
                    <div>
                      <div style={styles.entryCat}>{CAT_LABEL[e.category]}</div>
                      {e.note && <div style={styles.entryNote}>{e.note}</div>}
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={styles.entryAmt}>{money(e.amount)}</div>
                      <div style={styles.entryTime}>{range === "today" ? fmtTime(e.ts) : `${fmtDateShort(e.ts)} · ${fmtTime(e.ts)}`}</div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </>
        ) : (
          <>
            <div style={styles.heroCard}>
              <div style={styles.heroLabel}>Total owed to you</div>
              <div style={styles.heroAmount}>{money(totalOwed)}</div>
              <div style={{ fontSize: 12.5, color: "#6E6E73", marginTop: 6 }}>
                {activeOwes.length} {activeOwes.length === 1 ? "person" : "people"} pending
              </div>
            </div>

            <div style={styles.card}>
              <div style={styles.cardLabel}>Add who owes you</div>
              <input ref={owePersonRef} style={styles.input} type="text" placeholder="Name" value={owePersonInput}
                onChange={(e) => setOwePersonInput(e.target.value)} />
              <input style={{ ...styles.input, marginTop: 10 }} type="number" inputMode="decimal" placeholder="Amount"
                value={oweAmountInput} onChange={(e) => setOweAmountInput(e.target.value)} />
              <input style={{ ...styles.input, marginTop: 10 }} type="text" placeholder="Note (optional)" value={oweNoteInput}
                onChange={(e) => setOweNoteInput(e.target.value)} maxLength={40} />
              {oweError && <div style={styles.errorText}>{oweError}</div>}
              <button className="tap" style={{ ...styles.primaryBtn, marginTop: 14 }} onClick={addOwe}>Add</button>
            </div>

            <div style={styles.card}>
              <div style={styles.cardLabel}>Pending</div>
              {activeOwes.length === 0 ? (
                <div style={styles.emptyState}>Nobody owes you anything right now.</div>
              ) : (
                activeOwes.map((o, i, arr) => (
                  <div key={o.id} style={{ ...styles.entryRow, borderBottom: i === arr.length - 1 ? "none" : "1px solid #E5E5EA" }}>
                    <div>
                      <div style={styles.entryCat}>{o.person}</div>
                      {o.note && <div style={styles.entryNote}>{o.note}</div>}
                      <div style={styles.entryTime}>{fmtDateShort(o.ts)}</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={styles.entryAmt}>{money(o.amount)}</div>
                      <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                        <button className="tap" style={styles.miniBtnGreen} onClick={() => toggleSettled(o.id)}>Paid</button>
                        <button className="tap" style={styles.miniBtnGray} onClick={() => deleteOwe(o.id)}>✕</button>
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>

            {settledOwes.length > 0 && (
              <div style={styles.card}>
                <div style={styles.cardLabel}>Settled</div>
                {settledOwes.map((o, i, arr) => (
                  <div key={o.id} style={{ ...styles.entryRow, borderBottom: i === arr.length - 1 ? "none" : "1px solid #E5E5EA", opacity: 0.5 }}>
                    <div>
                      <div style={styles.entryCat}>{o.person}</div>
                      {o.note && <div style={styles.entryNote}>{o.note}</div>}
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ ...styles.entryAmt, textDecoration: "line-through" }}>{money(o.amount)}</div>
                      <div style={{ display: "flex", gap: 6, marginTop: 6, justifyContent: "flex-end" }}>
                        <button className="tap" style={styles.miniBtnGray} onClick={() => toggleSettled(o.id)}>Undo</button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </div>

      {showSettings && (
        <div style={styles.modalOverlay} onClick={() => setShowSettings(false)}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <div style={styles.modalTitle}>Daily budget</div>
            <div style={styles.modalSub}>Get a progress bar against today's total. Leave blank to turn off.</div>
            <input style={styles.input} type="number" inputMode="decimal" placeholder="₹ per day"
              value={budgetInput} onChange={(e) => setBudgetInput(e.target.value)} autoFocus />
            <div style={styles.modalBtnRow}>
              <button className="tap" style={styles.ghostBtn} onClick={() => setShowSettings(false)}>Cancel</button>
              <button className="tap" style={styles.primaryBtn} onClick={saveBudget}>Save</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const FONT = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";
const INK = "#1D1D1F";
const GRAY = "#6E6E73";
const GRAY_LIGHT = "#86868B";
const BG = "#F5F5F7";
const CARD = "#FFFFFF";
const BLUE = "#0071E3";

const globalCSS = `
  * { box-sizing: border-box; -webkit-font-smoothing: antialiased; }
  html, body, #root { height: 100%; }
  body { margin: 0; overscroll-behavior-y: contain; }
  input::placeholder { color: #A1A1A6; }
  input:focus { outline: none; border-color: ${BLUE} !important; box-shadow: 0 0 0 3px rgba(0,113,227,0.12); }
  .tap:active { transform: scale(0.97); opacity: 0.9; }
  button { font-family: ${FONT}; }
  .row { transition: background 0.12s ease; }
  .flash-total { animation: pulseTotal 0.4s ease; }
  @keyframes pulseTotal { 0% { opacity: 0.4; } 100% { opacity: 1; } }
  @media (prefers-reduced-motion: reduce) { .flash-total { animation: none !important; } }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-thumb { background: #D2D2D7; border-radius: 3px; }
`;

const styles = {
  page: { minHeight: "100vh", background: BG, fontFamily: FONT, padding: "24px 16px 60px" },
  container: { width: "100%", maxWidth: 440, margin: "0 auto" },
  topBar: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 },
  pageTitle: { fontSize: 28, fontWeight: 700, color: INK, letterSpacing: -0.5 },
  iconBtn: { background: CARD, border: "none", borderRadius: 10, height: 32, padding: "0 12px", fontSize: 12.5, fontWeight: 600, color: GRAY, cursor: "pointer", boxShadow: "0 1px 2px rgba(0,0,0,0.06)" },

  tabBar: { display: "flex", background: "#E8E8ED", borderRadius: 12, padding: 4, gap: 4, marginBottom: 14 },
  tabBtn: { flex: 1, padding: "10px 0", background: "none", border: "none", borderRadius: 9, fontSize: 14, fontWeight: 600, color: GRAY, cursor: "pointer" },
  tabBtnActive: { background: CARD, color: INK, boxShadow: "0 1px 3px rgba(0,0,0,0.1)" },

  heroCard: { background: CARD, borderRadius: 20, padding: "26px 24px", marginBottom: 14, boxShadow: "0 1px 3px rgba(0,0,0,0.04)" },
  heroLabel: { fontSize: 13, color: GRAY, fontWeight: 500 },
  heroAmount: { fontSize: 44, fontWeight: 700, color: INK, letterSpacing: -1, marginTop: 4 },
  budgetRow: { marginTop: 14 },
  progressTrack: { height: 6, background: "#E8E8ED", borderRadius: 3, overflow: "hidden" },
  progressFill: { height: "100%", borderRadius: 3, transition: "width 0.3s ease" },
  budgetText: { fontSize: 12.5, marginTop: 8, fontWeight: 500 },
  linkBtn: { marginTop: 14, background: "none", border: "none", color: BLUE, fontSize: 13, fontWeight: 600, cursor: "pointer", padding: 0 },

  card: { background: CARD, borderRadius: 20, padding: 22, marginBottom: 14, boxShadow: "0 1px 3px rgba(0,0,0,0.04)" },
  cardLabel: { fontSize: 15, fontWeight: 600, color: INK, marginBottom: 14 },

  presetRow: { display: "flex", gap: 8, flexWrap: "wrap" },
  pill: { padding: "9px 16px", background: BG, border: "none", borderRadius: 100, fontSize: 14, fontWeight: 600, color: INK, cursor: "pointer" },
  pillActive: { background: INK, color: "#fff" },

  input: { width: "100%", padding: "13px 14px", fontSize: 15, color: INK, background: BG, border: "1.5px solid transparent", borderRadius: 12, outline: "none", fontFamily: FONT },

  catGrid: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginTop: 12 },
  catPill: { padding: "10px 4px", background: BG, border: "none", borderRadius: 12, fontSize: 13, fontWeight: 600, color: GRAY, cursor: "pointer" },
  catPillActive: { background: BLUE, color: "#fff" },

  errorText: { fontSize: 12.5, color: "#FF3B30", marginTop: 8, textAlign: "center" },
  primaryBtn: { width: "100%", padding: "14px", background: BLUE, color: "#fff", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 600, cursor: "pointer" },
  ghostBtn: { flex: 1, padding: "14px", background: BG, color: INK, border: "none", borderRadius: 12, fontSize: 15, fontWeight: 600, cursor: "pointer" },

  segment: { display: "flex", background: BG, borderRadius: 10, padding: 3, gap: 2 },
  segmentItem: { flex: 1, padding: "8px 0", background: "none", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 600, color: GRAY, cursor: "pointer" },
  segmentItemActive: { background: CARD, color: INK, boxShadow: "0 1px 3px rgba(0,0,0,0.1)" },

  rangeTotalRow: { display: "flex", justifyContent: "space-between", alignItems: "baseline", marginTop: 16 },
  rangeTotalLabel: { fontSize: 13, color: GRAY },
  rangeTotalAmt: { fontSize: 20, fontWeight: 700, color: INK },

  barRow: { display: "flex", alignItems: "center", gap: 10, marginBottom: 10 },
  barLabel: { width: 50, fontSize: 12, color: GRAY, fontWeight: 500 },
  barTrack: { flex: 1, height: 6, background: "#E8E8ED", borderRadius: 3, overflow: "hidden" },
  barFill: { height: "100%", background: INK, borderRadius: 3 },
  barAmt: { width: 56, textAlign: "right", fontSize: 12.5, color: INK, fontWeight: 600 },

  emptyState: { textAlign: "center", fontSize: 13.5, color: GRAY_LIGHT, padding: "16px 10px" },
  entryRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "13px 2px", cursor: "pointer" },
  entryCat: { fontSize: 14.5, fontWeight: 600, color: INK },
  entryNote: { fontSize: 12.5, color: GRAY, marginTop: 2 },
  entryAmt: { fontSize: 15, fontWeight: 600, color: INK },
  entryTime: { fontSize: 11.5, color: GRAY_LIGHT, marginTop: 2 },

  miniBtnGreen: { background: "#34C759", color: "#fff", border: "none", borderRadius: 8, padding: "5px 10px", fontSize: 11.5, fontWeight: 600, cursor: "pointer" },
  miniBtnGray: { background: "#E8E8ED", color: GRAY, border: "none", borderRadius: 8, padding: "5px 10px", fontSize: 11.5, fontWeight: 600, cursor: "pointer" },

  modalOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, backdropFilter: "blur(4px)" },
  modal: { background: CARD, borderRadius: 20, padding: 26, width: "100%", maxWidth: 340, fontFamily: FONT },
  modalTitle: { fontSize: 18, fontWeight: 700, color: INK, marginBottom: 4 },
  modalSub: { fontSize: 13, color: GRAY, marginBottom: 16, lineHeight: 1.5 },
  modalBtnRow: { display: "flex", gap: 10, marginTop: 16 },
};
