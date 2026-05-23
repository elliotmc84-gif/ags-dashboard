import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer
} from "recharts";

const C = {
  bg: "#0D1117", surface: "#161B22", border: "#21262D",
  accent: "#2563EB", accentLt: "#3B82F6",
  text: "#E6EDF3", textSub: "#8B949E",
  green: "#3FB950", red: "#F85149",
  actual: "#3B82F6", budget: "#8B5CF6", py: "#F59E0B",
};

const fmtPct = v => v == null ? "—" : (v * 100).toFixed(1) + "%";
const fmtK = v => {
  if (v == null) return "—";
  const abs = Math.abs(v);
  const sign = v < 0 ? "-" : "";
  if (abs >= 1000000) return sign + "€" + (abs / 1000000).toFixed(2) + "M";
  if (abs >= 1000)    return sign + "€" + (abs / 1000).toFixed(1) + "k";
  return sign + "€" + Math.round(abs);
};

function serialToYearMonth(serial) {
  if (!serial || isNaN(serial)) return null;
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function parseMgmtAccount(wb) {
  const ws = wb.Sheets["SUMMARY"];
  if (!ws) throw new Error("No SUMMARY sheet found");
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Branch name
  let branchRaw = "";
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    for (const cell of rows[i]) {
      if (typeof cell === "string" && cell.includes("AGS") && cell.includes("|")) {
        branchRaw = cell.trim(); break;
      }
    }
    if (branchRaw) break;
  }
  let branchName = branchRaw;
  const cityMatch = branchRaw.match(/AGS\s+([A-Z\s]+)\s*\(/i);
  if (cityMatch) {
    branchName = "AGS " + cityMatch[1].trim().split(" ")
      .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(" ");
  }

  // Period
  let period = null;
  for (let i = 0; i < Math.min(12, rows.length); i++) {
    for (const cell of rows[i]) {
      if (typeof cell === "number" && cell > 40000 && cell < 55000) {
        period = serialToYearMonth(cell); break;
      }
    }
    if (period) break;
  }

  // Find row by label and extract Actual/Budget/PY YTD
  const extractRow = (label) => {
    for (const row of rows) {
      for (const cell of row) {
        if (typeof cell === "string" && cell.trim().toUpperCase().includes(label.toUpperCase())) {
          const nums = row.filter(v => typeof v === "number" && Math.abs(v) > 1);
          return { actual: nums[0] ?? null, budget: nums[2] ?? null, py: nums[4] ?? null };
        }
      }
    }
    return { actual: null, budget: null, py: null };
  };

  const revenue  = extractRow("TURNOVER");
  const gm       = extractRow("GROSS MARGIN");
  const opResult = extractRow("OPERATIONAL RESULT");

  const cost = {
    actual: revenue.actual != null && gm.actual != null ? revenue.actual - gm.actual : null,
    budget: revenue.budget != null && gm.budget != null ? revenue.budget - gm.budget : null,
    py:     revenue.py     != null && gm.py     != null ? revenue.py     - gm.py     : null,
  };

  return { branch: branchName, period, revenue, gm, cost, opResult, importedAt: new Date().toISOString() };
}

// Persistent storage
const STORE_KEY = "ags-records-v1";
function loadLocal() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY) || "{}"); } catch { return {}; }
}
function saveLocal(data) {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(data)); } catch {}
}

function KpiCard({ label, value, sub, delta }) {
  const up = delta >= 0;
  return (
    <div style={{
      background: C.surface, border: `1px solid ${C.border}`,
      borderRadius: 10, padding: "16px 18px", flex: 1, minWidth: 130
    }}>
      <div style={{ color: C.textSub, fontSize: 10, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6 }}>
        {label}
      </div>
      <div style={{ color: C.text, fontSize: 20, fontWeight: 700, fontFamily: "monospace" }}>
        {value}
      </div>
      {sub && <div style={{ color: C.textSub, fontSize: 11, marginTop: 3 }}>{sub}</div>}
      {delta != null && (
        <div style={{ marginTop: 5, display: "flex", alignItems: "center", gap: 4 }}>
          <span style={{
            fontSize: 11, fontWeight: 600,
            color: up ? C.green : C.red,
            background: up ? "#1a3323" : "#3d1a1a",
            padding: "2px 6px", borderRadius: 4
          }}>
            {up ? "▲" : "▼"} {fmtK(Math.abs(delta))}
          </span>
          <span style={{ color: C.textSub, fontSize: 10 }}>vs budget</span>
        </div>
      )}
    </div>
  );
}

function ChartTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: "#1C2128", border: `1px solid ${C.border}`,
      borderRadius: 8, padding: "10px 14px", fontSize: 12
    }}>
      <div style={{ color: C.textSub, marginBottom: 6, fontWeight: 600 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, display: "flex", gap: 8, marginBottom: 3 }}>
          <span>{p.name}:</span>
          <span style={{ fontFamily: "monospace" }}>{fmtK(p.value)}</span>
        </div>
      ))}
    </div>
  );
}

export default function App() {
  const [records, setRecords]     = useState(() => loadLocal());
  const [query, setQuery]         = useState("");
  const [branch, setBranch]       = useState(null);
  const [suggestions, setSugg]    = useState([]);
  const [metric, setMetric]       = useState("gm");
  const [dragging, setDragging]   = useState(false);
  const [toast, setToast]         = useState(null);
  const [view, setView]           = useState("dashboard"); // dashboard | overview

  const showToast = (msg, type = "ok") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

  const ingestFiles = useCallback(async (files) => {
    let current = loadLocal();
    let ok = 0, errs = [];
    for (const file of files) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const parsed = parseMgmtAccount(wb);
        if (!parsed.branch || !parsed.period) throw new Error("Could not read branch or period");
        if (!current[parsed.branch]) current[parsed.branch] = {};
        current[parsed.branch][parsed.period] = parsed;
        ok++;
      } catch (e) { errs.push(`${file.name}: ${e.message}`); }
    }
    saveLocal(current);
    setRecords({ ...current });
    if (ok && !errs.length) showToast(`${ok} file${ok > 1 ? "s" : ""} imported`);
    else if (ok) showToast(`${ok} imported, ${errs.length} failed`, "warn");
    else showToast(errs[0] || "Import failed", "err");
  }, []);

  const onDrop = useCallback(e => {
    e.preventDefault(); setDragging(false);
    const files = [...e.dataTransfer.files].filter(f => /\.xlsm?$/i.test(f.name));
    if (files.length) ingestFiles(files);
  }, [ingestFiles]);

  const allBranches = useMemo(() => Object.keys(records).sort(), [records]);

  const onQueryChange = e => {
    const v = e.target.value;
    setQuery(v);
    if (!v.trim()) { setSugg([]); setBranch(null); return; }
    const matches = allBranches.filter(b => b.toLowerCase().includes(v.toLowerCase()));
    setSugg(matches);
    const exact = allBranches.find(b => b.toLowerCase() === v.toLowerCase());
    if (exact) { setBranch(exact); setSugg([]); } else setBranch(null);
  };

  const selectBranch = b => { setQuery(b); setBranch(b); setSugg([]); setView("dashboard"); };

  const dashboard = useMemo(() => {
    if (!branch || !records[branch]) return null;
    const branchData = records[branch];
    const periods = Object.keys(branchData).sort();
    const latest = periods[periods.length - 1];
    const rec = branchData[latest];

    const ytd = {
      revenue: rec.revenue?.actual,
      gm: rec.gm?.actual,
      cost: rec.cost?.actual,
      opResult: rec.opResult?.actual,
    };
    ytd.gmPct = ytd.revenue ? ytd.gm / ytd.revenue : 0;

    const bud = { revenue: rec.revenue?.budget, gm: rec.gm?.budget };
    const py  = { revenue: rec.revenue?.py,     gm: rec.gm?.py };

    const chartData = periods.map(p => {
      const r = branchData[p];
      return {
        month: p.slice(5),
        actual: r.gm?.actual, budget: r.gm?.budget, py: r.gm?.py,
        actualRev: r.revenue?.actual, budgetRev: r.revenue?.budget, pyRev: r.revenue?.py,
        actualCost: r.cost?.actual,
      };
    });

    return { ytd, bud, py, chartData, latest, periods };
  }, [branch, records]);

  const overviewRows = useMemo(() => allBranches.map(b => {
    const periods = Object.keys(records[b]).sort();
    const latest = periods[periods.length - 1];
    const rec = records[b][latest];
    const rev = rec.revenue?.actual;
    const gm  = rec.gm?.actual;
    const bud = rec.revenue?.budget;
    return { branch: b, latest, rev, gm, gmPct: rev ? gm / rev : 0, vsBud: bud != null ? rev - bud : null };
  }), [allBranches, records]);

  const metricKeys = metric === "gm" ? ["actual","budget","py"]
    : metric === "revenue" ? ["actualRev","budgetRev","pyRev"]
    : ["actualCost"];
  const metricLabel = metric === "gm" ? "Gross Margin" : metric === "revenue" ? "Revenue" : "Cost";

  const hasData = allBranches.length > 0;

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Segoe UI',sans-serif", paddingBottom: 80 }}>
      <style>{`
        * { box-sizing: border-box; }
        input { outline: none; }
        .sugg:hover { background: #21262D !important; cursor: pointer; }
        .chip:hover { background: ${C.accent} !important; color: #fff !important; cursor: pointer; }
        .row-hover:hover { background: #1c2128 !important; cursor: pointer; }
        .mbtn:hover { opacity: 0.8; }
      `}</style>

      {/* Toast */}
      {toast && (
        <div style={{
          position: "fixed", top: 16, right: 16, zIndex: 999,
          background: toast.type === "err" ? "#3d1a1a" : toast.type === "warn" ? "#2d2010" : "#1a3323",
          border: `1px solid ${toast.type === "err" ? C.red : toast.type === "warn" ? C.amber : C.green}`,
          color: C.text, borderRadius: 8, padding: "10px 16px", fontSize: 13, fontWeight: 500,
          boxShadow: "0 4px 20px rgba(0,0,0,0.4)"
        }}>
          {toast.msg}
        </div>
      )}

      {/* Header */}
      <div style={{
        background: C.surface, borderBottom: `1px solid ${C.border}`,
        padding: "14px 20px", display: "flex", alignItems: "center", gap: 14,
        position: "sticky", top: 0, zIndex: 100
      }}>
        <div style={{ width: 6, height: 28, background: C.accent, borderRadius: 3, flexShrink: 0 }} />
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 15, fontWeight: 700 }}>AGS Network</div>
          <div style={{ fontSize: 10, color: C.textSub, letterSpacing: "0.06em", textTransform: "uppercase" }}>Branch Dashboard</div>
        </div>
        {hasData && (
          <div style={{ display: "flex", gap: 6 }}>
            {["dashboard","overview"].map(v => (
              <button key={v} className="mbtn" onClick={() => setView(v)} style={{
                padding: "6px 12px", borderRadius: 6, fontSize: 11, fontWeight: 600,
                border: "none", cursor: "pointer",
                background: view === v ? C.accent : C.border,
                color: view === v ? "#fff" : C.textSub,
              }}>{v.charAt(0).toUpperCase() + v.slice(1)}</button>
            ))}
          </div>
        )}
        <label style={{
          background: C.accent, color: "#fff", padding: "7px 14px",
          borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: "pointer", flexShrink: 0
        }}>
          + Upload
          <input type="file" accept=".xlsx,.xlsm" multiple style={{ display: "none" }}
            onChange={e => { ingestFiles([...e.target.files]); e.target.value = ""; }} />
        </label>
      </div>

      <div style={{ maxWidth: 900, margin: "0 auto", padding: "24px 16px 0" }}>

        {/* Drop zone (only when no data) */}
        {!hasData && (
          <div
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            style={{
              border: `2px dashed ${dragging ? C.accentLt : C.border}`,
              borderRadius: 14, padding: "60px 24px", textAlign: "center",
              background: dragging ? "#0d1f3c" : C.surface,
            }}
          >
            <div style={{ fontSize: 36, marginBottom: 12 }}>📂</div>
            <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Drop management account files here</div>
            <div style={{ color: C.textSub, fontSize: 13 }}>Accepts .xlsm and .xlsx — upload one or multiple at once</div>
          </div>
        )}

        {/* Overview table */}
        {hasData && view === "overview" && (
          <div>
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 16 }}>
              Network overview — {allBranches.length} branch{allBranches.length !== 1 ? "es" : ""} loaded
            </div>
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                    {["Branch","Period","Revenue YTD","GM YTD","GM %","vs Budget"].map(h => (
                      <th key={h} style={{ padding: "12px 16px", color: C.textSub, fontWeight: 600, textAlign: h === "Branch" ? "left" : "right", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {overviewRows.map((r, i) => (
                    <tr key={r.branch} className="row-hover"
                      onClick={() => selectBranch(r.branch)}
                      style={{ borderBottom: i < overviewRows.length - 1 ? `1px solid ${C.border}` : "none", background: "transparent", transition: "background 0.15s" }}>
                      <td style={{ padding: "12px 16px", fontWeight: 600 }}>{r.branch}</td>
                      <td style={{ padding: "12px 16px", color: C.textSub, textAlign: "right" }}>{r.latest}</td>
                      <td style={{ padding: "12px 16px", textAlign: "right", fontFamily: "monospace" }}>{fmtK(r.rev)}</td>
                      <td style={{ padding: "12px 16px", textAlign: "right", fontFamily: "monospace" }}>{fmtK(r.gm)}</td>
                      <td style={{ padding: "12px 16px", textAlign: "right", fontFamily: "monospace" }}>{fmtPct(r.gmPct)}</td>
                      <td style={{ padding: "12px 16px", textAlign: "right" }}>
                        {r.vsBud != null && (
                          <span style={{ color: r.vsBud >= 0 ? C.green : C.red, fontFamily: "monospace", fontWeight: 600 }}>
                            {r.vsBud >= 0 ? "▲" : "▼"} {fmtK(Math.abs(r.vsBud))}
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Dashboard view */}
        {hasData && view === "dashboard" && (
          <>
            {/* Search */}
            <div style={{ position: "relative", maxWidth: 440, marginBottom: 28 }}>
              <div style={{ position: "absolute", left: 13, top: "50%", transform: "translateY(-50%)", color: C.textSub, fontSize: 16, pointerEvents: "none" }}>⌕</div>
              <input value={query} onChange={onQueryChange} placeholder="Type a branch name..."
                style={{
                  width: "100%", padding: "12px 14px 12px 36px",
                  background: C.surface, border: `1px solid ${branch ? C.accent : C.border}`,
                  borderRadius: 10, color: C.text, fontSize: 15,
                }} />
              {suggestions.length > 0 && (
                <div style={{
                  position: "absolute", top: "100%", left: 0, right: 0, zIndex: 10,
                  background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, marginTop: 4, overflow: "hidden"
                }}>
                  {suggestions.map(s => (
                    <div key={s} className="sugg" onClick={() => selectBranch(s)}
                      style={{ padding: "11px 16px", fontSize: 14, borderBottom: `1px solid ${C.border}` }}>
                      {s}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Branch chips */}
            {!branch && (
              <div style={{ textAlign: "center", padding: "40px 0", color: C.textSub }}>
                <div style={{ fontSize: 28, marginBottom: 10 }}>🔍</div>
                <div style={{ fontSize: 13, marginBottom: 16 }}>
                  {allBranches.length} branch{allBranches.length !== 1 ? "es" : ""} loaded — select one to view
                </div>
                <div style={{ display: "flex", gap: 8, justifyContent: "center", flexWrap: "wrap" }}>
                  {allBranches.map(b => (
                    <span key={b} className="chip" onClick={() => selectBranch(b)}
                      style={{ padding: "6px 14px", borderRadius: 20, background: C.border, fontSize: 12, color: C.textSub, cursor: "pointer", transition: "all 0.15s" }}>
                      {b}
                    </span>
                  ))}
                </div>
              </div>
            )}

            {/* Dashboard */}
            {dashboard && branch && (
              <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>

                <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
                  <h2 style={{ margin: 0, fontSize: 24, fontWeight: 700 }}>{branch}</h2>
                  <span style={{ color: C.textSub, fontSize: 13 }}>YTD to {dashboard.latest}</span>
                </div>

                {/* KPIs */}
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <KpiCard label="Revenue YTD" value={fmtK(dashboard.ytd.revenue)}
                    sub={dashboard.bud.revenue ? `Budget: ${fmtK(dashboard.bud.revenue)}` : null}
                    delta={dashboard.bud.revenue != null ? dashboard.ytd.revenue - dashboard.bud.revenue : null} />
                  <KpiCard label="Gross Margin" value={fmtK(dashboard.ytd.gm)}
                    sub={dashboard.bud.gm ? `Budget: ${fmtK(dashboard.bud.gm)}` : null}
                    delta={dashboard.bud.gm != null ? dashboard.ytd.gm - dashboard.bud.gm : null} />
                  <KpiCard label="GM %" value={fmtPct(dashboard.ytd.gmPct)} />
                  {dashboard.py.revenue != null && (
                    <KpiCard label="Revenue vs PY"
                      value={fmtK(dashboard.ytd.revenue - dashboard.py.revenue)}
                      sub={`PY: ${fmtK(dashboard.py.revenue)}`}
                      delta={dashboard.ytd.revenue - dashboard.py.revenue} />
                  )}
                </div>

                {/* Chart */}
                {dashboard.chartData.length > 1 && (
                  <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "18px 16px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
                      <div style={{ fontWeight: 600, fontSize: 13 }}>Monthly trend (YTD)</div>
                      <div style={{ display: "flex", gap: 6 }}>
                        {[["gm","GM"],["revenue","Revenue"],["cost","Cost"]].map(([m, lbl]) => (
                          <button key={m} className="mbtn" onClick={() => setMetric(m)} style={{
                            padding: "4px 10px", borderRadius: 6, fontSize: 11, fontWeight: 600,
                            border: "none", cursor: "pointer",
                            background: metric === m ? C.accent : C.border,
                            color: metric === m ? "#fff" : C.textSub,
                          }}>{lbl}</button>
                        ))}
                      </div>
                    </div>
                    <ResponsiveContainer width="100%" height={240}>
                      <LineChart data={dashboard.chartData} margin={{ top: 4, right: 8, left: 0, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="month" stroke={C.textSub} tick={{ fontSize: 10 }} />
                        <YAxis stroke={C.textSub} tick={{ fontSize: 10 }} tickFormatter={fmtK} width={62} />
                        <Tooltip content={<ChartTooltip />} />
                        <Legend wrapperStyle={{ fontSize: 11, paddingTop: 10 }}
                          formatter={v => <span style={{ color: C.textSub }}>{v}</span>} />
                        <Line type="monotone" dataKey={metricKeys[0]} name={`Actual ${metricLabel}`}
                          stroke={C.actual} strokeWidth={2.5} dot={{ fill: C.actual, r: 3 }} connectNulls />
                        {metricKeys[1] && (
                          <Line type="monotone" dataKey={metricKeys[1]} name={`Budget ${metricLabel}`}
                            stroke={C.budget} strokeWidth={2} strokeDasharray="6 3" dot={false} connectNulls />
                        )}
                        {metricKeys[2] && (
                          <Line type="monotone" dataKey={metricKeys[2]} name={`PY ${metricLabel}`}
                            stroke={C.py} strokeWidth={2} strokeDasharray="3 3" dot={false} connectNulls />
                        )}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                )}

                {/* Single period note */}
                {dashboard.chartData.length === 1 && (
                  <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "18px 16px", color: C.textSub, fontSize: 13 }}>
                    Chart will appear once you upload more than one period for this branch.
                  </div>
                )}

              </div>
            )}
          </>
        )}

      </div>
    </div>
  );
}
