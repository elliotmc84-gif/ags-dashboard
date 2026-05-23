import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer
} from "recharts";

const C = {
  bg: "#F5F5F5", surface: "#FFFFFF", border: "#E0E0E0",
  accent: "#A51C30", accentLt: "#C0392B",
  text: "#1A1A1A", textSub: "#666666",
  green: "#1A7A3A", red: "#A51C30", amber: "#B45309",
  greenBg: "#E8F5EC", redBg: "#FDECEA",
  actual: "#A51C30", budget: "#4A4A8A", py: "#B45309",
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

// Parse the dashboard Excel template
// Actuals: Branch | Month | Revenue | Cost | Gross Margin | GM% | Operational Result | Current Topics
// Budget:  Branch | Month | Revenue_Budget | Cost_Budget | Gross_Margin_Budget | Operational Result Budget
// Prior Year: Branch | Month | Revenue_Prior Year | Fixed Costs_Prior Year | Gross_Margin_Prior Year | Operational Result Prior Year
function parseDashboardExcel(wb) {
  const toRows = name => {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { defval: null });
  };

  const actRows = toRows("Actuals");
  const budRows = toRows("Budget");
  const pyRows  = toRows("Prior Year");

  if (!actRows.length) throw new Error("No data found in the Actuals sheet.");

  const clean = v => (typeof v === "string" ? v.trim() : v);

  const makeGetter = (row) => {
    const keys = Object.keys(row);
    return (...hints) => {
      const k = keys.find(k => hints.some(h =>
        k.toLowerCase().replace(/[^a-z]/g,"").includes(h.toLowerCase().replace(/[^a-z]/g,""))
      ));
      return k ? clean(row[k]) : null;
    };
  };

  const actuals = actRows.map(row => {
    const get = makeGetter(row);
    const branch   = get("Branch");
    const month    = get("Month");
    const revenue  = Number(get("Revenue") || 0);
    const cost     = Number(get("Cost") || 0);
    const gm       = Number(get("GrossMargin","Gross") || revenue - cost);
    const opResult = Number(get("OperationalResult","OpResult","Operational") || 0);
    const topics   = get("CurrentTopics","Topics","Current") || "";
    return { branch: String(branch||"").trim(), month: String(month||"").trim(), revenue, cost, gm, opResult, topics };
  }).filter(r => r.branch && r.month);

  const budget = budRows.map(row => {
    const get = makeGetter(row);
    const branch   = get("Branch");
    const month    = get("Month");
    const revenue  = Number(get("RevenueBudget","Revenue") || 0);
    const cost     = Number(get("CostBudget","Cost") || 0);
    const gm       = Number(get("GrossMarginBudget","GrossMargin","Gross") || revenue - cost);
    const opResult = Number(get("OperationalResultBudget","OpResultBudget","OperationalResult","Operational") || 0);
    return { branch: String(branch||"").trim(), month: String(month||"").trim(), revenue, cost, gm, opResult };
  }).filter(r => r.branch && r.month);

  const priorYear = pyRows.map(row => {
    const get = makeGetter(row);
    const branch   = get("Branch");
    const month    = get("Month");
    const revenue  = Number(get("RevenuePrior","Revenue") || 0);
    const cost     = Number(get("FixedCostsPrior","CostPrior","Cost") || 0);
    const gm       = Number(get("GrossMarginPrior","GrossPrior","Gross") || revenue - cost);
    const opResult = Number(get("OperationalResultPrior","OpResultPrior","OperationalResult","Operational") || 0);
    return { branch: String(branch||"").trim(), month: String(month||"").trim(), revenue, cost, gm, opResult };
  }).filter(r => r.branch && r.month);

  return { actuals, budget, priorYear };
}

function KpiCard({ label, value, sub, sub2, delta, deltaPy, pct }) {
  const up   = delta >= 0;
  const upPy = deltaPy >= 0;
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
      {sub2 && <div style={{ color: C.textSub, fontSize: 11, marginTop: 1 }}>{sub2}</div>}
      {delta != null && (
        <div style={{ marginTop: 6 }}>
          <span style={{
            fontSize: 11, fontWeight: 600,
            color: up ? C.green : C.red,
            background: up ? "#E8F5EC" : "#FDECEA",
            padding: "2px 6px", borderRadius: 4
          }}>
            {up ? "▲" : "▼"} {pct ? fmtPct(Math.abs(delta)) : fmtK(Math.abs(delta))} vs bud
          </span>
        </div>
      )}
      {deltaPy != null && (
        <div style={{ marginTop: 4 }}>
          <span style={{
            fontSize: 11, fontWeight: 600,
            color: upPy ? C.green : C.red,
            background: upPy ? "#E8F5EC" : "#FDECEA",
            padding: "2px 6px", borderRadius: 4
          }}>
            {upPy ? "▲" : "▼"} {pct ? fmtPct(Math.abs(deltaPy)) : fmtK(Math.abs(deltaPy))} vs PY
          </span>
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

const STORAGE_KEY = "ags-dashboard-data-v1";

function saveToStorage(data, fileName) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ data, fileName }));
  } catch (e) {
    console.warn("Could not save to localStorage", e);
  }
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch { return null; }
}

export default function App() {
  const stored = loadFromStorage();
  const [data, setData]         = useState(stored?.data || null);
  const [fileName, setFileName] = useState(stored?.fileName || null);
  const [query, setQuery]       = useState("");
  const [branch, setBranch]     = useState(null);
  const [suggestions, setSugg]  = useState([]);
  const [metric, setMetric]     = useState("gm");
  const [dragging, setDragging] = useState(false);
  const [error, setError]       = useState(null);
  const [view, setView]         = useState("dashboard");

  const ingestFile = useCallback(async (file) => {
    setError(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const parsed = parseDashboardExcel(wb);
      setData(parsed);
      setFileName(file.name);
      saveToStorage(parsed, file.name);
      setBranch(null);
      setQuery("");
    } catch (e) {
      setError(e.message || "Could not read file. Make sure it's the AGS Dashboard Excel.");
    }
  }, []);

  const onDrop = useCallback(e => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) ingestFile(file);
  }, [ingestFile]);

  const allBranches = useMemo(() => {
    if (!data) return [];
    return [...new Set(data.actuals.map(r => r.branch))].sort();
  }, [data]);

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
    if (!data || !branch) return null;
    const actRows = data.actuals.filter(r => r.branch === branch);
    const budRows = data.budget.filter(r => r.branch === branch);
    const pyRows  = (data.priorYear || []).filter(r => r.branch === branch);

    const months = actRows.map(r => r.month).sort();
    const latestMonth = months[months.length - 1] || "";
    const currentYear = latestMonth.slice(0, 4) || String(new Date().getFullYear());

    const cyRows = actRows.filter(r => r.month.startsWith(currentYear));
    const cyBud  = budRows.filter(r => r.month.startsWith(currentYear));

    const sum = (rows, key) => rows.reduce((acc, r) => acc + (Number(r[key]) || 0), 0);

    const ytd = {
      revenue: sum(cyRows, "revenue"),
      cost:    sum(cyRows, "cost"),
      gm:      sum(cyRows, "gm"),
      opResult: sum(cyRows, "opResult"),
    };
    ytd.gmPct = ytd.revenue ? ytd.gm / ytd.revenue : 0;

    const bud = { revenue: sum(cyBud, "revenue"), gm: sum(cyBud, "gm"), opResult: sum(cyBud, "opResult") };
    bud.gmPct = bud.revenue ? bud.gm / bud.revenue : 0;
    const py  = { revenue: sum(pyRows, "revenue"), cost: sum(pyRows, "cost"), gm: sum(pyRows, "gm"), opResult: sum(pyRows, "opResult") };
    py.gmPct  = py.revenue ? py.gm / py.revenue : 0;

    const monthMap = {};
    cyRows.forEach(r => {
      monthMap[r.month] = { ...monthMap[r.month], actual: r.gm, actualRev: r.revenue, actualCost: r.cost, topics: r.topics };
    });
    cyBud.forEach(r => {
      monthMap[r.month] = { ...monthMap[r.month], budget: r.gm, budgetRev: r.revenue };
    });
    pyRows.forEach(r => {
      // PY rows use same month format; map them to current year for charting
      const key = currentYear + "-" + r.month.slice(5);
      monthMap[key] = { ...monthMap[key], py: r.gm, pyRev: r.revenue };
    });

    const chartData = Object.entries(monthMap)
      .sort(([a],[b]) => a.localeCompare(b))
      .map(([month, vals]) => ({ month: month.slice(5), ...vals }));

    const latestRow = [...cyRows].sort((a,b) => b.month.localeCompare(a.month))[0];
    const topics = latestRow?.topics
      ? latestRow.topics.split(";").map(t => t.trim()).filter(Boolean)
      : [];

    return { ytd, bud, py, chartData, topics, currentYear, latestMonth };
  }, [data, branch]);

  const overviewRows = useMemo(() => {
    if (!data) return [];
    return allBranches.map(b => {
      const actRows = data.actuals.filter(r => r.branch === b);
      const budRows = data.budget.filter(r => r.branch === b);
      const pyRows  = (data.priorYear || []).filter(r => r.branch === b);
      const months = actRows.map(r => r.month).sort();
      const latest = months[months.length - 1] || "";
      const currentYear = latest.slice(0, 4);
      const cyRows  = actRows.filter(r => r.month.startsWith(currentYear));
      const cyBud   = budRows.filter(r => r.month.startsWith(currentYear));
      const sum = (rows, key) => rows.reduce((acc, r) => acc + (Number(r[key]) || 0), 0);
      const rev     = sum(cyRows, "revenue");
      const gm      = sum(cyRows, "gm");
      const cost    = sum(cyRows, "cost");
      const budRev  = sum(cyBud, "revenue");
      const budGm   = sum(cyBud, "gm");
      const budCost = sum(cyBud, "cost");
      const budGmPct = budRev ? budGm / budRev : 0;
      const pyRev   = sum(pyRows, "revenue");
      const pyGm    = sum(pyRows, "gm");
      const pyCost  = sum(pyRows, "cost");
      const opResult    = sum(cyRows, "opResult");
      const opResultBud = sum(cyBud,  "opResult");
      const opResultPy  = sum(pyRows, "opResult");
      return {
        branch: b, latest, rev, gm, gmPct: rev ? gm / rev : 0, budGmPct, cost, budCost, pyCost,
        vsBudRev: budRev ? rev - budRev : null,
        vsBudGm:  budGm  ? gm  - budGm  : null,
        vsBudCost: budCost ? cost - budCost : null,
        pyRev, pyGm, vsPyRev: pyRev ? rev - pyRev : null, vsPyCost: pyCost ? cost - pyCost : null,
        opResult, opResultBud, opResultPy,
      };
    });
  }, [data, allBranches]);

  const regionalTotals = useMemo(() => {
    if (!overviewRows.length) return null;
    const sum = key => overviewRows.reduce((acc, r) => acc + (r[key] || 0), 0);
    const rev  = sum("rev");
    const gm   = sum("gm");
    const cost = sum("cost");
    const budRev  = overviewRows.reduce((acc, r) => acc + (r.vsBudRev != null ? r.rev - r.vsBudRev : 0), 0);
    const budGm   = overviewRows.reduce((acc, r) => acc + (r.vsBudGm  != null ? r.gm  - r.vsBudGm  : 0), 0);
    const budCost = overviewRows.reduce((acc, r) => acc + (r.vsBudCost != null ? r.cost - r.vsBudCost : 0), 0);
    const budGmPct = budRev ? budGm / budRev : 0;
    const pyRev = sum("pyRev"), pyGm = sum("pyGm"), pyCost = sum("pyCost");
    const pyGmPct  = pyRev ? pyGm / pyRev : 0;
    const periods = [...new Set(overviewRows.map(r => r.latest))].sort();
    return {
      rev, gm, gmPct: rev ? gm / rev : 0, budGmPct, pyGmPct, cost, budCost,
      vsBudRev:  overviewRows.some(r => r.vsBudRev  != null) ? sum("vsBudRev")  : null,
      vsBudCost: overviewRows.some(r => r.vsBudCost != null) ? sum("vsBudCost") : null,
      pyRev, pyGm, pyCost,
      opResult: sum("opResult"), opResultBud: sum("opResultBud"), opResultPy: sum("opResultPy"),
      periodRange: periods.length > 1 ? `${periods[0]} – ${periods[periods.length-1]}` : periods[0] || "",
      branchCount: overviewRows.length,
    };
  }, [overviewRows]);

  const metricKeys = metric === "gm" ? ["actual","budget","py"]
    : metric === "revenue" ? ["actualRev","budgetRev","pyRev"]
    : ["actualCost"];
  const metricLabel = metric === "gm" ? "Gross Margin" : metric === "revenue" ? "Revenue" : "Cost";

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'Segoe UI',sans-serif", paddingBottom: 80 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        body { background: #F5F5F5; }
        input { outline: none; }
        .sugg:hover { background: #F0F0F0 !important; cursor: pointer; }
        .chip:hover { background: ${C.accent} !important; color: #fff !important; cursor: pointer; }
        .row-hover:hover { background: #F9F9F9 !important; cursor: pointer; }
        .mbtn:hover { opacity: 0.8; }
      `}</style>

      {/* Header */}
      <div style={{
        background: C.surface, borderBottom: `3px solid ${C.accent}`,
        padding: "14px 20px", display: "flex", alignItems: "center", gap: 14,
        position: "sticky", top: 0, zIndex: 100, boxShadow: "0 1px 4px rgba(0,0,0,0.08)"
      }}>
        <div style={{ width: 4, height: 28, background: C.accent, borderRadius: 2, flexShrink: 0 }} />
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 15, fontWeight: 700 }}>AGS Network</div>
          <div style={{ fontSize: 10, color: C.textSub, letterSpacing: "0.06em", textTransform: "uppercase" }}>Branch Dashboard</div>
        </div>
        {data && (
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
          background: data ? C.border : C.accent,
          color: data ? C.textSub : "#fff",
          padding: "7px 14px", borderRadius: 7, fontSize: 12, fontWeight: 600, cursor: "pointer", flexShrink: 0
        }}>
          {data ? "↺ Replace" : "+ Load Excel"}
          <input type="file" accept=".xlsx,.xlsm" style={{ display: "none" }}
            onChange={e => { if (e.target.files[0]) ingestFile(e.target.files[0]); e.target.value = ""; }} />
        </label>
      </div>

      <div style={{ maxWidth: 900, margin: "0 auto", padding: "24px 16px 0" }}>

        {/* Drop zone */}
        {!data && (
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
            <div style={{ fontSize: 36, marginBottom: 12 }}>📊</div>
            <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Drop your AGS Dashboard Excel here</div>
            <div style={{ color: C.textSub, fontSize: 13, marginBottom: 20 }}>
              The dashboard template with Actuals and Budget sheets
            </div>
            <label style={{
              display: "inline-block", background: C.accent, color: "#fff",
              padding: "10px 24px", borderRadius: 8, fontWeight: 600, fontSize: 13, cursor: "pointer"
            }}>
              Browse file
              <input type="file" accept=".xlsx,.xlsm" style={{ display: "none" }}
                onChange={e => { if (e.target.files[0]) ingestFile(e.target.files[0]); e.target.value = ""; }} />
            </label>
          </div>
        )}

        {error && (
          <div style={{
            marginTop: 12, background: "#FDECEA", border: `1px solid ${C.red}`,
            borderRadius: 8, padding: "12px 16px", color: C.red, fontSize: 13
          }}>{error}</div>
        )}

        {/* File loaded indicator */}
        {data && fileName && (
          <div style={{
            marginBottom: 20, display: "flex", alignItems: "center", gap: 8,
            fontSize: 12, color: C.textSub
          }}>
            <span style={{ color: C.green }}>●</span>
            <span>{fileName} — {allBranches.length} branch{allBranches.length !== 1 ? "es" : ""}</span>
          </div>
        )}

        {/* Overview */}
        {data && view === "overview" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>

            {/* Regional totals */}
            {regionalTotals && (
              <div style={{ background: C.surface, border: `2px solid ${C.accent}`, borderRadius: 12, padding: "20px 20px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 16, flexWrap: "wrap", gap: 6 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>Europe — Regional Total</div>
                  <div style={{ fontSize: 11, color: C.amber }}>
                    ⚠ Periods vary by branch: {regionalTotals.periodRange}
                  </div>
                </div>

                {/* Revenue + Fixed Cost + GM% row */}
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
                  {[
                    { label: "Revenue YTD",   value: fmtK(regionalTotals.rev),  delta: regionalTotals.vsBudRev,  deltaPy: regionalTotals.pyRev  ? regionalTotals.rev  - regionalTotals.pyRev  : null, sub: regionalTotals.pyRev  ? `PY: ${fmtK(regionalTotals.pyRev)}`   : null, pct: false },
                    { label: "Fixed Cost YTD", value: fmtK(regionalTotals.cost), delta: regionalTotals.vsBudCost, deltaPy: regionalTotals.pyCost ? regionalTotals.cost - regionalTotals.pyCost : null, sub: regionalTotals.pyCost ? `PY: ${fmtK(regionalTotals.pyCost)}` : null, pct: false, invert: true },
                    { label: "GM %", value: fmtPct(regionalTotals.gmPct), delta: regionalTotals.budGmPct ? regionalTotals.gmPct - regionalTotals.budGmPct : null, deltaPy: regionalTotals.pyGmPct ? regionalTotals.gmPct - regionalTotals.pyGmPct : null, sub: regionalTotals.budGmPct ? `Budget: ${fmtPct(regionalTotals.budGmPct)}` : null, pct: true },
                  ].map(({ label, value, delta, deltaPy, sub, pct, invert }) => (
                    <div key={label} style={{
                      flex: 1, minWidth: 130, background: C.bg,
                      border: `1px solid ${C.border}`, borderRadius: 8, padding: "14px 16px"
                    }}>
                      <div style={{ color: C.textSub, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>{label}</div>
                      <div style={{ fontSize: 20, fontWeight: 700, fontFamily: "monospace", color: C.text }}>{value}</div>
                      {sub && <div style={{ color: C.textSub, fontSize: 11, marginTop: 3 }}>{sub}</div>}
                      {delta != null && (
                        <div style={{ marginTop: 5 }}>
                          <span style={{
                            fontSize: 11, fontWeight: 600,
                            color: (invert ? delta <= 0 : delta >= 0) ? C.green : C.red,
                            background: (invert ? delta <= 0 : delta >= 0) ? "#E8F5EC" : "#FDECEA",
                            padding: "2px 6px", borderRadius: 4, whiteSpace: "nowrap"
                          }}>
                            {delta >= 0 ? "▲" : "▼"} {pct ? fmtPct(Math.abs(delta)) : fmtK(Math.abs(delta))} vs bud
                          </span>
                        </div>
                      )}
                      {deltaPy != null && (
                        <div style={{ marginTop: 4 }}>
                          <span style={{
                            fontSize: 11, fontWeight: 600,
                            color: (invert ? deltaPy <= 0 : deltaPy >= 0) ? C.green : C.red,
                            background: (invert ? deltaPy <= 0 : deltaPy >= 0) ? "#E8F5EC" : "#FDECEA",
                            padding: "2px 6px", borderRadius: 4, whiteSpace: "nowrap"
                          }}>
                            {deltaPy >= 0 ? "▲" : "▼"} {pct ? fmtPct(Math.abs(deltaPy)) : fmtK(Math.abs(deltaPy))} vs PY
                          </span>
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                {/* Operational result row */}
                <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "14px 16px" }}>
                  <div style={{ color: C.textSub, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 12 }}>Operational Result</div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    {/* Actual */}
                    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px 14px", flex: 1, minWidth: 120 }}>
                      <div style={{ color: C.textSub, fontSize: 10, marginBottom: 4 }}>Actual YTD</div>
                      <div style={{ fontSize: 17, fontWeight: 700, fontFamily: "monospace", color: regionalTotals.opResult >= 0 ? C.green : C.red }}>
                        {fmtK(regionalTotals.opResult)}
                      </div>
                    </div>
                    {/* Budget */}
                    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px 14px", flex: 1, minWidth: 120 }}>
                      <div style={{ color: C.textSub, fontSize: 10, marginBottom: 4 }}>Budget YTD</div>
                      <div style={{ fontSize: 17, fontWeight: 700, fontFamily: "monospace", color: regionalTotals.opResultBud >= 0 ? C.green : C.red, marginBottom: 8 }}>
                        {fmtK(regionalTotals.opResultBud)}
                      </div>
                      <span style={{
                        fontSize: 11, fontWeight: 600, fontFamily: "monospace",
                        color: regionalTotals.opResult >= regionalTotals.opResultBud ? C.green : C.red,
                        background: regionalTotals.opResult >= regionalTotals.opResultBud ? "#E8F5EC" : "#FDECEA",
                        padding: "2px 7px", borderRadius: 4, whiteSpace: "nowrap"
                      }}>
                        {regionalTotals.opResult >= regionalTotals.opResultBud ? "▲" : "▼"} {fmtK(Math.abs(regionalTotals.opResult - regionalTotals.opResultBud))} vs bud
                      </span>
                    </div>
                    {/* Prior Year */}
                    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px 14px", flex: 1, minWidth: 120 }}>
                      <div style={{ color: C.textSub, fontSize: 10, marginBottom: 4 }}>Prior Year YTD</div>
                      <div style={{ fontSize: 17, fontWeight: 700, fontFamily: "monospace", color: regionalTotals.opResultPy >= 0 ? C.green : C.red, marginBottom: 8 }}>
                        {fmtK(regionalTotals.opResultPy)}
                      </div>
                      <span style={{
                        fontSize: 11, fontWeight: 600, fontFamily: "monospace",
                        color: regionalTotals.opResult >= regionalTotals.opResultPy ? C.green : C.red,
                        background: regionalTotals.opResult >= regionalTotals.opResultPy ? "#E8F5EC" : "#FDECEA",
                        padding: "2px 7px", borderRadius: 4, whiteSpace: "nowrap"
                      }}>
                        {regionalTotals.opResult >= regionalTotals.opResultPy ? "▲" : "▼"} {fmtK(Math.abs(regionalTotals.opResult - regionalTotals.opResultPy))} vs PY
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Branch table */}
            <div>
              <div style={{ fontWeight: 600, fontSize: 13, color: C.textSub, marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                Branch breakdown — click to open
              </div>
              <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                      {["Branch","Period","Revenue YTD","Fixed Cost","GM %","vs Rev Budget","Op Result"].map(h => (
                        <th key={h} style={{
                          padding: "11px 14px", color: C.textSub, fontWeight: 600,
                          textAlign: h === "Branch" ? "left" : "right",
                          fontSize: 10, textTransform: "uppercase", letterSpacing: "0.05em"
                        }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {overviewRows.map((r, i) => (
                      <tr key={r.branch} className="row-hover" onClick={() => selectBranch(r.branch)}
                        style={{ borderBottom: i < overviewRows.length - 1 ? `1px solid ${C.border}` : "none", background: "transparent", transition: "background 0.15s" }}>
                        <td style={{ padding: "11px 14px", fontWeight: 600 }}>{r.branch}</td>
                        <td style={{ padding: "11px 14px", color: C.textSub, textAlign: "right", fontSize: 11 }}>{r.latest}</td>
                        <td style={{ padding: "11px 14px", textAlign: "right", fontFamily: "monospace" }}>{fmtK(r.rev)}</td>
                        <td style={{ padding: "11px 14px", textAlign: "right", fontFamily: "monospace" }}>
                          {fmtK(r.cost)}
                          {r.vsBudCost != null && (
                            <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600,
                              color: r.vsBudCost <= 0 ? C.green : C.red }}>
                              {r.vsBudCost <= 0 ? "▼" : "▲"}{fmtK(Math.abs(r.vsBudCost))}
                            </span>
                          )}
                        </td>
                        <td style={{ padding: "11px 14px", textAlign: "right", fontFamily: "monospace" }}>
                          {fmtPct(r.gmPct)}
                          {r.budGmPct ? (
                            <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600,
                              color: r.gmPct >= r.budGmPct ? C.green : C.red }}>
                              {r.gmPct >= r.budGmPct ? "▲" : "▼"}{fmtPct(Math.abs(r.gmPct - r.budGmPct))}
                            </span>
                          ) : null}
                        </td>
                        <td style={{ padding: "11px 14px", textAlign: "right" }}>
                          {r.vsBudRev != null && (
                            <span style={{ color: r.vsBudRev >= 0 ? C.green : C.red, fontFamily: "monospace", fontWeight: 600 }}>
                              {r.vsBudRev >= 0 ? "▲" : "▼"} {fmtK(Math.abs(r.vsBudRev))}
                            </span>
                          )}
                        </td>
                        <td style={{ padding: "11px 14px", textAlign: "right", fontFamily: "monospace",
                          color: r.opResult >= 0 ? C.green : C.red, fontWeight: 600 }}>
                          {fmtK(r.opResult)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}

        {/* Dashboard */}
        {data && view === "dashboard" && (
          <>
            <div style={{ position: "relative", maxWidth: 440, marginBottom: 28 }}>
              <div style={{ position: "absolute", left: 13, top: "50%", transform: "translateY(-50%)", color: C.textSub, fontSize: 16, pointerEvents: "none" }}>⌕</div>
              <input value={query} onChange={onQueryChange} placeholder="Type a branch name..."
                style={{
                  width: "100%", padding: "12px 14px 12px 36px",
                  background: C.surface, border: `1px solid ${branch ? C.accent : C.border}`,
                  borderRadius: 10, color: C.text, fontSize: 15, fontFamily: "inherit",
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

            {!branch && (
              <div style={{ textAlign: "center", padding: "40px 0", color: C.textSub }}>
                <div style={{ fontSize: 28, marginBottom: 10 }}>🔍</div>
                <div style={{ fontSize: 13, marginBottom: 16 }}>Select a branch to view its dashboard</div>
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

            {dashboard && branch && (
              <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
                <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
                  <h2 style={{ margin: 0, fontSize: 24, fontWeight: 700 }}>{branch}</h2>
                  <span style={{ color: C.textSub, fontSize: 13 }}>YTD {dashboard.currentYear} to {dashboard.latestMonth}</span>
                </div>

                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <KpiCard label="Revenue YTD" value={fmtK(dashboard.ytd.revenue)}
                    sub={dashboard.bud.revenue ? `Budget: ${fmtK(dashboard.bud.revenue)}` : null}
                    sub2={dashboard.py.revenue ? `PY: ${fmtK(dashboard.py.revenue)}` : null}
                    delta={dashboard.bud.revenue ? dashboard.ytd.revenue - dashboard.bud.revenue : null}
                    deltaPy={dashboard.py.revenue ? dashboard.ytd.revenue - dashboard.py.revenue : null} />
                  <KpiCard label="Gross Margin" value={fmtK(dashboard.ytd.gm)}
                    sub={dashboard.bud.gm ? `Budget: ${fmtK(dashboard.bud.gm)}` : null}
                    sub2={dashboard.py.gm ? `PY: ${fmtK(dashboard.py.gm)}` : null}
                    delta={dashboard.bud.gm ? dashboard.ytd.gm - dashboard.bud.gm : null}
                    deltaPy={dashboard.py.gm ? dashboard.ytd.gm - dashboard.py.gm : null} />
                  <KpiCard label="GM %" value={fmtPct(dashboard.ytd.gmPct)}
                    sub={dashboard.bud.gmPct ? `Budget: ${fmtPct(dashboard.bud.gmPct)}` : null}
                    sub2={dashboard.py.gmPct ? `PY: ${fmtPct(dashboard.py.gmPct)}` : null}
                    delta={dashboard.bud.gmPct ? dashboard.ytd.gmPct - dashboard.bud.gmPct : null}
                    deltaPy={dashboard.py.gmPct ? dashboard.ytd.gmPct - dashboard.py.gmPct : null}
                    pct={true} />
                </div>

                {dashboard.chartData.length > 1 && (
                  <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "18px 16px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
                      <div style={{ fontWeight: 600, fontSize: 13 }}>Monthly trend</div>
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

                {/* P&L Result */}
                <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "18px 16px" }}>
                  <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 16 }}>Operational result</div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    {/* Actual */}
                    <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "14px 16px", flex: 1, minWidth: 130 }}>
                      <div style={{ color: C.textSub, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>Actual YTD</div>
                      <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "monospace", color: dashboard.ytd.opResult >= 0 ? C.green : C.red }}>
                        {fmtK(dashboard.ytd.opResult)}
                      </div>
                    </div>
                    {/* Budget */}
                    <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "14px 16px", flex: 1, minWidth: 130 }}>
                      <div style={{ color: C.textSub, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>Budget YTD</div>
                      <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "monospace", color: dashboard.bud.opResult >= 0 ? C.green : C.red, marginBottom: 8 }}>
                        {fmtK(dashboard.bud.opResult)}
                      </div>
                      <span style={{
                        fontSize: 11, fontWeight: 600, fontFamily: "monospace",
                        color: dashboard.ytd.opResult >= dashboard.bud.opResult ? C.green : C.red,
                        background: dashboard.ytd.opResult >= dashboard.bud.opResult ? "#E8F5EC" : "#FDECEA",
                        padding: "2px 7px", borderRadius: 4, whiteSpace: "nowrap"
                      }}>
                        {dashboard.ytd.opResult >= dashboard.bud.opResult ? "▲" : "▼"} {fmtK(Math.abs(dashboard.ytd.opResult - dashboard.bud.opResult))} vs bud
                      </span>
                    </div>
                    {/* Prior Year */}
                    <div style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "14px 16px", flex: 1, minWidth: 130 }}>
                      <div style={{ color: C.textSub, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>Prior Year YTD</div>
                      <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "monospace", color: dashboard.py.opResult >= 0 ? C.green : C.red, marginBottom: 8 }}>
                        {fmtK(dashboard.py.opResult)}
                      </div>
                      <span style={{
                        fontSize: 11, fontWeight: 600, fontFamily: "monospace",
                        color: dashboard.ytd.opResult >= dashboard.py.opResult ? C.green : C.red,
                        background: dashboard.ytd.opResult >= dashboard.py.opResult ? "#E8F5EC" : "#FDECEA",
                        padding: "2px 7px", borderRadius: 4, whiteSpace: "nowrap"
                      }}>
                        {dashboard.ytd.opResult >= dashboard.py.opResult ? "▲" : "▼"} {fmtK(Math.abs(dashboard.ytd.opResult - dashboard.py.opResult))} vs PY
                      </span>
                    </div>
                  </div>
                </div>

                {dashboard.topics.length > 0 && (
                  <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "18px 16px" }}>
                    <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 14 }}>Current topics</div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {dashboard.topics.map((t, i) => (
                        <div key={i} style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
                          <div style={{ width: 6, height: 6, background: C.accentLt, borderRadius: "50%", marginTop: 6, flexShrink: 0 }} />
                          <div style={{ fontSize: 13, color: C.text, lineHeight: 1.6 }}>{t}</div>
                        </div>
                      ))}
                    </div>
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
