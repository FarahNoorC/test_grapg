import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  ScatterChart, Scatter, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer
} from "recharts";

const COLORS = ["#378ADD","#1D9E75","#D85A30","#7F77DD","#BA7517","#D4537E","#639922","#E24B4A"];

const CHART_TYPES = [
  { id: "bar", label: "Bar" },
  { id: "line", label: "Line" },
  { id: "pie", label: "Pie" },
  { id: "scatter", label: "Scatter" },
];

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function ChartRenderer({ chartType, data, xKey, yKeys }) {
  if (!data || data.length === 0 || !xKey || yKeys.length === 0) return null;

  const commonProps = {
    data,
    margin: { top: 10, right: 20, left: 0, bottom: 40 }
  };

  if (chartType === "pie") {
    const flat = data.map(d => ({ name: d[xKey], value: parseFloat(d[yKeys[0]]) || 0 }));
    return (
      <ResponsiveContainer width="100%" height={320}>
        <PieChart>
          <Pie data={flat} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={110} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}>
            {flat.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
          </Pie>
          <Tooltip />
          <Legend />
        </PieChart>
      </ResponsiveContainer>
    );
  }

  if (chartType === "scatter") {
    return (
      <ResponsiveContainer width="100%" height={320}>
        <ScatterChart {...commonProps}>
          <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.08)" />
          <XAxis dataKey={xKey} name={xKey} tick={{ fontSize: 11 }} angle={-30} textAnchor="end" />
          <YAxis tick={{ fontSize: 11 }} />
          <Tooltip cursor={{ strokeDasharray: "3 3" }} />
          {yKeys.map((k, i) => (
            <Scatter key={k} name={k} data={data.map(d => ({ [xKey]: d[xKey], [k]: parseFloat(d[k]) || 0 }))} dataKey={k} fill={COLORS[i % COLORS.length]} />
          ))}
          <Legend />
        </ScatterChart>
      </ResponsiveContainer>
    );
  }

  if (chartType === "line") {
    return (
      <ResponsiveContainer width="100%" height={320}>
        <LineChart {...commonProps}>
          <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.08)" />
          <XAxis dataKey={xKey} tick={{ fontSize: 11 }} angle={-30} textAnchor="end" />
          <YAxis tick={{ fontSize: 11 }} />
          <Tooltip />
          <Legend />
          {yKeys.map((k, i) => <Line key={k} type="monotone" dataKey={k} stroke={COLORS[i % COLORS.length]} strokeWidth={2} dot={{ r: 3 }} />)}
        </LineChart>
      </ResponsiveContainer>
    );
  }

  return (
    <ResponsiveContainer width="100%" height={320}>
      <BarChart {...commonProps}>
        <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.08)" />
        <XAxis dataKey={xKey} tick={{ fontSize: 11 }} angle={-30} textAnchor="end" />
        <YAxis tick={{ fontSize: 11 }} />
        <Tooltip />
        <Legend />
        {yKeys.map((k, i) => <Bar key={k} dataKey={k} fill={COLORS[i % COLORS.length]} radius={[3,3,0,0]} />)}
      </BarChart>
    </ResponsiveContainer>
  );
}

export default function App() {
  const [tab, setTab] = useState("viz");

  // --- Viz tab state ---
  const [xlData, setXlData] = useState(null);
  const [xlCols, setXlCols] = useState([]);
  const [xlSheets, setXlSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [wbRef, setWbRef] = useState(null);
  const [chartType, setChartType] = useState("bar");
  const [xKey, setXKey] = useState("");
  const [yKeys, setYKeys] = useState([]);
  const [chartTitle, setChartTitle] = useState("My Chart");
  const fileInputRef = useRef();
  const chartRef = useRef();

  // --- Storage tab state ---
  const [storageCols, setStorageCols] = useState(["Name", "Value", "Notes"]);
  const [storageRows, setStorageRows] = useState([["", "", ""]]);
  const [newColName, setNewColName] = useState("");

  // --- Viz handlers ---
  function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "array" });
      setWbRef(wb);
      setXlSheets(wb.SheetNames);
      loadSheet(wb, wb.SheetNames[0]);
      setSelectedSheet(wb.SheetNames[0]);
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSheet(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
    setXlData(json);
    const cols = json.length > 0 ? Object.keys(json[0]) : [];
    setXlCols(cols);
    setXKey(cols[0] || "");
    setYKeys(cols.slice(1, 3));
  }

  function handleSheetChange(s) {
    setSelectedSheet(s);
    loadSheet(wbRef, s);
  }

  function toggleYKey(k) {
    setYKeys(prev => prev.includes(k) ? prev.filter(x => x !== k) : [...prev, k]);
  }

  function handleDownloadChart() {
    const svgEl = chartRef.current?.querySelector("svg");
    if (!svgEl) return;
    const serializer = new XMLSerializer();
    const svgStr = serializer.serializeToString(svgEl);
    const blob = new Blob([svgStr], { type: "image/svg+xml" });
    downloadBlob(blob, `${chartTitle.replace(/\s+/g, "_")}.svg`);
  }

  // --- Storage handlers ---
  function addRow() {
    setStorageRows(prev => [...prev, storageCols.map(() => "")]);
  }

  function removeRow(i) {
    setStorageRows(prev => prev.filter((_, idx) => idx !== i));
  }

  function updateCell(ri, ci, val) {
    setStorageRows(prev => prev.map((r, rIdx) => rIdx === ri ? r.map((c, cIdx) => cIdx === ci ? val : c) : r));
  }

  function addColumn() {
    const name = newColName.trim() || `Col ${storageCols.length + 1}`;
    setStorageCols(prev => [...prev, name]);
    setStorageRows(prev => prev.map(r => [...r, ""]));
    setNewColName("");
  }

  function removeColumn(ci) {
    setStorageCols(prev => prev.filter((_, i) => i !== ci));
    setStorageRows(prev => prev.map(r => r.filter((_, i) => i !== ci)));
  }

  function downloadStorageXLSX() {
    const ws = XLSX.utils.aoa_to_sheet([storageCols, ...storageRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data");
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([buf], { type: "application/octet-stream" }), "data_storage.xlsx");
  }

  function downloadStorageCSV() {
    const rows = [storageCols, ...storageRows].map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    downloadBlob(new Blob([rows], { type: "text/csv" }), "data_storage.csv");
  }

  const s = {
    app: { fontFamily: "system-ui, sans-serif", maxWidth: 860, margin: "0 auto", padding: "1rem" },
    header: { marginBottom: "1.25rem" },
    title: { fontSize: 22, fontWeight: 500, color: "var(--color-text-primary)", margin: 0 },
    sub: { fontSize: 13, color: "var(--color-text-secondary)", marginTop: 4 },
    tabs: { display: "flex", gap: 4, marginBottom: "1.5rem", borderBottom: "0.5px solid var(--color-border-tertiary)", paddingBottom: 0 },
    tab: (active) => ({
      padding: "8px 20px", fontSize: 14, fontWeight: active ? 500 : 400,
      color: active ? "var(--color-text-primary)" : "var(--color-text-secondary)",
      background: "transparent", border: "none", borderBottom: active ? "2px solid var(--color-text-primary)" : "2px solid transparent",
      cursor: "pointer", marginBottom: -1
    }),
    card: { background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "1.25rem", marginBottom: "1rem" },
    label: { fontSize: 13, fontWeight: 500, color: "var(--color-text-secondary)", marginBottom: 6, display: "block" },
    row: { display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" },
    select: { fontSize: 13, padding: "6px 10px", borderRadius: "var(--border-radius-md)", flex: 1, minWidth: 120 },
    btn: (variant = "default") => ({
      padding: "7px 16px", fontSize: 13, borderRadius: "var(--border-radius-md)", cursor: "pointer", fontWeight: 500,
      background: variant === "primary" ? "#378ADD" : variant === "success" ? "#1D9E75" : "var(--color-background-secondary)",
      color: variant === "primary" || variant === "success" ? "#fff" : "var(--color-text-primary)",
      border: variant === "primary" || variant === "success" ? "none" : "0.5px solid var(--color-border-secondary)"
    }),
    pill: (active) => ({
      padding: "4px 10px", fontSize: 12, borderRadius: 20, cursor: "pointer",
      background: active ? "#378ADD" : "var(--color-background-secondary)",
      color: active ? "#fff" : "var(--color-text-secondary)",
      border: active ? "none" : "0.5px solid var(--color-border-tertiary)"
    }),
    uploadBox: {
      border: "1.5px dashed var(--color-border-secondary)", borderRadius: "var(--border-radius-lg)",
      padding: "2rem", textAlign: "center", cursor: "pointer", marginBottom: "1rem",
      background: "var(--color-background-secondary)"
    },
    table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
    th: { background: "var(--color-background-secondary)", padding: "8px 10px", fontWeight: 500, textAlign: "left", borderBottom: "0.5px solid var(--color-border-tertiary)", fontSize: 12, color: "var(--color-text-secondary)" },
    td: { padding: "6px 8px", borderBottom: "0.5px solid var(--color-border-tertiary)" },
    input: { width: "100%", fontSize: 13, padding: "4px 6px", borderRadius: "var(--border-radius-md)", boxSizing: "border-box" },
  };

  return (
    <div style={s.app}>
      <h2 style={{ display: "none" }}>Data Studio — visualization and storage tool</h2>
      <div style={s.header}>
        <h1 style={s.title}>Data Studio</h1>
        <p style={s.sub}>Upload Excel files to visualize · Build and export custom datasets</p>
      </div>

      <div style={s.tabs}>
        <button style={s.tab(tab === "viz")} onClick={() => setTab("viz")}>📊 Visualizer</button>
        <button style={s.tab(tab === "store")} onClick={() => setTab("store")}>🗄️ Data Storage</button>
      </div>

      {/* ===== VIZ TAB ===== */}
      {tab === "viz" && (
        <div>
          {/* Upload */}
          <div style={s.card}>
            <label style={s.label}>Upload Excel file (.xlsx / .xls / .csv)</label>
            <div style={s.uploadBox} onClick={() => fileInputRef.current.click()}>
              <div style={{ fontSize: 28, marginBottom: 8 }}>📂</div>
              <div style={{ fontSize: 14, color: "var(--color-text-secondary)" }}>Click to upload or drag & drop</div>
              <div style={{ fontSize: 12, color: "var(--color-text-tertiary)", marginTop: 4 }}>Supports .xlsx, .xls, .csv</div>
            </div>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={handleFileUpload} />
            {xlSheets.length > 1 && (
              <div style={{ marginTop: 8 }}>
                <label style={s.label}>Sheet</label>
                <select style={s.select} value={selectedSheet} onChange={e => handleSheetChange(e.target.value)}>
                  {xlSheets.map(s => <option key={s}>{s}</option>)}
                </select>
              </div>
            )}
          </div>

          {xlData && xlData.length > 0 && (
            <>
              {/* Chart config */}
              <div style={s.card}>
                <div style={s.row}>
                  <div style={{ flex: 1, minWidth: 160 }}>
                    <label style={s.label}>Chart type</label>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      {CHART_TYPES.map(ct => (
                        <button key={ct.id} style={s.pill(chartType === ct.id)} onClick={() => setChartType(ct.id)}>{ct.label}</button>
                      ))}
                    </div>
                  </div>
                  <div style={{ flex: 1, minWidth: 140 }}>
                    <label style={s.label}>X axis / Label column</label>
                    <select style={s.select} value={xKey} onChange={e => setXKey(e.target.value)}>
                      {xlCols.map(c => <option key={c}>{c}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: 2, minWidth: 200 }}>
                    <label style={s.label}>Y axis columns (select multiple)</label>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      {xlCols.filter(c => c !== xKey).map(c => (
                        <button key={c} style={s.pill(yKeys.includes(c))} onClick={() => toggleYKey(c)}>{c}</button>
                      ))}
                    </div>
                  </div>
                </div>
                <div style={{ marginTop: 12, display: "flex", gap: 12, alignItems: "center" }}>
                  <div style={{ flex: 1 }}>
                    <label style={s.label}>Chart title</label>
                    <input style={{ ...s.input, width: "auto", minWidth: 200 }} value={chartTitle} onChange={e => setChartTitle(e.target.value)} />
                  </div>
                </div>
              </div>

              {/* Chart output */}
              <div style={s.card}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <span style={{ fontWeight: 500, fontSize: 15 }}>{chartTitle}</span>
                  <button style={s.btn("primary")} onClick={handleDownloadChart}>⬇ Download SVG</button>
                </div>
                <div ref={chartRef}>
                  <ChartRenderer chartType={chartType} data={xlData} xKey={xKey} yKeys={yKeys} />
                </div>
                <div style={{ marginTop: 8, fontSize: 12, color: "var(--color-text-tertiary)" }}>
                  {xlData.length} rows · {xlCols.length} columns
                </div>
              </div>

              {/* Data preview */}
              <div style={s.card}>
                <label style={s.label}>Data preview (first 10 rows)</label>
                <div style={{ overflowX: "auto" }}>
                  <table style={s.table}>
                    <thead>
                      <tr>{xlCols.map(c => <th key={c} style={s.th}>{c}</th>)}</tr>
                    </thead>
                    <tbody>
                      {xlData.slice(0, 10).map((row, i) => (
                        <tr key={i}>
                          {xlCols.map(c => <td key={c} style={s.td}>{String(row[c] ?? "")}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}

          {!xlData && (
            <div style={{ ...s.card, textAlign: "center", padding: "3rem", color: "var(--color-text-tertiary)" }}>
              <div style={{ fontSize: 40, marginBottom: 12 }}>📈</div>
              <div style={{ fontSize: 15 }}>Upload an Excel or CSV file to get started</div>
            </div>
          )}
        </div>
      )}

      {/* ===== STORAGE TAB ===== */}
      {tab === "store" && (
        <div>
          <div style={s.card}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <span style={{ fontWeight: 500, fontSize: 15 }}>Data Table ({storageRows.length} rows)</span>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button style={s.btn()} onClick={addRow}>+ Add row</button>
                <button style={s.btn("success")} onClick={downloadStorageXLSX}>⬇ Download XLSX</button>
                <button style={s.btn()} onClick={downloadStorageCSV}>⬇ Download CSV</button>
              </div>
            </div>

            {/* Add column */}
            <div style={{ display: "flex", gap: 8, marginBottom: 14, alignItems: "center" }}>
              <input
                style={{ ...s.input, width: 160 }}
                placeholder="New column name"
                value={newColName}
                onChange={e => setNewColName(e.target.value)}
                onKeyDown={e => e.key === "Enter" && addColumn()}
              />
              <button style={s.btn()} onClick={addColumn}>+ Add column</button>
            </div>

            {/* Table */}
            <div style={{ overflowX: "auto" }}>
              <table style={s.table}>
                <thead>
                  <tr>
                    <th style={{ ...s.th, width: 32 }}>#</th>
                    {storageCols.map((col, ci) => (
                      <th key={ci} style={s.th}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span>{col}</span>
                          {storageCols.length > 1 && (
                            <button onClick={() => removeColumn(ci)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--color-text-tertiary)", fontSize: 11, padding: 0 }}>✕</button>
                          )}
                        </div>
                      </th>
                    ))}
                    <th style={{ ...s.th, width: 40 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {storageRows.map((row, ri) => (
                    <tr key={ri}>
                      <td style={{ ...s.td, color: "var(--color-text-tertiary)", fontSize: 11, textAlign: "center" }}>{ri + 1}</td>
                      {row.map((cell, ci) => (
                        <td key={ci} style={s.td}>
                          <input
                            style={s.input}
                            value={cell}
                            onChange={e => updateCell(ri, ci, e.target.value)}
                            placeholder="—"
                          />
                        </td>
                      ))}
                      <td style={s.td}>
                        <button onClick={() => removeRow(ri)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--color-text-tertiary)", fontSize: 14 }}>🗑</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div style={{ marginTop: 12, fontSize: 12, color: "var(--color-text-tertiary)" }}>
              {storageCols.length} columns · {storageRows.length} rows · Click any cell to edit
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
