import { useState, useRef, useCallback, useEffect, useReducer } from "react";
import * as XLSX from "xlsx";

/* ─── Config ─────────────────────────────────────────────────────────────────
   In development the Vite proxy forwards /api → http://localhost:3001
   In production set VITE_API_URL to your deployed server URL, e.g.
     VITE_API_URL=https://xl-forge-api.onrender.com
─────────────────────────────────────────────────────────────────────────── */
const API = import.meta.env.VITE_API_URL || "";

/* ─── Design tokens (WCAG 2.1 AA) ───────────────────────────────────────── */
const C = {
  bg:       "#0B0D14",
  surface:  "#131620",
  surface2: "#1A1E2E",
  border:   "#252A3E",
  border2:  "#303654",
  accent:   "#3B82F6",
  accentHi: "#60A5FA",
  green:    "#22C55E",
  red:      "#F87171",
  yellow:   "#FBBF24",
  text:     "#F1F3FA",
  text2:    "#8892B0",
  text3:    "#4A527A",
  cellBg:   "#0F1119",
  cellSel:  "#1E2D4A",
  headBg:   "#0D1020",
  headText: "#7B88AA",
};

/* ─── Global CSS ─────────────────────────────────────────────────────────── */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --accent: ${C.accent}; --accent-hi: ${C.accentHi};
  --green: ${C.green};   --red: ${C.red};
}
body { background:${C.bg}; color:${C.text}; font-family:'Plus Jakarta Sans',sans-serif; }

:focus-visible { outline:2px solid var(--accent); outline-offset:2px; border-radius:4px; }
:focus:not(:focus-visible) { outline:none; }

::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background:transparent; }
::-webkit-scrollbar-thumb { background:${C.border2}; border-radius:99px; }

.skip-link {
  position:fixed; top:-100%; left:16px; z-index:9999;
  padding:8px 16px; background:var(--accent); color:#fff;
  font-weight:700; border-radius:0 0 8px 8px; text-decoration:none;
  font-size:13px; font-family:'Plus Jakarta Sans',sans-serif; transition:top .15s;
}
.skip-link:focus { top:0; }

/* Toolbar buttons */
.tbtn {
  display:inline-flex; align-items:center; gap:6px;
  padding:7px 14px; border-radius:8px; border:1px solid ${C.border2};
  background:${C.surface2}; color:${C.text2}; font-size:12px; font-weight:600;
  font-family:inherit; cursor:pointer; white-space:nowrap;
  transition:background .15s, color .15s, border-color .15s;
}
.tbtn:hover, .tbtn:focus-visible { background:${C.border2}; color:${C.text}; border-color:${C.accent}60; }
.tbtn.danger:hover  { background:#3A1010; color:var(--red); border-color:var(--red)60; }
.tbtn.success { background:#14301A; color:var(--green); border:1px solid ${C.green}40; }
.tbtn.primary { background:var(--accent); color:#fff; border-color:transparent; }
.tbtn.primary:hover, .tbtn.primary:focus-visible { background:var(--accent-hi); }
.tbtn:disabled { opacity:.4; cursor:not-allowed; pointer-events:none; }

/* Grid */
.grid-wrap { overflow:auto; flex:1; min-height:0; background:${C.cellBg}; }
.grid-table {
  border-collapse:collapse; font-family:'JetBrains Mono',monospace;
  font-size:13px; table-layout:fixed; min-width:100%;
}
.grid-table th {
  position:sticky; top:0; z-index:2; background:${C.headBg}; color:${C.headText};
  font-size:11px; font-weight:600; font-family:'Plus Jakarta Sans',sans-serif;
  padding:0; text-align:center; border-right:1px solid ${C.border};
  border-bottom:2px solid ${C.border2}; user-select:none; letter-spacing:.05em;
}
.grid-table th.row-num-head { width:48px; min-width:48px; max-width:48px; left:0; z-index:3; }
.grid-table th .th-inner { padding:8px 10px; display:flex; align-items:center; justify-content:center; gap:4px; white-space:nowrap; }
.grid-table td {
  border-right:1px solid ${C.border}; border-bottom:1px solid ${C.border};
  padding:0; color:${C.text}; vertical-align:middle; min-width:120px; max-width:240px;
}
.grid-table td.row-num {
  position:sticky; left:0; z-index:1; background:${C.headBg}; color:${C.headText};
  font-size:11px; font-family:'Plus Jakarta Sans',sans-serif;
  text-align:center; padding:0 8px; min-width:48px; max-width:48px; width:48px;
  border-right:2px solid ${C.border2}; user-select:none;
}
.grid-table tr:hover td { background:${C.surface}; }
.grid-table tr:hover td.row-num { background:${C.headBg}; }

.cell-display {
  padding:7px 10px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
  width:100%; height:100%; display:block; cursor:cell; min-height:33px; line-height:1.4;
}
.cell-display.selected { background:${C.cellSel}; box-shadow:inset 0 0 0 2px var(--accent); }
.cell-input {
  width:100%; height:100%; padding:7px 10px; background:${C.cellSel};
  border:none; outline:2px solid var(--accent); color:${C.text};
  font-family:'JetBrains Mono',monospace; font-size:13px; line-height:1.4;
  min-height:33px; box-shadow:0 4px 20px rgba(59,130,246,.25);
}

/* Sheet tabs */
.sheet-tabs {
  display:flex; gap:2px; padding:0 16px;
  background:${C.surface}; border-bottom:1px solid ${C.border}; overflow-x:auto;
}
.sheet-tab {
  padding:9px 16px 8px; font-size:12px; font-weight:600; cursor:pointer;
  border:none; background:transparent; color:${C.text2};
  border-bottom:2px solid transparent; white-space:nowrap; font-family:inherit;
  transition:color .15s, border-color .15s;
}
.sheet-tab:hover { color:${C.text}; }
.sheet-tab.active { color:${C.accentHi}; border-bottom-color:var(--accent); }

/* AI Panel */
.ai-panel { background:${C.surface}; border-top:1px solid ${C.border}; display:flex; flex-direction:column; }
.ai-header { display:flex; align-items:center; justify-content:space-between; padding:12px 20px; border-bottom:1px solid ${C.border}; }
.ai-body { padding:14px 20px; display:flex; flex-direction:column; gap:10px; }
.ai-textarea {
  width:100%; background:${C.surface2}; border:1px solid ${C.border2};
  border-radius:10px; padding:10px 14px; font-size:13px; color:${C.text};
  font-family:'Plus Jakarta Sans',sans-serif; resize:vertical; min-height:70px;
  line-height:1.6; caret-color:var(--accent); transition:border-color .2s;
}
.ai-textarea:focus { border-color:var(--accent); outline:none; }
.ai-footer { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }

/* Confirm dialog */
.confirm-overlay {
  position:fixed; inset:0; z-index:1000; background:rgba(0,0,0,.75);
  backdrop-filter:blur(4px); display:flex; align-items:center; justify-content:center; padding:20px;
}
.confirm-box {
  background:${C.surface}; border:1px solid ${C.border2}; border-radius:16px;
  padding:28px 32px; max-width:560px; width:100%; box-shadow:0 24px 64px rgba(0,0,0,.6);
}
.confirm-changes {
  background:${C.surface2}; border:1px solid ${C.border}; border-radius:10px;
  padding:14px 16px; margin:14px 0; font-size:13px; font-family:'JetBrains Mono',monospace;
  color:${C.text2}; max-height:200px; overflow-y:auto; line-height:1.7;
}

/* Log */
.log-panel {
  background:${C.surface2}; border-top:1px solid ${C.border};
  padding:8px 20px; font-size:11px; font-family:'JetBrains Mono',monospace;
  color:${C.text3}; max-height:90px; overflow-y:auto;
}
.log-line { padding:2px 0; line-height:1.6; }
.log-line.success { color:var(--green); }
.log-line.error   { color:var(--red); }
.log-line.info    { color:${C.accentHi}; }

/* Status bar */
.status-bar {
  padding:5px 20px; background:${C.surface}; border-top:1px solid ${C.border};
  font-size:11px; color:${C.text3}; display:flex; gap:20px; align-items:center;
  font-family:'JetBrains Mono',monospace; flex-wrap:wrap;
}

/* Loading screen */
.loading-screen {
  flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:20px;
}

/* Example chips */
.example-chip {
  background:${C.surface2}; border:1px solid ${C.border2}; color:${C.text2};
  border-radius:6px; padding:4px 10px; font-size:11px; cursor:pointer;
  font-family:inherit; transition:all .15s; white-space:nowrap;
}
.example-chip:hover, .example-chip:focus-visible {
  background:${C.border}; color:${C.text}; border-color:var(--accent)60;
}

/* Saved badge */
.saved-badge {
  font-size:10px; font-weight:700; letter-spacing:.06em; text-transform:uppercase;
  padding:2px 8px; border-radius:4px;
}

/* Spinner */
@keyframes spin { to { transform:rotate(360deg); } }
.spinner {
  width:13px; height:13px; border:2px solid transparent;
  border-top-color:currentColor; border-radius:50%;
  animation:spin .7s linear infinite; display:inline-block;
}
@keyframes fadeUp { from{opacity:0;transform:translateY(6px)} to{opacity:1;transform:translateY(0)} }
.fade-up { animation:fadeUp .2s ease forwards; }

@media (prefers-reduced-motion:reduce) {
  *,*::before,*::after { animation-duration:.001ms!important; transition-duration:.001ms!important; }
}
@media (max-width:640px) {
  .ai-footer { flex-direction:column; align-items:stretch; }
  .status-bar { gap:10px; }
}
`;

/* ─── Helpers ────────────────────────────────────────────────────────────── */
const colLabel = (i) => {
  let s = ""; i++;
  while (i > 0) { s = String.fromCharCode(64 + (i % 26 || 26)) + s; i = Math.floor((i - 1) / 26); }
  return s;
};
const normalize = (data) => {
  const w = Math.max(0, ...data.map((r) => r.length));
  return data.map((r) => { const row = [...r]; while (row.length < w) row.push(""); return row; });
};

const EXAMPLES = [
  "Add a Total row at the bottom summing all numeric columns",
  "Sort rows by the first column alphabetically",
  "Add a new column called 'Status' with value 'Pending' for all rows",
  "Rename the first column header to 'ID'",
];

/* ─── State reducer ──────────────────────────────────────────────────────── */
function reducer(state, action) {
  switch (action.type) {
    case "LOAD_SUCCESS": {
      const { sheetNames, sheets, fileName, fileSize, lastModified } = action;
      const activeSheet = sheetNames[0];
      return {
        ...state,
        sheetNames, sheets, fileName, fileSize, lastModified,
        activeSheet,
        data: normalize(sheets[activeSheet] || [[]]),
        undoStack: [], redoStack: [], dirty: false, loaded: true,
      };
    }
    case "SWITCH_SHEET": {
      const sheets = { ...state.sheets, [state.activeSheet]: state.data };
      const data   = normalize(sheets[action.sheet] || [[]]);
      return { ...state, sheets, activeSheet: action.sheet, data, dirty: true };
    }
    case "SET_CELL": {
      const data = state.data.map((r) => [...r]);
      data[action.row][action.col] = action.value;
      return { ...state, data, undoStack: [...state.undoStack, state.data], redoStack: [], dirty: true };
    }
    case "BULK_SET": {
      const data = normalize(action.data);
      return { ...state, data, undoStack: [...state.undoStack, state.data], redoStack: [], dirty: true };
    }
    case "ADD_ROW": {
      const data = [...state.data, Array(state.data[0]?.length || 1).fill("")];
      return { ...state, data, undoStack: [...state.undoStack, state.data], redoStack: [], dirty: true };
    }
    case "ADD_COL": {
      const data = state.data.map((r) => [...r, ""]);
      return { ...state, data, undoStack: [...state.undoStack, state.data], redoStack: [], dirty: true };
    }
    case "DEL_ROW": {
      if (state.data.length <= 1) return state;
      const data = state.data.filter((_, i) => i !== action.row);
      return { ...state, data, undoStack: [...state.undoStack, state.data], redoStack: [], dirty: true };
    }
    case "DEL_COL": {
      if ((state.data[0]?.length || 0) <= 1) return state;
      const data = state.data.map((r) => r.filter((_, i) => i !== action.col));
      return { ...state, data, undoStack: [...state.undoStack, state.data], redoStack: [], dirty: true };
    }
    case "UNDO": {
      if (!state.undoStack.length) return state;
      const undoStack = [...state.undoStack];
      const data = undoStack.pop();
      return { ...state, data, undoStack, redoStack: [...state.redoStack, state.data], dirty: true };
    }
    case "REDO": {
      if (!state.redoStack.length) return state;
      const redoStack = [...state.redoStack];
      const data = redoStack.pop();
      return { ...state, data, redoStack, undoStack: [...state.undoStack, state.data], dirty: true };
    }
    case "MARK_SAVED":
      return { ...state, dirty: false, lastModified: action.lastModified };
    default:
      return state;
  }
}

const initState = {
  loaded: false, sheetNames: [], sheets: {}, activeSheet: null,
  data: [], fileName: "", fileSize: 0, lastModified: null,
  undoStack: [], redoStack: [], dirty: false,
};

/* ─── Cell component ─────────────────────────────────────────────────────── */
function Cell({ value, rowIdx, colIdx, selected, onSelect, onChange }) {
  const [editing, setEditing] = useState(false);
  const [draft,   setDraft]   = useState("");
  const inputRef = useRef(null);

  useEffect(() => { if (editing) inputRef.current?.focus(); }, [editing]);

  const commit = () => { setEditing(false); onChange(rowIdx, colIdx, draft); };
  const cancel = () => setEditing(false);

  const handleCellKey = (e) => {
    if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); setDraft(String(value ?? "")); setEditing(true); }
    if ((e.key === "Delete" || e.key === "Backspace") && !editing) onChange(rowIdx, colIdx, "");
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) { setDraft(e.key); setEditing(true); }
  };

  return (
    <td role="gridcell" aria-colindex={colIdx + 2} aria-rowindex={rowIdx + 2} aria-selected={selected}>
      {editing ? (
        <input
          ref={inputRef}
          className="cell-input"
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); commit(); }
            if (e.key === "Escape") cancel();
            if (e.key === "Tab") { e.preventDefault(); commit(); }
          }}
          aria-label={`Cell ${colLabel(colIdx)}${rowIdx + 1}, editing`}
        />
      ) : (
        <div
          className={`cell-display${selected ? " selected" : ""}`}
          tabIndex={0}
          role="button"
          aria-label={`Cell ${colLabel(colIdx)}${rowIdx + 1}: ${String(value ?? "") || "empty"}`}
          onClick={() => onSelect(rowIdx, colIdx)}
          onDoubleClick={() => { setDraft(String(value ?? "")); setEditing(true); }}
          onKeyDown={handleCellKey}
          onFocus={() => onSelect(rowIdx, colIdx)}
          title={String(value ?? "")}
        >
          {String(value ?? "")}
        </div>
      )}
    </td>
  );
}

/* ─── Confirm dialog ─────────────────────────────────────────────────────── */
function ConfirmDialog({ plan, onConfirm, onCancel }) {
  const ref = useRef(null);
  useEffect(() => { ref.current?.focus(); }, []);
  return (
    <div className="confirm-overlay" role="dialog" aria-modal="true" aria-labelledby="cd-title"
      onClick={(e) => e.target === e.currentTarget && onCancel()}>
      <div className="confirm-box fade-up" ref={ref} tabIndex={-1}>
        <h2 id="cd-title" style={{ fontSize: 17, fontWeight: 700, color: C.text, marginBottom: 8 }}>Review AI Changes</h2>
        <p style={{ fontSize: 13, color: C.text2 }}>These changes will be applied to the spreadsheet:</p>
        <div className="confirm-changes" role="list">
          {plan.steps?.map((s, i) => (
            <div key={i} role="listitem" style={{ display: "flex", gap: 8 }}>
              <span style={{ color: C.accent, flexShrink: 0 }}>▸</span>
              <span>{s.description}</span>
            </div>
          ))}
        </div>
        <p style={{ fontSize: 12, color: C.text3, marginBottom: 20 }}>Summary: {plan.summary}</p>
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button className="tbtn" onClick={onCancel}>Cancel</button>
          <button className="tbtn primary" onClick={onConfirm}>✓ Apply Changes</button>
        </div>
      </div>
    </div>
  );
}

/* ─── Main App ───────────────────────────────────────────────────────────── */
export default function App() {
  const [state,      dispatch]    = useReducer(reducer, initState);
  const [selCell,    setSelCell]  = useState(null);
  const [prompt,     setPrompt]   = useState("");
  const [aiPlan,     setAiPlan]   = useState(null);
  const [aiRunning,  setAiRunning]= useState(false);
  const [saving,     setSaving]   = useState(false);
  const [saveStatus, setSaveStatus]= useState(null); // "saved" | "error"
  const [loadError,  setLoadError]= useState(null);
  const [logs,       setLogs]     = useState([]);
  const [liveMsg,    setLiveMsg]  = useState("");

  const promptRef = useRef(null);

  const announce = useCallback((m) => setLiveMsg(m), []);
  const addLog   = useCallback((msg, type = "default") => {
    setLogs((p) => [...p.slice(-49), { msg, type, ts: Date.now() }]);
    if (type !== "default") announce(msg);
  }, [announce]);

  /* ── Load file from server on mount ── */
  useEffect(() => {
    const load = async () => {
      try {
        addLog("Loading spreadsheet from server…", "info");
        const res  = await fetch(`${API}/api/file`);
        if (!res.ok) throw new Error(`Server returned ${res.status}`);
        const json = await res.json();
        dispatch({ type: "LOAD_SUCCESS", ...json });
        addLog(`Loaded: ${json.fileName} (${json.sheetNames.length} sheet(s))`, "success");
        announce(`Spreadsheet loaded: ${json.fileName}`);
      } catch (err) {
        setLoadError(err.message);
        addLog(`Failed to load: ${err.message}`, "error");
      }
    };
    load();
  }, []);

  /* ── Keyboard shortcuts ── */
  useEffect(() => {
    const h = (e) => {
      if (!state.loaded) return;
      if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) { e.preventDefault(); dispatch({ type: "UNDO" }); announce("Undo"); }
      if ((e.ctrlKey || e.metaKey) && (e.key === "y" || (e.key === "z" && e.shiftKey))) { e.preventDefault(); dispatch({ type: "REDO" }); announce("Redo"); }
      if ((e.ctrlKey || e.metaKey) && e.key === "s") { e.preventDefault(); saveToServer(); }
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [state.loaded, state.data, state.sheets, state.sheetNames, state.activeSheet]);

  /* ── Build current sheets payload ── */
  const buildPayload = useCallback(() => {
    const sheets = { ...state.sheets, [state.activeSheet]: state.data };
    return { sheetNames: state.sheetNames, sheets };
  }, [state.sheets, state.activeSheet, state.data, state.sheetNames]);

  /* ── Save to server ── */
  const saveToServer = async () => {
    if (saving) return;
    setSaving(true);
    setSaveStatus(null);
    try {
      const res  = await fetch(`${API}/api/save`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(buildPayload()),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Save failed");
      dispatch({ type: "MARK_SAVED", lastModified: json.lastModified });
      setSaveStatus("saved");
      addLog("Saved to server successfully.", "success");
      announce("Changes saved to server.");
      setTimeout(() => setSaveStatus(null), 3000);
    } catch (err) {
      setSaveStatus("error");
      addLog(`Save failed: ${err.message}`, "error");
      announce(`Save failed: ${err.message}`);
    }
    setSaving(false);
  };

  /* ── Run AI prompt ── */
  const runPrompt = async () => {
    if (!state.loaded || !prompt.trim() || aiRunning) return;
    setAiRunning(true);
    announce("Running AI transformation, please wait.");
    try {
      const csvPreview = state.data.slice(0, 25).map((r) => r.join(",")).join("\n");
      const system = `You are an expert spreadsheet transformation engine.
Respond ONLY with a valid JSON object — no markdown, no explanation.
{
  "steps": [{ "action": string, "description": string, ...fields }],
  "summary": string
}
Actions: set_cell(row,col,value) | add_row(position,values[]) | delete_row(row) | add_column(header,fill) | delete_column(col) | rename_column(col,newName) | sort(col,direction,hasHeader) | filter_delete(col,operator,value) | replace_all(col,find,replace,transform) | multiply_column(col,factor)
Sheet: ${state.data.length} rows x ${state.data[0]?.length || 0} cols.
CSV preview:\n${csvPreview}`;

      const resp = await fetch(`${API}/api/claude`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system,
          messages: [{ role: "user", content: prompt }],
        }),
      });
      const json = await resp.json();
      const raw  = json.content?.map((b) => b.text || "").join("");
      let plan;
      try { plan = JSON.parse(raw.replace(/```json\n?/g,"").replace(/```\n?/g,"").trim()); }
      catch { addLog("Could not parse AI response. Try rephrasing.", "error"); setAiRunning(false); return; }
      addLog(`AI plan ready: ${plan.summary}`, "info");
      setAiPlan(plan);
    } catch (err) { addLog(`AI error: ${err.message}`, "error"); }
    setAiRunning(false);
  };

  /* ── Apply AI step ── */
  const applyStep = (data, step) => {
    const d = data.map((r) => [...r]);
    if (step.action === "set_cell") {
      while (d.length <= step.row) d.push([]);
      while (d[step.row].length <= step.col) d[step.row].push("");
      d[step.row][step.col] = step.value;
    } else if (step.action === "add_row") {
      const pos = step.position === "end" ? d.length : (step.position ?? d.length);
      d.splice(pos, 0, step.values || Array(d[0]?.length || 1).fill(""));
    } else if (step.action === "delete_row") {
      d.splice(step.row, 1);
    } else if (step.action === "add_column") {
      if (!d.length) d.push([]);
      d[0].push(step.header || "New Column");
      for (let i = 1; i < d.length; i++) d[i].push(step.fill ?? "");
    } else if (step.action === "delete_column") {
      d.forEach((r) => r.splice(step.col, 1));
    } else if (step.action === "rename_column") {
      if (d[0]) d[0][step.col] = step.newName;
    } else if (step.action === "sort") {
      const hdr  = step.hasHeader !== false ? [d[0]] : [];
      const rows = (step.hasHeader !== false ? d.slice(1) : [...d]).sort((a, b) => {
        const [av, bv] = [a[step.col]??"", b[step.col]??""];
        const cmp = typeof av === "number" && typeof bv === "number" ? av-bv : String(av).localeCompare(String(bv));
        return step.direction === "desc" ? -cmp : cmp;
      });
      return normalize([...hdr, ...rows]);
    } else if (step.action === "filter_delete") {
      return normalize(d.filter((row, i) => {
        if (i === 0 && step.hasHeader !== false) return true;
        const sv = String(row[step.col] ?? "");
        if (step.operator === "equals")    return sv !== String(step.value ?? "");
        if (step.operator === "empty")     return sv.trim() !== "";
        if (step.operator === "not_empty") return sv.trim() === "";
        if (step.operator === "contains")  return !sv.includes(step.value ?? "");
        return true;
      }));
    } else if (step.action === "replace_all") {
      d.forEach((row) => {
        (step.col === -1 ? row.map((_,i)=>i) : [step.col]).forEach((ci) => {
          if (ci >= row.length) return;
          let v = String(row[ci] ?? "");
          if (step.find !== undefined) v = v.replaceAll(String(step.find), String(step.replace ?? ""));
          if (step.transform === "uppercase") v = v.toUpperCase();
          if (step.transform === "lowercase") v = v.toLowerCase();
          row[ci] = v;
        });
      });
    } else if (step.action === "multiply_column") {
      for (let i = 1; i < d.length; i++) {
        const v = parseFloat(d[i][step.col]);
        if (!isNaN(v)) d[i][step.col] = parseFloat((v * step.factor).toFixed(6));
      }
    }
    return normalize(d);
  };

  const confirmPlan = () => {
    let d = state.data;
    for (const step of aiPlan.steps || []) {
      try { d = applyStep(d, step); } catch (err) { addLog(`Warning: ${err.message}`, "error"); }
    }
    dispatch({ type: "BULK_SET", data: d });
    addLog(`Applied ${aiPlan.steps?.length || 0} AI operation(s). Don't forget to save!`, "success");
    announce(`Changes applied. ${aiPlan.summary}`);
    setAiPlan(null);
    setPrompt("");
  };

  /* ─── Render ─────────────────────────────────────────────────────────── */
  const { loaded, sheetNames, activeSheet, data, fileName, lastModified,
          undoStack, redoStack, dirty } = state;
  const numRows = data.length;
  const numCols = data[0]?.length || 0;

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text,
                  fontFamily: "'Plus Jakarta Sans',sans-serif", display: "flex", flexDirection: "column" }}>
      <style>{CSS}</style>
      <a href="#main-content" className="skip-link">Skip to main content</a>
      <div role="status" aria-live="polite" aria-atomic="true"
        style={{ position:"absolute", width:1, height:1, overflow:"hidden", clip:"rect(0,0,0,0)", whiteSpace:"nowrap" }}>
        {liveMsg}
      </div>

      {/* ── HEADER ── */}
      <header style={{ borderBottom:`1px solid ${C.border}`, padding:"0 24px",
                       background:C.surface, display:"flex", alignItems:"center",
                       justifyContent:"space-between", gap:16, minHeight:56, flexWrap:"wrap" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <span aria-hidden="true" style={{ fontSize:20 }}>⚡</span>
          <span style={{ fontWeight:700, fontSize:16 }}>XL-Forge</span>
          {fileName && (
            <span style={{ fontSize:12, color:C.text3, borderLeft:`1px solid ${C.border2}`, paddingLeft:12, marginLeft:4 }}>
              {fileName}
            </span>
          )}
        </div>

        {loaded && (
          <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
            {/* Save status */}
            {saveStatus === "saved" && (
              <span className="saved-badge" style={{ background:"#14301A", color:C.green }}>✓ Saved</span>
            )}
            {saveStatus === "error" && (
              <span className="saved-badge" style={{ background:"#301414", color:C.red }}>✕ Save failed</span>
            )}
            {dirty && !saveStatus && (
              <span className="saved-badge" style={{ background:`${C.yellow}18`, color:C.yellow }}>● Unsaved changes</span>
            )}
            <button className="tbtn" onClick={() => { dispatch({ type:"UNDO" }); }} disabled={!undoStack.length} aria-label="Undo (Ctrl+Z)">↩ Undo</button>
            <button className="tbtn" onClick={() => { dispatch({ type:"REDO" }); }} disabled={!redoStack.length} aria-label="Redo (Ctrl+Y)">↪ Redo</button>
            <button
              className="tbtn primary"
              onClick={saveToServer}
              disabled={saving || !dirty}
              aria-label={saving ? "Saving…" : "Save changes to server (Ctrl+S)"}
              aria-busy={saving}
            >
              {saving ? <><span className="spinner" aria-hidden="true" /> Saving…</> : "💾 Save"}
            </button>
          </div>
        )}
      </header>

      {/* ── MAIN ── */}
      <main id="main-content" style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>

        {/* Loading state */}
        {!loaded && !loadError && (
          <div className="loading-screen">
            <span className="spinner" style={{ width:32, height:32, borderWidth:3, color:C.accent }} aria-hidden="true" />
            <p style={{ color:C.text2, fontSize:14 }}>Loading spreadsheet from server…</p>
          </div>
        )}

        {/* Error state */}
        {loadError && (
          <div className="loading-screen">
            <div style={{ textAlign:"center", maxWidth:440 }}>
              <div style={{ fontSize:36, marginBottom:16 }} aria-hidden="true">⚠️</div>
              <h2 style={{ fontSize:18, fontWeight:700, marginBottom:8, color:C.text }}>Could not load file</h2>
              <p style={{ fontSize:13, color:C.text2, marginBottom:20 }}>{loadError}</p>
              <p style={{ fontSize:12, color:C.text3 }}>Make sure the backend server is running:<br />
                <code style={{ color:C.accentHi }}>cd server && npm start</code>
              </p>
              <button className="tbtn primary" style={{ marginTop:20 }}
                onClick={() => { setLoadError(null); window.location.reload(); }}>
                ↺ Retry
              </button>
            </div>
          </div>
        )}

        {/* Editor */}
        {loaded && (
          <div style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>

            {/* Toolbar */}
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
                          padding:"7px 20px", background:C.surface2, borderBottom:`1px solid ${C.border}`,
                          gap:10, flexWrap:"wrap" }}>
              <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                <button className="tbtn" onClick={() => { dispatch({ type:"ADD_ROW" }); addLog("Added row","info"); }} aria-label="Add empty row at bottom">+ Row</button>
                <button className="tbtn" onClick={() => { dispatch({ type:"ADD_COL" }); addLog("Added column","info"); }} aria-label="Add empty column on right">+ Column</button>
                {selCell && <>
                  <button className="tbtn danger" onClick={() => { dispatch({ type:"DEL_ROW", row:selCell.row }); addLog(`Deleted row ${selCell.row+1}`,"info"); setSelCell(null); }}
                    aria-label={`Delete row ${selCell.row+1}`}>✕ Row {selCell.row+1}</button>
                  <button className="tbtn danger" onClick={() => { dispatch({ type:"DEL_COL", col:selCell.col }); addLog(`Deleted col ${colLabel(selCell.col)}`,"info"); setSelCell(null); }}
                    aria-label={`Delete column ${colLabel(selCell.col)}`}>✕ Col {colLabel(selCell.col)}</button>
                </>}
              </div>
              {lastModified && (
                <span style={{ fontSize:11, color:C.text3 }}>
                  Last saved: {new Date(lastModified).toLocaleString()}
                </span>
              )}
            </div>

            {/* Sheet tabs */}
            <nav className="sheet-tabs" aria-label="Sheets" role="tablist">
              {sheetNames.map((n) => (
                <button key={n} role="tab" aria-selected={n === activeSheet}
                  className={`sheet-tab${n === activeSheet ? " active" : ""}`}
                  onClick={() => dispatch({ type:"SWITCH_SHEET", sheet:n })}
                  aria-label={`Switch to sheet ${n}`}>{n}</button>
              ))}
            </nav>

            {/* Grid */}
            <div className="grid-wrap" role="grid"
              aria-label={`Spreadsheet: ${numRows} rows, ${numCols} columns`}
              aria-rowcount={numRows} aria-colcount={numCols}>
              <table className="grid-table" aria-label={activeSheet}>
                <thead>
                  <tr role="row">
                    <th className="row-num-head" scope="col" aria-label="Row numbers"><div className="th-inner">#</div></th>
                    {(data[0]||[]).map((_,ci) => (
                      <th key={ci} scope="col" aria-label={`Column ${colLabel(ci)}`}>
                        <div className="th-inner">{colLabel(ci)}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.map((row, ri) => (
                    <tr key={ri} role="row" aria-rowindex={ri+2}>
                      <td className="row-num" role="rowheader" aria-label={`Row ${ri+1}`}>{ri+1}</td>
                      {row.map((cell, ci) => (
                        <Cell key={`${ri}-${ci}`} value={cell} rowIdx={ri} colIdx={ci}
                          selected={selCell?.row===ri && selCell?.col===ci}
                          onSelect={(r,c) => setSelCell({row:r,col:c})}
                          onChange={(r,c,v) => dispatch({type:"SET_CELL",row:r,col:c,value:v})}
                        />
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Status bar */}
            <div className="status-bar" aria-live="polite">
              <span>Rows: <strong>{numRows}</strong></span>
              <span>Cols: <strong>{numCols}</strong></span>
              {selCell && <span>Cell: <strong>{colLabel(selCell.col)}{selCell.row+1}</strong></span>}
              <span>Undo: <strong>{undoStack.length}</strong></span>
              <span style={{ marginLeft:"auto", color: dirty ? C.yellow : C.green, fontWeight:600 }}>
                {dirty ? "● Unsaved" : "✓ Up to date"}
              </span>
            </div>

            {/* AI Panel */}
            <section className="ai-panel" aria-labelledby="ai-heading">
              <div className="ai-header">
                <h2 id="ai-heading" style={{ fontSize:13, fontWeight:700, color:C.text, display:"flex", alignItems:"center", gap:8 }}>
                  <span aria-hidden="true">🤖</span> AI Assistant
                </h2>
                <p style={{ fontSize:11, color:C.text3 }}>Describe a change — review before applying</p>
              </div>
              <div className="ai-body">
                <label htmlFor="ai-prompt" style={{ fontSize:11, fontWeight:600, letterSpacing:".06em", textTransform:"uppercase", color:C.text2 }}>
                  Your instruction
                </label>
                <textarea id="ai-prompt" ref={promptRef} className="ai-textarea" rows={2}
                  placeholder="e.g. Sort rows by Store Name alphabetically…"
                  value={prompt} onChange={(e) => setPrompt(e.target.value)}
                  onKeyDown={(e) => { if ((e.ctrlKey||e.metaKey) && e.key==="Enter") { e.preventDefault(); runPrompt(); } }}
                  aria-describedby="ai-hint"
                />
                <div className="ai-footer">
                  <div style={{ display:"flex", gap:6, flexWrap:"wrap" }} role="group" aria-label="Example prompts">
                    {EXAMPLES.slice(0,3).map((ex) => (
                      <button key={ex} className="example-chip"
                        onClick={() => { setPrompt(ex); promptRef.current?.focus(); }}
                        aria-label={`Use: ${ex}`}>
                        {ex.length>44 ? ex.slice(0,44)+"…" : ex}
                      </button>
                    ))}
                  </div>
                  <button className="tbtn primary" onClick={runPrompt}
                    disabled={!prompt.trim()||aiRunning}
                    aria-disabled={!prompt.trim()||aiRunning} aria-busy={aiRunning}
                    aria-label={aiRunning?"Processing…":"Ask AI (Ctrl+Enter)"}>
                    {aiRunning ? <><span className="spinner" aria-hidden="true"/> Processing…</> : "✦ Ask AI"}
                  </button>
                </div>
                <p id="ai-hint" style={{ position:"absolute", width:1, height:1, overflow:"hidden", clip:"rect(0,0,0,0)", whiteSpace:"nowrap" }}>
                  Describe changes in plain English. Press Ctrl+Enter or click Ask AI. You will review changes before they are applied.
                </p>
              </div>
            </section>

            {/* Log */}
            {logs.length > 0 && (
              <div className="log-panel" role="log" aria-label="Operation log" aria-live="off" tabIndex={0}>
                {logs.map((l, i) => (
                  <div key={i} className={`log-line ${l.type}`}>
                    <span style={{ color:C.text3, marginRight:8 }}>
                      {new Date(l.ts).toLocaleTimeString([],{hour:"2-digit",minute:"2-digit",second:"2-digit"})}
                    </span>
                    {l.msg}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </main>

      {/* Confirm dialog */}
      {aiPlan && (
        <ConfirmDialog plan={aiPlan} onConfirm={confirmPlan}
          onCancel={() => { setAiPlan(null); addLog("AI changes discarded.","info"); }} />
      )}
    </div>
  );
}
