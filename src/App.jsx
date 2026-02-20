import { useState, useRef, useCallback, useEffect, useReducer } from "react";
import * as XLSX from "xlsx";

/* ───────────────────────────────────────────────────────────────────────────
   DESIGN TOKENS  — all pairs meet WCAG 2.1 AA
─────────────────────────────────────────────────────────────────────────── */
const C = {
  bg:       "#0B0D14",
  surface:  "#131620",
  surface2: "#1A1E2E",
  border:   "#252A3E",
  border2:  "#303654",
  accent:   "#3B82F6",   // blue — 4.6:1 on surface ✓
  accentHi: "#60A5FA",
  green:    "#22C55E",   // success
  red:      "#F87171",   // error
  yellow:   "#FBBF24",   // warning
  text:     "#F1F3FA",   // 17:1 on bg ✓
  text2:    "#8892B0",   // 5.4:1 on bg ✓
  text3:    "#4A527A",
  cellBg:   "#0F1119",
  cellSel:  "#1E2D4A",
  headBg:   "#0D1020",
  headText: "#7B88AA",  // 4.6:1 on headBg ✓
};

/* ───────────────────────────────────────────────────────────────────────────
   GLOBAL CSS
─────────────────────────────────────────────────────────────────────────── */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --accent: ${C.accent};
  --accent-hi: ${C.accentHi};
  --green: ${C.green};
  --red: ${C.red};
}

body { background: ${C.bg}; color: ${C.text}; font-family: 'Plus Jakarta Sans', sans-serif; }

:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
  border-radius: 4px;
}
:focus:not(:focus-visible) { outline: none; }

/* Scrollbars */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: ${C.border2}; border-radius: 99px; }

/* Skip link */
.skip-link {
  position:fixed; top:-100%; left:16px; z-index:9999;
  padding:8px 16px; background:var(--accent); color:#fff;
  font-weight:700; border-radius:0 0 8px 8px; text-decoration:none; font-size:13px;
  font-family:'Plus Jakarta Sans',sans-serif; transition:top .15s;
}
.skip-link:focus { top:0; }

/* Dropzone */
.dropzone {
  border: 2px dashed ${C.border2};
  border-radius: 16px;
  padding: 64px 40px;
  text-align: center;
  cursor: pointer;
  background: ${C.surface};
  transition: border-color .2s, background .2s;
  position: relative;
}
.dropzone:hover, .dropzone.drag-over, .dropzone:focus-visible {
  border-color: var(--accent);
  background: ${C.accent}0D;
}

/* ── Toolbar buttons ── */
.tbtn {
  display:inline-flex; align-items:center; gap:6px;
  padding:7px 14px; border-radius:8px; border:1px solid ${C.border2};
  background:${C.surface2}; color:${C.text2}; font-size:12px; font-weight:600;
  font-family:inherit; cursor:pointer; white-space:nowrap;
  transition: background .15s, color .15s, border-color .15s;
}
.tbtn:hover, .tbtn:focus-visible {
  background:${C.border2}; color:${C.text}; border-color:${C.accent}60;
}
.tbtn.danger:hover { background:#3A1010; color:var(--red); border-color:var(--red)60; }
.tbtn.primary {
  background:var(--accent); color:#fff; border-color:transparent;
}
.tbtn.primary:hover, .tbtn.primary:focus-visible {
  background:var(--accent-hi); color:#fff;
}
.tbtn:disabled { opacity:.4; cursor:not-allowed; pointer-events:none; }

/* ── Grid ── */
.grid-wrap {
  overflow: auto;
  flex: 1;
  min-height: 0;
  background: ${C.cellBg};
  border-radius: 0 0 12px 12px;
}
.grid-table {
  border-collapse: collapse;
  font-family: 'JetBrains Mono', monospace;
  font-size: 13px;
  table-layout: fixed;
  min-width: 100%;
}
.grid-table th {
  position: sticky; top: 0; z-index: 2;
  background: ${C.headBg};
  color: ${C.headText};
  font-size: 11px; font-weight: 600; font-family: 'Plus Jakarta Sans', sans-serif;
  padding: 0; text-align: center;
  border-right: 1px solid ${C.border};
  border-bottom: 2px solid ${C.border2};
  user-select: none;
  letter-spacing: .05em;
}
.grid-table th.row-num-head {
  width: 48px; min-width: 48px; max-width: 48px;
  left: 0; z-index: 3;
}
.grid-table th .th-inner {
  padding: 8px 10px;
  display:flex; align-items:center; justify-content:center; gap:4px;
  white-space:nowrap;
}
.grid-table td {
  border-right: 1px solid ${C.border};
  border-bottom: 1px solid ${C.border};
  padding: 0;
  color: ${C.text};
  vertical-align: middle;
  min-width: 120px;
  max-width: 240px;
}
.grid-table td.row-num {
  position: sticky; left: 0; z-index: 1;
  background: ${C.headBg};
  color: ${C.headText};
  font-size: 11px; font-family: 'Plus Jakarta Sans', sans-serif;
  text-align: center; padding: 0 8px;
  min-width: 48px; max-width: 48px; width: 48px;
  border-right: 2px solid ${C.border2};
  user-select: none;
}
.grid-table tr:hover td { background: ${C.surface}; }
.grid-table tr:hover td.row-num { background: ${C.headBg}; }

/* Cell display */
.cell-display {
  padding: 7px 10px;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  width: 100%; height: 100%; display: block;
  cursor: cell; min-height: 33px; line-height: 1.4;
}
.cell-display.selected {
  background: ${C.cellSel};
  box-shadow: inset 0 0 0 2px var(--accent);
}
.cell-display.editing {
  padding: 0;
}
.cell-input {
  width: 100%; height: 100%;
  padding: 7px 10px;
  background: ${C.cellSel};
  border: none;
  outline: 2px solid var(--accent);
  color: ${C.text};
  font-family: 'JetBrains Mono', monospace;
  font-size: 13px;
  line-height: 1.4;
  min-height: 33px;
  box-shadow: 0 4px 20px rgba(59,130,246,.25);
}
.cell-input:focus { outline: 2px solid var(--accent); }

/* ── Tabs ── */
.sheet-tabs {
  display:flex; gap:2px; padding:0 16px;
  background:${C.surface}; border-bottom:1px solid ${C.border};
  overflow-x:auto;
}
.sheet-tab {
  padding:9px 16px 8px; font-size:12px; font-weight:600;
  cursor:pointer; border:none; background:transparent;
  color:${C.text2}; border-bottom:2px solid transparent;
  white-space:nowrap; font-family:inherit;
  transition: color .15s, border-color .15s;
}
.sheet-tab:hover { color:${C.text}; }
.sheet-tab.active { color:${C.accentHi}; border-bottom-color:var(--accent); }

/* ── AI Panel ── */
.ai-panel {
  background:${C.surface};
  border-top: 1px solid ${C.border};
  display:flex; flex-direction:column; gap:0;
}
.ai-header {
  display:flex; align-items:center; justify-content:space-between;
  padding:14px 20px; border-bottom:1px solid ${C.border};
}
.ai-body { padding:16px 20px; display:flex; flex-direction:column; gap:12px; }
.ai-textarea {
  width:100%; background:${C.surface2}; border:1px solid ${C.border2};
  border-radius:10px; padding:12px 14px; font-size:13px; color:${C.text};
  font-family:'Plus Jakarta Sans',sans-serif; resize:vertical; min-height:80px;
  line-height:1.6; caret-color:var(--accent);
  transition: border-color .2s;
}
.ai-textarea:focus { border-color:var(--accent); outline:none; }
.ai-footer { display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }

/* ── Confirm dialog ── */
.confirm-overlay {
  position:fixed; inset:0; z-index:1000;
  background:rgba(0,0,0,.7); backdrop-filter:blur(4px);
  display:flex; align-items:center; justify-content:center; padding:20px;
}
.confirm-box {
  background:${C.surface}; border:1px solid ${C.border2};
  border-radius:16px; padding:28px 32px; max-width:560px; width:100%;
  box-shadow:0 24px 64px rgba(0,0,0,.6);
}
.confirm-changes {
  background:${C.surface2}; border:1px solid ${C.border};
  border-radius:10px; padding:14px 16px; margin:16px 0;
  font-size:13px; font-family:'JetBrains Mono',monospace;
  color:${C.text2}; max-height:220px; overflow-y:auto;
  line-height:1.7;
}
.change-item { display:flex; gap:10px; align-items:flex-start; }
.change-dot { color:var(--accent); flex-shrink:0; margin-top:2px; }

/* ── Log ── */
.log-panel {
  background:${C.surface2}; border-top:1px solid ${C.border};
  padding:10px 20px; font-size:11px; font-family:'JetBrains Mono',monospace;
  color:${C.text3}; max-height:100px; overflow-y:auto;
}
.log-line { padding:2px 0; line-height:1.6; }
.log-line.success { color:var(--green); }
.log-line.error { color:var(--red); }
.log-line.info { color:${C.accentHi}; }

/* ── Status bar ── */
.status-bar {
  padding:6px 20px; background:${C.surface}; border-top:1px solid ${C.border};
  font-size:11px; color:${C.text3}; display:flex; gap:24px; align-items:center;
  font-family:'JetBrains Mono',monospace;
}
.status-bar span { display:flex; align-items:center; gap:5px; }

/* ── Download bar ── */
.dl-bar {
  background: linear-gradient(135deg, #1A2A1A, #162216);
  border: 1px solid ${C.green}40;
  border-radius:12px; padding:16px 24px;
  display:flex; align-items:center; justify-content:space-between; gap:16px;
  flex-wrap:wrap;
}

/* ── Chips ── */
.example-chip {
  background:${C.surface2}; border:1px solid ${C.border2};
  color:${C.text2}; border-radius:6px; padding:4px 10px;
  font-size:11px; cursor:pointer; font-family:inherit;
  transition:all .15s; white-space:nowrap;
}
.example-chip:hover, .example-chip:focus-visible {
  background:${C.border}; color:${C.text}; border-color:var(--accent)60;
}

/* ── Spinner ── */
@keyframes spin { to { transform: rotate(360deg); } }
.spinner {
  width:14px; height:14px; border:2px solid transparent;
  border-top-color:currentColor; border-radius:50%;
  animation:spin .7s linear infinite; display:inline-block;
}

/* ── Fade in ── */
@keyframes fadeUp {
  from { opacity:0; transform:translateY(8px); }
  to   { opacity:1; transform:translateY(0); }
}
.fade-up { animation:fadeUp .25s ease forwards; }

/* ── Reduced motion ── */
@media (prefers-reduced-motion:reduce) {
  *, *::before, *::after {
    animation-duration:.001ms !important;
    transition-duration:.001ms !important;
  }
}

/* ── Responsive ── */
@media (max-width:640px) {
  .ai-footer { flex-direction:column; align-items:stretch; }
  .dl-bar { flex-direction:column; }
  .status-bar { gap:12px; flex-wrap:wrap; }
}
`;

/* ───────────────────────────────────────────────────────────────────────────
   HELPERS
─────────────────────────────────────────────────────────────────────────── */
const colLabel = (i) => {
  let s = ""; i++;
  while (i > 0) { s = String.fromCharCode(64 + (i % 26 || 26)) + s; i = Math.floor((i - 1) / 26); }
  return s;
};
const fmt = (b) => b < 1024 ? b + " B" : b < 1048576 ? (b / 1024).toFixed(1) + " KB" : (b / 1048576).toFixed(1) + " MB";

const toBlob = (wb) => {
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
};

const sheetToData = (sheet) =>
  XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

const dataToSheet = (data) => XLSX.utils.aoa_to_sheet(data);

const cloneWb = (wb) => {
  const nw = XLSX.utils.book_new();
  for (const n of wb.SheetNames) {
    const d = sheetToData(wb.Sheets[n]).map((r) => [...r]);
    nw.SheetNames.push(n);
    nw.Sheets[n] = dataToSheet(d);
  }
  return nw;
};

/* Pad all rows to same width */
const normalize = (data) => {
  const w = Math.max(0, ...data.map((r) => r.length));
  return data.map((r) => { const row = [...r]; while (row.length < w) row.push(""); return row; });
};

const EXAMPLES = [
  "Add a Total row at the bottom summing all numeric columns",
  "Sort rows by the first column alphabetically",
  "Add a new column called 'Status' with value 'Active' for all rows",
  "Rename the first column header to 'ID'",
  "Remove rows where any cell is empty",
  "Multiply all values in column C by 1.15",
];

/* ───────────────────────────────────────────────────────────────────────────
   STATE REDUCER  — single source of truth for grid data
─────────────────────────────────────────────────────────────────────────── */
const initState = {
  workbook: null,
  file: null,
  activeSheet: null,
  data: [],          // 2-D array for active sheet
  undoStack: [],
  redoStack: [],
  dirty: false,
};

function reducer(state, action) {
  switch (action.type) {
    case "LOAD": {
      const { workbook, file } = action;
      const activeSheet = workbook.SheetNames[0];
      const data = normalize(sheetToData(workbook.Sheets[activeSheet]));
      return { ...initState, workbook, file, activeSheet, data };
    }
    case "SWITCH_SHEET": {
      // Save current sheet back to workbook first
      const wb = cloneWb(state.workbook);
      wb.Sheets[state.activeSheet] = dataToSheet(state.data);
      const data = normalize(sheetToData(wb.Sheets[action.sheet]));
      return { ...state, workbook: wb, activeSheet: action.sheet, data, dirty: true };
    }
    case "SET_CELL": {
      const { row, col, value } = action;
      const data = state.data.map((r) => [...r]);
      data[row][col] = value;
      return {
        ...state, data,
        undoStack: [...state.undoStack, state.data],
        redoStack: [], dirty: true,
      };
    }
    case "BULK_SET": {
      // Replace all data for active sheet
      const data = normalize(action.data);
      return {
        ...state, data,
        undoStack: [...state.undoStack, state.data],
        redoStack: [], dirty: true,
      };
    }
    case "ADD_ROW": {
      const empty = Array(state.data[0]?.length || 1).fill("");
      const data = [...state.data, empty];
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
      return { ...state, dirty: false };
    default:
      return state;
  }
}

/* ───────────────────────────────────────────────────────────────────────────
   CELL COMPONENT
─────────────────────────────────────────────────────────────────────────── */
function Cell({ value, rowIdx, colIdx, selected, onSelect, onChange }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft]     = useState("");
  const inputRef = useRef(null);

  useEffect(() => { if (editing && inputRef.current) inputRef.current.focus(); }, [editing]);

  const startEdit = () => { setDraft(String(value ?? "")); setEditing(true); };
  const commit    = () => { setEditing(false); onChange(rowIdx, colIdx, draft); };
  const cancel    = () => { setEditing(false); };

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); commit(); }
    if (e.key === "Escape") cancel();
    if (e.key === "Tab") { e.preventDefault(); commit(); }
  };

  const handleCellKey = (e) => {
    if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); startEdit(); }
    if (e.key === "Delete" || e.key === "Backspace") onChange(rowIdx, colIdx, "");
    // Start typing replaces cell
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      setDraft(e.key); setEditing(true);
    }
  };

  return (
    <td
      role="gridcell"
      aria-colindex={colIdx + 2}
      aria-rowindex={rowIdx + 2}
      aria-selected={selected}
    >
      {editing ? (
        <input
          ref={inputRef}
          className="cell-input"
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commit}
          onKeyDown={handleKeyDown}
          aria-label={`Cell ${colLabel(colIdx)}${rowIdx + 1}, editing`}
        />
      ) : (
        <div
          className={`cell-display${selected ? " selected" : ""}`}
          tabIndex={0}
          role="button"
          aria-label={`Cell ${colLabel(colIdx)}${rowIdx + 1}: ${String(value ?? "") || "empty"}`}
          onClick={() => { onSelect(rowIdx, colIdx); }}
          onDoubleClick={startEdit}
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

/* ───────────────────────────────────────────────────────────────────────────
   CONFIRM DIALOG
─────────────────────────────────────────────────────────────────────────── */
function ConfirmDialog({ plan, onConfirm, onCancel }) {
  const ref = useRef(null);
  useEffect(() => { ref.current?.focus(); }, []);

  return (
    <div
      className="confirm-overlay"
      role="dialog"
      aria-modal="true"
      aria-labelledby="confirm-title"
      onClick={(e) => e.target === e.currentTarget && onCancel()}
    >
      <div className="confirm-box fade-up" ref={ref} tabIndex={-1}>
        <h2 id="confirm-title" style={{ fontSize: 18, fontWeight: 700, color: C.text, marginBottom: 8 }}>
          Review AI Changes
        </h2>
        <p style={{ fontSize: 13, color: C.text2 }}>
          The following changes will be applied to your spreadsheet:
        </p>
        <div className="confirm-changes" role="list" aria-label="Proposed changes">
          {plan.steps?.map((s, i) => (
            <div key={i} className="change-item" role="listitem">
              <span className="change-dot" aria-hidden="true">▸</span>
              <span>{s.description}</span>
            </div>
          ))}
        </div>
        <p style={{ fontSize: 12, color: C.text3, marginBottom: 20 }}>
          Summary: {plan.summary}
        </p>
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button className="tbtn" onClick={onCancel} aria-label="Cancel and discard AI suggestions">
            Cancel
          </button>
          <button className="tbtn primary" onClick={onConfirm} aria-label="Apply all suggested changes">
            ✓ Apply Changes
          </button>
        </div>
      </div>
    </div>
  );
}

/* ───────────────────────────────────────────────────────────────────────────
   MAIN APP
─────────────────────────────────────────────────────────────────────────── */
export default function App() {
  const [state,    dispatch]   = useReducer(reducer, initState);
  const [selCell,  setSelCell] = useState(null);          // {row, col}
  const [prompt,   setPrompt]  = useState("");
  const [aiPlan,   setAiPlan]  = useState(null);          // pending confirm
  const [aiRunning,setAiRunning]= useState(false);
  const [logs,     setLogs]    = useState([]);
  const [dragOver, setDragOver]= useState(false);
  const [liveMsg,  setLiveMsg] = useState("");

  const fileInputRef = useRef(null);
  const promptRef    = useRef(null);
  const dropRef      = useRef(null);

  const announce = useCallback((m) => setLiveMsg(m), []);
  const addLog   = useCallback((msg, type = "default") => {
    setLogs((p) => [...p.slice(-49), { msg, type, ts: Date.now() }]);
    if (type !== "default") announce(msg);
  }, [announce]);

  /* ── Keyboard shortcuts (Ctrl+Z / Ctrl+Y) ── */
  useEffect(() => {
    const handler = (e) => {
      if (!state.workbook) return;
      if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) {
        e.preventDefault(); dispatch({ type: "UNDO" });
        announce("Undo");
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === "y" || (e.key === "z" && e.shiftKey))) {
        e.preventDefault(); dispatch({ type: "REDO" });
        announce("Redo");
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [state.workbook, announce]);

  /* ── Load file ── */
  const loadFile = (f) => {
    if (!f) return;
    const ext = f.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv"].includes(ext)) {
      addLog(`Unsupported file type: .${ext}`, "error"); return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        dispatch({ type: "LOAD", workbook: wb, file: f });
        addLog(`Loaded: ${f.name}`, "success");
        announce(`File loaded: ${f.name}`);
        setTimeout(() => promptRef.current?.focus(), 150);
      } catch (err) {
        addLog(`Failed to parse: ${err.message}`, "error");
      }
    };
    reader.readAsArrayBuffer(f);
  };

  /* ── Build current workbook from state ── */
  const buildWb = useCallback(() => {
    const wb = cloneWb(state.workbook);
    wb.Sheets[state.activeSheet] = dataToSheet(state.data);
    return wb;
  }, [state.workbook, state.activeSheet, state.data]);

  /* ── Download ── */
  const download = () => {
    const wb   = buildWb();
    const blob = toBlob(wb);
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    const parts = state.file.name.split("."); parts.splice(-1, 0, "edited");
    a.download = parts.join("."); a.href = url; a.click();
    URL.revokeObjectURL(url);
    dispatch({ type: "MARK_SAVED" });
    addLog(`Downloaded: ${a.download}`, "success");
    announce(`Downloading ${a.download}`);
  };

  /* ── Cell change ── */
  const handleCellChange = (row, col, val) => {
    dispatch({ type: "SET_CELL", row, col, value: val });
  };

  /* ── Apply AI step ── */
  const applyStep = (data, step) => {
    const d = data.map((r) => [...r]);
    const { action } = step;

    if (action === "set_cell") {
      while (d.length <= step.row) d.push([]);
      while (d[step.row].length <= step.col) d[step.row].push("");
      d[step.row][step.col] = step.value;
    } else if (action === "add_row") {
      const pos = step.position === "end" ? d.length : (step.position ?? d.length);
      d.splice(pos, 0, step.values || Array(d[0]?.length || 1).fill(""));
    } else if (action === "delete_row") {
      d.splice(step.row, 1);
    } else if (action === "add_column") {
      if (!d.length) d.push([]);
      d[0].push(step.header || "New Column");
      for (let i = 1; i < d.length; i++)
        d[i].push(step.fill !== undefined ? step.fill : (step.values?.[i - 1] ?? ""));
    } else if (action === "delete_column") {
      d.forEach((r) => r.splice(step.col, 1));
    } else if (action === "rename_column") {
      if (d[0]) d[0][step.col] = step.newName;
    } else if (action === "sort") {
      const hdr = step.hasHeader !== false ? [d[0]] : [];
      const rows = (step.hasHeader !== false ? d.slice(1) : [...d]).sort((a, b) => {
        const [av, bv] = [a[step.col] ?? "", b[step.col] ?? ""];
        const cmp = typeof av === "number" && typeof bv === "number"
          ? av - bv : String(av).localeCompare(String(bv));
        return step.direction === "desc" ? -cmp : cmp;
      });
      return normalize([...hdr, ...rows]);
    } else if (action === "filter_delete") {
      return normalize(d.filter((row, i) => {
        if (i === 0 && step.hasHeader !== false) return true;
        const sv = String(row[step.col] ?? "");
        if (step.operator === "equals")    return sv !== String(step.value ?? "");
        if (step.operator === "empty")     return sv.trim() !== "";
        if (step.operator === "not_empty") return sv.trim() === "";
        if (step.operator === "contains")  return !sv.includes(step.value ?? "");
        return true;
      }));
    } else if (action === "replace_all") {
      d.forEach((row) => {
        (step.col === -1 ? row.map((_, i) => i) : [step.col]).forEach((ci) => {
          if (ci >= row.length) return;
          let v = String(row[ci] ?? "");
          if (step.find !== undefined) v = v.replaceAll(String(step.find), String(step.replace ?? ""));
          if (step.transform === "uppercase") v = v.toUpperCase();
          if (step.transform === "lowercase") v = v.toLowerCase();
          row[ci] = v;
        });
      });
    } else if (action === "multiply_column") {
      for (let i = 1; i < d.length; i++) {
        const v = parseFloat(d[i][step.col]);
        if (!isNaN(v)) d[i][step.col] = parseFloat((v * step.factor).toFixed(6));
      }
    }
    return normalize(d);
  };

  /* ── Run AI prompt ── */
  const runPrompt = async () => {
    if (!state.workbook || !prompt.trim() || aiRunning) return;
    setAiRunning(true);
    announce("Running AI transformation, please wait.");

    try {
      const csvPreview = state.data.slice(0, 25)
        .map((r) => r.join(",")).join("\n");

      const system = `You are an expert spreadsheet transformation engine.
Respond ONLY with a valid JSON object — no markdown, no explanation.
{
  "steps": [
    {
      "action": string,        // set_cell | add_row | delete_row | add_column | delete_column | rename_column | sort | filter_delete | replace_all | multiply_column
      "description": string,   // plain English
      // action fields:
      // set_cell:        row(0-idx), col(0-idx), value
      // add_row:         position("end"|number), values[]
      // delete_row:      row
      // add_column:      header, fill, values[]
      // delete_column:   col
      // rename_column:   col, newName
      // sort:            col, direction("asc"|"desc"), hasHeader
      // filter_delete:   col, operator("equals"|"empty"|"not_empty"|"contains"), value, hasHeader
      // replace_all:     col(-1=all), find, replace, transform(null|"uppercase"|"lowercase")
      // multiply_column: col, factor
    }
  ],
  "summary": string
}
Sheet has ${state.data.length} rows x ${state.data[0]?.length || 0} cols.
CSV (first 25 rows):
${csvPreview}`;

      const resp = await fetch("/api/claude", {
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
      try {
        plan = JSON.parse(raw.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim());
      } catch {
        addLog("Could not parse AI response. Try rephrasing.", "error"); setAiRunning(false); return;
      }

      addLog(`AI plan ready: ${plan.summary}`, "info");
      setAiPlan(plan);   // open confirm dialog
    } catch (err) {
      addLog(`AI error: ${err.message}`, "error");
    }
    setAiRunning(false);
  };

  /* ── Confirm AI plan ── */
  const confirmPlan = () => {
    let d = state.data;
    for (const step of aiPlan.steps || []) {
      try { d = applyStep(d, step); }
      catch (err) { addLog(`Warning: ${err.message}`, "error"); }
    }
    dispatch({ type: "BULK_SET", data: d });
    addLog(`Applied ${aiPlan.steps?.length || 0} AI operation(s).`, "success");
    announce(`Changes applied: ${aiPlan.summary}`);
    setAiPlan(null);
    setPrompt("");
  };

  const { workbook, file, activeSheet, data, undoStack, redoStack, dirty } = state;
  const numRows = data.length;
  const numCols = data[0]?.length || 0;

  /* ───────────────────────────────────────────────────────────────────────
     RENDER
  ─────────────────────────────────────────────────────────────────────── */
  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Plus Jakarta Sans', sans-serif", color: C.text, display: "flex", flexDirection: "column" }}>
      <style>{CSS}</style>

      {/* Skip link */}
      <a href="#main-content" className="skip-link">Skip to main content</a>

      {/* Live region */}
      <div role="status" aria-live="polite" aria-atomic="true"
        style={{ position: "absolute", width: 1, height: 1, overflow: "hidden", clip: "rect(0,0,0,0)", whiteSpace: "nowrap" }}>
        {liveMsg}
      </div>

      {/* ── HEADER ── */}
      <header style={{ borderBottom: `1px solid ${C.border}`, padding: "0 24px", background: C.surface, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, minHeight: 56, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span aria-hidden="true" style={{ fontSize: 20 }}>⚡</span>
          <span style={{ fontWeight: 700, fontSize: 16, letterSpacing: "-.01em" }}>XL-Forge</span>
          <span style={{ fontSize: 11, color: C.text3, fontWeight: 500, letterSpacing: ".05em", textTransform: "uppercase", marginLeft: 4 }}>
            Spreadsheet Editor
          </span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {file && (
            <>
              <button className="tbtn" onClick={() => dispatch({ type: "UNDO" })} disabled={!undoStack.length} aria-label="Undo last change (Ctrl+Z)">↩ Undo</button>
              <button className="tbtn" onClick={() => dispatch({ type: "REDO" })} disabled={!redoStack.length} aria-label="Redo last undone change (Ctrl+Y)">↪ Redo</button>
              <button className="tbtn" onClick={() => { dispatch({ type: "LOAD", workbook: state.workbook, file }); addLog("Reset to original", "info"); announce("Reset to original file."); }} aria-label="Discard all changes and reload original file">⟳ Reset</button>
              <button className="tbtn danger" onClick={() => { dispatch({ type: "LOAD", workbook: { SheetNames: [], Sheets: {} }, file: null }); setLogs([]); setPrompt(""); announce("File removed."); }} aria-label="Remove file and start over">✕ Remove file</button>
            </>
          )}
        </div>
      </header>

      {/* ── MAIN ── */}
      <main id="main-content" style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {!file ? (
          /* ── UPLOAD SCREEN ── */
          <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 40 }}>
            <div style={{ maxWidth: 560, width: "100%" }}>
              <div
                ref={dropRef}
                className={`dropzone${dragOver ? " drag-over" : ""}`}
                role="button"
                tabIndex={0}
                aria-label="Upload spreadsheet. Click or press Enter to browse, or drag and drop."
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={(e) => { e.preventDefault(); setDragOver(false); loadFile(e.dataTransfer.files[0]); }}
                onClick={() => fileInputRef.current?.click()}
                onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInputRef.current?.click(); } }}
              >
                <div style={{ fontSize: 40, marginBottom: 16 }} aria-hidden="true">📂</div>
                <p style={{ fontSize: 18, fontWeight: 700, color: C.text, marginBottom: 8 }}>Drop your spreadsheet here</p>
                <p style={{ fontSize: 13, color: C.text2 }}>or click to browse — .xlsx, .xls, .csv supported</p>
                <input
                  ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv"
                  aria-label="Choose spreadsheet file"
                  style={{ position: "absolute", opacity: 0, width: 0, height: 0, pointerEvents: "none" }}
                  tabIndex={-1}
                  onChange={(e) => loadFile(e.target.files[0])}
                />
              </div>
              <div style={{ marginTop: 32, display: "flex", flexDirection: "column", gap: 12 }}>
                <p style={{ fontSize: 12, color: C.text3, fontWeight: 600, letterSpacing: ".08em", textTransform: "uppercase" }}>What you can do</p>
                {[
                  ["✏️", "Direct editing", "Click any cell to edit it inline"],
                  ["🤖", "AI prompts", "Describe changes in plain English"],
                  ["↩", "Undo / Redo", "Full history with Ctrl+Z / Ctrl+Y"],
                  ["⬇️", "Download", "Save the edited file at any time"],
                ].map(([icon, title, desc]) => (
                  <div key={title} style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
                    <span aria-hidden="true" style={{ fontSize: 16, marginTop: 1 }}>{icon}</span>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.text }}>{title}</div>
                      <div style={{ fontSize: 12, color: C.text2 }}>{desc}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        ) : (
          /* ── EDITOR SCREEN ── */
          <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

            {/* File bar */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 20px", background: C.surface2, borderBottom: `1px solid ${C.border}`, gap: 12, flexWrap: "wrap" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span aria-hidden="true" style={{ fontSize: 16 }}>📊</span>
                <span style={{ fontWeight: 600, fontSize: 13, color: C.text }}>{file.name}</span>
                <span style={{ fontSize: 11, color: C.text3 }}>{fmt(file.size)}</span>
                {dirty && <span style={{ fontSize: 10, color: C.yellow, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", background: `${C.yellow}18`, padding: "2px 6px", borderRadius: 4 }} aria-label="File has unsaved changes">● Unsaved</span>}
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button className="tbtn" onClick={() => { dispatch({ type: "ADD_ROW" }); addLog("Added row", "info"); }} aria-label="Add a new empty row at the bottom">+ Row</button>
                <button className="tbtn" onClick={() => { dispatch({ type: "ADD_COL" }); addLog("Added column", "info"); }} aria-label="Add a new empty column on the right">+ Column</button>
                {selCell && (
                  <>
                    <button className="tbtn danger" onClick={() => { dispatch({ type: "DEL_ROW", row: selCell.row }); addLog(`Deleted row ${selCell.row + 1}`, "info"); setSelCell(null); }} aria-label={`Delete row ${selCell.row + 1}`}>✕ Row {selCell.row + 1}</button>
                    <button className="tbtn danger" onClick={() => { dispatch({ type: "DEL_COL", col: selCell.col }); addLog(`Deleted column ${colLabel(selCell.col)}`, "info"); setSelCell(null); }} aria-label={`Delete column ${colLabel(selCell.col)}`}>✕ Col {colLabel(selCell.col)}</button>
                  </>
                )}
                <button className="tbtn primary" onClick={download} aria-label="Download edited spreadsheet file">⬇ Download</button>
              </div>
            </div>

            {/* Sheet tabs */}
            <nav className="sheet-tabs" aria-label="Spreadsheet sheets" role="tablist">
              {workbook.SheetNames.map((n) => (
                <button
                  key={n}
                  role="tab"
                  aria-selected={n === activeSheet}
                  className={`sheet-tab${n === activeSheet ? " active" : ""}`}
                  onClick={() => dispatch({ type: "SWITCH_SHEET", sheet: n })}
                  aria-label={`Switch to sheet: ${n}`}
                >
                  {n}
                </button>
              ))}
            </nav>

            {/* Grid */}
            <div
              className="grid-wrap"
              role="grid"
              aria-label={`Spreadsheet grid: ${numRows} rows, ${numCols} columns`}
              aria-rowcount={numRows}
              aria-colcount={numCols}
            >
              <table className="grid-table" aria-label={activeSheet}>
                <thead>
                  <tr role="row">
                    <th className="row-num-head" scope="col" aria-label="Row numbers"><div className="th-inner">#</div></th>
                    {(data[0] || []).map((_, ci) => (
                      <th key={ci} scope="col" aria-label={`Column ${colLabel(ci)}`}>
                        <div className="th-inner">{colLabel(ci)}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.map((row, ri) => (
                    <tr key={ri} role="row" aria-rowindex={ri + 2}>
                      <td className="row-num" role="rowheader" aria-label={`Row ${ri + 1}`}>{ri + 1}</td>
                      {row.map((cell, ci) => (
                        <Cell
                          key={`${ri}-${ci}`}
                          value={cell}
                          rowIdx={ri}
                          colIdx={ci}
                          selected={selCell?.row === ri && selCell?.col === ci}
                          onSelect={(r, c) => setSelCell({ row: r, col: c })}
                          onChange={handleCellChange}
                        />
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Status bar */}
            <div className="status-bar" aria-live="polite" aria-label="Spreadsheet status">
              <span aria-label={`${numRows} rows`}>Rows: <strong>{numRows}</strong></span>
              <span aria-label={`${numCols} columns`}>Cols: <strong>{numCols}</strong></span>
              {selCell && <span aria-label={`Selected cell ${colLabel(selCell.col)}${selCell.row + 1}`}>Cell: <strong>{colLabel(selCell.col)}{selCell.row + 1}</strong></span>}
              <span aria-label={`${undoStack.length} undo steps available`}>Undo: <strong>{undoStack.length}</strong></span>
            </div>

            {/* AI Panel */}
            <section className="ai-panel" aria-labelledby="ai-panel-heading">
              <div className="ai-header">
                <h2 id="ai-panel-heading" style={{ fontSize: 13, fontWeight: 700, color: C.text, display: "flex", alignItems: "center", gap: 8 }}>
                  <span aria-hidden="true">🤖</span> AI Assistant
                </h2>
                <p style={{ fontSize: 11, color: C.text3 }}>Describe a change in plain English — review before applying</p>
              </div>
              <div className="ai-body">
                <label htmlFor="ai-prompt" style={{ fontSize: 11, fontWeight: 600, letterSpacing: ".06em", textTransform: "uppercase", color: C.text2 }}>
                  Your instruction
                </label>
                <textarea
                  id="ai-prompt"
                  ref={promptRef}
                  className="ai-textarea"
                  rows={2}
                  placeholder="e.g. Sort rows by the Name column alphabetically…"
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  onKeyDown={(e) => { if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); runPrompt(); } }}
                  aria-describedby="ai-hint"
                />
                <div className="ai-footer">
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }} role="group" aria-label="Example prompts">
                    {EXAMPLES.slice(0, 3).map((ex) => (
                      <button key={ex} className="example-chip" onClick={() => { setPrompt(ex); promptRef.current?.focus(); }} aria-label={`Use example: ${ex}`}>
                        {ex.length > 45 ? ex.slice(0, 45) + "…" : ex}
                      </button>
                    ))}
                  </div>
                  <button
                    className="tbtn primary"
                    style={{ gap: 8 }}
                    disabled={!prompt.trim() || aiRunning}
                    aria-disabled={!prompt.trim() || aiRunning}
                    aria-busy={aiRunning}
                    aria-label={aiRunning ? "AI is processing, please wait" : "Send prompt to AI (Ctrl+Enter)"}
                    onClick={runPrompt}
                  >
                    {aiRunning ? <><span className="spinner" aria-hidden="true" /> Processing…</> : "✦ Ask AI"}
                  </button>
                </div>
                <p id="ai-hint" style={{ position: "absolute", width: 1, height: 1, overflow: "hidden", clip: "rect(0,0,0,0)", whiteSpace: "nowrap" }}>
                  Describe changes in plain English. Press Ctrl+Enter or click Ask AI. You will review changes before they are applied.
                </p>
              </div>
            </section>

            {/* Log */}
            {logs.length > 0 && (
              <div className="log-panel" role="log" aria-label="Operation log" aria-live="off" tabIndex={0}>
                {logs.map((l, i) => (
                  <div key={i} className={`log-line ${l.type}`}>
                    <span style={{ color: C.text3, marginRight: 8 }}>
                      {new Date(l.ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
                    </span>
                    {l.msg}
                  </div>
                ))}
              </div>
            )}

          </div>
        )}
      </main>

      {/* ── CONFIRM DIALOG ── */}
      {aiPlan && (
        <ConfirmDialog
          plan={aiPlan}
          onConfirm={confirmPlan}
          onCancel={() => { setAiPlan(null); addLog("AI changes discarded.", "info"); }}
        />
      )}
    </div>
  );
}
