/**
 * XL-Forge Backend Server
 * 
 * Hosts the Excel file and exposes REST API endpoints:
 *   GET  /api/file        → returns all sheet data as JSON
 *   POST /api/save        → receives updated sheet data, writes back to Excel file
 *   POST /api/claude      → proxies AI requests (keeps API key secret)
 *   GET  /api/health      → health check
 */

const express    = require("express");
const cors       = require("cors");
const multer     = require("multer");
const XLSX       = require("xlsx");
const path       = require("path");
const fs         = require("fs");

const app  = express();
const PORT = process.env.PORT || 3001;

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// ── File path — this is the Excel file stored on the server ──────────────────
// Place your .xlsx file in the /server/data/ folder and set the name here.
const DATA_DIR  = path.join(__dirname, "data");
const FILE_NAME = process.env.EXCEL_FILE || "store-data.xlsx";
const FILE_PATH = path.join(DATA_DIR, FILE_NAME);

// Create data dir if missing
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// Create a sample file if none exists (so the app works out of the box)
if (!fs.existsSync(FILE_PATH)) {
  console.log(`No Excel file found at ${FILE_PATH}. Creating sample file...`);
  const wb = XLSX.utils.book_new();

  const sampleData = [
    ["Store ID", "Store Name",       "Manager",        "City",        "Region",  "Monthly Target ($)", "Actual Sales ($)", "Status"    ],
    ["S001",     "Downtown Central", "Alice Johnson",  "New York",    "East",    150000,                142000,             "Review"    ],
    ["S002",     "West Side Mall",   "Bob Martinez",   "Los Angeles", "West",    120000,                135000,             "Verified"  ],
    ["S003",     "Northgate Plaza",  "Carol White",    "Chicago",     "Midwest", 100000,                98000,              "Review"    ],
    ["S004",     "Eastfield Centre", "David Lee",      "Houston",     "South",   110000,                115000,             "Verified"  ],
    ["S005",     "Riverside Market", "Eva Patel",      "Phoenix",     "West",    90000,                 87000,              "Review"    ],
    ["S006",     "Lakeside Store",   "Frank Brown",    "Philadelphia","East",    130000,                128000,             "Pending"   ],
    ["S007",     "Summit Square",    "Grace Kim",      "San Antonio", "South",   95000,                 101000,             "Verified"  ],
    ["S008",     "Metro Junction",   "Henry Davis",    "San Diego",   "West",    105000,                99000,              "Pending"   ],
    ["S009",     "Pinewood Corner",  "Iris Chen",      "Dallas",      "South",   115000,                120000,             "Verified"  ],
    ["S010",     "Cedarwood Mall",   "James Wilson",   "San Jose",    "West",    125000,                118000,             "Review"    ],
  ];

  const ws = XLSX.utils.aoa_to_sheet(sampleData);
  // Column widths
  ws["!cols"] = [
    { wch: 10 }, { wch: 20 }, { wch: 18 }, { wch: 14 },
    { wch: 10 }, { wch: 22 }, { wch: 18 }, { wch: 12 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, "Stores");

  const notesData = [
    ["Sheet Notes"],
    [""],
    ["This file is managed via XL-Forge web editor."],
    ["Store managers can verify and update their information through the web UI."],
    ["Last updated: " + new Date().toLocaleDateString()],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(notesData), "Notes");
  XLSX.writeFile(wb, FILE_PATH);
  console.log(`Sample file created: ${FILE_PATH}`);
}

// ── Helper: read workbook from disk ──────────────────────────────────────────
const readWorkbook = () => XLSX.readFile(FILE_PATH);

// ── Helper: convert workbook to JSON-serialisable structure ──────────────────
const wbToJson = (wb) => {
  const result = {};
  for (const name of wb.SheetNames) {
    result[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], {
      header: 1, defval: "",
    });
  }
  return { sheetNames: wb.SheetNames, sheets: result };
};

// ── GET /api/health ───────────────────────────────────────────────────────────
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", file: FILE_NAME, time: new Date().toISOString() });
});

// ── GET /api/file ─────────────────────────────────────────────────────────────
// Returns all sheet data as JSON. The frontend loads this on startup.
app.get("/api/file", (req, res) => {
  try {
    const wb   = readWorkbook();
    const data = wbToJson(wb);
    const stat = fs.statSync(FILE_PATH);
    res.json({
      ...data,
      fileName: FILE_NAME,
      fileSize: stat.size,
      lastModified: stat.mtime,
    });
  } catch (err) {
    console.error("Error reading file:", err);
    res.status(500).json({ error: "Could not read Excel file: " + err.message });
  }
});

// ── POST /api/save ────────────────────────────────────────────────────────────
// Body: { sheets: { SheetName: [[row],[row],...] }, sheetNames: [...] }
// Writes the data back to the Excel file on disk.
app.post("/api/save", (req, res) => {
  try {
    const { sheets, sheetNames } = req.body;
    if (!sheets || !sheetNames) {
      return res.status(400).json({ error: "Missing sheets or sheetNames in request body" });
    }

    // Read existing workbook to preserve formatting/metadata where possible
    const wb = XLSX.utils.book_new();
    for (const name of sheetNames) {
      const data = sheets[name] || [[]];
      wb.SheetNames.push(name);
      wb.Sheets[name] = XLSX.utils.aoa_to_sheet(data);
    }

    // Backup current file before overwriting
    const backupPath = FILE_PATH.replace(".xlsx", `_backup_${Date.now()}.xlsx`);
    if (fs.existsSync(FILE_PATH)) {
      fs.copyFileSync(FILE_PATH, backupPath);
      // Keep only the last 5 backups
      const backups = fs.readdirSync(DATA_DIR)
        .filter((f) => f.includes("_backup_"))
        .sort();
      if (backups.length > 5) {
        fs.unlinkSync(path.join(DATA_DIR, backups[0]));
      }
    }

    XLSX.writeFile(wb, FILE_PATH);
    const stat = fs.statSync(FILE_PATH);
    console.log(`File saved: ${FILE_PATH} (${stat.size} bytes) at ${new Date().toISOString()}`);

    res.json({
      success: true,
      message: "File saved successfully",
      lastModified: stat.mtime,
      fileSize: stat.size,
    });
  } catch (err) {
    console.error("Error saving file:", err);
    res.status(500).json({ error: "Could not save file: " + err.message });
  }
});

// ── POST /api/claude ──────────────────────────────────────────────────────────
// Proxy for Anthropic API — keeps ANTHROPIC_API_KEY on the server only.
app.post("/api/claude", async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY is not set on the server" });
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(req.body),
    });
    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err) {
    console.error("Claude API error:", err);
    res.status(500).json({ error: err.message });
  }
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅ XL-Forge server running on http://localhost:${PORT}`);
  console.log(`   Excel file: ${FILE_PATH}`);
  console.log(`   API endpoints:`);
  console.log(`     GET  http://localhost:${PORT}/api/file`);
  console.log(`     POST http://localhost:${PORT}/api/save`);
  console.log(`     POST http://localhost:${PORT}/api/claude\n`);
});
