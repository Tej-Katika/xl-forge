# ⚡ XL-Forge — Spreadsheet Editor

Edit Excel files directly in the browser with an interactive grid **and** AI-powered prompts.  
Built with React + Vite. Deploy in minutes to Vercel (free tier).

---

## ✨ Features

| Feature | Description |
|---|---|
| 📂 Upload | Drag-and-drop or click-to-browse for `.xlsx`, `.xls`, `.csv` |
| ✏️ Inline editing | Click any cell to edit it; Tab/Enter to confirm |
| ➕ Rows & Columns | Add or delete rows/columns with one click |
| ↩ Undo / Redo | Full history — `Ctrl+Z` / `Ctrl+Y` |
| 🤖 AI prompts | Describe changes in plain English; review before applying |
| 📑 Multi-sheet | Switch between sheets via tabs |
| ⬇️ Download | Save the modified file at any time |
| ♿ Accessible | WCAG 2.1 AA — keyboard nav, screen-reader announcements, skip link |

---

## 🗂 Project structure

```
xl-forge/
├── api/
│   └── claude.js          ← Vercel serverless proxy (keeps API key secret)
├── src/
│   ├── main.jsx            ← React entry point
│   └── App.jsx             ← Main app (copy excel-editor.jsx here)
├── index.html
├── vite.config.js
├── vercel.json
├── package.json
├── .env.example
└── .gitignore
```

---

## 🚀 Local development (5 minutes)

### 1. Install dependencies

```bash
npm install
```

### 2. Add your Anthropic API key

```bash
cp .env.example .env
# Open .env and paste your key:
# ANTHROPIC_API_KEY=sk-ant-...
```

Get a key at [console.anthropic.com](https://console.anthropic.com).

### 3. Update the API call in App.jsx

The app currently calls the Anthropic API directly (fine for local testing).  
Before deploying, change the `fetch` URL in `runPrompt()` from:

```js
// ❌ Direct call — exposes API key in the browser
fetch("https://api.anthropic.com/v1/messages", {
  headers: { "Content-Type": "application/json" },
  ...
})
```

to:

```js
// ✅ Goes through your secure serverless proxy
fetch("/api/claude", {
  headers: { "Content-Type": "application/json" },
  ...
})
```

Also remove the `"x-api-key"` header from the browser-side request — the proxy adds it server-side.

### 4. Start the dev server

```bash
npm run dev
# Open http://localhost:5173
```

---

## ☁️ Deploy to Vercel (recommended — free)

Vercel hosts the React frontend **and** the `/api/claude.js` serverless function in one place.

### Step 1 — Push to GitHub

```bash
git init
git add .
git commit -m "initial commit"
gh repo create xl-forge --public --push   # or use github.com manually
```

### Step 2 — Import to Vercel

1. Go to [vercel.com](https://vercel.com) → **Add New Project**
2. Click **Import** next to your GitHub repo
3. Framework preset: **Vite** (auto-detected)
4. Click **Deploy** — Vercel builds and deploys automatically

### Step 3 — Add the API key as an Environment Variable

1. In the Vercel dashboard → your project → **Settings → Environment Variables**
2. Add:
   - **Name:** `ANTHROPIC_API_KEY`
   - **Value:** `sk-ant-...`
   - **Environment:** Production + Preview + Development
3. Click **Save** then **Redeploy**

Your app is now live at `https://xl-forge-xxxx.vercel.app` 🎉

Every `git push` triggers an automatic redeploy.

---

## 🌐 Alternative: Deploy to Netlify

```bash
npm run build          # creates dist/
```

1. Go to [netlify.com](https://netlify.com) → **Add new site → Import from Git**
2. Build command: `npm run build`
3. Publish directory: `dist`
4. **Site settings → Environment variables → Add:** `ANTHROPIC_API_KEY`

For the serverless proxy on Netlify, rename `api/claude.js` → `netlify/functions/claude.js` and adjust the export format:

```js
// netlify/functions/claude.js
exports.handler = async (event) => {
  // same body as api/claude.js, return { statusCode, body }
};
```

---

## 🔒 Security notes

- **Never** commit `.env` or your API key to git (`.gitignore` covers this)
- The `/api/claude.js` proxy ensures `ANTHROPIC_API_KEY` lives only on the server
- The app processes all files **locally in the browser** — no file data is ever uploaded to a server
- Only the plain-text CSV preview (first 25 rows) is sent to the AI for context

---

## 🛠 Customisation tips

| Goal | What to change |
|---|---|
| Limit rows sent to AI | Change `slice(0, 25)` in `runPrompt()` |
| Change AI model | Edit `model:` in `runPrompt()` |
| Add more example prompts | Edit the `EXAMPLES` array |
| Change colour scheme | Edit the `C` tokens at the top of `App.jsx` |
| Add authentication | Use [Vercel Auth](https://vercel.com/docs/security/vercel-auth) or [Clerk](https://clerk.com) |

---

## ♿ Accessibility

- WCAG 2.1 AA colour contrast on all text
- Full keyboard navigation (Tab, Enter, F2, Arrow keys, Delete, Ctrl+Z/Y)
- ARIA grid roles, live regions, `aria-label` on every control
- Skip-to-content link
- `prefers-reduced-motion` respected
- Screen reader announcements on file load, errors, AI completion
