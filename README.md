# Excel-Ai-Assistant

An open-source AI assistant that lives inside Excel. Type what you want in plain English and it edits your workbook: formulas, formatting, sorting, tables, charts and PivotTables.

You bring your own [Gemini API key](https://aistudio.google.com/apikey). There is no backend and no account.

> **Status:** early and experimental. Core flows work, but not every Excel action has been tested on every Excel version. Use **Undo last** and keep a backup of important workbooks.

## What it does

Open the **AI Assistant** pane from the Home tab and type a command:

- "Add a Total column that multiplies Qty by Price"
- "Bold the header row and fill it light blue"
- "Sort by the second column, largest first"
- "Make a column chart from my selection"
- "Create a pivot table showing total Revenue by Region"

The model reads your selection, calls tools, and the add-in applies them to the open workbook. Each step shows in the pane as a ✓ or ⚠ line, followed by a short summary.

## Features

| Area | Tools |
|---|---|
| Read | `get_workbook_info`, `read_range` |
| Write | `write_cells` (values and formulas), `clear_range` |
| Format | `format_range` (bold, italic, colours, number format, alignment, wrap), `autofit_columns` |
| Data | `sort_range`, `create_table` |
| Visuals | `create_chart`, `create_pivot_table` |

Also: a model picker (listed from your key), a **Stop** button, and **Undo last** for the previous command.

**Undo covers** cell contents and formulas, sorts, tables, charts and PivotTables. **It does not restore formatting.**

## How it works

```
You type a command
      │
      ▼
React task pane (Office.js) ──► builds a compact summary of your selection
      │                         (headers, column types, a few sample rows)
      ▼
Gemini (function calling) ──► decides which tools to call
      │
      ▼
Tool arguments validated with Zod ──► executed with the Excel JavaScript API
      │
      ▼
Results go back to Gemini until it finishes (max 12 steps)
```

Design choices:

- **The model never edits cells directly.** It requests tools; the add-in validates every argument (A1 address format, hex colours, a cap of 2,000 rows / 20,000 cells per write) before touching Excel.
- **Only a summary is sent**, not the whole workbook. The model reads more ranges on demand.
- **Cell contents are treated as data, not instructions**, to reduce prompt-injection risk from spreadsheet content.
- **Pure logic is separate from Excel code** (agent loop, tool schemas, pivot field matching, undo log), so it is unit tested without Excel.

## Privacy

- Your Gemini key is stored in this add-in's browser storage on your computer only, and sent only to Google's Gemini API.
- When you send a command, the selection summary and any ranges the model reads are sent to Google for processing. Do not use it with data you are not allowed to share with Google.
- No telemetry, no server of ours.

## Tech stack

TypeScript, React, Fluent UI, Vite, Office.js (Excel API), [`@google/genai`](https://www.npmjs.com/package/@google/genai), Zod, Vitest.

## Getting started

### Requirements

- Windows or macOS with **desktop Microsoft Excel** (Microsoft 365; PivotTables need Excel API 1.8+)
- Node.js 20+ (developed on Node 24)
- A Gemini API key

### Install and run

```bash
git clone https://github.com/ifahimkhan/Excel-Ai-Assistant.git
cd Excel-Ai-Assistant
npm install

npm run certs      # one time: create and trust a localhost HTTPS certificate
npm run dev        # serves the add-in at https://localhost:3100
```

In a second terminal:

```bash
npm run sideload   # opens Excel with the add-in loaded
```

In Excel, go to **Home → AI Assistant**, enter your Gemini key, and start typing commands.

### Scripts

| Command | What it does |
|---|---|
| `npm run dev` | Vite dev server (HTTPS, port 3100) |
| `npm run build` | Typecheck and production build |
| `npm test` | Run unit tests |
| `npm run typecheck` | TypeScript check only |
| `npm run certs` | Install the localhost dev certificate |
| `npm run sideload` / `npm run stop` | Load / unload the add-in in desktop Excel |
| `npm run validate` | Validate `manifest.xml` |

## Project structure

```
manifest.xml            Office add-in manifest (ribbon button, task pane URL)
src/
  agent/                Model <-> tool loop, tool schemas, Gemini client, prompt
  excel/                Excel executors, selection summary, undo log, pivot field matching
  ui/                   React components (key gate, chat, app shell)
  lib/                  Key storage
scripts/make-icons.mjs  Generates placeholder icons
```

## Troubleshooting

- **Button missing after sideload:** close all Excel windows and run `npm run sideload` once. Check the Home tab's overflow menu on narrow windows.
- **Blank pane:** make sure `npm run dev` is running and `https://localhost:3100` loads in a browser. Re-run `npm run certs` if the certificate is not trusted.
- **"Port 3100 is already in use":** another dev server is running; stop it or change the port in `vite.config.ts` and `manifest.xml`.
- **Manifest changes not showing:** Excel caches the manifest. Restart Excel and sideload again.
- **Quota or rate-limit errors:** these come from Gemini; wait or use a different model from the dropdown.

## Limitations and roadmap

- Undo does not restore formatting
- The model sees your selection summary, not the entire workbook, unless it reads ranges itself
- Desktop Excel is the tested target; Excel on the web is untested
- Ideas: formatting undo, conditional formatting, data cleaning helpers, multi-sheet awareness, other model providers

## Contributing

Issues and pull requests are welcome. Please run `npm run typecheck && npm test` before submitting, and add tests for any new pure logic.

## License

[MIT](LICENSE)
