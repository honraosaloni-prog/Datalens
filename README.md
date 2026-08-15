DataLens

AI-powered, privacy-first data analysis — entirely on your device.

🔗 Live Demo

DataLens is a single-file React application that lets you upload a data file and instantly get column profiling, visual dashboards, automatic insights, a data-quality score, and a natural-language chat interface — all without the file ever leaving your browser. There is no server, no API call, and no network request involved in analyzing your data.

Why DataLens

Most "AI data analyst" tools require uploading your file to a server before you can ask a single question about it. DataLens takes the opposite approach: every byte of parsing, statistics, chart rendering, and Q&A happens inside the browser tab you're using right now. Close the tab and nothing about your data persists anywhere.

Features
File support
CSV — parsed with PapaParse
Excel (.xlsx, .xls) — parsed with SheetJS
JSON — arrays of objects, or an object containing an array
TXT — auto-detects delimited data, otherwise treated as line-based text
PDF — best-effort text extraction for simple, uncompressed text PDFs (see Limitations)
Drag-and-drop or click-to-browse upload
Built-in sample dataset for a no-file quick start
Automatic analysis

For every column, DataLens detects:

Data type — numeric, categorical, date, boolean, text
Missing values — count and percentage
Uniqueness — unique value count, top categories
Outliers — via the 1.5×IQR rule for numeric columns
Duplicate rows — exact row-level duplicates across the dataset
Correlations — Pearson r between all numeric column pairs
Trends — direction of numeric columns across row order
Dashboard

Charts adapt to whatever is in your dataset:

Column type breakdown (pie)
Missing values by column (bar)
Distribution histogram for any numeric column (selectable)
Category frequency for any categorical column (selectable)
Correlation matrix (up to 8 numeric columns)
Data preview table with per-column type badges

Every chart has a one-click PNG export button.

Insights & recommendations

A dedicated tab lists plain-language, auto-generated observations (e.g. missing-value warnings, correlation callouts, trend detection) and actionable recommendations (e.g. "remove duplicate rows," "investigate outliers in X").

Data quality report

A weighted quality score (0–100) built from:

Completeness — share of non-missing cells
Uniqueness — share of non-duplicate rows
Validity — share of numeric values that aren't statistical outliers

Includes a per-column breakdown table.

Ask DataLens (chat)

A natural-language interface answers questions about your data — missing values, duplicates, correlations, outliers, trends, quality score, or stats for a specific column — by reasoning over the locally computed analysis. This is a rule-based local engine, not a call to an external AI API, which is what keeps the "your data never leaves the device" guarantee intact even in chat.

Export
PDF report — opens the browser print dialog with a formatted report; choose "Save as PDF"
JSON report — full machine-readable report (stats, insights, correlations, recommendations)
Cleaned CSV — original data with exact duplicate rows removed
PNG — any dashboard chart, individually
Design
Light and dark mode
Responsive layout (sidebar nav on desktop, scrollable tab bar on mobile)
Animated "on-device" privacy indicator
Smooth load/transition animations
Getting started

DataLens is a single React component with no required backend. These steps use Vite, and work identically on your machine or in a GitHub Codespace.

1. Scaffold a React app
bash
npm create vite@latest . -- --template react
npm install

(the . scaffolds into the current folder — say "Ignore files and continue" if prompted)

2. Install DataLens' dependencies
bash
npm install papaparse xlsx recharts lucide-react
3. Install Tailwind CSS (v4)
bash
npm install -D tailwindcss @tailwindcss/postcss

Create postcss.config.js at the project root:

js
export default {
  plugins: {
    '@tailwindcss/postcss': {},
  },
}

Replace the entire contents of src/index.css with just:

css
@import "tailwindcss";
4. Add the component

Create src/DataLens.jsx and paste in the full DataLens code. Replace src/App.jsx with:

jsx
import DataLensApp from './DataLens';

export default function App() {
  return <DataLensApp />;
}
5. Run it
bash
npm run dev -- --host

Open the printed local URL (or, in a Codespace, the forwarded port under the Ports tab).

No environment variables, API keys, or backend services are required.

How it works
Stage	What happens
Upload	File is read with the browser's FileReader / File.arrayBuffer() API — never transmitted
Parse	Format-specific parser (PapaParse, SheetJS, JSON.parse, or a custom PDF text extractor) converts the file into rows and columns in memory
Profile	Each column is typed and statistically summarized (mean, median, std dev, quartiles, outliers, frequency tables)
Analyze	Dataset-level checks run: duplicate detection, pairwise correlation, trend detection, quality scoring
Present	Charts, insight text, and the quality gauge are derived from the analysis object and rendered with Recharts
Chat	Questions are matched against keywords and column names, then answered directly from the analysis object already in memory
Limitations
PDF parsing is best-effort. True PDF extraction requires a dedicated library (e.g. pdf.js) that isn't bundled here. DataLens uses a lightweight regex-based extractor that only works on simple, uncompressed, text-based PDFs. Scanned/image PDFs or PDFs with compressed content streams will show a clear on-screen notice asking you to export as TXT or CSV instead.
Large files: datasets are capped at 20,000 rows for in-browser performance; larger files are truncated with a visible notice.
Chat is heuristic, not generative. It answers from the precomputed statistics using keyword and column-name matching — it won't produce free-form analysis outside what the app already computes.
Single-sheet Excel: only the first sheet of a multi-sheet workbook is analyzed (noted on screen if applicable).
Privacy

All parsing, statistics, chart rendering, and chat happen with in-browser JavaScript. DataLens makes no fetch/XHR calls with your file's contents, uses no localStorage for your data, and has no server component. Closing the tab discards everything

## Screenshots

<img width="1365" height="654" alt="output1" src="https://github.com/user-attachments/assets/ae8e7e9e-28fd-4c11-8867-52275da705d4" />
<img width="1365" height="654" alt="output12" src="https://github.com/user-attachments/assets/c2ca2faf-236b-42e9-ba78-26f4a0135132" />
<img width="1353" height="656" alt="output2" src="https://github.com/user-attachments/assets/59e0e58a-6acb-435d-a9b0-a01a42133650" />
<img width="1365" height="656" alt="output4" src="https://github.com/user-attachments/assets/e707376f-4b90-4bd6-a46a-24ae4d9040c2" />
<img width="1365" height="656" alt="output7" src="https://github.com/user-attachments/assets/456bd9c4-26c9-4356-8c27-ab31c3cd87d4" />
<img width="1365" height="655" alt="output9" src="https://github.com/user-attachments/assets/a8934886-6684-4c98-a5b4-40dd4974bae7" />
<img width="1365" height="652" alt="output10" src="https://github.com/user-attachments/assets/a830da74-46cd-465d-8293-09597247b32b" />
<img width="1365" height="656" alt="output11" src="https://github.com/user-attachments/assets/9715b5ff-6c1e-4cab-b909-bb38079ebde4" />
<img width="1365" height="656" alt="output5" src="https://github.com/user-attachments/assets/55fb6fd1-af51-4d8c-8577-a82013991ca4" />
<img width="1365" height="658" alt="output6" src="https://github.com/user-attachments/assets/93bae6f2-0875-498a-8ab7-ad0cdcca3f9c" />




