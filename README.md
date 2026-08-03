# DataLens
AI-powered, privacy-first data analysis — entirely on your device.
DataLens is a single-file React application that lets you upload a data file and instantly
get column profiling, visual dashboards, automatic insights, a data-quality score, and a
natural-language chat interface — all without the file ever leaving your browser. There
is no server, no API call, and no network request involved in analyzing your data.

# Why DataLens
Most "AI data analyst" tools require uploading your file to a server before you can ask a
single question about it. DataLens takes the opposite approach: every byte of parsing,
statistics, chart rendering, and Q&A happens inside the browser tab you're using right
now. Close the tab and nothing about your data persists anywhere.

#Features
  File support
    CSV — parsed with PapaParse
    Excel ( .xlsx , .xls )— parsed with SheetJS
    JSON — arrays of objects, or an object containing an array
    TXT — auto-detects delimited data, otherwise treated as line-based text
    PDF — best-eort text extraction for simple, uncompressed text PDFs (see Limitations)
    Drag-and-dropor click-to-browse upload
    Built-in sample dataset for a no-file quick start

  Automatic analysis
    For every column, DataLens detects:
        Data type — numeric, categorical, date, boolean, text
        Missing values — count and percentage
        Uniqueness — unique value count, topcategories
        Outliers — via the 1.5×IQR rule for numeric columns
        Duplicate rows — exact row-level duplicates across the dataset
        Correlations — Pearson r between all numeric column pairs
        Trends — direction of numeric columns across row orde
