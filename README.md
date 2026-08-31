# Scopus DOI Lookup for Google Sheets

A Google Apps Script that fetches Scopus metadata — journal, title, publication
date, citation count, first author, ISSN, open-access status — for a list of
DOIs, directly inside Google Sheets.

<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/64363461-5b20-4082-aaac-a3c0f1b5cb9e" />



## Features
- **Live mode** — paste a DOI into column A and the row fills in automatically.
- **Batch mode** — process an entire column of existing DOIs at once, fetched
  in parallel chunks.
- Handles blank rows, malformed DOIs, rate limiting, and API errors without
  breaking the sheet.

## Setup
1. Open your Google Sheet → **Extensions → Apps Script**.
2. Delete any existing code in `Code.gs` and paste in the contents of
   [`scopus_tool.gs`](./scopus_tool.gs).
3. Save, then reload the spreadsheet. A **Scopus Tool** menu appears.
4. The first time you use the menu, Google will prompt you to authorize the
   script — this is expected, accept it. After that, live auto-fetch installs
   itself automatically; there's nothing else to enable.

## Usage
- **Live mode**: type or paste a DOI into column A, starting at row 2.
  Metadata fills in to the right automatically.
- **Batch mode**: **Scopus Tool → 🔄 Run Manual Batch (All Rows)** processes
  every DOI already in column A.

## About the embedded API key
This script ships with a working Elsevier API key for convenience, so it
works out of the box with no setup. That key draws against the maintainer's
Elsevier quota and rate limits — for heavy or production use, get your own
free key at [dev.elsevier.com](https://dev.elsevier.com) and swap it into the
`API_KEY` constant near the top of the script.

## License
MIT — see [LICENSE](./LICENSE).
