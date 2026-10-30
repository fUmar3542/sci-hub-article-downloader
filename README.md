# Sci-Hub Article Downloader

Playwright automation that reads DOIs from an Excel file, opens each one on
`https://sci-hub.reatk.com/{DOI}`, and downloads the article PDF if it loaded.
Articles that don't load are skipped. Every row prints `DOWNLOADED`, `SKIPPED`
or `ALREADY EXISTS` to the console.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

## Usage

```bash
python scihub_downloader.py Articles_to_Download.xlsx
```

| Option       | Description                                          |
|--------------|------------------------------------------------------|
| `--out DIR`  | Output folder (default: `downloads`)                 |
| `--headless` | Run without a visible browser window                 |
| `--limit N`  | Only process the first N rows (good for a test run)  |
| `--min-delay S` / `--max-delay S` | Random wait between articles (default 3-7 s) |
| `--post-click-delay MIN MAX` | Wait after clicking download before moving on (default 1 2 s) |
| `--fast` | Minimal waiting (0.5-1.5 s between articles, no mouse movement, no breaks) |
| `--skip-doi-check` | Don't verify DOIs against doi.org first        |

Try `--limit 5` first to confirm it works on your machine.

## Excel format

Expected columns (first sheet, header in row 1):

| Column                          | Used for                          |
|---------------------------------|-----------------------------------|
| `MasterID`                      | Console/log labels                |
| `Filename to save as (exact)`   | Name of the saved PDF             |
| `DOI`                           | Appended to the Sci-Hub URL       |

Other columns are ignored. Rows with a blank DOI are skipped.

## How it works

1. Clean the DOI (strips `https://doi.org/` etc.) and check it exists via the
   doi.org registry. Malformed or non-existent DOIs are skipped and never sent
   to Sci-Hub. (If the check itself fails due to a network error, it tries
   Sci-Hub anyway.)
2. Open `https://sci-hub.reatk.com/{DOI}` and pause as if reading the page.
3. Wait up to 20 s for the embedded PDF or the "↓ save" button.
   If neither appears, the row is skipped.
4. Move the mouse / scroll a little, hover the save button, click it and save
   the download under the Excel filename.
5. If that fails, fetch the PDF directly from the embed URL
   (Chrome's built-in PDF viewer toolbar can't be clicked by Playwright).

## Human-like pacing

- Default: random 3-7 s between articles, a short pause after the page opens,
  a little mouse movement and a hover before clicking save, and a 15-40 s break
  every 10-20 articles.
- `--fast`: 0.5-1.5 s between articles, no mouse movement, no breaks.
- After clicking save it only waits 1-2 s, then moves on to the next DOI while the
  PDF keeps downloading in the background. Each file is confirmed/saved once the
  next page has loaded (or at the end), so the transfer time overlaps with the
  next article's lookup instead of adding to it.
- Repeated DOIs in the sheet are copied locally instead of fetched again.
- Tune with `--min-delay` / `--max-delay`. If the site starts blocking or
  showing captchas, slow down again.

## Output

- `downloads/*.pdf` - the saved articles
- `downloads/results.csv` - one line per row: `MasterID, DOI, status, detail`

Re-running is safe: files that already exist are skipped, so you can resume
after an interruption.

## Troubleshooting

- **Everything is skipped:** the site layout may differ. Open one DOI in the
  browser, inspect the save button/embed, and update `SAVE_SELECTOR` /
  `PDF_SELECTOR` at the top of the script.
- **Captcha or blocking:** run without `--headless` and solve it manually.
- **Slow site:** raise `LOAD_TIMEOUT` / `DOWNLOAD_TIMEOUT` in the script.

## Disclaimer

Make sure your use complies with copyright law and the policies that apply to
you or your institution.
