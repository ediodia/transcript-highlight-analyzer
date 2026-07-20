# transcript-highlight-analyzer

Tools for a Line-By-Line qualitative coding workflow on Google Docs transcripts: one script applies codebook highlight colors to matching keywords, the other reads those highlights back out, counts them, and writes results into a Google Sheet and local reports.

## What's in here

### `tag_transcripts.py`
Applies `Codebook.xlsx` highlight colors to matching keywords across every Google Doc transcript in the Line-By-Line folder.

```bash
python tag_transcripts.py                # whole folder
python tag_transcripts.py --doc-id XXXXX  # single doc
```

### `count_highlights.py`
Reads the highlights already present in each transcript (no LLM calls — purely reads existing formatting) and:
1. Writes one row per highlighted segment into the **Analysis** tab of the "Thematic Analysis Sheet" (found by name, expected in the Customer Discovery folder).
2. Writes total appearance counts per code into the **Frequency** tab of that same sheet.
3. Writes local `line_by_line_highlight_counts.md` / `.xlsx` reports, rolled up by research group (Ethicality, Feasibility, Viability, Desirability).

Re-runs are **idempotent** — it checks existing `(Interview ID, Quote)` pairs in the Analysis tab before appending, so running it twice won't duplicate rows.

```bash
python count_highlights.py                 # whole Line-By-Line folder
python count_highlights.py --doc-id XXXXX   # single doc
python count_highlights.py --dry-run        # skip writing to the Sheet, just produce local reports
python count_highlights.py --sheet-id XXXXX # point directly at a spreadsheet instead of searching by name
```

## Why category-level counts are reliable, but code-level counts are best-effort

In `Codebook.xlsx`, every code's highlight color is inherited from its row's fill in column A, and that fill is shared across **all** codes in the same research group (e.g. Ethics, Wellbeing, Inequality, Agency & Empowerment all share the same green). So a highlighted quote's color tells you its **category** unambiguously, but not which of the ~10 codes in that category it is.

- **Category counts** (Theme column, category summary) are exact, derived straight from color.
- **Code-level counts** (Segment ID, Frequency tab) are best-effort: if the highlighted text contains exactly one codebook keyword from that category, the highlight is attributed to that code. If it contains several keywords, or none, it's left as `(unspecified)` in the Analysis row and excluded from the Frequency tab rather than guessed.

## Setup

```bash
pip install openpyxl google-api-python-client google-auth-oauthlib
```

You'll need a Google Cloud project with the Docs, Drive, and Sheets APIs enabled, and an OAuth `credentials.json` (Desktop app type) downloaded from that project. The first run opens a browser to authorize and saves the resulting `token.json` locally.

> **`credentials.json` and `token.json` are excluded from this repo via `.gitignore`** — they're per-user secrets and should never be committed. Each person running these scripts needs their own.

**Note:** if you've used an older version of `count_highlights.py`, delete `token.json` before re-running — this version needs write access to Sheets (`spreadsheets` scope), which an older token won't have been granted.

## Required scopes

| Script | Scopes |
|---|---|
| `tag_transcripts.py` | `documents`, `drive.readonly` |
| `count_highlights.py` | `documents.readonly`, `drive.readonly`, `spreadsheets` |

## Project structure

```
.
├── tag_transcripts.py                  # Applies codebook colors to transcripts
├── count_highlights.py                 # Reads highlights, counts, writes to Sheet + local reports
├── Codebook.xlsx                       # Code definitions, categories, and their colors
├── Thematic Analysis Sheet.xlsx        # Local copy — live version lives in Google Sheets
├── line_by_line_highlight_counts.md    # Generated report (Markdown)
├── line_by_line_highlight_counts.xlsx  # Generated report (Excel, 2-3 tabs)
└── Qualitative_Findings.pdf            # Findings write-up
```

## Codebook format expected

`Codebook.xlsx`, `Sheet1`:
- Legend at `E2:E5` — category names (e.g. "Ethicality", "Feasibility", "Viability", "Desirability"), with each cell's fill color defining that category's color.
- Codes starting at row 9: `Code ID | Code Name | Operational Definition | Included when... | Excluded when... | Example Quote`, with the Code ID cell's fill matched to its nearest legend color.

## License

Add a license of your choice (MIT is a common default for research tooling).
