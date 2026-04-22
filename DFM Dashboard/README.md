# DFM Factory Dashboard

Static DFM site generated from `DFM 2026.xlsx`.

## Files

- `index.html`: presentation dashboard for modernization
- `data.html`: separate data management page with filters and CRUD
- `styles.css`: TV-focused styling
- `app.js`: shared analytics, filters, CRUD, and local storage persistence
- `seed-data.js`: bundled workbook data for offline use
- `scripts/build_seed_data.py`: rebuilds `seed-data.js` from the workbook using Python standard library only
- `.github/workflows/sync-onedrive.yml`: GitHub Actions workflow for Power Automate / OneDrive refresh

## Run

Open `index.html` directly in a browser, or serve the folder locally:

```bash
cd /Users/jwtan/Downloads/Codex
python3 -m http.server 8000
```

Then open:

- `http://localhost:8000/index.html` for the presentation dashboard
- `http://localhost:8000/data.html` for the editable data page

## GitHub Setup

Push this project to GitHub, then add this repository secret:

- `ONEDRIVE_XLSX_URL`: direct download URL for `DFM 2026.xlsx` in OneDrive

The included workflow listens for:

- `repository_dispatch` with event name `sync_onedrive`
- manual runs from the GitHub Actions tab

Power Automate can trigger this workflow using the GitHub connector action:

- `Create a repository dispatch event`
- `Event name`: `sync_onedrive`

## Verified seed totals

- Construction rows: `2151`
- Unique `season + style` keys: `209`
- Unique FG total by `season + style`: `2,673,984`
- Remaining workbook FG conflict: `SP26 / IO1266` with `499` and `492`

## Notes

- FG is counted once per `season + style`, not once per construction-code row.
- Add/edit/delete actions are stored in browser `localStorage`.
- Reset returns the dashboard to the imported Excel seed.
- The main dashboard intentionally does not show raw data records.
