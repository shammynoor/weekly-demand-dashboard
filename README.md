# Weekly Demand Dashboard — GitHub Pages

## Deploy

1. In `https://github.com/shammynoor/weekly-demand-dashboard`, select **Add file → Upload files**.
2. Upload **index.html** and **app.js** from this package to the repository root. Commit changes.
3. GitHub Settings → Pages → Deploy from branch → `main` / `(root)`.
4. Visit `https://shammynoor.github.io/weekly-demand-dashboard/`.
5. Sign in to the **personal Microsoft account** that owns the CSVs; select the folder using **Choose folder**.

## Microsoft registration

- Application (client) ID: `ba2bc895-4213-447c-ba62-d7edec6e8173`
- Supported account types: Personal Microsoft accounts
- Redirect URI (SPA): `https://shammynoor.github.io/weekly-demand-dashboard/`
- Delegated Microsoft Graph permissions: `Files.Read` and `Files.ReadWrite.AppFolder`
- No client secret. Browser authentication uses MSAL with PKCE.

## Data behaviour

- Reads all `.csv` files **directly in** the chosen OneDrive folder, not subfolders.
- Refresh button forces reload; background check every 5 minutes **while the page remains open**.
- Only CSV files and metadata are fetched; nothing is uploaded to GitHub.
- Rows are aggregated in memory; the dashboard does not save raw CSV content to browser storage.
- Category assignments are saved to the OneDrive app folder and restored after Microsoft sign-in on each computer. The selected CSV folder is remembered per browser. If settings sync fails, a status message appears on the Manage categories page.
- CSV format: `Appointment Date`, `Appointment Time`, `Booked Time to Slot Time` (minutes), `Slot Type`, `Availability`. Blank appointment dates/times are forward-filled per file. Only `Booked` rows with valid lead times are included. Booking week is calculated from appointment date/time minus lead time.
- Duplicate appointments appearing in multiple CSV exports **will be double-counted**. Use non-overlapping CSV exports or deduplicate them before adding to the folder.
- `Urgent` and `Routine` default slot classification is inferred from slot names; review and correct under **Manage categories**. You can import existing category settings using JSON export from your old dashboard.
- Bank holiday dates currently cover 2025 and 2026 England/Wales; update `holidayDates` in `app.js` for later years.
- Public GitHub Pages means **website code is public**, not OneDrive files. Anyone with the URL can view the blank website, but Microsoft authentication is needed to read OneDrive.
- External libraries are loaded from jsDelivr CDN.

## Troubleshooting

- 404: Pages deployment may take several minutes; ensure both files are at repository root.
- `redirect_uri_mismatch`: verify the Microsoft SPA redirect URI exactly matches the website URL (including trailing slash).
- `Access denied`: check `Files.Read` delegated permission and sign in to the Hotmail account that owns the OneDrive folder.
- Blank chart: review slot category assignments, source CSV column names, or duplicate report files.

## Category behaviour

- Assign slot types on **Manage categories** to Urgent, Routine or Uncategorised.
- No automatic keyword assignment is applied. Uncategorised is explicitly retained and excluded from both charts.
- Category changes update charts immediately and are saved to the app-only OneDrive settings file `weekly-demand-categories.json`.
- Existing local settings are retained if cloud settings do not yet exist. Change an assignment once to save them to OneDrive.
- Export a JSON backup before replacing the old website files.
