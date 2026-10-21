# Weekly Demand Dashboard — GitHub Pages

Upload `index.html` and `app.js` to the **root** of https://github.com/shammynoor/weekly-demand-dashboard. GitHub Pages should deploy from `main` / `(root)`.

## Important: Microsoft redirect URI

In Microsoft Entra → App registrations → Weekly Demand Dashboard → Authentication (Preview), add this **Single-page application** redirect URI:

`https://shammynoor.github.io/weekly-demand-dashboard/index.html`

This is important because the root URL was displaying the README rather than the application. Keep the existing root redirect URI if desired, but add the `/index.html` URI before testing the update.

## Using the dashboard

- Unauthenticated users see only a sign-in screen.
- After sign-in, the dashboard loads directly without the OneDrive connection panel. A small Sign out button remains available.
- The CSV source folder is saved in the private OneDrive app settings alongside category assignments.
- **First use of this version:** open it on the **same computer and browser** where you previously selected the OneDrive folder. Sign in there once to migrate the remembered folder to OneDrive settings. After that, other computers load the same folder automatically.
- If no folder has been saved anywhere, go to Manage categories → Configure source folder (one-time setup). This is a fallback only.
- Refresh data manually with the top-right button or wait for the five-minute automatic refresh while the page is open.
- Category assignments sync to the same Microsoft account's OneDrive app folder.

## Privacy

The public GitHub Pages site contains only HTML and JavaScript, not appointment CSV data. OneDrive CSVs are fetched in browser memory after authentication. The app requests delegated `Files.Read` and `Files.ReadWrite.AppFolder`. Never commit tokens, passwords, patient data or CSV files to GitHub.

## Known limitations

The selected source folder is tied to the signed-in Microsoft account. Multiple simultaneous category edits on different devices may overwrite each other. Browser testing with Microsoft sign-in and live OneDrive is still required. Overlapping CSV exports may double-count appointments.


## Login fix (October 2026)
Uses the MSAL Browser 2.38.4 UMD script path instead of the previous 3.x CDN URL, and guards the sign-in button if authentication has not initialised. Upload both index.html and app.js. Clear the browser cache with Ctrl+F5.


Green toggle update: switches are pill-shaped, green when enabled and grey when disabled.
\n\n## Category synchronisation fix\nAll named categories (including Follow up) and their slot-type assignments are now stored in the OneDrive app settings file. On the original computer, old browser-only categories can be migrated from the previous version. Open that computer first after deploying.\n