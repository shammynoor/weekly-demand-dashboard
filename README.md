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
\n\n## Category synchronisation fix\nAll named categories (including Follow up) and their slot-type assignments are now stored in the OneDrive app settings file. On the original computer, old browser-only categories can be migrated from the previous version. Open that computer first after deploying.\n\n\n## Manual category saving\nChanges to category names and slot-type assignments are stored in the current browser immediately, but are **not automatically uploaded**. On Manage Categories, click **Save** and wait for the success message. Other computers pick up the latest saved settings on a full page refresh after sign-in. The existing one-time source-folder setup still saves its folder reference to OneDrive.\n
Default: Include Bank Holiday Weeks is Off; bank-holiday weeks are excluded until switched On.


## Sign-in persistence fix
Microsoft authentication now uses localStorage to remember the signed-in account between tabs and browser restarts. Background OneDrive token renewal failures no longer trigger an unexpected sign-in redirect; they display a status error instead. Login initialization is guarded against duplicate calls. Browsers that block site storage or clear it on exit may still require sign-in. No access tokens are included in the GitHub files.

## Follow-up graph
A third blue Follow-up appointments chart appears below Urgent. Assign slot types to the **Follow up** category in Manage Categories and click **Save** to share across devices. The chart uses the same week exclusions, comparison modes and shared Y-axis scale as the other charts.

### Follow-up chart visibility fix
The blue chart now has its own Y-axis scale, so small follow-up volumes are visible. If no slot types are assigned to Follow up, the chart shows an explanation rather than an unexplained flat line. Assign the relevant slot types in Manage Categories and click Save.

Current incomplete week: when Show current week is enabled, the last point is hollow and the connecting segment dashed on all three charts.

## Clinician type categories
Manage Categories now has two independent dimensions: Urgency (Urgent/Routine/Follow up) and Clinician type (user-created, e.g. GP/Nurse). Each CSV slot type can be assigned to one value in each dimension. Both dimensions save together to the same private OneDrive settings JSON when Save is clicked; old settings without clinician types remain compatible. Clinician types do not change existing charts yet.

## Consult mode categories
Manage Categories now includes a third independent user-defined classification: Consult mode. Add names such as Telephone or Face to face, assign slot types, then click Save. The settings are stored in the same OneDrive app settings JSON as urgency and clinician type. Existing graphs are unchanged.


## Drag-and-drop category assignment

Manage categories now shows draggable slot-type chips and a drop box for every category, including Uncategorised. Creating a new category immediately creates a box. Drag a chip between boxes to assign it. Alternatively, select a chip (click or Enter), then click or press Enter on a destination box. The existing Save button writes all three category dimensions to OneDrive; no data schema migration is required.


## Sign-in stability update
Keeps the dashboard visible during background OneDrive loading and restores the cached Microsoft account when a UI update runs. Versioned JavaScript URL avoids stale browser caches.


## Drag-and-drop repair
Fixed a JavaScript scoping error: the category drag-and-drop renderer was incorrectly declared inside the CSV loading function, preventing the editor from rendering reliably. Authentication display is now latched once an account is established so a background render cannot replace graphs with the sign-in screen. 


## Category tabs
Manage categories now has Urgency, Clinician type and Consult mode tabs. Only one drag-and-drop board is displayed at a time. Save remains available beneath the selected board.


## Rename categories
Each named drag-and-drop category box now has a Rename button. Renaming preserves its slot assignments in all three classification types. Click Save to synchronise to OneDrive.


Category cards now have matching compact Rename and Delete buttons. Slot-count badges have been removed.


## Clinician type chart filter
The main charts page now has a clinician type dropdown populated from the user-created Clinician type categories. Select All clinician types (default), any named type, or Uncategorised. All three charts and their axes/averages recalculate using only matching slot types. The selection is a local view filter, not saved to OneDrive; clinician category definitions and assignments remain synced as before.


## Chart fit and refresh feedback
Charts now size to the available panel width on initial render, window resizing, and switching back from Manage categories. Refresh data displays an animated progress indicator and file-by-file progress, then success or failure beside the button.


## Current week visual update
The incomplete current-week segment uses closely spaced dots and a sampled cubic curve based on adjacent weekly values, retaining the hollow current-week endpoint. Applies to all three charts only when Show current week is enabled.


## Clinician type switches
The chart page shows a separate Clinician type filter panel alongside the existing Filters panel. Every user-created clinician category has its own on/off switch, initially on. Switches filter all three charts. Uncategorised is not shown and its slots are excluded from these filtered charts. Switch state remains while navigating pages, and renaming a clinician type preserves its switch setting.


## Weekly/monthly charts
The Filters box has a Monthly data switch. Monthly totals use actual booking dates and the current-period label changes to Show current month. All three graphs update. Bank-holiday-week exclusions continue to apply.


### Month or week level data
The toggle is labelled Month or week level data, with Week and Month indicators. In Month mode Include Bank Holiday Weeks is forced on and disabled. Returning to Week mode restores the user's previous bank-holiday toggle setting.


## Seasonal backgrounds
Very light grey bands show winter (Dec–Feb), spring (Mar–May), summer (Jun–Aug), and autumn (Sep–Nov) on all three charts, in both weekly/monthly and superimposed/continuous views.


## Shade seasons switch
The Filters box has a Shade seasons switch, on by default. Turning it off removes seasonal background bands from all three charts; turning it on restores them.


## Fix 22 October 2026
Continuous timeline x-axis range is explicitly constrained to displayed periods so season background rectangles cannot extend the axis. The Shade seasons toggle state is now stored outside refresh() and is not reset to ON during every redraw. Both modes and all three charts use the setting.


## Urgent Appointment Availability
Third navigation page with a daily stacked bar chart, using Appointment Date directly and stacking the CSV Availability values. It includes only slot types assigned to the Urgent category. The fixed six-week window runs Monday two weeks before the current week through Sunday three weeks after. Data refresh updates the chart.
\n\n## Availability chart refinement\nUses only the exact Slot Type 'urgent appts' (case-insensitive). Monday–Friday only. Booked (red) at the base, Held Back = Blocked + Embargoed (dark grey), Available (white) on top. Light grey plot background, weekday-only X labels, six week labels above each group.\n

## Urgent category filter correction
The availability chart filters by the Urgent category assignments in Manage Categories, rather than an exact slot type name. Booked, Blocked/Embargoed and Available stacks and weekday display are unchanged.


## Urgency category selector (22 October 2026)
Urgent Appointment Availability has a selectable Urgency category filter, populated from user-created urgency categories, defaulting to Urgent. It does not filter on literal slot-type names. All slot types (including those with no booked rows) appear in Manage categories. The availability chart shows an explanation when the selected category has no matching data.


## Sign-in screen stability
The signed-in dashboard view is latched after Microsoft authentication, so chart redraws, category updates and background data refreshes cannot switch back to the login screen. Only explicit Sign out clears the latch. Microsoft Graph token failures remain visible as errors rather than silently displaying the login page.


## Refresh diagnostics
A OneDrive data-fetch error is shown with its actual reason rather than only 'Refresh failed'. Chart rendering errors are isolated from loading so the dashboard remains signed in and displays the specific chart error. Unhandled background errors are displayed in the status area. These changes aid diagnosing a persistent sign-in/refresh problem without masking it.


## Render and login stability
Removed repeated Plotly redraws during a single CSV load and removed the redraw/resize feedback loop. The charts now render once per data refresh and respond to ordinary window resizing. MSAL session restoration avoids showing the sign-in screen while cached credentials are being restored. Errors remain visible in the dashboard status.


## Persistent dashboard session
The dashboard's visible signed-in state now survives unexpected page reloads in the same tab, independently of chart refresh and CSV parsing. Actual Microsoft access still requires a valid MSAL account. A failed session restoration is reported in the dashboard rather than silently returning to the sign-in page. The Sign out button clears this tab's dashboard view marker.


## Fix: startup crash
Initialise `selectedAvailabilityUrgency` before the initial call to `renderEditor()`; previously it was declared afterwards, causing a temporal-dead-zone ReferenceError and preventing normal dashboard startup. Updated app.js cache version.


## Availability chart spacing
Six weeks are shown with Saturday and Sunday as blank spacing columns (no weekend tick labels or bars). Solid dark grey lines mark each new week. A vertical dotted dark grey line marks today's London-local date. The urgency category dropdown remains in place.


Availability chart styling: separators between Saturday and Sunday; reduced bar gaps; muted red booked bars; 80%-transparent pale red shading before today.


## Group Urgency
Manage Categories includes Detailed urgency (existing slot-type assignments) and Group Urgency (new drag-and-drop grouping of detailed urgency categories). Group categories and assignments are persisted in browser settings and the existing OneDrive category JSON, without changing existing chart category definitions.


## Urgency Group availability filter
The appointment availability chart dropdown now lists Urgency Group categories, not detailed urgency categories. Counts include slots whose detailed urgency category is assigned to the selected Urgency Group. The management tab is labelled Urgency Group.


## Availability Extras
The availability chart now shows the detailed urgency category named Extras as a separate purple stack between Held Back and Available (all statuses of Extras are grouped here without double counting). Week annotations use the same normal-weight font styling as x-axis labels.


Availability visual update: circular left-aligned legend, lighter Held Back and week separators, fine gaps between stack segments, slightly wider bars, taller plot with y-axis headroom, concise daily tooltips and grey totals above each weekday bar.


## Appointment availability: two detailed-urgency charts
The availability page now has two independent, identically styled six-week plots. The upper plot uses Detailed urgency categories Reds and Extras, with Extras counted as a separate purple stack. The lower plot uses only Detailed urgency category Priority, with Booked, Held Back (Blocked + Embargoed) and Available. These charts do not depend on the Urgency Group selector; the selector has been removed from this page. Assignments continue to come from Manage categories.


## Availability chart refinement
Both panels share the same y-axis range. Their legends are vertically centered on the left, weekday names appear rotated inside each weekday bar, and explanatory footnotes are removed (diagnostics remain when data is missing). Priority booked slots are pastel yellow.


## Availability chart refinements
Weekday abbreviations are now vertical x-axis tick labels; Saturdays and Sundays retain empty spacing with no labels. The week separators remain in the weekend gaps. The Priority chart uses pastel yellow shading before Today, while Reds uses light red. Bars are maximised within their category slots, and the gap between the two charts is reduced.


## Chart layout v2
This version adds a visible Chart layout v2 label on the availability page to confirm that GitHub Pages has served the new HTML. The JavaScript reference is uniquely versioned to avoid old cached code. Day ticks are explicitly blank for weekend dates; labels are short, vertical and on the x-axis. Plot height and surrounding HTML spacing have been reduced to bring the two charts together. Past shading for Priority is pastel yellow.

## Availability status breakdown
Both availability charts now show Blocked and Embargoed as separate stacked segments and legend entries, with dark and light grey colours respectively. Extras remain a separate purple category regardless of availability status. The total number of slots and shared Y-axis scaling are unchanged.


## CSV continuation-row fix (22 October 2026)
Rows with blank Slot Type inherit the most recent nonblank Slot Type within the same Appointment Date. Rows without a recognised Availability (Booked, Available, Blocked, Embargoed) are not counted as slots. Date changes reset inherited slot type and time. Both the daily availability charts and weekly booking aggregation use the corrected slot type. This may change counts; verify against your export and category assignments.


## Seven-field CSV fill-down correction
Within each CSV, the import independently fills down blank Appointment Date, Appointment Time, User Type of the Session Holder of the Session, Slot Type, Availability, Town of the Location of the Session, and Full Name of the Session Holder of the Session. The fill-down resets at the beginning of each file. Rows containing only patient-detail continuation data (without an explicit availability, slot type, appointment time, or booked-to-slot lead) are not counted as additional appointment slots. This avoids duplicate counting of secondary patient-detail rows.


## Counting continuation rows (October 2026 correction)
Every non-empty imported CSV row is counted after filling down the seven appointment fields. The previous hasSlotEvidence exclusion has been removed. On the supplied Q4 export, Tuesday 27 October now gives 22 Blocked and 108 Embargoed RED-named rows, matching the Excel figures. The booking lead-time chart still requires a valid booking lead time to calculate booking dates.


## Clinician categories
Detailed Clinician Types is based on the filled-down CSV field `User Type of the Session Holder of the Session` (not Slot Type). Clinician Groups groups those detailed categories. Both persist in browser and OneDrive settings. Weekly chart clinician filters use session-holder user type.
