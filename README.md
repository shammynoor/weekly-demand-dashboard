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


## Weekly Demand Chart filters
Urgent appointments use the `Urgent` Urgency Group (Detailed urgency -> Urgency Group); clinician filter switches are from Clinician Groups (CSV session-holder user type -> Detailed Clinician Types -> Clinician Groups). Routine and Follow up remain mapped to their Detailed urgency categories. Unassigned clinician types are excluded.


## Consult mode filters
The Weekly Demand Chart now has a third filter card, Consult mode, next to Filters and Clinician Groups. Each toggle is populated from Manage categories > Consult mode and filters all three demand charts using the slot type's consult-mode assignment. All switches start on; unassigned slot types are excluded until assigned a consult mode.


## Routine Appointment Availability
New navigation tab: weekly stacked slot counts from the appointment date (not the booking date), for last week, current week, and the following six weeks. Includes only slots whose Detailed urgency category is assigned to the Routine Urgency Group. Booked pastel green, Blocked dark grey, Embargoed light grey, Available white; weekly totals and hover breakdowns.


## Routine chart: Today and independent filters
The Routine Availability page has its own Clinician Groups and Consult mode switches, initially enabled, filtering availability records through saved category mappings. The dotted Today line lies 5%, 25%, 45%, 65%, or 85% across the current week for Monday through Friday, with the plot to its left shaded 50%-opacity grey. Weekends clamp to the Friday position.

Routine chart styling: hidden X tick labels, 8 week headers, dotted vertical week dividers, vertical left circular legend, lighter stack separators, outline-free Available segments, darker plot background and thicker baseline.


## Booking Times tab
Calculated booking timestamps from appointment date/time minus `Booked Time to Slot Time` (minutes) are grouped into 96 15-minute bins across all loaded CSV files. A bar histogram shows count and a line on a secondary 0–100% axis shows the running cumulative share. Independent urgency group, clinician group, and consult mode switches use saved category assignments. Only booked appointments with valid booking lead time and all three group assignments are included. Existing charts remain unchanged.


## Booking Times refinements
The chart shows only 06:00–21:00 with the left count axis hidden; the 0–100% cumulative axis remains on the right. The running percentage is calculated over all bookings throughout the entire day. A dotted vertical line drops from the first 15-minute bin reaching 50% to the bottom of the chart; if that point is outside the displayed range, no marker is drawn.


## Urgent availability detailed-category toggles
The previous Reds & Extras / Priority panels are replaced by one chart and right-hand switches for every Detailed urgency category assigned to the Urgent Urgency Group. Booked portions use category-specific pastel colours (Red red, Extras purple, Priority yellow). Blocked, Embargoed and Available remain separate status categories and are summed over enabled urgency categories. The legend uses circles for enabled categories plus availability statuses.


## Urgent chart display update
Compact urgency toggle panel with category-coloured switches; circular availability/category legend centred beneath the plot; darker plot background and outline-free Available bars. Independent range sliders choose 0–3 weeks before and 0–4 weeks after the current week (defaults: 1 week back, 3 weeks ahead).


## Single dual-handle week range slider
Urgent Availability now uses one shared range track with two independently draggable handles for start (-3 to 0 weeks) and end (0 to +4 weeks), defaulting to last week through three weeks ahead. The highlighted track shows the selected interval.


## Urgent chart right sidebar filters
Added Clinician Groups (pastel green) and Consult mode (pastel blue) toggle panels above Urgent categories in the same 155px sidebar column. Filters are independent of Weekly Demand and Routine chart switches, and apply to all status segments. The chart column retains its previous width.


## Routine availability visual alignment
Routine Appointment Availability now uses a single dual-handle slider (-2 to +7 weeks, default -1 to +6), the same 155px right sidebar as urgent availability, stacked Clinician Groups and Consult mode toggles (pastel green and blue), and a centred legend below the chart. Booked slots use deeper pastel green #81c6a0. The existing daily-relative Today marker and shading remain.


## Availability range and filter safeguards
Routine dual slider keeps start within -2 to 0 (This Week) and end within +1 (Next Week) to +7; chart cannot be reduced below two weeks. Both availability charts have larger, lower Today labels and 50%-opacity grey shading above bars. Urgent and Routine availability toggle groups prevent deselecting their last enabled option.


### Verified corrective patch
- Routine dual range: start -2..0, end 1..7 (same 9-step rail), enforced on every drag.
- Both Today labels 16px inside chart; historical overlay rgba(128,136,146,0.50) with Plotly layer above.
- Urgent detailed categories, urgent clinician/consult filters and routine clinician/consult filters each retain at least one enabled option.
- Changed app.js cache-busting URL to v4.
\n\nRoutine chart: Weeks/Days segmented switch beside range slider, daily bars use appointment dates and preserve all status and side filters. Weekly bars widened. Urgent weekly separators dotted.\n\n\nUrgent availability has Weeks/Days switch beside slider (defaults Days). Both chart granularity controls use compact dark-grey styling. Availability sidebar filters grey out zero-data options based on selected date range and active filters.\n\n\nCorrected switch style: matching Timeline switch UI with 'by Day or Week' and Day/Week value on both availability pages. Fixed urgent weekly Plotly category axis and hover totals to use aggregated weekly buckets.\n

Restored compact Days/Weeks segmented controls on both availability pages, matching dark-grey theme.


Days/Weeks segmented controls: subtle inset shadow on the selected dark-grey button for both pages.
\n\n## Routine chart recovery
Fixed a circular dependency in auto-disabled clinician/consult filters that could leave all options disabled and the Routine chart empty. Eligibility now checks each dimension independently and automatically recovers available selections. Kept both charts and existing Day/Week shadow styling.\n

## Fix: circular filter lockout
Source-data eligibility is evaluated independently for each filter group, not against switches disabled in other groups. If all selected filters would hide every matching appointment, the last switch action is reverted. An existing all-off state is recovered by enabling available options. Applies to urgent and routine charts.
\n\n## Manage people\nNew navigation tab listing distinct CSV Full Name of the Session Holder of the Session values, with User Type of the Session Holder of the Session and a Locum Yes/No selector (default No). Records are deduplicated by name + user type, sorted by name, and searchable. Locum selections are stored locally and synced to the existing private OneDrive settings JSON using Save to OneDrive. No changes to appointment counts or filtering.\n

Fix: declare peopleAggregate in loadFolder before passing it to aggregateCSV. People are now collected from every parsed CSV row after fill-down, deduplicated case-insensitively by full name; user type is the most frequent recorded type for that name. Existing appointment count logic is unchanged.


### 2026-10-23 critical refresh repair
Declare peopleAggregate in loadFolder before CSV aggregation, count occurrences per person/type, and deduplicate by person name (case-insensitive) selecting the most frequently observed user type. Correct name whitespace normalisation. This fixes ReferenceError: peopleAggregate is not defined.

## Manage people: clinician type and locum
- Sorted by User Type then Name.
- Third column is an editable dropdown of Detailed Clinician Types categories.
- Preselects from the existing User Type → Detailed Clinician Types assignments, falling back to exact name matches; otherwise Unassigned.
- Fourth column is Locum Yes/No. Conservatively suggests Yes only for explicit 'locum', 'agency', 'bank', or 'temporary' in user type; otherwise No. These are guesses, not verified employment statuses.
- Both selections are saved to browser storage and to the existing private OneDrive settings on Save. Explicit selections override guesses.


## Person-specific clinician categorisation
All charts now use the named session holder's Assigned Clinician Type from Manage people as their detailed clinician category. This takes priority over the CSV User Type field. The original CSV user type remains visible in Manage people and is used as fallback when no valid person-specific category exists. The person name is retained in each aggregated appointment record to prevent different clinicians from being merged before filtering. Changing a person's category updates visible charts without reimporting CSVs.

## Manage people — clinician group boxes
Manage people now shows a box for each existing Clinician Group, containing its assigned Detailed Clinician Types. Each box has an Add type field and button. A final Add Group box creates another Clinician Group. Types can be reassigned to another group using their dropdown. All changes update the existing clinician group/category data model, the category editors, and the charts; they are stored locally and sync to the existing private OneDrive settings when saved.

## Compact clinician group organiser
Manage people displays up to six clinician-group boxes per row on wide screens (five at narrower widths). Detailed clinician types are draggable between groups. Each type has rename (pencil) and delete (×) actions; rename updates person and CSV user-type assignments, and deletion clears those assignments. Add Group is a single button that reveals an input only when clicked. Each group has a compact Add type button. Changes are stored locally and sync to OneDrive with the existing Save button.


## Clinician Group rename/delete controls
Every main group heading now has a pencil (rename) and × (delete). Renaming updates the group's name and all detailed clinician type group assignments. Deleting requires confirmation and leaves its detailed types intact but unassigned, so person clinician types are preserved. Both actions update local settings and OneDrive sync on Save.


Manage people table filters: a Clinician Groups dropdown next to search (All or any saved group) and an Unassigned toggle that shows only people with no valid clinician-group assignment. When Unassigned is on the group dropdown is disabled. Filters use the effective Assigned Clinician Type, including manual overrides, and refresh when people or group assignments change.

## Detailed clinician type colours
Each Detailed Clinician Type has a stable pastel colour swatch in Manage people. Click its dot to open the browser's native colour picker. Colours are saved in localStorage and private OneDrive settings (`clinicianTypeColours`). The Manage people table now includes Clinician Group as column four, with a dot showing the person's assigned detailed clinician type colour; Locum is column five. Clinician-group filters on Urgent and Routine availability charts use the first alphabetically assigned type's colour for the group, while category/status colours remain unchanged.

## Group colours (corrected)
One colour per Clinician Group, shown as a larger clickable circle in the group header. No colours on individual Detailed Clinician Types. More saturated pastel defaults. The Manage People Clinician Group column and the Urgent/Routine clinician-group toggles use the same group colour. Group colours persist in local browser settings and private OneDrive settings as `clinicianGroupColours`; renaming/deleting groups carries over or removes the group colour. Prior per-type colours are ignored.


## Manage categories simplified
Removed the Detailed Clinician Types and Clinician Groups tabs and editors from the visible Manage categories page. Their management remains available under Manage people. Internal legacy editor nodes are retained invisibly for compatibility with existing JS event bindings and refresh logic; the underlying clinician assignments and saved OneDrive settings are unchanged.

## Weekly Demand chart performance
Reduced repeated CSV scans and person lookup overhead. A named-person → user-type map avoids repeated `peopleRecords.find` searches; clinician and filter results are cached per chart refresh; urgency category totals for all three charts are calculated in one pass rather than repeated scans. Rapid switch changes are coalesced into one animation-frame refresh. Chart definitions and appointment-counting rules are unchanged.


Navigation: chart tabs are left aligned; Slot settings and People settings are right aligned in a muted lavender-grey style, with responsive wrapping. The page IDs and chart logic are unchanged.

## Replace CSV from desktop by drag and drop
Drag exactly one `.csv` onto any dashboard page. Only a file with the same case-insensitive filename as an existing CSV directly in the configured OneDrive folder is accepted. Both CSVs must have compatible appointment-export headings in identical order; malformed, empty, or wrong-extension files are rejected. The dashboard requires confirmation before replacing the existing OneDrive item by its ID, then refreshes the data. This does NOT edit the separate private app-folder settings JSON containing people, clinician groups, colours, urgency/slot assignments or other settings. Upload requires delegated Microsoft Graph Files.ReadWrite permission, requested on sign-in; a Microsoft consent prompt may be shown. Files above 200 MB are rejected (simple Graph upload endpoint has a 250 MB limit). No new filenames or new files are allowed.

## Routine Appointment Availability visual refresh
Updated only the Routine Availability Plotly styling: cleaner off-white chart background, subtly stronger green booked bars, refined blue-grey statuses, softer gridlines and borders, more readable totals and headers, modest bar spacing, and a blue Today line. Historical shading stays in front of the bars with alpha 0.20 (80% transparent). The Today x-position calculation, slider, daily/weekly behaviour, appointment aggregation, all filters and saved settings are unchanged.
