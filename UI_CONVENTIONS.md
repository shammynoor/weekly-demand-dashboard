# UI conventions

## Mini toggles

Match the compact Clinician Groups sidebar switches: approximately 32.4 × 18 CSS px, with a 14 px thumb and 12 px labels. Use actual dimensions, no scale transform or negative margin. Arrange related controls horizontally beneath their slider. Use 4 px between switch and label, 12 px between the separate 'Show in:' prompt and 'Days', and 32 px between the Days/Weeks control and Group unavailable appts. Allow whole controls to wrap on narrow screens. Preserve accessible labels and keyboard focus. Reuse `.mini-toggle-row` in `index.html`.


## Historic Booking Activity smoothing

Use a compact discrete slider in the right column of the chart Display options panel, beneath the Data and Timeline dropdowns. Label its four stops Off, Low, Medium and High, without repeating the active value beside Smoothing. Retain native range keyboard behaviour, visible focus and an accessible value label. Reuse the existing blue rail/white-rimmed thumb styling.

Use an aligned 65 px label column and 9 px gap to the short smoothing track. Stack the toggle and dropdown/smoothing columns on narrow screens; do not use a full-width smoothing slider or negative spacing.


## Current-period range bar

Use a translucent capped vertical range bar in the series colour to show an indicative final-total range for the incomplete week/month. Keep completed historical lines solid and use a fine dotted connector to the predicted finish. Keep the report timestamp and indicative prediction coverage in hover details. Hover details distinguish actual bookings so far from estimated finishing bounds. If the report is stale or historical coverage is insufficient, explain why no range can be shown. The show-current control owns this visibility.

The range bar has no permanent numeric or upper/lower text labels. Keep figures in hover details. Smoothing uses a two-row layout: title centred with the track, four tick labels below. Align each tick centre with the thumb travel, including the half-thumb inset at both ends.


## Fine whisker with current-period circle

Use a 1.1 px whisker in the series colour at 65% opacity, short 3 px caps, and a 7 px open circle in the same series colour for the observed count. Show no permanent numbers beside either marker. Keep each series pair in one column and keep all series at the same current-period x coordinate; overlap is allowed. Recompute spacing on resize. Distinguish observation and estimate in hover details.

Smoothing must share the neighbouring toolbar label font family, size, weight and line height. Keep the label/track row centred with the toggle labels, reserving tick-label space below it.


## Current-period continuity

Use the historical booking-date/month axis for current observations and finishing ranges. Join the last completed historical point to the finishing estimate with a fine 1.35 px dotted segment (1 px dots, 3 px gaps). Keep its endpoint, observed circle, whisker and caps at one shared x position. All current-period series share the same booking date/month position; whiskers may overlap. Use 1.1 px whiskers at 65% opacity with 3 px caps and indicative 70% prediction coverage. Keep the full historical axis and mean lines continuous; do not reserve a separate prediction lane. Line widths relative to the latest displayed year are 2.6 px, 2.0 px for the previous year, 1.5 px two years earlier, and 1.2 px for all older years. Earlier-year opacity remains unchanged.

## Secondary toggle-box expansion

Save an explicit expand/collapse choice before changing the DOM. Restore saved expansion state once per account/state, without replaying button clicks after every render. A minus click must hide the filter stack, show its selection summary, change to plus and remain collapsed after redraw on every analytics page.

## Personal UI preferences and historical date ranges

Keep sliders, filters, display options, secondary-panel expansion, forecast tabs, settings workspace tabs/sort, Day View dates/details and the routine-capacity target private to the signed-in account in this browser. Do not put personal UI choices in the shared practice-settings file. Restore both month-range handles atomically after data bounds are available, migrate legacy individual slider values, and keep a user's saved range when refreshed data temporarily narrows the available bounds. An account with no saved choice uses the original defaults.

Historical slider limits and their To labels use the latest imported Slot Report timestamp in Europe/London. Show just the refresh date for its partial day (for example, 27 Oct 2026), without End. Earlier complete months retain End of month labels. Missing report timestamps must not imply a complete period or fall back to the wall-clock date.


## Slate blue navigation

Use the slate blue (#3c6487) bar with 36 px desktop buttons and 4 px vertical bar padding. Retain the existing module icons at 21 px, use 14 px module names and 9 px IQ superscripts with a 1 px gap and no preceding whitespace. Use a white active module with blue text. Keep 44 px touch targets and the existing in-flow mobile dropdowns.


## Average capsules

Place small Average capsules inside the plot, along their dotted average lines. Use horizontal staggering and collision-aware nearby rows on crowded narrow plots, recalculating after Plotly redraw/resize. Use explicit HTML pill backgrounds with fully rounded ends, 10 px text and 4 px by 9 px padding. The background is 75% opaque, with a 2 px fully opaque strip exactly where the average line passes behind the pill. Keep category names in accessible labels. Do not show the category legend, the Year-on-year comparison subtitle or the Fine whiskers explanatory block. Retain year legends for overlaid historical years.


## Darwin responsive header

Use Darwin Insight IQ with IQ in superscript, retaining the original logo mark. The compact branding/account row is soft slate (#527595); the menu row stays slate blue (#3c6487). Display the saved personal name (falling back to the Microsoft name/email) above Administrator or View only. Show Last clinical system download with the date and time of the latest successfully imported Slot Report, formatted in Europe/London; do not substitute the refresh-check time or a Signposting timestamp. Keep refresh and account/sign-out actions available.

Above 1100 px keep the icon-led navigation expanded. At 1100 px and below use a Menu toggle with a 29 px settings cog at the right end of the same menu row. Open the main navigation as wrapping compact pills, with bounded dropdown submenus; close it on selection, outside click, Escape or crossing the breakpoint. Opening Menu must not change the selected page. Do not show a page-name label beside the hamburger. Keep name, role and timestamp visible on phones and preserve 44 px touch targets.


## Navigation pill spacing

Use fully rounded primary pills (999 px radius), 14 px desktop labels, 20 px icons, 40 px height and 8 px gaps. Inactive pills retain a subtle white outline; the selected pill uses a white fill. The expanded hamburger menu uses content-width pills rather than full-width rows. Its submenus use 12 px text, 32 px desktop minimum rows and 40 px touch minimum rows, with capped height and scrolling. Clamp dropdown positions to the navigation bounds when opened or resized.


While refreshing/updating, replace the clinical-download caption and timestamp with progress in that same header region, then restore the imported-file timestamp when finished. Keep refresh errors visible in the feedback region. Retain a small down-chevron immediately to the right of each primary dropdown label.


## Waiting-page recovery and download age

Waiting Time Ranges and Median Waiting Times use the shared capacityUrgencyCategory resolver for both direct and detailed-to-group mappings. Keep month-range helpers in global shared scope so either waiting page can render without first opening the historical chart. Primary appointment types remain available when secondary filters conflict or no data matches.

Under Last clinical system download, display Refreshed with hours and minutes when less than two hours old, otherwise whole hours. Above two hours, show Consider EMIS update in bold. Recalculate once a minute from the Slot Report timestamp, and hide the age/reminder during refresh progress. CSV replacement progress uses the same top header region and clears on successful completion or cancellation; upload/refresh errors remain visible.


## Urgent Forecast bank-holiday policy

Appointment Number Suggestions must not read the analytics day-of-week selection. Fit normal daily suggestions and weekly allocations using the selected historical window with bank-holiday weeks excluded. Use the separate before/after-bank-holiday models for holiday-adjacent suggestion dates, retaining all available adjacent-day history. Analytics Monday–Friday automatically exclude bank-holiday weeks; before/after-bank-holiday analytics automatically include them and use all available adjacent-day history. Remove the Ignore bank holiday weeks checkbox; obsolete saved values have no effect. Prediction cache keys exclude the analytics weekday.


## Shared chart layout memory

See [CHART_LAYOUT.md](CHART_LAYOUT.md) for the user-approved title-left, Display-options-right, timeline-below, full-width-chart layout. Historic Booking Activity uses this theme first. Preserve the flow-based options panel, tight toggle column, aligned dropdown/smoothing column, rounded open dropdown menus, non-overlapping tick labels and per-account preferences when applying it to other requested charts.


## Historical timeline and line finish

Match the capacity timeline dimensions: 30 px track container, 8 px rail and 24 px white-rimmed handles, without the oversized outlined container. Fill geometry must use the actual 24 px thumb diameter. Historical years alternate between neutral grey #fafafa and #f1f1f1 at full opacity; apply a subtle neutral shadow only to historical data-line paths, preserving average lines and uncertainty markers.

Historical reading grid: major horizontal guides #b0b0b0, lighter minor guides #e3e3e3, sparse quarterly vertical guides #e4e4e4 and 2 px axis lines. Keep the original 62 px left margin and Y-axis domain/anchor; do not enable axis automargins. Place the weekly/monthly appointment title beside the Y-axis, not above the chart.
