# UI conventions

## Mini toggles

Match the compact Clinician Groups sidebar switches: approximately 32.4 × 18 CSS px, with a 14 px thumb and 12 px labels. Use actual dimensions, no scale transform or negative margin. Arrange related controls horizontally beneath their slider. Use 4 px between switch and label, 12 px between the separate 'Show in:' prompt and 'Days', and 32 px between the Days/Weeks control and Group unavailable appts. Allow whole controls to wrap on narrow screens. Preserve accessible labels and keyboard focus. Reuse `.mini-toggle-row` in `index.html`.


## Historic Booking Activity smoothing

Use a compact discrete slider in the chart display-options row, beside Combine selected lines. Label its four stops Off, Low, Medium and High, without repeating the active value beside Smoothing. Retain native range keyboard behaviour, visible focus and an accessible value label. Reuse the existing blue rail/white-rimmed thumb styling.

Keep 6 px between the Smoothing title and the bar. Add 24 px of separation before the smoothing group in addition to the toolbar gap, wrapping cleanly on narrow screens.


## Current-period range bar

Use a translucent capped vertical range bar in the series colour to show an indicative final-total range for the incomplete week/month. Keep completed historical lines solid and use a fine dotted connector to the predicted finish. Keep the report timestamp and indicative prediction coverage in hover details. Hover details distinguish actual bookings so far from estimated finishing bounds. If the report is stale or historical coverage is insufficient, explain why no range can be shown. The show-current control owns this visibility.

The range bar has no permanent numeric or upper/lower text labels. Keep figures in hover details. Smoothing uses a two-row layout: title centred with the track, four tick labels below. Align each tick centre with the thumb travel, including the half-thumb inset at both ends.


## Fine whisker with current-period circle

Use a 1.1 px whisker in the series colour at 65% opacity, short 3 px caps, and a 7 px open circle in the same series colour for the observed count. Show no permanent numbers beside either marker. Keep each series pair in one column and keep all series at the same current-period x coordinate; overlap is allowed. Recompute spacing on resize. Distinguish observation and estimate in hover details.

Smoothing must share the neighbouring toolbar label font family, size, weight and line height. Keep the label/track row centred with the toggle labels, reserving tick-label space below it.


## Current-period continuity

Use the historical booking-date/month axis for current observations and finishing ranges. Join the last completed historical point to the finishing estimate with a fine 1.35 px dotted segment (1 px dots, 3 px gaps). Keep its endpoint, observed circle, whisker and caps at one shared x position. All current-period series share the same booking date/month position; whiskers may overlap. Use 1.1 px whiskers at 65% opacity with 3 px caps and indicative 70% prediction coverage. Keep the full historical axis and mean lines continuous; do not reserve a separate prediction lane. Current-year line width is 2.75 px; earlier-year lines are thinner and more transparent.

## Secondary toggle-box expansion

Save an explicit expand/collapse choice before changing the DOM. Restore saved expansion state once per account/state, without replaying button clicks after every render. A minus click must hide the filter stack, show its selection summary, change to plus and remain collapsed after redraw on every analytics page.

## Personal UI preferences and historical date ranges

Keep sliders, filters, display options, secondary-panel expansion, forecast tabs, settings workspace tabs/sort, Day View dates/details and the routine-capacity target private to the signed-in account in this browser. Do not put personal UI choices in the shared practice-settings file. Restore both month-range handles atomically after data bounds are available, migrate legacy individual slider values, and keep a user's saved range when refreshed data temporarily narrows the available bounds. An account with no saved choice uses the original defaults.

Historical slider limits and their To labels use the latest imported Slot Report timestamp in Europe/London. Show just the refresh date for its partial day (for example, 27 Oct 2026), without End. Earlier complete months retain End of month labels. Missing report timestamps must not imply a complete period or fall back to the wall-clock date.


## Slate blue navigation

Use the slate blue (#3c6487) bar with 36 px desktop buttons and 4 px vertical bar padding. Retain the existing module icons at 21 px, use 14 px module names and 9 px IQ superscripts with a 1 px gap and no preceding whitespace. Use a white active module with blue text. Keep 44 px touch targets and the existing in-flow mobile dropdowns.


## Average capsules

Place small Average capsules inside the plot, along their dotted average lines. Use horizontal staggering and collision-aware nearby rows on crowded narrow plots, recalculating after Plotly redraw/resize. Use explicit HTML pill backgrounds with fully rounded ends, 10 px text and 4 px by 9 px padding. The background is 75% opaque, with a 2 px fully opaque strip exactly where the average line passes behind the pill. Keep category names in accessible labels. Do not show the category legend, the Year-on-year comparison subtitle or the Fine whiskers explanatory block. Retain year legends for overlaid historical years.
