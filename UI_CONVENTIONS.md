# UI conventions

## Mini toggles

Match the compact Clinician Groups sidebar switches: approximately 32.4 × 18 CSS px, with a 14 px thumb and 12 px labels. Use actual dimensions, no scale transform or negative margin. Arrange related controls horizontally beneath their slider. Use 4 px between switch and label, 12 px between the separate 'Show in:' prompt and 'Days', and 32 px between the Days/Weeks control and Group unavailable appts. Allow whole controls to wrap on narrow screens. Preserve accessible labels and keyboard focus. Reuse `.mini-toggle-row` in `index.html`.


## Historic Booking Activity smoothing

Use a compact discrete slider in the chart display-options row, beside Combine selected lines. Label its four stops Off, Low, Medium and High, without repeating the active value beside Smoothing. Retain native range keyboard behaviour, visible focus and an accessible value label. Reuse the existing blue rail/white-rimmed thumb styling.

Keep 6 px between the Smoothing title and the bar. Add 24 px of separation before the smoothing group in addition to the toolbar gap, wrapping cleanly on narrow screens.


## Current-period range bar

Use a translucent capped vertical range bar in the series colour to show an indicative final-total range for the incomplete week/month. Keep historical lines solid, draw no precise projected endpoint, and label the shaded area and report timestamp next to the chart. Hover details distinguish actual bookings so far from estimated finishing bounds. If the report is stale or historical coverage is insufficient, explain why no range can be shown. The show-current control owns this visibility.

The range bar has no permanent numeric or upper/lower text labels. Keep figures in hover details. Smoothing uses a two-row layout: title centred with the track, four tick labels below. Align each tick centre with the thumb travel, including the half-thumb inset at both ends.


## Fine whisker with current-period circle

Use a 1.2 px whisker in the series colour at 40% opacity, short 4 px caps, and an 8 px open circle in the same series colour for the observed count. Show no permanent numbers beside either marker. Keep each series pair in one column and offset separate pairs horizontally by up to 6 px where space permits; clamp the group inside the plot. Recompute spacing on resize. Distinguish observation and estimate in hover details.

Smoothing must share the neighbouring toolbar label font family, size, weight and line height. Keep the label/track row centred with the toggle labels, reserving tick-label space below it.


## Current-period continuity

Use the historical booking-date/month axis for current observations and finishing ranges. Join the last completed historical point to the finishing estimate with a dashed segment. Keep its endpoint, observed circle, whisker and caps at one shared x position. For multiple current-period series, use a small horizontal nudge (up to 6 px between columns, 24 px overall), recomputed on resize. Keep the full historical axis and mean lines continuous; do not reserve a separate prediction lane. Current-year line width is 2.75 px; earlier-year lines are thinner and more transparent.

## Secondary toggle-box expansion

Save an explicit expand/collapse choice before changing the DOM. Restore saved expansion state once per account/state, without replaying button clicks after every render. A minus click must hide the filter stack, show its selection summary, change to plus and remain collapsed after redraw on every analytics page.
