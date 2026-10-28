# Approved analytics chart layout

This is the user-approved shared layout for Darwin Insight IQ charts. Apply it to additional charts when requested; keep their own controls and calculations.

- Title at top left and a compact **Display options** button at top right.
- Full-width date/timeline range immediately below the heading, then the chart using the available width. Do not move controls into a permanent right sidebar.
- Options start collapsed. Opening them inserts a panel between the heading and range; the panel pushes the range/chart down instead of covering them.
- Toggles form a tight, aligned left column. Dropdowns and a short smoothing slider form the right column. At 620 px and below stack these groups; never shrink controls into overlapping columns.
- Dropdown **opened menus**, rather than just their closed buttons, have rounded corners and rounded option rows. Preserve visible focus and keyboard operation.
- Smoothing labels align to the four thumb positions, including half-thumb insets. A short track has its own row height for tick labels; do not stretch it across the panel.
- Keep the date range, control panel and plotting area in separate flow regions, with explicit spacing. Retain the existing graph height and maximise plot width. Allow top margin for year legends only when required.

## Components

The styles in `index.html` use `.analytics-chart-layout`, `.analytics-chart-heading`, `.chart-display-options-button`, `.chart-display-options-panel`, `.chart-options-grid`, `.chart-options-toggles` and `.chart-options-fields`. A toggle button's `aria-controls` references the panel ID; its `aria-expanded` matches panel visibility. `initializeChartDisplayOptions()` binds these components idempotently. Open/closed state is saved privately per signed-in account and page.

A native select marked `data-chart-select` uses the shared rounded listbox menu while retaining its normal change events, values and saved preference. The current historic Week/Month and Timeline selects instead use `data-chart-proxy` and `data-checked-value` to adapt existing checkbox models; they deliberately do not create a second saved preference. Keep those original models and handlers intact.

Reusable sizing: title 21 px; control labels 12 px; options grid up to 650 px; 22 px between columns; 23 px fine-pointer toggle rows, 32 px on touch; dropdown menus 13 px radius and option rows 9 px radius. Current historical smoothing uses a 14 px thumb and 7 px tick-label inset. Native keyboard controls, Escape dismissal and outside-click dismissal remain available.

Historic Booking Activity is the first implementation. Its controls are Include Bank Holiday Weeks, Adjust for List Size, Show current week/month, Combine selected lines, Data, Timeline and Smoothing. Other charts can have different options while retaining this layout.
