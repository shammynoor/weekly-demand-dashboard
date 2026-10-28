# UI conventions

## Mini toggles

Match the compact Clinician Groups sidebar switches: approximately 32.4 × 18 CSS px, with a 14 px thumb and 12 px labels. Use actual dimensions, no scale transform or negative margin. Arrange related controls horizontally beneath their slider. Use 4 px between switch and label, 12 px between the separate 'Show in:' prompt and 'Days', and 32 px between the Days/Weeks control and Group unavailable appts. Allow whole controls to wrap on narrow screens. Preserve accessible labels and keyboard focus. Reuse `.mini-toggle-row` in `index.html`.


## Historic Booking Activity smoothing

Use a compact discrete slider in the chart display-options row, beside Combine selected lines. Label its four stops Off, Low, Medium and High, without repeating the active value beside Smoothing. Retain native range keyboard behaviour, visible focus and an accessible value label. Reuse the existing blue rail/white-rimmed thumb styling.

Keep 6 px between the Smoothing title and the bar. Add 24 px of separation before the smoothing group in addition to the toolbar gap, wrapping cleanly on narrow screens.
