# UI conventions

## Mini toggles

The user's final mini-toggle reference is the compact Clinician Groups sidebar switch (GP / Physio / Mental Health screenshot), not the larger 60 × 36 px switch. Match its displayed size: approximately 32.4 × 18 CSS px, with a 14 px thumb and 12 px labels. Use actual dimensions with no scale transform or negative margin. Keep related controls side by side below the associated slider, with clear switch-to-label spacing; allow whole controls to wrap on narrow screens. Preserve accessible labels and keyboard focus. Reuse `.mini-toggle-row` in `index.html` for future mini-toggle requests.
