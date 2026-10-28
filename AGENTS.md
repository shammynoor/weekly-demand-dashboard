# Chart layout guidance

Before changing an analytics chart layout, read [CHART_LAYOUT.md](CHART_LAYOUT.md) and [UI_CONVENTIONS.md](UI_CONVENTIONS.md). They record the user's approved reusable layout and control theme. Apply that layout only to charts the user asks to migrate.

Run the relevant regression checks with `node --test tests/*.test.cjs`. Preserve existing calculations, control IDs and per-account UI preferences when rearranging a chart.
