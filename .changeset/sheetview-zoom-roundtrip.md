---
"@alosha/xlsx": patch
---

fix: round-trip sheet view zoom scale

`renderSheetViews` (`src/xlsx/parts/worksheet.ts`) only emitted `showGridLines` and `<pane>` on the
`<sheetView>` element, silently dropping `zoomScale`; `parseSheetViews`
(`src/xlsx/read/worksheet.ts`) only read those two back. A view set to `{ zoomScale: 150 }`
round-tripped to an empty `views` array, contradicting the README's claim that views (including
zoom) round-trip. `renderSheetViews` now emits `zoomScale`/`zoomScaleNormal` (mirrored, since we
don't track a separate normal-view zoom) whenever the zoom isn't the default 100, in both the
frozen/split and plain branches; `parseSheetViews` reads `zoomScale` back, coercing to a number and
dropping it (along with the rest of the view, if otherwise empty) when it's the default 100. The
streaming writer's prologue calls the same `renderSheetViews`, so it inherits the fix without
changes.
