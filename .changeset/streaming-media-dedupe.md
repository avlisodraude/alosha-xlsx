---
"@alosha/xlsx": patch
---

fix: dedupe repeated workbook media in the streaming writer

`WorkbookWriter` wrote duplicate ZIP entries when the same workbook-registered image was placed on
multiple sheets. Repro: `wb.addImage` once, `addImage` the returned id on two sheets, each
`ws.commit()` independently pushes the same `xl/media/imageN.ext` `ExtraPart` — `WorkbookWriter.addExtraPart`
wrote it into the archive unconditionally every time, producing two ZIP local-file-headers (and two
central-directory records) for one part name, an OPC violation. The buffered writer never had this bug:
`collectWorksheetParts` gathers parts into a `Map` keyed by path, so a repeat silently overwrites the
same key. `addExtraPart` now tracks written part paths in a `Set` and skips a repeat — its content is
identical by construction (same `imageId` resolves to the same registry buffer).
