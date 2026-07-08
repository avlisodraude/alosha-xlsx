---
"@alosha/xlsx": patch
---

fix: guard numeric character reference decoding against RangeError on out-of-range code points

`decodeEntities` in `src/xml/xml-reader.ts` guarded `&#nn;`/`&#xhh;` numeric character references
with `Number.isFinite` before calling `String.fromCodePoint`, but that function throws `RangeError`
for code points beyond `0x10FFFF` or lone surrogate halves (`0xD800`-`0xDFFF`) — neither of which
`Number.isFinite` rejects. A corrupt or crafted `.xlsx` part (e.g. `&#x110000;` in
`sharedStrings.xml`) crashed `readWorkbookBuffer` with an untyped `RangeError`, bypassing the typed
error contract in `src/errors.ts`. `decodeEntities` now validates the code point range before
converting and falls back to the original match text when it's invalid, the same fallback already
used for unknown named entities.
