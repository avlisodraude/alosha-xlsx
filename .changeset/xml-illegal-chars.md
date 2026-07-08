---
"@alosha/xlsx": patch
---

fix: sanitize characters illegal in XML 1.0 when writing cell text and attributes

`escapeXmlText`/`escapeXmlAttr` only escaped `& < > " '`. Cell text containing control characters
(e.g. `\x0B` vertical tab, common when pasting from PDFs), lone surrogate halves, or the
U+FFFE/U+FFFF noncharacters was written verbatim into `sharedStrings.xml`/inline strings and
attribute values, producing XML that parsers and Excel reject. These characters are now re-encoded
using Excel's own `_xHHHH_` convention (rather than stripped), and the reader decodes the same
token back, so values round-trip through our own writer/reader intact.
