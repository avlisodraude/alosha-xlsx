---
"@alosha/xlsx": patch
---

fix: escape embedded double quotes in conditional-formatting text-rule formulae

`containsText`/`notContainsText`/`beginsWith`/`endsWith` conditional-formatting rules generate an
Excel `<formula>` that interpolates the rule's `text` into a string literal (e.g.
`SEARCH("...",A1)`). A `text` value containing a `"` was interpolated unescaped, producing an
invalid formula that made Excel prompt to repair the file on open. Embedded quotes are now doubled
(`"` → `""`) per Excel string-literal syntax. The `text=` XML attribute was already escaped
correctly and is unaffected.
