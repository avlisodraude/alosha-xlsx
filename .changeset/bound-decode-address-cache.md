---
"@alosha/xlsx": patch
---

Bound the address-decode memo so reading large or many workbooks no longer leaks memory.

`decodeAddress`'s cache is keyed by full cell address, so distinct cells (up to 16,384 × 1,048,576)
grew a process-global `Map` without limit — reading big or numerous workbooks interned every address
forever. It is now capped with FIFO eviction (`DECODE_ADDRESS_CACHE_MAX`), keeping the hot-path
memoization for realistic working sets while bounding memory. The column-letter/number caches are
unchanged (already bounded by Excel's 16,384-column cap). No public API or output change.
