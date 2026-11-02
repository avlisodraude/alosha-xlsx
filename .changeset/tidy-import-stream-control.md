---
"@alosha/xlsx": minor
---

Add opt-in structured import warnings and strict known-loss checks. Add streaming export cancellation, awaitable row production, and byte-based backpressure with a default 16 MiB hard output queue ceiling. Synchronous producers exceeding the ceiling now fail explicitly; use writeRows, addRowAsync, or commitAsync for large exports.
