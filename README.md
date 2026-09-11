# SheetUtils

The helpers every Google Apps Script project ends up rewriting, written once, batched and
safe on large sheets. One file, no dependencies, no library ID to add.

```js
SheetUtils.readObjects('Orders')                  // -> [{Email:'a@b.com', Amount: 12}, ...]
SheetUtils.writeObjects('Clean', rows)            // headers + rows in one call
SheetUtils.dedupe('Leads', ['Email'])             // remove duplicate rows by key columns
SheetUtils.mergeSheets(['Jan','Feb','Mar'], 'All')// stack sheets with the same headers
SheetUtils.upsert('Customers', 'Email', rows)     // update matching rows, append new ones
SheetUtils.withLock(function(){ ... })            // stop two triggers colliding
SheetUtils.retry(function(){ return UrlFetchApp.fetch(url); }, 3)
```

## Install

Extensions → Apps Script → add a file called `SheetUtils.gs`, paste
[SheetUtils.gs](SheetUtils.gs) in, save. That's it.

## Why it exists

Three things break most Apps Script code once a sheet gets big:

- **Cell-by-cell reads and writes.** `getRange().getValue()` in a loop is the single most
  common reason a script that worked on 200 rows dies on 20,000. Everything here reads and
  writes in one batched call.
- **Concurrent triggers.** Two triggers firing at once will happily corrupt each other's
  work. `withLock` serialises them.
- **Flaky external calls.** `retry` does exponential backoff instead of failing the run.

## Licence

Free for personal and commercial use, including client work. No attribution required.

---

If you want the rest of what I use — a batch runner that checkpoints and re-schedules
itself past the 6-minute limit, a quota-aware Gmail mail merge, a Gmail-to-Sheet logger,
a Forms-to-PDF pipeline and a REST-to-Sheet loader — they're in a paid toolkit here:
https://ishubham8.gumroad.com/l/wfxzs

Questions or a bug: open an issue.
