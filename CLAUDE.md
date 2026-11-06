# Macro Tracker

Google Apps Script web app for tracking calories and protein. Deployed with clasp; `rootDir` is `src/`, so only files there reach Apps Script.

## Layout

- `src/Code.js`: server code (`doGet`, `getData`, `addEntry`, `deleteEntry`, `addWeight`, `deleteWeight`)
- `src/Index.html`: the web app UI
- `src/appsscript.json`: Apps Script manifest

## Data

Data lives in the **Macro Tracker** Google Sheet (`SHEET_ID` in `src/Code.js`), which has the tabs **Today**, **Stats**, **Log** and **Weight**.

`src/Code.js` reads and writes these tabs:

- **Stats**: settings, read from `B4` (calorie target), `B7` (protein) and `B8` (maintenance)
- **Log**: one row per entry, with a header in row 1. Columns are Date | Time | Food | Calories | Protein.
- **Weight**: one row per weigh-in in `A:B` (Date | Weight), with a header in row 1. `E:F` is a read-only block of Adaptive TDEE formulas in the same rows, read from `F5:F12`.

Only ever write to `A:B` on **Weight**. Never call `deleteRow()` there, because it would remove part of the `E:F` block. `deleteWeight` moves the rows below up by value instead.

It doesn't reference **Today**.

## After every change

Run these in order:

```sh
git add <changed files>
git commit -m "<summary>"
git push
clasp push --force
clasp create-deployment --deploymentId AKfycbyl5JMrX9YZwGrEsDz9qj1rVhRXNPkT1gjlBhh6mZsDv8y40AwSHW5BgSFLhFM1AQlA --description "<summary>"
```

**Never deploy without `--deploymentId`.** Doing so creates a new deployment, which changes the web app URL.
