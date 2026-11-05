# Macro Tracker

Google Apps Script web app for tracking calories and protein. Deployed with clasp; `rootDir` is `src/`, so only files there reach Apps Script.

## Layout

- `src/Code.js`: server code (`doGet`, `getData`, `addEntry`, `deleteEntry`)
- `src/Index.html`: the web app UI
- `src/appsscript.json`: Apps Script manifest

## Data

Data lives in the **Macro Tracker** Google Sheet (`SHEET_ID` in `src/Code.js`), which has the tabs **Today**, **Stats** and **Log**.

`src/Code.js` reads and writes these tabs:

- **Stats**: settings, read from `B4` (calorie target), `B7` (protein) and `B8` (maintenance)
- **Log**: one row per entry, with a header in row 1. Columns are Date | Time | Food | Calories | Protein.

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
