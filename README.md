# Macro Tracker App

A dark-themed calorie and protein tracker that runs as a Google Apps Script web app. All data lives in a Google Sheet, so there's nothing to host.

## Features
- **Today:** calorie ring split by food, calories left, protein progress
- **Calendar:** monthly view showing target hits and misses
- **Log:** every entry by day, with delete and a link to the source sheet
- **Quick add:** log food from any page

## Stack
Google Apps Script, Google Sheets, plain HTML/CSS/JS, deployed with [clasp](https://github.com/google/clasp).

## Deploy
```
clasp push --force
clasp create-deployment --deploymentId <DEPLOYMENT_ID>
```
