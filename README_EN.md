# Shift Priority — Full English Guide

Version 5.38 — September 2026
Author: Omid Farhadnia
License: MIT (open source — free to use, modify and share; keep author credit)

---

## What this app does

Built for TOK Transit drivers, this tool reads a crew shift schedule Excel
file and ranks every crew (group) against your personal preferences.

How it works:
1. Upload the shift Excel file (e.g. `Mobility Runs_Sep_06_2026.xlsx`).
2. The app auto-detects the Crew #, Type (4DAY/5DAY), AM/PM shift, the seven
   weekday columns, and the total-hours column.
3. Each day's work region (Newmarket, Richmond Hill, Caldari, Stouffville,
   Maple, or standby/RPT) is detected from the Excel cell's fill color plus
   the STF/MRG/RPT text codes in the run number.
4. You build a personal priority list: e.g. "PM shift matters most, then
   4-day weeks, then Newmarket, then having a standby day." Up to 10
   criteria, in any order you choose.
5. Ranking is a strict priority chain (lexicographic sort): crews are first
   sorted by criterion #1; ties are broken by criterion #2, then #3, and so
   on down the list — never a blended average.
6. Results: the top 5 matches get gold/silver/bronze medal cards, the rest
   are listed below, each showing the exact hours worked per day per region.
7. You can select multiple crews to compare side by side, and export the
   report to print/PDF or Excel.

## Supported languages
Persian, English, Hindi — switch from the top of the screen.

## What each system needs to run it

### Windows — two options
- **Quick mode (no permanent install):** just needs Node.js (`run.bat` will
  guide you through installing it if missing).
- **Real .exe installer:** run `build-exe.bat`; the first run needs Node.js
  and some libraries, but the final installer (inside the `release` folder)
  can be installed on any Windows PC with no Node.js required there.

### iPhone (iOS)
Apple does not allow installing a raw app file (.ipa) without an Apple
Developer account ($99/year) plus a Mac and Xcode. The practical solution is
the PWA: open the app in Safari, tap Share → "Add to Home Screen." Internet
is only needed the first time; after that it works offline too.

### Android
- Simple: same PWA method (Chrome → "Add to Home screen").
- Advanced (real .apk): requires Android Studio on a computer (free, but a
  few GB download). Exact commands are in `README.txt` in this folder.

### General requirements
- Node.js version 18 or newer (to run/build on a computer).
- Libraries: React, Vite, xlsx (SheetJS), lucide-react — all installed
  automatically via `npm install`.
- For the Windows .exe: Electron and electron-builder (also via npm).
- Internet is only needed once, to install the libraries. After that the
  app runs fully offline — no data is ever sent to a server; everything is
  processed locally on your own computer or phone.

## Project file structure
```
shift-priority-app/
├── src/App.jsx        full application source code (React)
├── src/main.jsx         React entry point
├── index.html            main HTML page + PWA settings
├── public/                icons, manifest.json, service worker
├── electron/main.js     Windows desktop entry point
├── package.json           dependency list and scripts
├── run.bat                  quick run (dev mode)
├── build-exe.bat         builds the Windows installer
├── LICENSE                  MIT license text
├── README.txt              technical per-platform install guide
└── README_FA.md / README_EN.md   this guide, in both languages
```

## Privacy
No data (the Excel file, your priorities, or the results) is ever sent to
any server; everything runs inside your own browser/app.

---
If you hit any issue, keep the exact error text and ask.
