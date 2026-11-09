# Morvarid Business Suite

React + TypeScript + Vite application for egg import batch management, container lookup, Excel import and PDF reporting.

## Requirements
- Node.js 20.19+ or 22.12+
- npm

## Run locally
```bash
npm install
npm run dev
```

## Build
```bash
npm run build
```

## Current implementation
- Responsive glassmorphism business workspace.
- Batch entry, editing and deletion with QAR cost/profit calculations.
- Container/invoice/batch lookup.
- Excel (.xlsx, .xls) and CSV preview/import.
- PDF batch register and financial report.
- Browser-local persistence using localStorage.

## Important limitations
- Data currently lives in the current browser profile. It is not yet a Google Sheets backup and will not automatically appear on another device.
- Google OAuth and live Sheets synchronization are not enabled. Configure a Google Cloud OAuth consent/client and a secure Apps Script or server integration before enabling account authorization or remote writes.
- The current Excel importer supports common Morvarid column headings. Review the preview and validate imported records before relying on them for accounting.
- Financial calculations use gross profit = sales - (purchase egg + shipping + driver + storage + damage + forklift - discount). Confirm this formula against the business's approved accounting rules before production use.

## Scripts
- `npm run dev`
- `npm run build`
- `npm run preview`
