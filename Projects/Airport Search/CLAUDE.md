# Airport Search - Project Guide

## Overview
نظام بحث المطار لمتابعة رحلات الحجاج - جزء من منظومة إكرام الضيف.
Web App مبني على Google Apps Script مع واجهة HTML عربية (RTL).

## Architecture
- **Runtime:** Google Apps Script V8
- **Database:** Google Sheets - Spreadsheet ID: `1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s`
- **Sheet Name:** `رحلة الحاج ` (مع مسافة في النهاية)
- **Deployment:** Web App - `ANYONE_ANONYMOUS` access, `USER_DEPLOYING` execution
- **Auth:** Session-based via CacheService (UUID tokens, 6-hour duration)
- **Data Caching:** Chunked storage (90KB chunks) for large datasets
- **Timezone:** Asia/Riyadh
- **Push tool:** `clasp` CLI v3.3.0

## Files
| File | Purpose |
|---|---|
| `Code.js` | Backend - auth, data fetching, filtering, search, Excel export |
| `Search.html` | Main search interface - filters, stats, flight tables, modals |
| `Login.html` | Login page with SHA-256 auth |
| `Index.html` | Airport selection (Madinah/Jeddah + hall selection) |
| `appsscript.json` | Config - Drive API v2, Asia/Riyadh timezone |

## Key Functions (Code.js)
- `doGet()` / `doPost()` - Web app entry points, routing
- `getAllData()` - Fetches all data from sheet with caching
- `quickSearch(query)` - Searches ALL data (not filtered by airport/hall), max 50 results
- `searchByPassport(query)` - Wrapper for quickSearch
- `exportToExcel(data, groupBy)` - Creates temp Spreadsheet, exports as xlsx, returns Drive link
- `formatSheet_()` - Batch formatting with setBackgrounds/setFontColors
- `buildSheetByDay_()` / `buildSheetByFlight_()` / `buildSheetByDestination_()` / `buildSheetByHotel_()` - Batch data writers
- `clearCache()` - Clears all cache keys including chunks

## Column Indices (COL object)
- `COL.NAME` = 7 (اسم الحاج)
- `COL.PASSPORT` = 8 (رقم الجواز)

## Recent Updates (2026-03-13)
### 1. Excel Export - Batch Write Optimization
- Converted all 4 `buildSheetBy*` functions from row-by-row `setValues` to batch `setValues` with 2D arrays
- Converted `formatSheet_` from individual cell color operations to batch `setBackgrounds()` and `setFontColors()`
- **Why:** Row-by-row writes caused race conditions with `SpreadsheetApp.flush()`, leading to table not scrolling down and save issues

### 2. Dead Code Cleanup
- Removed unused `loadDepartureCities()` function and its call in init

### 3. Filter Persistence (sessionStorage)
- Added `FILTER_IDS` array, `FILTER_STORAGE_KEY` per airport/hall
- `saveFilters()` called in `doSearch()` before API call
- `restoreFilters()` called in `setTodayDefault()` before defaulting to today
- `clearAllFilters()` also clears sessionStorage

### 4. Last Refresh Time Display
- `updateLastRefreshTime()` called in `onSearchSuccess` callback

### 5. Auto-Refresh Cache Clear
- `startAutoRefresh` now calls `google.script.run.clearCache()` before `doSearch()` in interval callback

## Deployment Notes
- After `clasp push`, must create **New Deployment** in Apps Script for changes to go live
- Old deployments keep serving old code

## Known Behavior
- `quickSearch` searches ALL data without airport/hall filtering (potential improvement)
- Data is cached server-side; auto-refresh clears cache before fetching
