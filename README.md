# PharmaAlloc — Pharmaceutical Inventory Allocation System

A modern, single-page web application that replicates the full logic of the `Pharma_Allocation_Based.xlsx` Excel template. Designed for inventory managers, pharmacists, and supply chain teams.

---

## ✨ Features

| Feature | Details |
|---|---|
| **Data Input** | Upload CSV/Excel or paste data for Branch CDSS, Central Warehouse, and National Branches |
| **Allocation Calculations** | Full replication of all Excel formulas (Branch Remaining Need, National Remaining Need, Recommended Allocation, Allocation %, Status, Overstock Flag, Comments, Suggested Redistribution) |
| **Interactive Dashboard** | 8 KPI cards, Allocation Status donut chart, Top 10 Urgent Items bar chart, Top 10 Urgent/Overstock tables |
| **Allocation Report** | Full sortable/filterable table with all 18 columns |
| **Material Code Editor** | Add/edit/remove codes manually, paste a list, or auto-detect from loaded data |
| **Search & Filter** | By material code, description, status, type (ZME/ZMS), overstock flag |
| **Export** | Export Allocation Report as CSV or Excel (.xlsx) with a KPI summary sheet |
| **Dark Mode** | Toggle-able, persists across sessions |
| **Persistence** | Data survives page refresh (stored in localStorage) |
| **Responsive** | Works on desktop, tablet, and mobile |

---

## 📐 Formula Reference

All calculations mirror the Excel formulas exactly:

| Column | Formula |
|---|---|
| Description | `VLOOKUP(code, Branch_CDSS, col2)` |
| Material Type | `VLOOKUP(code, Branch_CDSS, col4)` |
| Branch SOH | `VLOOKUP(code, Branch_CDSS, col3)` |
| Central SOH | `VLOOKUP(code, Central_Warehouse, col3)` |
| National SOH | `VLOOKUP(code, National_Branches, col3)` |
| AMC | `VLOOKUP(code, Central_Warehouse, col5)` |
| MOS Central | `ROUND(Central_SOH / AMC, 1)` |
| Branch Forecast | `VLOOKUP(code, Branch_CDSS, col7)` |
| Delivered Qty | `VLOOKUP(code, Branch_CDSS, col8)` |
| **Branch Remaining Need** | `Forecast − (Branch_SOH + Delivered)` |
| **National Remaining Need** | `NatForecast − (NatSOH + NatDelivered − CentralSOH)` |
| **Recommended Allocation** | `IF(BranchNeed≤0, 0, IF(NatNeed≤0, MIN(BranchNeed,CentralSOH), MAX(0, MIN(BranchNeed, CentralSOH, ROUND(BranchNeed/ABS(NatNeed)×CentralSOH, 0))))` |
| **Allocation %** | `RecommendedAllocation / BranchForecast` |
| **Allocation Status** | Critical Shortage / No Need / Full / Rationed / Partial |
| **Overstock Flag** | `OVERSTOCK` if MOS > 9, else `OK` |
| Comments | Automated human-readable comment |
| Suggested Redistribution | `ROUND(BranchForecast / NatForecast × CentralSOH, 0)` (Overstock only) |

---

## 🚀 Deployment to GitHub Pages

### Option 1: Direct Upload (simplest)

1. Create a new GitHub repository (e.g. `pharma-alloc`)
2. Upload all three files: `index.html`, `style.css`, `script.js`
3. Go to **Settings → Pages**
4. Under **Source**, select `main` branch and `/ (root)` folder
5. Click **Save**
6. Your app is live at `https://YOUR_USERNAME.github.io/pharma-alloc/`

### Option 2: Git CLI

```bash
git init
git add .
git commit -m "Initial commit — PharmaAlloc"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/pharma-alloc.git
git push -u origin main
```

Then enable GitHub Pages as above.

### Option 3: GitHub Desktop

1. Open GitHub Desktop → File → New Repository
2. Copy all files into the local folder
3. Commit and push
4. Enable Pages in repository settings

---

## 📁 File Structure

```
pharma-alloc/
├── index.html      # Single-page app markup, navigation, tab structure
├── style.css       # All styles, CSS variables, dark mode, responsive rules
├── script.js       # All logic: calculations, rendering, charts, export
└── README.md       # This file
```

---

## 📋 CSV Data Format

All three source sheets use the same 8-column format:

| Column | Field |
|---|---|
| 1 | Material Code |
| 2 | Description |
| 3 | Stock on Hand |
| 4 | Material Type Code (ZME / ZMS) |
| 5 | AMC (Average Monthly Consumption) |
| 6 | MOS (Months of Stock) |
| 7 | Forecast Qty |
| 8 | Delivered Qty |

The header row is auto-detected and skipped. Column order must match exactly.

---

## 🛠 Extending the App

- **Add new KPIs**: edit `updateDashboard()` in `script.js`
- **Add new chart**: add a `<canvas>` in `index.html`, initialize in `initCharts()`, update in `updateDashboard()`
- **Change allocation logic**: edit `calcRow()` in `script.js`
- **Add new columns**: add to `calcRow()` return object, add `<th>/<td>` in HTML, add to `exportAsExcel()`
- **Change colors**: edit CSS variables at the top of `style.css`

---

## 🔒 Privacy

All data is processed **entirely in the browser**. Nothing is sent to any server. The app works fully offline once loaded.

---

## 📦 Dependencies (CDN, no install needed)

| Library | Purpose |
|---|---|
| Chart.js 4.4 | Dashboard charts |
| PapaParse 5.4 | CSV parsing |
| SheetJS (xlsx) | Excel import/export |
| Google Fonts (DM Sans / DM Mono) | Typography |

All loaded from CDN — no build step required.
