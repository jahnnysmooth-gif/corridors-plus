# Corridor's Plus ➕ — Project Plan

## Project Summary

A mobile-first, single-file HTML app built for a NYC hallway contractor to track man-hours by work item and floor. No internet connection required. Runs in any browser or as a PWA on an iPhone home screen.

---

## What Has Been Built

### Core App Structure
- [x] Single-file HTML app (`hallway-tracker.html`) — no build step, no dependencies
- [x] `localStorage` persistence (key: `hallway_tracker_v2`)
- [x] PWA meta tags for iOS "Add to Home Screen"
- [x] Three-screen navigation: Floors → Floor → Item
- [x] Header with back button and Export button (visible on all screens)
- [x] Branding: **Corridor's Plus ➕** with white plus sign

### Floors Screen
- [x] Summary stats: total floors, total man-hours, total items
- [x] Floor cards showing name, start date, estimated end date (start + 5 weeks), man-hours, and item progress
- [x] Add Floor with name and start date
- [x] Delete Floor
- [x] Global "Log Today" — log hours across all active floors in one modal
- [x] Backup data (downloads `.json`)
- [x] Restore data (uploads `.json`)

### Floor Screen
- [x] Color-coded 3-column tile grid of work items
- [x] Item status indicators: Not Started (grey) / In Progress (amber) / Done (green)
- [x] Progress bar showing done items vs total
- [x] Add custom work item
- [x] Floor-level "Log Today" — log hours for all items on this floor
- [x] Delete floor button

### Item Screen
- [x] Status toggle buttons (Not Started / In Progress / Done)
- [x] Stats: total man-hours, total sessions
- [x] Inline quick log (no popup) with date chips (Today / Yesterday / Calendar)
- [x] Separate Skilled and Laborer counts, each with their own hours field
- [x] Live man-hours preview before submitting
- [x] History table: Date | Skilled | Skld Hrs | Laborer | Lab Hrs | Man-Hrs | Delete
- [x] Delete individual log entries
- [x] Delete item button

### Default Work Items (auto-added to every new floor)
1. Demo
2. Ceiling Skim Coat
3. Wall Skim Coat
4. Wall Prime
5. Ceiling Prime
6. Ceiling Paint
7. Crown/Raceway
8. Wallpaper
9. Baseboard Installation

### Labor Tracking
- [x] Skilled workers tracked separately from Laborers
- [x] Each type has its own hours field (skilled may work 10 hrs, laborers 8 hrs on same day)
- [x] Man-hours formula: `(skilled × skilledHours) + (laborer × laborerHours)`
- [x] Full backward compatibility with older entry formats

### Export
- [x] Floor selector: specific floor or **All Floors**
- [x] Period: Daily / Weekly (Mon–Sun) / Monthly
- [x] **PDF** — opens browser print dialog (Save as PDF)
  - Single floor: item breakdown + grand total
  - All Floors: each floor is a named section with its own total + overall grand total, page break between floors
- [x] **CSV** — downloads spreadsheet file
  - Columns: FLOOR, ITEM, DATE, SKILLED, SKILLED HRS, LABORER, LABORER HRS, MAN-HOURS
  - Per-item TOTAL rows in all caps
  - Works for single floor or all floors in one file

---

## Potential Future Improvements

### Reporting
- [ ] Weekly summary view on the Floors screen — show total man-hours logged this week across all floors
- [ ] Per-item bar chart on the floor screen showing man-hours distribution
- [ ] Running cost estimate (if hourly rates for skilled vs laborer are added)

### Worker Management
- [ ] Add named workers instead of just counts (e.g., "Mike — Skilled")
- [ ] Track which workers were on each floor each day
- [ ] Worker-level summary report

### Floor Management
- [ ] Mark a floor as complete (archive it but keep data)
- [ ] Reorder floors manually (drag or up/down arrows)
- [ ] Floor notes field for job-site observations

### Work Items
- [ ] Reorder items on a floor
- [ ] Mark an item's estimated hours (target vs actual comparison)
- [ ] Duplicate items across floors

### Usability
- [ ] Search / filter items on the floor screen
- [ ] Swipe to delete log entries on mobile
- [ ] Undo last log entry
- [ ] Dark mode

### Data & Sync
- [ ] Google Drive or iCloud sync so multiple people (foreman, office) see the same data
- [ ] Export all floors, all time (no date range filter) in one tap
- [ ] Auto-backup on a schedule

---

## Technical Notes

- All data lives in the browser's `localStorage` — clearing browser data will erase it. **Always back up before switching phones or browsers.**
- The backup `.json` file is the only way to transfer data between devices currently.
- The app is a single file — to update it, replace the file and restore from backup if needed.
- Three legacy log entry formats exist in old data; `entryManHours()` and `entryDisplay()` handle all of them transparently.
