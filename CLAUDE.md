# Hallway Tracker — CLAUDE.md

## Project Overview

Single-file HTML app for a NYC hallway contractor to track man-hours per work item per floor. No build step, no dependencies, no server — open `hallway-tracker.html` directly in a browser or add to iPhone home screen (PWA).

## File

`hallway-tracker.html` — everything lives here: HTML, CSS, and vanilla JS in one file.

## Data Model

Stored in `localStorage` under the key `hallway_tracker_v2`.

```js
state = {
  floors: [{
    id,           // uid()
    name,         // e.g. "Floor 3"
    startDate,    // YYYY-MM-DD
    items: [{
      id,
      name,
      status,     // 'pending' | 'active' | 'done'
      hoursLog: [{
        date,          // YYYY-MM-DD
        skilled,       // number of skilled workers
        skilledHours,  // hours worked by skilled
        laborer,       // number of laborers
        laborerHours   // hours worked by laborers
      }]
    }]
  }]
}
```

### Backward compatibility

Old entries may use different shapes. Always use these helpers — never read `entry.hours` or `entry.workers` directly:

- `entryManHours(e)` — returns total man-hours for an entry (handles all 3 legacy formats)
- `entryDisplay(e)` — returns `{ skilled, skilledHours, laborer, laborerHours }` normalized for display
- `itemManHours(item)` — sums all entries for an item

**Three legacy formats in the wild:**
1. `{ workers, hours }` — original, single worker count with shared hours
2. `{ skilled, laborer, hours }` — first split, shared hours between types
3. `{ skilled, skilledHours, laborer, laborerHours }` — current format

## Man-Hours Formula

`man-hours = (skilled × skilledHours) + (laborer × laborerHours)`

## Navigation / Screens

Three screens, driven by `view = { screen, floorId, itemId }`:

| Screen | Description |
|--------|-------------|
| `floors` | Home — floor cards + global stats + Backup/Restore |
| `floor` | Tile grid of work items for one floor |
| `item` | Quick log form + history table for one item |

Navigate with `navigate(screen, floorId?, itemId?)`. Always call `render()` after state changes.

## Default Work Items (per new floor)

Demo, Ceiling Skim Coat, Wall Skim Coat, Wall Prime, Ceiling Prime, Ceiling Paint, Crown/Raceway, Wallpaper, Baseboard Installation

## Key Constants / Config

- `DEFAULT_ITEMS` — array of default item names added to every new floor
- `TILE_COLORS` — 11 color pairs cycling by item index on the floor tile grid
- Floor duration shown on cards: **5 weeks** from start date (`addWeeks(startDate, 5)`)
- `persist()` — call after every state mutation to save to localStorage
- `uid()` — generates unique IDs

## Date Handling

- **Stored:** `YYYY-MM-DD` (ISO)
- **Displayed:** `MM-DD-YYYY` via `displayDate(iso)`
- **PDF/export short form:** `MM/DD/YY` via `shortDate(iso)`
- Never store or display dates in any other format

## Item Status

- `pending` — grey dot / "Not Started"
- `active` — amber dot / "In Progress" (auto-set on first log entry)
- `done` — green dot / "Done"

## Quick Log (item screen)

Fields: Skilled count + Skilled Hrs + Laborer count + Laborer Hrs → preview shows computed man-hours.

IDs: `ql_skilled`, `ql_skilled_hours`, `ql_laborer`, `ql_laborer_hours`, `ql_preview`, `ql_date`

Validation: each worker type entered must have its own hours > 0.

## Log Today Modal (global, all floors)

Loops over all floors and items. Each item row has four inputs:
- `lt_s_${fi}_${idx}` — skilled count
- `lt_sh_${fi}_${idx}` — skilled hours
- `lt_l_${fi}_${idx}` — laborer count
- `lt_lh_${fi}_${idx}` — laborer hours

Only saves rows where `skilled × skilledHours + laborer × laborerHours > 0`.

## Export

Accessed from any screen via the Export button in the header.

**Options:**
- **Floor:** individual floor or "All Floors"
- **Format:** PDF (browser print dialog) or CSV
- **Period:** Daily / Weekly (Mon–Sun) / Monthly

**PDF — All Floors:** each floor is a separate section with its own heading and floor total; page break between floors; grand total at the bottom.

**CSV columns:** FLOOR, ITEM, DATE, SKILLED, SKILLED HRS, LABORER, LABORER HRS, MAN-HOURS

## Backup / Restore

Buttons on the floors screen. Backup downloads `hallway_backup_YYYY-MM-DD.json`. Restore reads a `.json` file and replaces `state` entirely (prompts for confirmation first).

## Branding

Header title: `Corridor's Plus ➕` — the ➕ emoji is white via `filter:brightness(0) invert(1)`.

## PWA

Meta tags in `<head>` allow "Add to Home Screen" on iOS. No service worker.

## What NOT to do

- Do not add a build system, bundler, or external libraries
- Do not split into multiple files
- Do not change the `localStorage` key (`hallway_tracker_v2`) without adding a migration
- Do not read `entry.hours` or `entry.workers` directly — always use `entryManHours()` / `entryDisplay()`
- Do not remove backward-compat logic in `entryManHours()` and `entryDisplay()` — old data in localStorage depends on it
