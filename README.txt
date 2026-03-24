# BIS Planner

A color-coded Best-In-Slot gear planner for TBC Classic. Generates an interactive Google Sheets spreadsheet per class with stat comparison, current equipment tracking, and item search — powered by LoonBestInSlot data and Wowhead stats.

### Supported Classes

| Class | Specs |
|---|---|
| Druid | Balance, Bear, Cat, Restoration |
| Hunter | Beast Mastery, Marksmanship, Survival |
| Mage | Arcane, Fire, Frost |
| Paladin | Holy, Protection, Retribution |
| Priest | Holy, Shadow |
| Rogue | Dps |
| Shaman | Elemental, Enhancement, Restoration |
| Warlock | Affliction, Demonology, Destruction |
| Warrior | Arms, Fury, Protection |

---

## For Users

### Quick Start

1. Upload the `.xlsx` file from the `output/` folder to Google Drive (e.g., `Paladin_BIS_Planner.xlsx`)
2. Open it with Google Sheets
3. Open `apps_script.gs` in a text editor, copy the entire contents
4. In Google Sheets go to **Extensions > Apps Script**
5. Delete any existing code, paste the script, press **Ctrl+S** to save
6. Close the Apps Script tab — you're done

The spreadsheet is now fully interactive. When you share it (File > Share, or File > Make a copy), the script travels with it — recipients don't need to set up anything.

### Spreadsheet Tabs

Each generated `.xlsx` includes these tabs (in order):

| Tab | Description |
|---|---|
| **General Info** | Static overview of how the workbook and Apps Script behave (first tab). |
| **Current Equipment** | Enter/track your currently equipped items per spec. |
| **BIS Planner** | All BIS and pre-BIS items for every spec on a single sheet, with stat comparison. |
| **ItemDB** (hidden) | TBC item stat database used by lookups and dropdowns. |
| **GemDB** (hidden) | Gem stats JSON for ideal gem scoring in Apps Script. |

### BIS Planner Sheet

All BIS and pre-BIS items for the class, with every spec combined on one sheet.

**Columns (left to right):**

| Column | Description |
|---|---|
| **Interest** | Dropdown depends on **Gear Type** (C): **Ring** / **Trinket** rows get `Pass`, `Consider`, `Need`, `Equipped 1`, `Equipped 2` (no plain `Equipped`). All other rows get `Pass`, `Consider`, `Need`, `Equipped` (no `Equipped 1`/`2`). Set in the exported `.xlsx` and re-applied on open in Google Sheets by Apps Script. |
| **Spec** | Spec name (e.g., Protection, Holy, Retribution) |
| **Gear Type** | Equipment slot (Head, Shoulder, Back, Chest, etc.) |
| **Name** | Item name |
| **Phase** | BIS rank label (PreRaid, BIS, Alt, Alt Mit, etc.) |
| **Acquisition Type** | How to obtain (Dungeon Drop, Reputation, Quest, PvP, etc.) |
| **Quest** | Quest name (blank for non-quest items) |
| **Dungeon** | Dungeon name (blank for non-dungeon items) |
| **Difficulty** | Normal or Heroic (blank for non-dungeon items) |
| **Stats** | Item stats (e.g., `142 Armor; 24 Sta; 23 Int`) and, when present, on-use / equip **Attributes** text after a single space |
| **Comparison** | Stat differences vs equipped gear (column **K**). **Google Sheets:** Apps Script builds comparison text from **Current Equipment** + ItemDB, then writes colored rich text to **K** (`+`/`−`), Ring/Trinket slot headers, and Attributes. **Excel:** column **K** stays blank unless you open the file in Google Sheets with the script. |
| **Notes** | Source details (boss name, badge cost, reputation, etc.) — last planner column (**M**). |

**Features:**

- **Interest dropdown**: Lists are **row-specific** (see table above). **Equipped** / **Equipped 1** / **Equipped 2** update Current Equipment where applicable and bold the cell. Only one planner row per spec + slot bucket can hold each equipped state; picking a new item clears the previous row in that bucket.
- **Comparison column**: The script fills **K** on open and when Current Equipment or planner data change (equipped items are read from **Current Equipment**, not from extra planner columns). Ring/Trinket rows use **Slot 1 Comparison:** / **Slot 2 comparison** and **Equipped slot 1:** / **Equipped slot 2** before on-use text; the script colors headers and splits inline **Use:** / **Equip:** onto new lines when present. **Regenerate the workbook** after pulling updates so the sheet layout matches the script (A–M only on BIS Planner).
- **Filter/Sort**: Use the header dropdowns to filter by Spec, Gear Type, Acquisition Type, etc.
- **Frozen columns**: Interest, Spec, Gear Type, and Name (columns A-D) stay visible while scrolling.
- **Color-coded rows**: Each row is colored by acquisition type (see Row Colors below).

### Current Equipment Sheet

Tracks what you're currently wearing, with separate sections for each spec (e.g., `--- PROTECTION ---`, `--- HOLY ---`). **Ring** and **Trinket** use two rows each: **Ring 1**, **Ring 2**, **Trinket 1**, **Trinket 2** (planner Gear Type stays a single **Ring** or **Trinket** column).

**Columns:**

| Column | Description |
|---|---|
| **Gear Type** | Equipment slot (Head, Shoulder, Back, **Ring 1**, **Ring 2**, **Trinket 1**, **Trinket 2**, etc.) |
| **Item Name** | Editable cell with a per-slot dropdown of 3,600+ TBC items |
| **Stat columns** | Armor, DPS, Str, Agi, Sta, Int, Spi, Healing, Spell Dmg, AP, MP5, Defense, Dodge, Parry, Block Rating, Block Value, Hit, Crit, Spell Hit, Spell Crit, Haste |

**Features:**

- Gear Type and Item Name columns (A-B) are frozen while scrolling.
- Each gear slot has a dropdown with all items of that slot type from the item database.
- Stat columns auto-populate via `VLOOKUP` when an item name is entered.
- Manually entering an item that exists in the BIS Planner will auto-set its Interest to **Equipped** / **Equipped 1** / **Equipped 2** to match that CE row (via Apps Script).
- Clearing or replacing an item auto-clears the matching equipped Interest on the BIS Planner (via Apps Script).

### Row Colors

| Acquisition Type | Color |
|---|---|
| Auction House | Cyan |
| Reputation | Green |
| Quest / Quest (Dung) | Gold |
| PvP | Pink |
| Dungeon Token | Grey |
| Dungeon Drop | Per-dungeon color (normal = light pastel, heroic = deep saturated) |

---

## For Developers

### Requirements

- **Python 3.6+**
- **openpyxl** (`pip install openpyxl`)

### Apps Script caches

Document Cache keys are per spreadsheet; execution-scoped memos (ItemDB name column, Current Equipment weights-row map) reset on the next Apps Script run.

### Folder Structure

```
BIS Planner/
├── generate_bis_planner.py       # generator script
├── apps_script.gs                # Google Apps Script (paste into Sheets once)
├── data/                         # item IDs, Wowhead cache, Pawn JSON, ItemSources
│   ├── ItemSources.lua           # item source / acquisition database
│   ├── item_database.json        # cached item stats (generated by --build-db)
│   └── pawn_scales_tbc.json      # regenerated from AddonReference/Pawn on each Excel export
├── output/                       # all generated CSV + XLSX files
│   ├── Paladin_BIS_Planner.csv
│   ├── Paladin_BIS_Planner.xlsx
│   ├── Warrior_BIS_Planner.csv
│   ├── Warrior_BIS_Planner.xlsx
│   └── ...
├── README.md
├── AddonReference/               # upstream addon files used by the generator
│   ├── Loon/                     # LoonBestInSlot Lua guide files (all 9 classes)
│   │   ├── DruidBalance.lua
│   │   ├── DruidBear.lua
│   │   ├── ...
│   │   └── WarriorProtection.lua
│   └── Pawn/                     # only ClassicHawsJon.lua (HawsJon TBC scales — same idea as Loon/*.lua)
```

### Two-Step Workflow

**Step 1: Build the item database** (one-time, ~5 min)

Fetches item stats from Wowhead for all items in `data/ItemSources.lua` using 4 parallel workers. Results are cached in `data/item_database.json` — subsequent runs only fetch new items.

```bash
python3 generate_bis_planner.py --build-db
```

For a quick test with only one class's items (~1 min):

```bash
python3 generate_bis_planner.py --build-db --test paladin
```

**Step 2: Generate a class spreadsheet** (instant, no API calls)

```bash
python3 generate_bis_planner.py paladin
python3 generate_bis_planner.py warrior
python3 generate_bis_planner.py druid
python3 generate_bis_planner.py --all          # every class (CSV + xlsx in output/)
python3 generate_bis_planner.py all            # same as --all
```

Output goes to `output/{Class}_BIS_Planner.csv` and `output/{Class}_BIS_Planner.xlsx`. The workbook includes a **Pawn weights (TBC)** sheet (HawsJon scales from `AddonReference/Pawn/ClassicHawsJon.lua`) when that class has matching Pawn data; `data/pawn_scales_tbc.json` is refreshed at the same time.

### After changing tooltip parsing or spell-stat normalization

If you edit `parse_tooltip_to_dict`, `normalize_spell_stats`, or other Wowhead parsing logic, rebuild the cached database and regenerate spreadsheets so `data/item_database.json` and ItemDB `VLOOKUP` columns stay aligned:

```bash
python3 generate_bis_planner.py --build-db
python3 generate_bis_planner.py --all
```

Use `--build-db --test paladin` for a faster partial rebuild while iterating.

### Phase Support

The `--phase` flag controls which content tiers to include. Phases are cumulative, matching the in-game addon's `LBIS.CurrentPhase` behavior:

```bash
python3 generate_bis_planner.py paladin              # Phase 0 = PreRaid only (default)
python3 generate_bis_planner.py paladin --phase 1    # PreRaid + Phase 1 items
python3 generate_bis_planner.py paladin --phase 2    # PreRaid + Phase 1 + Phase 2
```

Phase labels in the spreadsheet match the addon:

- **Phase 0** items show **"PreRaid"**
- **Phase 99** items show **"PrePatch"**
- Items from earlier phases show "Alt" instead of "BIS" (e.g., a Phase 0 "BIS" becomes "Alt" when generating with `--phase 1`)
- Items from the current max phase keep their original label (BIS, Alt, Alt Mit, etc.)

### Other Options

```bash
python3 generate_bis_planner.py paladin --output path.csv  # custom output path
python3 generate_bis_planner.py paladin --no-excel         # CSV only, skip xlsx
```

### Updating Lua Files

When LoonBestInSlot updates:

1. Copy the updated Lua files into `AddonReference/Loon/` and `data/`:

```bash
cp /path/to/LoonBestInSlot/Guides/*.lua  AddonReference/Loon/
cp /path/to/LoonBestInSlot/DB/ItemSources.lua  data/
```

To refresh Pawn weights, copy `ClassicHawsJon.lua` from the Pawn addon into `AddonReference/Pawn/` (replace the existing file).

2. Rebuild the database (only fetches new items):

```bash
python3 generate_bis_planner.py --build-db
```

3. Regenerate whichever classes you need:

```bash
python3 generate_bis_planner.py paladin
python3 generate_bis_planner.py warrior
```

4. Upload the new `.xlsx` to Google Sheets and re-add the Apps Script.

### Troubleshooting

- **"No module named openpyxl"**: `pip install openpyxl`
- **SSL errors**: The script bypasses SSL verification for the Wowhead API. Check your Python SSL module: `python3 -c "import ssl; print(ssl.OPENSSL_VERSION)"`
- **404 errors**: Some items may not exist in Wowhead's TBC Classic database. Their stats will be empty.
- **Rate limiting (429)**: The script backs off automatically. Increase `FETCH_DELAY` in the script if needed.
- **"No guide files found"**: Make sure the Lua files for that class are in `AddonReference/Loon/`.

---

Created by Steven Bennett 2026
