#!/usr/bin/env python3
"""
BIS Planner — TBC Classic Gear Spreadsheet Generator
Created by Steven Bennett 2026

Two-step workflow:
  1. python3 generate_bis_planner.py --build-db         # fetch item stats from Wowhead (once)
  2. python3 generate_bis_planner.py paladin             # generate CSV + Excel (no API calls)

Supported classes: druid, hunter, mage, paladin, priest, rogue, shaman, warlock, warrior

Requires: openpyxl  (pip install openpyxl)
"""

import argparse
import concurrent.futures
import csv
import json
import os
import re
import ssl
import sys
import threading
import time
import urllib.error
import urllib.request

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GUIDES_DIR = os.path.join(SCRIPT_DIR, "Guides")
DB_DIR = os.path.join(SCRIPT_DIR, "DB")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
ITEM_DB_PATH = os.path.join(SCRIPT_DIR, "item_database.json")
GEAR_ORDER = [
    "Head", "Shoulder", "Back", "Chest", "Wrist", "Hands",
    "Waist", "Legs", "Feet", "Neck", "Ring", "Trinket",
    "Main Hand", "Off Hand", "Two Hand", "Ranged/Relic",
]
ACQ_ORDER = [
    "Auction House", "Dungeon Drop", "Dungeon Token",
    "Quest", "Quest (Dung)", "Reputation", "PvP",
    "World Drop", "Raid Drop", "Vendor", "Other",
]

DUNGEON_KEYWORDS = [
    "Shattered Halls", "Steamvault", "Botanica", "Shadow Labyrinth",
    "Auchenai Crypts", "Black Morass", "Slave Pens", "Mechanar",
    "Arcatraz", "Hellfire Ramparts", "Blood Furnace", "Underbog",
    "Mana-Tombs", "Mana Tombs", "Sethekk Halls", "Old Hillsbrad",
    "Caverns of Time", "Tempest Keep",
]

LOCATION_ABBREVS = {
    "The Shattered Halls": "SH", "Shattered Halls": "SH",
    "The Steamvault": "SV", "Steamvault": "SV",
    "The Botanica": "Bot", "Botanica": "Bot",
    "Shadow Labyrinth": "SL", "Auchenai Crypts": "AC",
    "The Black Morass": "BM", "Black Morass": "BM",
    "The Slave Pens": "SP", "Slave Pens": "SP",
    "The Mechanar": "Mech", "Mechanar": "Mech",
    "The Arcatraz": "Arc", "Arcatraz": "Arc",
    "Hellfire Ramparts": "Ramp",
    "The Blood Furnace": "BF", "Blood Furnace": "BF",
    "The Underbog": "UB", "Underbog": "UB",
    "Mana-Tombs": "MT", "Mana Tombs": "MT",
    "Sethekk Halls": "Seth",
    "Old Hillsbrad Foothills": "OHF",
    "Opening of the Dark Portal": "BM",
    "Caverns of Time": "CoT", "Tempest Keep": "TK",
    "Netherstorm": "NS", "Blade's Edge Mountains": "BEM",
    "Shadowmoon Valley": "SMV", "Nagrand": "Nag",
    "Zangarmarsh": "ZM", "Terokkar Forest": "TF",
    "Hellfire Peninsula": "HFP", "Auchindoun": "Auch",
    "Blackrock Spire": "BRS", "Naxxramas": "Naxx",
    "Ahn'Qiraj": "AQ", "Temple of Ahn'Qiraj": "AQ40",
    "Zul'Gurub": "ZG",
}

WOWHEAD_DATA_ENV = 5
FETCH_DELAY = 0.15
FETCH_WORKERS = 4

STAT_COLUMNS = [
    "Armor", "DPS", "Str", "Agi", "Sta", "Int", "Spi",
    "Healing", "Spell Dmg", "AP", "MP5",
    "Defense", "Dodge", "Parry", "Block Rating", "Block Value",
    "Hit", "Crit", "Spell Hit", "Spell Crit", "Haste",
]

CSV_FIELDS = [
    "Spec", "Gear Type", "Phase", "Name", "Acquisition Type",
    "Quest", "Dungeon", "Difficulty", "Stats", "Notes",
]

SHEET_FIELDS = [
    "Interest", "Spec", "Gear Type", "Name", "Phase", "Acquisition Type",
    "Quest", "Dungeon", "Difficulty", "Stats", "Equip",
    "CmpRaw", "Comparison", "Notes",
]

INTEREST_OPTIONS = ["Pass", "Consider", "Need", "Equipped"]

SLOT_MAP = {
    "Head": "Head", "Shoulder": "Shoulder", "Back": "Back",
    "Chest": "Chest", "Wrist": "Wrist", "Hands": "Hands",
    "Waist": "Waist", "Legs": "Legs", "Feet": "Feet",
    "Neck": "Neck", "Ring": "Ring", "Trinket": "Trinket",
    "Main Hand": "Main Hand", "Off Hand": "Off Hand",
    "Two Hand": "Two Hand", "Ranged/Relic": "Ranged/Relic",
    "Main Hand~Off Hand": "Main Hand",
}

# Wowhead tooltip slot strings -> our slot names
TOOLTIP_SLOT_MAP = {
    "Head": "Head", "Shoulder": "Shoulder", "Back": "Back",
    "Chest": "Chest", "Robe": "Chest", "Wrist": "Wrist",
    "Hands": "Hands", "Waist": "Waist", "Legs": "Legs",
    "Feet": "Feet", "Neck": "Neck", "Finger": "Ring",
    "Trinket": "Trinket", "Main Hand": "Main Hand",
    "Off Hand": "Off Hand", "One-Hand": "Main Hand",
    "Two-Hand": "Two Hand", "Held In Off-hand": "Off Hand",
    "Ranged": "Ranged/Relic", "Relic": "Ranged/Relic",
    "Thrown": "Ranged/Relic", "Shield": "Off Hand",
}

SKIP_SOURCE_TYPES = {"Alchemy", "Transmute", "Conjured"}


def log(msg):
    print(msg, file=sys.stderr)


# ============================================================
# 1. PARSE ITEM SOURCES DB
# ============================================================

def parse_item_sources():
    filepath = os.path.join(DB_DIR, "ItemSources.lua")
    sources = {}
    with open(filepath, "r") as f:
        for line in f:
            m = re.match(
                r'\s*\[(\d+)\]\s*=\s*\{\s*'
                r'Name\s*=\s*"([^"]*)".*?'
                r'SourceType\s*=\s*LBIS\.L\["([^"]*)"\].*?'
                r'Source\s*=\s*(.*?),\s*'
                r'SourceNumber\s*=\s*"([^"]*)".*?'
                r'SourceLocation\s*=\s*(.*?),\s*'
                r'SourceFaction\s*=\s*"([^"]*)"',
                line,
            )
            if not m:
                continue

            def extract_val(raw):
                lm = re.match(r'LBIS\.L\["([^"]*)"\]', raw)
                if lm:
                    return lm.group(1)
                sm = re.match(r'"([^"]*)"', raw)
                if sm:
                    return sm.group(1)
                if ".." in raw:
                    parts = re.findall(r'LBIS\.L\["([^"]*)"\]|"([^"]*)"', raw)
                    return " ".join(p[0] or p[1] for p in parts).strip()
                return raw.strip('"')

            sources[int(m.group(1))] = {
                "name": m.group(2),
                "source_type": m.group(3),
                "source": extract_val(m.group(4).strip()),
                "source_number": m.group(5),
                "source_location": extract_val(m.group(6).strip()),
                "faction": m.group(7),
            }
    return sources


# ============================================================
# 2. PARSE GUIDE LUA FILES
# ============================================================

def parse_guide_items(filepath, spec_name, max_phase):
    """Parse items from phases 0..max_phase (cumulative), plus phase 99 (PrePatch).

    Applies the same label transforms as the in-game addon:
      - Phase 0  -> "PreRaid"
      - Phase 99 -> "PrePatch"
      - Phase < max_phase -> replace "BIS" with "Alt"
      - Phase == max_phase -> keep original label
    Duplicates across phases use the highest phase's label.
    """
    phases_to_scan = list(range(max_phase + 1)) + [99]
    raw_by_phase = {p: [] for p in phases_to_scan}

    with open(filepath, "r") as f:
        content = f.read()

    for p in phases_to_scan:
        phase_var = f"spec{p}"
        for m in re.finditer(
            rf'LBIS:AddItem\({phase_var},\s*"(\d+)",\s*LBIS\.L\["([^"]*)"\],\s*"([^"]*)"\)\s*--(.+)',
            content,
        ):
            raw_by_phase[p].append({
                "item_id": int(m.group(1)),
                "gear_type": m.group(2),
                "bis_rank": m.group(3),
                "name": m.group(4).strip(),
            })

    seen = {}
    for p in phases_to_scan:
        for item in raw_by_phase[p]:
            if p == 0:
                label = "PreRaid"
            elif p == 99:
                label = "PrePatch"
            elif p < max_phase:
                label = item["bis_rank"].replace("BIS", "Alt")
            else:
                label = item["bis_rank"]

            key = (item["item_id"], item["gear_type"])
            if key in seen:
                if p != 99 and p > seen[key]["_phase"]:
                    seen[key]["bis_rank"] = label
                    seen[key]["_phase"] = p
            else:
                seen[key] = {
                    "spec": spec_name,
                    "item_id": item["item_id"],
                    "gear_type": item["gear_type"],
                    "bis_rank": label,
                    "name": item["name"],
                    "_phase": p,
                }

    return [
        {k: v for k, v in entry.items() if k != "_phase"}
        for entry in seen.values()
    ]


def parse_all_guide_items():
    """Parse all guide Lua files to get item_id -> slot mapping."""
    id_to_slot = {}
    id_to_name = {}
    for fname in os.listdir(GUIDES_DIR):
        if not fname.endswith(".lua"):
            continue
        path = os.path.join(GUIDES_DIR, fname)
        with open(path, "r") as f:
            for line in f:
                m = re.match(
                    r'LBIS:AddItem\(spec\d+,\s*"(\d+)",\s*LBIS\.L\["([^"]*)"\].*?\)\s*--(.+)',
                    line.strip(),
                )
                if m:
                    iid = int(m.group(1))
                    slot = SLOT_MAP.get(m.group(2), m.group(2))
                    id_to_slot.setdefault(iid, slot)
                    id_to_name[iid] = m.group(3).strip()
    return id_to_slot, id_to_name


# ============================================================
# 2b. CLASS / SPEC DISCOVERY
# ============================================================

def _extract_spec_name(filepath):
    """Extract spec display name from the first RegisterSpec call in a guide file."""
    with open(filepath, "r") as f:
        for line in f:
            m = re.match(
                r'.*LBIS:RegisterSpec\(LBIS\.L\["[^"]*"\],\s*LBIS\.L\["([^"]*)"\]', line
            )
            if m:
                return m.group(1)
    return None


def discover_class_guides(class_name):
    """Auto-discover spec guides for a class by scanning Guides/.

    Returns (spec_order, guides_list) where guides_list is
    [(spec_name, filename), ...] sorted alphabetically by spec.
    """
    class_title = class_name.strip().title()

    guides = []
    for fname in sorted(os.listdir(GUIDES_DIR)):
        if fname.startswith(class_title) and fname.endswith(".lua"):
            path = os.path.join(GUIDES_DIR, fname)
            spec_name = _extract_spec_name(path)
            if spec_name:
                guides.append((spec_name, fname))

    if not guides:
        available = set()
        for fname in os.listdir(GUIDES_DIR):
            if fname.endswith(".lua"):
                m = re.match(r"([A-Z][a-z]+)", fname)
                if m:
                    available.add(m.group(1).lower())
        log(f"ERROR: No guide files found for class '{class_title}' in {GUIDES_DIR}")
        log(f"  Available classes: {', '.join(sorted(available))}")
        sys.exit(1)

    spec_order = [g[0] for g in guides]
    log(f"Discovered {class_title}: {', '.join(spec_order)} ({len(guides)} specs)")
    return spec_order, guides


def _detect_max_phase(filepath):
    """Detect the highest phase number registered in a guide file."""
    max_p = 0
    with open(filepath, "r") as f:
        for line in f:
            m = re.match(r'local\s+spec(\d+)\s*=\s*LBIS:RegisterSpec', line)
            if m:
                p = int(m.group(1))
                if p != 99 and p > max_p:
                    max_p = p
    return max_p


# ============================================================
# 3. TOOLTIP FETCHING + STAT PARSING
# ============================================================

def _make_ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def fetch_tooltip(item_id, ssl_ctx, retries=3):
    url = f"https://nether.wowhead.com/tooltip/item/{item_id}?dataEnv={WOWHEAD_DATA_ENV}&locale=0"
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
                "Accept": "application/json",
            })
            with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                return json.loads(resp.read().decode("utf-8")).get("tooltip", "")
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(3 * (attempt + 1))
            elif e.code == 404:
                return None
            else:
                time.sleep(1)
        except Exception:
            time.sleep(1)
    return None


def parse_tooltip_to_dict(html):
    """Parse tooltip HTML into (stats_dict, special_str, slot_str, item_level)."""
    if not html:
        return {}, "", None, 0

    stats = {}

    armor_m = re.search(r"(\d+)\s+Armor", html)
    if armor_m:
        stats["Armor"] = int(armor_m.group(1))

    dps_m = re.search(r"([\d.]+) damage per second", html, re.I)
    if dps_m:
        stats["DPS"] = float(dps_m.group(1))

    for pattern, key in [
        (r"\+(\d+)\s+Strength", "Str"),
        (r"\+(\d+)\s+Agility", "Agi"),
        (r"\+(\d+)\s+Stamina", "Sta"),
        (r"\+(\d+)\s+Intellect", "Int"),
        (r"\+(\d+)\s+Spirit", "Spi"),
    ]:
        m = re.search(pattern, html)
        if m:
            stats[key] = int(m.group(1))

    equip_patterns = [
        (r"[Ii](?:ncreases?|mproves?) (?:your )?defense rating by\s*(\d+)", "Defense"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?dodge rating by\s*(\d+)", "Dodge"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?parry rating by\s*(\d+)", "Parry"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?(?:shield )?block rating by\s*(\d+)", "Block Rating"),
        (r"[Ii](?:ncreases?|mproves?) the block value of your shield by\s*(\d+)", "Block Value"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?resilience rating by\s*(\d+)", "Resilience"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?(?:melee and ranged )?hit rating by\s*(\d+)", "Hit"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?(?:melee and ranged )?critical strike rating by\s*(\d+)", "Crit"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?spell hit rating by\s*(\d+)", "Spell Hit"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?spell critical strike rating by\s*(\d+)", "Spell Crit"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?(?:melee )?haste rating by\s*(\d+)", "Haste"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?expertise rating by\s*(\d+)", "Expertise"),
        (r"[Ii](?:ncreases?|mproves?) (?:your )?armor penetration rating by\s*(\d+)", "ArPen"),
        (r"[Ii](?:ncreases?|mproves?) attack power by\s*(\d+)", "AP"),
        (r"[Ii](?:ncreases?|mproves?) damage and healing done by magical spells and effects by (?:up to )?(\d+)", "Spell Dmg/Heal"),
        (r"[Rr]estores? (\d+) mana per 5 sec", "MP5"),
        (r"[Rr]estores? (\d+) health per 5 sec", "HP5"),
    ]
    for pattern, key in equip_patterns:
        m = re.search(pattern, html)
        if m:
            stats[key] = int(m.group(1))

    heal_m = re.search(r"[Ii]ncreases? healing done by (?:up to )?(\d+)(?:.*?damage done by (?:up to )?(\d+))?", html)
    if heal_m:
        stats["Healing"] = int(heal_m.group(1))
        if heal_m.lastindex >= 2 and heal_m.group(2):
            stats["Spell Dmg"] = int(heal_m.group(2))

    specials = []
    use_m = re.search(r"Use:\s*(?:<[^>]*>)*\s*([^<]+?)(?:\.\s*\(|<)", html)
    if use_m:
        eff = use_m.group(1).strip().rstrip(".")
        if 5 < len(eff) < 200:
            specials.append(f"Use: {eff}")

    proc_m = re.search(r"Chance on (?:hit|melee hit|spell critical hit|critical hit):\s*(?:<[^>]*>)*\s*([^<]+?)(?:\.\s*|<)", html)
    if proc_m:
        eff = proc_m.group(1).strip().rstrip(".")
        if len(eff) > 5:
            specials.append(f"Proc: {eff}")

    equip_special = re.findall(r"Equip:\s*(?:<[^>]*>)*\s*([^<]+?)(?:\.|<)", html)
    for eff in equip_special:
        eff_clean = eff.strip().rstrip(".")
        already = any(
            kw in eff_clean.lower()
            for kw in [
                "defense", "dodge", "parry", "block", "resilience", "hit",
                "critical", "haste", "expertise", "armor penetration",
                "attack power", "damage and healing", "healing done",
                "mana per 5", "health per 5", "damage done",
            ]
        )
        if not already and 10 < len(eff_clean) < 200:
            specials.append(f"Equip: {eff_clean}")

    special_str = "; ".join(specials).replace("&nbsp;", " ").replace("  ", " ").strip()

    slot = None
    slot_m = re.search(r'<table[^>]*><tr><td>([A-Za-z][A-Za-z -]+?)</td>', html)
    if slot_m:
        raw_slot = slot_m.group(1).strip()
        slot = TOOLTIP_SLOT_MAP.get(raw_slot)

    ilvl = 0
    ilvl_m = re.search(r"Item Level\s*(?:<!--[^>]*-->)?\s*(\d+)", html)
    if ilvl_m:
        ilvl = int(ilvl_m.group(1))

    normalize_spell_stats(stats)

    return stats, special_str, slot, ilvl


def normalize_spell_stats(stats):
    """Map spell tooltip data to Healing + Spell Dmg only (no stored Spell Dmg/Heal key).

    Parse may still set Spell Dmg/Heal from combined green text; split lines from heal_m win on
    conflict. Combined-only tooltips become equal Healing and Spell Dmg; _spell_from_combined_green
    marks that case for idempotent reloads.
    """
    if not stats:
        return

    if "Spell Dmg/Heal" in stats and (
        "Healing" in stats or "Spell Dmg" in stats
    ):
        stats.pop("_spell_from_combined_green", None)
        del stats["Spell Dmg/Heal"]

    if "Spell Dmg/Heal" in stats and "Healing" not in stats and "Spell Dmg" not in stats:
        v = stats["Spell Dmg/Heal"]
        stats["Healing"] = v
        stats["Spell Dmg"] = v
        del stats["Spell Dmg/Heal"]
        stats["_spell_from_combined_green"] = True

    stats.pop("Spell Dmg/Heal", None)


def stats_dict_to_string(stats, attributes=""):
    """Convert a stats dict + attribute string to the readable Stats column.

    Stat lines use '; ' between pairs; two newlines separate the stat block from
    on-use / equip attribute text when both are present.
    """
    parts = []
    for key in STAT_COLUMNS:
        val = stats.get(key)
        if val:
            parts.append(f"{val} {key}")
    result = "; ".join(parts)
    if attributes:
        at = attributes.strip()
        if at:
            if result:
                result += "\n\n" + at
            else:
                result = at
    return result


# ============================================================
# 4. BUILD ITEM DATABASE (--build-db)
# ============================================================

_fetch_counter_lock = threading.Lock()
_fetch_counter = [0, 0]  # [done, total]


def _fetch_one_item(item_id, name, ssl_ctx):
    tooltip = fetch_tooltip(item_id, ssl_ctx)
    with _fetch_counter_lock:
        _fetch_counter[0] += 1
        done, total = _fetch_counter
        if done % 50 == 0 or done == total:
            log(f"  [{done}/{total}] ...")
    time.sleep(FETCH_DELAY)
    if not tooltip:
        return item_id, None
    stats, special, slot, ilvl = parse_tooltip_to_dict(tooltip)
    return item_id, {"name": name, "slot": slot, "stats": stats, "special": special, "ilvl": ilvl}


def build_item_database(test_mode=False, test_class=None):
    existing_db = {}
    if os.path.exists(ITEM_DB_PATH):
        with open(ITEM_DB_PATH, "r") as f:
            existing_db = json.load(f)
        log(f"Loaded existing database: {len(existing_db)} items")

    guide_slots, guide_names = parse_all_guide_items()
    log(f"Parsed guide files: {len(guide_slots)} unique items with known slots")

    sources = parse_item_sources()
    log(f"Parsed ItemSources.lua: {len(sources)} items")

    if test_mode:
        test_ids = set()
        if test_class:
            _, guides = discover_class_guides(test_class)
        else:
            guides = [
                (None, f)
                for f in os.listdir(GUIDES_DIR) if f.endswith(".lua")
            ]
        for _, filename in guides:
            path = os.path.join(GUIDES_DIR, filename)
            if os.path.exists(path):
                with open(path, "r") as f:
                    for line in f:
                        m = re.match(r'LBIS:AddItem\(spec\d+,\s*"(\d+)"', line.strip())
                        if m:
                            test_ids.add(int(m.group(1)))
        label = test_class.title() if test_class else "all guide"
        items_to_fetch = {
            iid: guide_names.get(iid, sources.get(iid, {}).get("name", f"Item {iid}"))
            for iid in test_ids
        }
        log(f"TEST MODE: only fetching {len(items_to_fetch)} {label} items")
    else:
        items_to_fetch = {}
        for iid, info in sources.items():
            if info["source_type"] in SKIP_SOURCE_TYPES:
                continue
            src = info.get("source", "")
            if info["source_type"] == "Profession" and any(
                skip in src for skip in ("Alchemy", "Cooking", "First Aid")
            ):
                continue
            items_to_fetch[iid] = info["name"]
        for iid, name in guide_names.items():
            if iid not in items_to_fetch:
                items_to_fetch[iid] = name

    to_fetch = {iid: name for iid, name in items_to_fetch.items() if str(iid) not in existing_db}
    log(f"Items to fetch from Wowhead: {len(to_fetch)} (skipping {len(items_to_fetch) - len(to_fetch)} already cached)")

    if to_fetch:
        _fetch_counter[0] = 0
        _fetch_counter[1] = len(to_fetch)
        ssl_ctx = _make_ssl_ctx()

        with concurrent.futures.ThreadPoolExecutor(max_workers=FETCH_WORKERS) as pool:
            futures = {
                pool.submit(_fetch_one_item, iid, name, ssl_ctx): iid
                for iid, name in to_fetch.items()
            }
            for future in concurrent.futures.as_completed(futures):
                iid, result = future.result()
                if result:
                    existing_db[str(iid)] = result

    db = {}
    for str_id, entry in existing_db.items():
        iid = int(str_id)
        slot = entry.get("slot")
        if not slot and iid in guide_slots:
            entry["slot"] = guide_slots[iid]
            slot = entry["slot"]
        if slot:
            db[str_id] = entry

    for entry in db.values():
        normalize_spell_stats(entry.get("stats") or {})

    with open(ITEM_DB_PATH, "w") as f:
        json.dump(db, f, separators=(",", ":"))

    log(f"\nSaved {len(db)} gear items -> {ITEM_DB_PATH}")
    return db


def load_item_database():
    if not os.path.exists(ITEM_DB_PATH):
        log(f"ERROR: {ITEM_DB_PATH} not found. Run --build-db first.")
        sys.exit(1)
    with open(ITEM_DB_PATH, "r") as f:
        data = json.load(f)
    for entry in data.values():
        if isinstance(entry.get("stats"), dict):
            normalize_spell_stats(entry["stats"])
    return data


# ============================================================
# 5. CLASSIFY ACQUISITION
# ============================================================

def _abbrev_location(loc):
    if not loc:
        return loc
    clean = loc.replace(" (H)", "").replace(" (N)", "").replace(" (Heroic)", "").replace(" (Normal)", "")
    return LOCATION_ABBREVS.get(clean, clean)


def _is_dungeon_quest(loc):
    if not loc:
        return False
    return any(kw.lower() in loc.lower() for kw in DUNGEON_KEYWORDS)


def _shorten_dungeon(name):
    return name[4:] if name.startswith("The ") else name


def classify_acquisition(src_info):
    """Returns (acq_type, quest, dungeon, difficulty, notes)."""
    if not src_info:
        return "Other", "", "", "", ""

    st = src_info["source_type"]
    source = src_info["source"]
    loc = src_info["source_location"]

    if st == "Profession" or st == "Transmute":
        return "Auction House", "", "", "", ""

    if st == "Quest":
        abbr = _abbrev_location(loc)
        kind = "Quest (Dung)" if _is_dungeon_quest(loc) else "Quest"
        return kind, f"{source} ({abbr})", "", "", ""

    if st == "Drop":
        if source == "World Drop" or not loc:
            return "Auction House", "", "", "", "World Drop"
        if loc in ("World Bosses", "World Boss"):
            return "World Drop", "", "", "", source
        is_heroic = any(h in loc for h in ("(H)", "(Heroic)"))
        clean_loc = re.sub(r"\s*\((H|N|Heroic|Normal)\)", "", loc)
        dungeon = _shorten_dungeon(clean_loc)
        difficulty = "Heroic" if is_heroic else "Normal"
        return "Dungeon Drop", "", dungeon, difficulty, source

    if st == "Dungeon Token":
        cost = src_info["source_number"]
        if "Badge of Justice" in source:
            note = f"{cost} Badges of Justice"
        elif "Spirit Shard" in source:
            note = f"{cost} Spirit Shards"
        else:
            note = f"{source} x{cost}"
        return "Dungeon Token", "", "", "", note

    if st == "Reputation":
        return "Reputation", "", "", "", f"{source} - {loc}"

    if st == "PvP":
        return "PvP", "", "", "", ""

    if st == "Vendor":
        return "Vendor", "", "", "", loc

    if st == "Tier Token":
        return "Raid Drop", "", "", "", f"{source} ({loc})"

    if st == "Token":
        return "Dungeon Token", "", "", "", source

    if st == "Conjured":
        return "Other", "", "", "", source

    return "Other", "", "", "", ""


# ============================================================
# 6. GENERATE CSV
# ============================================================

def generate_csv(max_phase, class_name, output_csv):
    spec_order, guides = discover_class_guides(class_name)

    item_db = load_item_database()
    id_to_stats = {}
    for str_id, entry in item_db.items():
        id_to_stats[int(str_id)] = entry

    sources = parse_item_sources()
    log(f"Parsed ItemSources.lua: {len(sources)} items")

    all_items = []
    for spec_name, filename in guides:
        path = os.path.join(GUIDES_DIR, filename)
        if not os.path.exists(path):
            log(f"  WARNING: {path} not found, skipping")
            continue
        items = parse_guide_items(path, spec_name, max_phase)
        all_items.extend(items)
        log(f"  {spec_name}: {len(items)} items (phases 0-{max_phase})")

    rows = []
    for item in all_items:
        src = sources.get(item["item_id"])
        acq_type, quest, dungeon, difficulty, notes = classify_acquisition(src)
        db_entry = id_to_stats.get(item["item_id"], {})
        stats_d = db_entry.get("stats", {})
        attributes = db_entry.get("special", "")
        stats_str = stats_dict_to_string(stats_d, attributes)
        rows.append({
            "Spec": item["spec"],
            "Gear Type": item["gear_type"],
            "Phase": item["bis_rank"],
            "Name": item["name"],
            "Acquisition Type": acq_type,
            "Quest": quest,
            "Dungeon": dungeon,
            "Difficulty": difficulty,
            "Stats": stats_str,
            "Notes": notes,
        })

    sort_key = _make_sort_key(spec_order)
    rows.sort(key=sort_key)

    os.makedirs(os.path.dirname(output_csv), exist_ok=True)
    with open(output_csv, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)

    log(f"Wrote {len(rows)} rows -> {output_csv}")
    return rows, spec_order


def _make_sort_key(spec_order):
    def _sort_key(row):
        s = spec_order.index(row["Spec"]) if row["Spec"] in spec_order else 99
        g = GEAR_ORDER.index(row["Gear Type"]) if row["Gear Type"] in GEAR_ORDER else 99
        a = ACQ_ORDER.index(row["Acquisition Type"]) if row["Acquisition Type"] in ACQ_ORDER else 99
        return (s, g, a, row["Name"])
    return _sort_key


# ============================================================
# 7. EXCEL EXPORT
# ============================================================

ACQ_COLORS = {
    "Auction House": "80DEEA", "Reputation": "A5D6A7",
    "Quest": "FFE082", "Quest (Dung)": "FFE082",
    "PvP": "F48FB1", "Dungeon Token": "BDBDBD",
    "World Drop": "EEEEEE", "Raid Drop": "D7CCC8",
    "Vendor": "EEEEEE", "Other": "EEEEEE",
}

DUNGEON_COLORS_NORMAL = {
    "Blood Furnace": "FFD5D5", "Hellfire Ramparts": "FFE0CC",
    "Shattered Halls": "FFEABB", "Mechanar": "FFF5BB",
    "Botanica": "E8FFCC", "Arcatraz": "FFF0CC",
    "Auchenai Crypts": "F5CCF0", "Mana Tombs": "D5CCF5",
    "Sethekk Halls": "E0CCF5", "Shadow Labyrinth": "E8CCEE",
    "Slave Pens": "CCE5FF", "Underbog": "CCEED8",
    "Steamvault": "CCD5F5",
    "Old Hillsbrad Foothills": "EEE4CC", "Black Morass": "CCE8D8",
    "Dire Maul": "F0EADD", "Molten Core": "FFE0D0",
    "Blackwing Lair": "D8DEE8", "Ahn'Qiraj": "F5F0CC",
    "Naxxramas": "D0EEEE", "Terokkar Forest": "DEE8CC",
}
DUNGEON_COLORS_HEROIC = {
    "Blood Furnace": "E87070", "Hellfire Ramparts": "E89060",
    "Shattered Halls": "D8A040", "Mechanar": "C8B030",
    "Botanica": "70B848", "Arcatraz": "D89838",
    "Auchenai Crypts": "C850A8", "Mana Tombs": "7858C8",
    "Sethekk Halls": "9858C8", "Shadow Labyrinth": "A848B8",
    "Slave Pens": "4090D8", "Underbog": "40A878",
    "Steamvault": "4868C8",
    "Old Hillsbrad Foothills": "B89050", "Black Morass": "48A880",
    "Dire Maul": "A89868", "Molten Core": "D07050",
    "Blackwing Lair": "7888A8", "Ahn'Qiraj": "B0A040",
    "Naxxramas": "48A8B0", "Terokkar Forest": "80A848",
}


def _get_row_color(row):
    acq = row["Acquisition Type"]
    dungeon = row.get("Dungeon", "")
    difficulty = row.get("Difficulty", "")

    if acq == "Dungeon Drop" and dungeon:
        dname = dungeon.replace("Mana-Tombs", "Mana Tombs")
        is_heroic = difficulty == "Heroic"
        palette = DUNGEON_COLORS_HEROIC if is_heroic else DUNGEON_COLORS_NORMAL
        color = palette.get(dname, "D89838" if is_heroic else "FFF0CC")
        return color, is_heroic

    return ACQ_COLORS.get(acq), False


def export_xlsx(csv_path, spec_order):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        log("WARNING: openpyxl not installed. Skipping Excel export.")
        log("  Install with: pip install openpyxl")
        return

    item_db = load_item_database()

    all_rows = []
    with open(csv_path, "r") as f:
        for row in csv.DictReader(f):
            r = dict(row)
            legacy_sp = (r.pop("Special", None) or "").strip()
            if legacy_sp:
                st = (r.get("Stats") or "").strip()
                r["Stats"] = (st + " " + legacy_sp).strip() if st else legacy_sp
            all_rows.append(r)

    if not all_rows:
        log("No rows to export.")
        return

    header_fill = PatternFill("solid", fgColor="37474F")
    header_font = Font(bold=True, size=11, color="FFFFFF")
    section_fill = PatternFill("solid", fgColor="546E7A")
    section_font = Font(bold=True, size=12, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="BDBDBD"),
        right=Side(style="thin", color="BDBDBD"),
        top=Side(style="thin", color="BDBDBD"),
        bottom=Side(style="thin", color="BDBDBD"),
    )
    wrap_align = Alignment(wrap_text=True, vertical="top")
    intro_title_align = Alignment(wrap_text=False, vertical="center", horizontal="left")
    intro_title_font = Font(bold=True, size=14, color="212121")
    intro_body_font = Font(size=10, color="424242")
    intro_label_font = Font(bold=True, size=10, color="424242")
    intro_label_align = Alignment(wrap_text=False, vertical="center", horizontal="left")

    wb = Workbook()
    wb.remove(wb.active)

    # -- ItemDB sheet (hidden) --
    ws_db = wb.create_sheet(title="ItemDB")
    db_headers = ["Name", "Slot"] + STAT_COLUMNS + ["Attributes"]
    db_num_cols = len(db_headers)
    for ci, h in enumerate(db_headers, 1):
        ws_db.cell(row=1, column=ci, value=h)

    items_by_slot = {}
    db_row = 2
    for str_id, entry in sorted(item_db.items(), key=lambda x: x[1].get("name", "")):
        name = entry.get("name", "")
        slot = entry.get("slot", "")
        stats = entry.get("stats", {})
        special = entry.get("special", "")

        ws_db.cell(row=db_row, column=1, value=name)
        ws_db.cell(row=db_row, column=2, value=slot)
        for si, stat_key in enumerate(STAT_COLUMNS):
            val = stats.get(stat_key)
            if val:
                ws_db.cell(row=db_row, column=3 + si, value=val)
        ws_db.cell(row=db_row, column=3 + len(STAT_COLUMNS), value=special)

        items_by_slot.setdefault(slot, []).append(name)
        db_row += 1

    # Per-slot item lists (for dropdown data validation references)
    slot_list_cols = {}
    slot_col_start = db_num_cols + 2
    for si, gear_type in enumerate(GEAR_ORDER):
        col = slot_col_start + si
        ws_db.cell(row=1, column=col, value=gear_type)
        slot_items = sorted(items_by_slot.get(gear_type, []))
        for ri, item_name in enumerate(slot_items):
            ws_db.cell(row=2 + ri, column=col, value=item_name)
        slot_list_cols[gear_type] = (get_column_letter(col), len(slot_items))

    ws_db.sheet_state = "hidden"
    total_db_items = db_row - 2
    log(f"  ItemDB: {total_db_items} items")

    db_end_col_letter = get_column_letter(db_num_cols)
    itemdb_last_row = max(2, db_row - 1)
    itemdb_range_bounded = f"ItemDB!A$2:{db_end_col_letter}${itemdb_last_row}"

    # -- Current Equipment sheet (first tab) --
    ws_ce = wb.create_sheet(title="Current Equipment")
    ws_ce.sheet_properties.tabColor = "455A64"

    ce_headers = ["Gear Type", "Item Name"] + STAT_COLUMNS
    ce_ncol = len(ce_headers)
    ce_last_col = get_column_letter(ce_ncol)

    ce_intro_para1 = (
        "Add your current equipment to see comparisons by typing in item name. "
        "Note: Stats and comparisons will take a few seconds to update."
    )
    ce_intro_para2 = (
        'If an item in the BIS Planner sheet is set to "Equipped" in the Interest column, '
        "it will automatically update here. If your item is not on the list, you can manually "
        "enter the stats (integers only)."
    )
    # Rows 1–2: title A1:B2; C1:C2 "Instructions:"; D row1 / D row2 = one paragraph each (no vertical merge of body)
    CE_HEADER_ROW = 3
    ce_label_col = 3
    ce_text_start_col = 4
    ce_text_start_letter = get_column_letter(ce_text_start_col)

    ws_ce.merge_cells("A1:B2")
    ws_ce.cell(row=1, column=1, value="Current Equipment")
    ws_ce.cell(row=1, column=1).font = intro_title_font
    ws_ce.cell(row=1, column=1).alignment = intro_title_align

    ws_ce.merge_cells(f"{get_column_letter(ce_label_col)}1:{get_column_letter(ce_label_col)}2")
    lab_ce = ws_ce.cell(row=1, column=ce_label_col, value="Instructions:")
    lab_ce.font = intro_label_font
    lab_ce.alignment = intro_label_align

    ws_ce.merge_cells(f"{ce_text_start_letter}1:{ce_last_col}1")
    p1 = ws_ce.cell(row=1, column=ce_text_start_col, value=ce_intro_para1)
    p1.font = intro_body_font
    p1.alignment = wrap_align

    ws_ce.merge_cells(f"{ce_text_start_letter}2:{ce_last_col}2")
    p2 = ws_ce.cell(row=2, column=ce_text_start_col, value=ce_intro_para2)
    p2.font = intro_body_font
    p2.alignment = wrap_align

    for ci, h in enumerate(ce_headers, 1):
        cell = ws_ce.cell(row=CE_HEADER_ROW, column=ci, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    ws_ce.column_dimensions["A"].width = 14
    ws_ce.column_dimensions["B"].width = 34
    ws_ce.column_dimensions["C"].width = 16
    for ci in range(4, 3 + len(STAT_COLUMNS)):
        ws_ce.column_dimensions[get_column_letter(ci)].width = 9
    ws_ce.freeze_panes = f"C{CE_HEADER_ROW + 1}"

    ce_spec_rows = {}
    ce_row = CE_HEADER_ROW + 1
    for spec_name in spec_order:
        cell = ws_ce.cell(row=ce_row, column=1, value=f"--- {spec_name.upper()} ---")
        cell.font = section_font
        cell.fill = section_fill
        for ci in range(2, len(ce_headers) + 1):
            c = ws_ce.cell(row=ce_row, column=ci)
            c.fill = section_fill
        ce_row += 1

        spec_start = ce_row
        for gear_type in GEAR_ORDER:
            ws_ce.cell(row=ce_row, column=1, value=gear_type)

            for si, stat_key in enumerate(STAT_COLUMNS):
                db_col = 3 + si
                formula = (
                    f'=IFERROR(VLOOKUP(B{ce_row},{itemdb_range_bounded},'
                    f'{db_col},FALSE),"")'
                )
                ws_ce.cell(row=ce_row, column=3 + si, value=formula)

            col_info = slot_list_cols.get(gear_type)
            if col_info:
                sl_col, sl_count = col_info
                if sl_count > 0:
                    dv = DataValidation(
                        type="list",
                        formula1=f"ItemDB!${sl_col}$2:${sl_col}${sl_count + 1}",
                        allow_blank=True,
                    )
                    dv.prompt = f"Select or type a {gear_type} item"
                    dv.promptTitle = gear_type
                    ws_ce.add_data_validation(dv)
                    dv.add(ws_ce.cell(row=ce_row, column=2))

            for ci in range(1, len(ce_headers) + 1):
                ws_ce.cell(row=ce_row, column=ci).border = thin_border

            ce_row += 1

        ce_spec_rows[spec_name] = (spec_start, ce_row - 1)
        ce_row += 1

    credit_font = Font(size=8, italic=True, color="999999")
    credit_cell = ws_ce.cell(row=ce_row + 1, column=1, value="Created by Steven Bennett 2026")
    credit_cell.font = credit_font

    log(f"  Current Equipment: {len(spec_order)} specs x {len(GEAR_ORDER)} slots")

    # -- Single BIS Planner sheet --
    # A=Interest B=Spec C=Gear Type D=Name E=Phase F=Acq Type
    # G=Quest H=Dungeon I=Difficulty J=Stats K=Equip L=CmpRaw M=Comparison N=Notes
    ws = wb.create_sheet(title="BIS Planner")
    ws.sheet_properties.tabColor = "1565C0"

    equip_col_letter = _col_letter(SHEET_FIELDS.index("Equip") + 1)
    cmp_raw_col_letter = _col_letter(SHEET_FIELDS.index("CmpRaw") + 1)

    col_widths = {
        "A": 12, "B": 12, "C": 13, "D": 28, "E": 10, "F": 20,
        "G": 16, "H": 13, "I": 12, "J": 34, "K": 2, "L": 2, "M": 34, "N": 14,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions[equip_col_letter].hidden = True
    ws.column_dimensions[cmp_raw_col_letter].hidden = True

    bp_ncol = len(SHEET_FIELDS)
    bp_last_col = get_column_letter(bp_ncol)
    bp_intro_row1 = (
        "Click the down arrow on a given column to filter or sort values. "
        'If no values are displayed, click the "Remove Filter/Filter" button on the Sheets '
        "toolbar twice to reset the filters."
    )
    bp_intro_row2 = (
        "Select field in interest column and filter most wanted items. "
        "Note: When selecting Equipped, sheet will take a few seconds to update."
    )
    # Rows 1–2: title A1:B2; E1:E2 "Instructions:"; F1:N1 and F2:N2 instruction lines (text from F1 / F2)
    BP_HEADER_ROW = 3
    bp_text_start_col = 6
    bp_text_start_letter = get_column_letter(bp_text_start_col)

    ws.merge_cells("A1:B2")
    ws.cell(row=1, column=1, value="BIS Planner")
    ws.cell(row=1, column=1).font = intro_title_font
    ws.cell(row=1, column=1).alignment = intro_title_align

    ws.merge_cells("E1:E2")
    lab_bp = ws.cell(row=1, column=5, value="Instructions:")
    lab_bp.font = intro_label_font
    lab_bp.alignment = intro_label_align

    ws.merge_cells(f"{bp_text_start_letter}1:{bp_last_col}1")
    b1 = ws.cell(row=1, column=bp_text_start_col, value=bp_intro_row1)
    b1.font = intro_body_font
    b1.alignment = wrap_align

    ws.merge_cells(f"{bp_text_start_letter}2:{bp_last_col}2")
    b2 = ws.cell(row=2, column=bp_text_start_col, value=bp_intro_row2)
    b2.font = intro_body_font
    b2.alignment = wrap_align

    for ci, field in enumerate(SHEET_FIELDS, 1):
        cell = ws.cell(row=BP_HEADER_ROW, column=ci, value=field)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = wrap_align

    ws.freeze_panes = f"E{BP_HEADER_ROW + 1}"
    filter_end = get_column_letter(len(SHEET_FIELDS))

    interest_dv = DataValidation(
        type="list",
        formula1=f'"{",".join(INTEREST_OPTIONS)}"',
        allow_blank=True,
    )
    ws.add_data_validation(interest_dv)

    for i, data_row in enumerate(all_rows):
        excel_row = i + BP_HEADER_ROW + 1
        color_hex, dark_bg = _get_row_color(data_row)
        fill = PatternFill("solid", fgColor=color_hex) if color_hex else None
        font = Font(color="FFFFFF", size=10) if dark_bg else Font(size=10)

        for ci, field in enumerate(SHEET_FIELDS, 1):
            if field == "Equip":
                formula = _build_equip_name_formula(
                    excel_row, ce_spec_rows, spec_order
                )
                cell = ws.cell(row=excel_row, column=ci, value=formula)
            elif field == "CmpRaw":
                formula = _build_cmp_raw_formula(
                    excel_row,
                    equip_col_letter,
                    itemdb_range_bounded,
                    db_num_cols,
                )
                cell = ws.cell(row=excel_row, column=ci, value=formula)
            elif field == "Comparison":
                cell = ws.cell(
                    row=excel_row,
                    column=ci,
                    value=f"={cmp_raw_col_letter}{excel_row}",
                )
            elif field == "Interest":
                cell = ws.cell(row=excel_row, column=ci, value="")
                interest_dv.add(cell)
            else:
                cell = ws.cell(row=excel_row, column=ci, value=data_row.get(field, ""))
            cell.border = thin_border
            cell.alignment = wrap_align
            cell.font = font
            if fill:
                cell.fill = fill

    bp_last_row = BP_HEADER_ROW + len(all_rows)
    ws.auto_filter.ref = f"A{BP_HEADER_ROW}:{filter_end}{bp_last_row}"

    log(f"  BIS Planner: {len(all_rows)} items")

    # Tab order: Current Equipment, BIS Planner, ItemDB (hidden)
    target_order = ["Current Equipment", "BIS Planner", "ItemDB"]
    for idx, name in enumerate(target_order):
        if name in wb.sheetnames:
            current_idx = wb.sheetnames.index(name)
            wb.move_sheet(name, offset=idx - current_idx)

    xlsx_path = csv_path.rsplit(".", 1)[0] + ".xlsx"
    wb.save(xlsx_path)
    log(f"Excel -> {xlsx_path} ({len(all_rows)} items, {total_db_items} in ItemDB)")


def _build_equip_name_formula(row, ce_spec_rows, spec_order):
    """Equipped item name from Current Equipment (one IFS+INDEX+MATCH per row)."""
    ce = "'Current Equipment'"
    equipped_checks = []
    for spec_name in spec_order:
        ce_start, ce_end = ce_spec_rows[spec_name]
        equipped_checks.append(
            f'B{row}="{spec_name}",'
            f'INDEX({ce}!B{ce_start}:B{ce_end},'
            f'MATCH(C{row},{ce}!A{ce_start}:A{ce_end},0))'
        )
    return f'=IFERROR(IFS({",".join(equipped_checks)}),"")'


def _build_cmp_raw_formula(row, helper_col_letter, itemdb_range, attr_col_num):
    """Plain-text comparison + equipped Attributes (ItemDB last col); Google Sheets Apps Script adds colors in Comparison column."""
    parts = []
    for si, stat_key in enumerate(STAT_COLUMNS):
        db_col_num = 3 + si
        item_stat = (
            f'IFERROR(VLOOKUP(D{row},{itemdb_range},{db_col_num},FALSE),0)'
        )
        equip_stat = (
            f'IFERROR(VLOOKUP({helper_col_letter}{row},{itemdb_range},'
            f'{db_col_num},FALSE),0)'
        )
        diff = f"{item_stat}-{equip_stat}"
        part = f'IF({diff}<>0,TEXT({diff},"+#;-#")&" {stat_key}","")'
        parts.append(part)

    tj = f'TEXTJOIN(", ",TRUE,{",".join(parts)})'
    v_attr = (
        f'IFERROR(VLOOKUP({helper_col_letter}{row},{itemdb_range},'
        f'{attr_col_num},FALSE),"")'
    )
    combined = (
        f'IF(AND({tj}="",{v_attr}=""),"",'
        f'IF({v_attr}="",{tj},IF({tj}="",{v_attr},{tj}&CHAR(10)&CHAR(10)&{v_attr})))'
    )
    return (
        f'=IF({helper_col_letter}{row}="","Current Equipment Not Specified",{combined})'
    )


def _col_letter(n):
    """1-indexed column number to letter(s)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ============================================================
# MAIN
# ============================================================


def list_guide_classes():
    """Lowercase class names that have at least one guide file in Guides/."""
    classes = set()
    for fname in os.listdir(GUIDES_DIR):
        if fname.endswith(".lua"):
            m = re.match(r"([A-Z][a-z]+)", fname)
            if m:
                classes.add(m.group(1).lower())
    return sorted(classes)


def main():
    parser = argparse.ArgumentParser(
        description="BIS Planner — TBC Classic Gear Spreadsheet Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Workflow:\n"
            "  1. python3 generate_bis_planner.py --build-db       # fetch item stats (once)\n"
            "  2. python3 generate_bis_planner.py paladin           # generate Paladin sheet\n"
            "     python3 generate_bis_planner.py --all              # generate every class\n"
            "     python3 generate_bis_planner.py all               # same as --all\n"
            "\n"
            "  Classes: druid, hunter, mage, paladin, priest, rogue, shaman, warlock, warrior\n"
            "  Phase 0 (default) = PreRaid. Phase 1 = PreRaid + Phase 1 items (cumulative).\n"
        ),
    )
    parser.add_argument("class_name", nargs="?", default=None,
                        help="Class to generate (e.g., paladin, warrior, druid), or 'all' for every class")
    parser.add_argument("--all", action="store_true",
                        help="Generate CSV and Excel for every class that has guide files")
    parser.add_argument("--build-db", action="store_true",
                        help="Fetch item stats from Wowhead and build item_database.json")
    parser.add_argument("--test", action="store_true",
                        help="With --build-db: only fetch items for the specified class (or all guide items)")
    parser.add_argument("--phase", type=int, default=0,
                        help="Max phase to include (cumulative). 0=PreRaid (default), 1=Phase 1, etc.")
    parser.add_argument("--output", type=str, default=None, help="Output CSV path override")
    parser.add_argument("--no-excel", action="store_true", help="Skip .xlsx generation")
    args = parser.parse_args()

    if args.build_db:
        build_item_database(test_mode=args.test, test_class=args.class_name)
        return

    run_all = args.all or (
        args.class_name is not None and args.class_name.strip().lower() == "all"
    )
    if run_all:
        if args.output:
            log("ERROR: --output cannot be used with --all or 'all' (each class writes its own file).")
            sys.exit(1)
        classes = list_guide_classes()
        log(f"Generating all {len(classes)} classes: {', '.join(classes)}")
        for cn in classes:
            class_title = cn.strip().title()
            output_csv = os.path.join(OUTPUT_DIR, f"{class_title}_BIS_Planner.csv")
            rows, spec_order = generate_csv(args.phase, cn, output_csv)
            if not args.no_excel:
                export_xlsx(output_csv, spec_order)
        log("Done (all classes).")
        return

    if not args.class_name:
        available = list_guide_classes()
        log("ERROR: Please specify a class name.")
        log(f"  Available classes: {', '.join(available)}")
        log(f"  Example: python3 generate_bis_planner.py paladin")
        log(f"  Or all: python3 generate_bis_planner.py --all")
        sys.exit(1)

    class_title = args.class_name.strip().title()
    output_csv = args.output or os.path.join(OUTPUT_DIR, f"{class_title}_BIS_Planner.csv")

    rows, spec_order = generate_csv(args.phase, args.class_name, output_csv)

    if not args.no_excel:
        export_xlsx(output_csv, spec_order)


if __name__ == "__main__":
    main()
