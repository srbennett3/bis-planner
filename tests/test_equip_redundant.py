"""Tests for stripping Equip: text redundant with STAT_COLUMNS display stats."""

import os
import sys

import pytest

SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import generate_bis_planner as gbp


def test_combined_spell_power_redundant_when_both_stats_match():
    stats = {
        "Sta": 24,
        "Healing": 34,
        "Spell Dmg": 34,
        "Spell Crit": 12,
    }
    special = (
        "Equip: Increases damage and healing done by magical spells and effects by up to 34"
    )
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert out == ""
    full = gbp.stats_dict_to_string(stats, special)
    assert "Equip:" not in full
    assert "34 Healing" in full
    assert "34 Spell Dmg" in full


def test_combined_spell_power_kept_when_number_mismatches():
    stats = {"Healing": 34, "Spell Dmg": 33}
    special = (
        "Equip: Increases damage and healing done by magical spells and effects by up to 34"
    )
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert "Equip:" in out
    assert "up to 34" in out


def test_combined_spell_power_kept_when_only_one_of_healing_spell_dmg_in_stats():
    stats = {"Spell Dmg": 34}
    special = (
        "Equip: Increases damage and healing done by magical spells and effects by up to 34"
    )
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert "Equip:" in out


def test_school_specific_shadow_kept_despite_spell_dmg_stat():
    stats = {"Spell Dmg": 50}
    special = "Equip: Increases damage done by Shadow spells and effects by up to 50"
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert out == special


def test_proc_like_equip_kept():
    stats = {"Spell Dmg": 159, "Healing": 159}
    special = (
        "Equip: Your harmful spells have a chance to increase your spell haste rating "
        "by 320 for 6 secs. (Proc chance: 10%, 45s cooldown)"
    )
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert out == special


def test_multi_segment_one_redundant_one_kept():
    stats = {"Healing": 34, "Spell Dmg": 34}
    special = (
        "Equip: Increases damage and healing done by magical spells and effects by up to 34; "
        "Equip: Allows the bearer to see into the ghost world while in Shadowmoon Valley"
    )
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert "magical spells" not in out
    assert "Shadowmoon Valley" in out
    assert out.count("Equip:") == 1


def test_attack_power_redundant():
    stats = {"AP": 80}
    special = "Equip: Increases attack power by 80"
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert out == ""


def test_use_and_proc_segments_unchanged():
    stats = {"Healing": 40, "Spell Dmg": 40}
    special = (
        "Use: Restore 123 mana; "
        "Equip: Increases damage and healing done by magical spells and effects by up to 40; "
        "Proc: Chance on hit: do something"
    )
    out = gbp.strip_equip_redundant_with_display_stats(stats, special)
    assert "Use:" in out
    assert "Proc:" in out
    assert "magical spells" not in out
