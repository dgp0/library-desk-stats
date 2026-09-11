"""Turn the LimeSurvey export into two files the Desk Stats sheet can take.

Input: the LimeSurvey results export (one row per campus per week, one
column per category, mode and weekday; gate counts per block and weekday).
Output, next to this script unless told otherwise:
  history-log.csv   rows shaped like the Log tab
  history-gate.csv  rows shaped like the Gate tab
  history-report.txt what was dropped or changed, so nothing is silent

Usage: python3 limesurvey_to_history.py <export.csv> [output folder]
"""

from __future__ import annotations

import csv
import datetime as dt
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
CAMPUS_RENAME = {"Moore Co.": "Moore County"}
CATEGORY_RENAME = {"General Reference": "General Ref", "Technical Support": "Tech Support"}
MODE_RENAME = {"In-Person": "In person", "Remote": "Remote"}
LOG_HEADERS = ["When", "Date", "Hour", "Weekday", "Campus", "Category", "Mode", "Count", "Recorded by", "Source", "Tap id"]
GATE_HEADERS = ["Date", "Campus", "Block", "Count", "Recorded by", "Source", "When entered"]
_CELL_RE = re.compile(r"\[([^\]]+)\]\[(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)\]\s*$")
_CAT_SPLIT = re.compile(r"(Aiding|Software|Limited|All questions|Performing|Provision)")


def parse_date(text: str) -> dt.date | None:
    """Accept 2025-10-19, 2025-10-19 00:00:00, or 10/19/2025."""
    s = (text or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    return None


def sunday_of(d: dt.date) -> dt.date:
    return d - dt.timedelta(days=(d.weekday() + 1) % 7)


def parse_count(text: str) -> tuple[int | None, str | None]:
    """A whole number, or the leading number of a note like '12 approx'; text-only becomes blank."""
    s = (text or "").strip()
    if s == "":
        return None, None
    if re.fullmatch(r"\d+", s):
        return int(s), None
    m = re.match(r"^(\d+)", s)
    if m:
        return int(m.group(1)), f"kept the number {m.group(1)} out of '{s}'"
    return None, f"left blank, was '{s}'"


def load(path: Path) -> tuple[list[str], list[list[str]]]:
    text = path.read_text(encoding="utf-8-sig")
    first = text.splitlines()[0]
    delim = "\t" if first.count("\t") > first.count(",") else ","
    rows = list(csv.reader(text.splitlines(), delimiter=delim))
    return rows[0], rows[1:]


def column_map(header: list[str]) -> list[tuple[int, str, str | None, str, int]]:
    cols: list[tuple[int, str, str | None, str, int]] = []
    for i, h in enumerate(header):
        m = _CELL_RE.search(h)
        if not m:
            continue
        day = DAY_NAMES.index(m.group(2))
        if h.startswith("Gate Count"):
            cols.append((i, "gate", None, m.group(1), day))
        else:
            cat = _CAT_SPLIT.split(h)[0].replace("&amp;", "&").strip()
            cols.append((i, "tap", CATEGORY_RENAME.get(cat, cat), MODE_RENAME.get(m.group(1), m.group(1)), day))
    return cols


def convert(export: Path, out_dir: Path) -> None:
    header, data = load(export)
    cols = column_map(header)

    def col(prefix: str) -> int:
        return next(i for i, h in enumerate(header) if h.startswith(prefix))

    camp_i, from_i, to_i = col("Campus Location"), col("Date From"), col("Date To")
    notes: list[str] = []
    responses = []
    for r in data:
        rid = r[0].strip()
        campus = CAMPUS_RENAME.get(r[camp_i].strip(), r[camp_i].strip()) if len(r) > camp_i else ""
        start, end = parse_date(r[from_i]) if len(r) > from_i else None, parse_date(r[to_i]) if len(r) > to_i else None
        if not campus:
            notes.append(f"response {rid}: no campus, skipped")
            continue
        if not start:
            notes.append(f"response {rid} ({campus}): no start date, skipped")
            continue
        if end and (end - start).days > 7:
            notes.append(f"response {rid} ({campus}): end date {end} is {(end - start).days} days after start {start}; the week of {start} was used")
        if end and end < start:
            notes.append(f"response {rid} ({campus}): end date {end} is before start {start}; the week of {start} was used")
        week = sunday_of(start)
        cells = {}
        for i, kind, cat, mode_or_block, day in cols:
            value, note = parse_count(r[i] if i < len(r) else "")
            if note:
                notes.append(f"response {rid} ({campus}, {kind} {mode_or_block} {DAY_NAMES[day]}): {note}")
            if value is not None:
                cells[(kind, cat, mode_or_block, day)] = value
        responses.append({"id": rid, "campus": campus, "week": week, "cells": cells})

    # One entry per campus-week: the one with more numbers wins; ties go to the later id.
    by_week: dict[tuple[str, dt.date], list[dict]] = defaultdict(list)
    for resp in responses:
        by_week[(resp["campus"], resp["week"])].append(resp)
    kept = []
    for (campus, week), group in sorted(by_week.items()):
        group.sort(key=lambda x: (len(x["cells"]), int(x["id"]) if x["id"].isdigit() else 0))
        winner = group[-1]
        for loser in group[:-1]:
            notes.append(f"campus-week {campus} {week}: response {loser['id']} dropped ({len(loser['cells'])} numbers), response {winner['id']} kept ({len(winner['cells'])} numbers)")
        kept.append(winner)

    log_rows, gate_rows = [], []
    totals = Counter()
    for resp in kept:
        src = f"LimeSurvey {resp['id']}"
        for (kind, cat, mode_or_block, day), value in sorted(resp["cells"].items(), key=lambda kv: (kv[0][3], kv[0][0], str(kv[0][1]), kv[0][2])):
            date = resp["week"] + dt.timedelta(days=day)
            if kind == "tap":
                log_rows.append(["", date.isoformat(), "", DAY_NAMES[day], resp["campus"], cat, mode_or_block, value, "", src, ""])
                totals[("taps", resp["campus"])] += value
            else:
                gate_rows.append([date.isoformat(), resp["campus"], mode_or_block, value, "", src, ""])
                totals[("gate", resp["campus"])] += value

    out_dir.mkdir(parents=True, exist_ok=True)
    with (out_dir / "history-log.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(LOG_HEADERS)
        w.writerows(log_rows)
    with (out_dir / "history-gate.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(GATE_HEADERS)
        w.writerows(gate_rows)
    weeks = sorted({r["week"] for r in kept})
    report = [
        f"Source: {export}",
        f"Responses in the export: {len(data)}; usable: {len(responses)}; after removing double entries: {len(kept)}",
        f"Weeks covered: {weeks[0]} to {weeks[-1]} ({len(weeks)} distinct weeks)",
        f"Log rows written: {len(log_rows)}; Gate rows written: {len(gate_rows)}",
        "Totals by campus:",
    ]
    for campus in sorted({r["campus"] for r in kept}):
        report.append(f"  {campus}: questions {totals[('taps', campus)]}, gate {totals[('gate', campus)]}")
    report.append("Notes (everything dropped or changed):")
    report.extend("  " + n for n in notes)
    (out_dir / "history-report.txt").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("usage: limesurvey_to_history.py <export.csv> [output folder]")
    convert(Path(sys.argv[1]), Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent)
