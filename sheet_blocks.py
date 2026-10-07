# -*- coding: utf-8 -*-
"""
Turns any report-style worksheet into a list of renderable "blocks" so every
sheet in the workbook can become a dashboard tab automatically.

Block types produced (all JSON-serialisable):
  kpi    {"t":"kpi","title","sub","items":[{"l","v","f","s"}]}
  table  {"t":"table","title","sub","groups":[..]|None,"cols":[{"h","f"}],"rows":[[..]],"notes":[..]}
  pairs  {"t":"pairs","title","sub","items":[{"h"} | {"l","x"}]}          (commentary / insights)
  note   {"t":"note","title","lines":[..]}

Column formats ("f"): txt, num, pct, inr, lakh, date, time, plain
"""
import math
import re
from datetime import datetime, date, time, timedelta

ERR_STRS = {"#VALUE!", "#DIV/0!", "#N/A", "#REF!", "#NAME?", "#NULL!", "#NUM!"}
SKIP_TITLE_RE = re.compile(r"^(sheets in this|verdict thresholds|data notes)", re.I)
KPI_HEADERS = {"key indicator", "key indicators", "key figures", "key figure"}


def _clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.replace("\r", "").strip()
        if s in ERR_STRS or s == "":
            return None
        return s
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    return v


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_time(v):
    return isinstance(v, (time, timedelta))


def _is_date(v):
    return isinstance(v, (datetime, date))


def _nn(r):
    return [i for i, v in enumerate(r) if v is not None]


def _slug(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def _label(h):
    return re.sub(r"\s+", " ", str(h)).strip() if h is not None else ""


def _val(v, fmt):
    """Serialise a cell for JSON according to its column format."""
    if v is None:
        return None
    if _is_date(v):
        return (v.date() if isinstance(v, datetime) else v).isoformat()
    if isinstance(v, timedelta):
        return round(v.total_seconds(), 1)
    if isinstance(v, time):
        return v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6
    if _is_num(v):
        if fmt in ("plain",):
            return int(v) if float(v).is_integer() else v
        return int(v) if float(v).is_integer() and abs(v) < 1e15 else round(float(v), 6)
    return str(v)


def _col_format(header, values):
    h = _label(header).lower()
    vals = [v for v in values if v is not None]
    if not vals:
        return "txt"
    if all(_is_date(v) for v in vals):
        return "date"
    if all(_is_time(v) for v in vals):
        return "time"
    nums = [v for v in vals if _is_num(v)]
    if len(nums) < max(1, int(0.6 * len(vals))):
        return "txt"
    if re.search(r"(^|\b)(sr|sl)\.?\s*(no)?\.?$|pincode|^year$|^id$", h):
        return "plain"
    if "lakh" in h:
        return "lakh"
    if "₹" in h or re.search(r"\bcost\b|refund amount|spend|\bfee\b|amount", h):
        return "inr"
    if re.search(r"%|share|percent|\bctr\b|bounce|\brate\b|growth", h):
        return "pct"
    return "num"


def _kpi_fmt(label, value):
    if _is_date(value):
        return "date"
    if not _is_num(value):
        return "txt"
    l = label.lower()
    if "lakh" in l:
        return "lakh"
    if "₹" in l or re.search(r"\brefund\b|\bamount\b|\bcost\b|\bspend\b", l):
        return "inr"
    if re.search(r"%|share|percent|\brate\b", l) and abs(value) <= 5:
        return "pct"
    return "num"


def _split_blocks(rows):
    blocks, cur = [], []
    for r in rows:
        if all(v is None for v in r):
            if cur:
                blocks.append(cur)
                cur = []
        else:
            cur.append(r)
    if cur:
        blocks.append(cur)
    return blocks


def _single_str_rows(rows):
    return all(len(_nn(r)) == 1 and isinstance(r[_nn(r)[0]], str) for r in rows)


def _titleish(lines):
    """A block of only single-cell strings that reads like a section heading."""
    if not lines or len(lines) > 3:
        return False
    first = lines[0]
    if len(first) > 90 or first.startswith(("•", "-", "*")) or first.lower().startswith("note"):
        return False
    return True


def _parse_data_block(rows, pending):
    """rows: list of equal-length lists of cleaned cells with >=1 non-empty cell.
    Returns a list of blocks (a side-by-side layout yields several)."""
    W = max(len(r) for r in rows)
    rows = [list(r) + [None] * (W - len(r)) for r in rows]

    lead = []
    while rows and len(_nn(rows[0])) == 1 and isinstance(rows[0][_nn(rows[0])[0]], str):
        i = _nn(rows[0])[0]
        lead.append((i, rows[0][i]))
        rows.pop(0)

    title = lead[0][1] if lead else (pending[0] if pending else "")
    sub = (pending[1] if (pending and len(pending) > 1 and not lead) else "")
    base_col = lead[0][0] if lead else 0
    sub_lines = [t for (ci, t) in lead[1:] if ci == base_col]
    captions_all = [(ci, t) for (ci, t) in lead[1:] if ci != base_col]
    if sub_lines:
        sub = " · ".join(sub_lines)

    # validation helper rows ("Check: ...") are noise
    rows = [r for r in rows
            if not (any(isinstance(v, str) and v.strip().lower().startswith("check") for v in r) and len(_nn(r)) <= 3)]
    if not rows:
        if lead:
            lines = [t for _, t in lead]
            return [{"t": "title", "title": lines[0], "sub": " · ".join(lines[1:])}]
        return []

    used = [j for j in range(W) if any(r[j] is not None for r in rows)]
    if len(rows) <= 2:
        segs = [list(range(used[0], used[-1] + 1))]          # label row + value row = one horizontal KPI strip
    else:
        segs = []
        for j in used:
            if segs and j == segs[-1][-1] + 1:
                segs[-1].append(j)
            else:
                segs.append([j])

    # pass 1: decide which segments survive
    kept = []
    for seg in segs:
        sub_rows = [[r[j] for j in seg] for r in rows]
        sub_rows = [r for r in sub_rows if any(v is not None for v in r)]
        if not sub_rows:
            continue
        if len(segs) > 1 and len(seg) == 1 and not any(_is_num(r[0]) for r in sub_rows):
            continue                                   # text-only helper/legend column
        captions = [t for (ci, t) in captions_all if seg[0] <= ci <= seg[-1]]
        if captions and re.search(r"feeds chart|^ranked", captions[0], re.I):
            continue                                   # Excel chart-feeding helper table
        kept.append((seg, sub_rows, captions))

    out = []
    multi = len(kept) > 1
    for si, (seg, sub_rows, captions) in enumerate(kept):
        seg_title, seg_sub = title, sub
        if captions:
            seg_title = captions[0]
            seg_sub = sub if si == 0 else ""
        elif multi:
            hdr0 = next((v for v in sub_rows[0] if isinstance(v, str)), "")
            srlike = bool(re.match(r"^(sr|sl|#)[.]?\s*(no)?[.]?$", hdr0, re.I))
            if si == 0 and (srlike or not title):
                seg_title = title or hdr0
            else:
                seg_title = f"{title} — {hdr0}" if title else hdr0
            seg_sub = sub if si == 0 else ""
        if SKIP_TITLE_RE.search(seg_title or ""):
            continue
        blk = _classify(sub_rows, seg_title, seg_sub)
        if blk:
            out.append(blk)
    return out


def _classify(rows, title, sub):
    ncols = len(rows[0])
    # trailing single-cell rows = notes
    notes = []
    while len(rows) > 1 and len(_nn(rows[-1])) == 1 and isinstance(rows[-1][_nn(rows[-1])[0]], str):
        notes.insert(0, rows[-1][_nn(rows[-1])[0]])
        rows = rows[:-1]

    # commentary: 1-2 string columns, long text in col 2, optional single-cell sub-heads
    def is_pairs(rs):
        pair_rows = []
        for r in rs:
            k = _nn(r)
            if len(k) == 1 and isinstance(r[k[0]], str):
                continue
            if len(k) == 2 and all(isinstance(r[i], str) for i in k):
                pair_rows.append(len(r[k[1]]))
                continue
            return False
        return len(pair_rows) >= 2 and (sum(pair_rows) / len(pair_rows) > 40 or sum(1 for n in pair_rows if n > 40) >= 0.6 * len(pair_rows))

    if is_pairs(rows):
        items = []
        for idx, r in enumerate(rows):
            k = _nn(r)
            if len(k) == 1:
                items.append({"h": r[k[0]]})
            elif idx == 0 and len(r[k[1]]) <= 45:
                continue                                # header-like first row of a text table
            else:
                items.append({"l": r[k[0]].lstrip("•· ").strip(), "x": r[k[1]]})
        return {"t": "pairs", "title": title, "sub": sub, "items": items}

    if all(len(_nn(r)) == 1 for r in rows) and all(isinstance(v, str) for r in rows for v in r if v is not None):
        return {"t": "note", "title": title, "lines": [r[_nn(r)[0]] for r in rows] + notes}

    # header / group-row detection
    def all_str(r):
        k = _nn(r)
        return len(k) >= 1 and all(isinstance(r[i], str) for i in k)

    groups = None
    header = None
    data = rows
    if len(rows) >= 2 and all_str(rows[1]) and all_str(rows[0]) and len(_nn(rows[0])) < len(_nn(rows[1])) \
            and len(_nn(rows[1])) == ncols and len(rows) >= 3:
        g, cur = [], ""
        for v in rows[0]:
            if v is not None:
                cur = v
            g.append(cur)
        groups, header, data = g, rows[1], rows[2:]
    elif len(rows) >= 2 and all_str(rows[0]) and len(_nn(rows[0])) >= 2 and (
            any(_is_num(v) or _is_date(v) for v in rows[1][1:]) or ncols >= 4):
        header, data = rows[0], rows[1:]

    # KPI forms -------------------------------------------------------------
    # horizontal: label row + value row
    if header is not None and groups is None and len(data) == 1 and 2 <= ncols <= 8 and \
            all(isinstance(header[i], str) for i in _nn(header)) and any(_is_num(v) or _is_date(v) or v is None for v in data[0]):
        items = []
        for j in range(ncols):
            if header[j] is None:
                continue
            v = data[0][j]
            items.append({"l": _label(header[j]), "v": _val(v, "num"), "f": _kpi_fmt(_label(header[j]), v), "s": ""})
        if items:
            return {"t": "kpi", "title": title, "sub": sub, "items": items}

    # vertical: label | value [| pct-or-note]
    hdr_is_kpi = header is not None and groups is None and isinstance(header[0], str) and header[0].strip().lower() in KPI_HEADERS
    if (header is None or hdr_is_kpi) and 2 <= ncols <= 3 and len(data) <= 16 and \
            all(isinstance(r[0], str) for r in data):
        items = []
        for r in data:
            v = r[1]
            sub_s = ""
            if ncols == 3 and r[2] is not None:
                sub_s = (f"{r[2]*100:.1f}% of total" if _is_num(r[2]) and abs(r[2]) <= 1.5 else str(r[2]))
            items.append({"l": _label(r[0]), "v": _val(v, "num"), "f": _kpi_fmt(_label(r[0]), v), "s": sub_s})
        return {"t": "kpi", "title": title, "sub": sub, "items": items}

    # table ---------------------------------------------------------------
    if header is None:
        header = [f"Col {i+1}" for i in range(ncols)]
    heads = [_label(h) if h is not None else "" for h in header]
    cols_vals = list(zip(*data)) if data else [[] for _ in heads]
    fmts = [_col_format(h, cv) for h, cv in zip(heads, cols_vals)]
    # the first column is a label column unless it is numeric
    # Percent columns whose cached values are all 0/blank (formulas never calculated) are rebuilt from the counts beside them
    def _is_agg(r):
        return any(isinstance(v, str) and re.search(r"total", v, re.I) for v in r[:2])
    recomputed = False
    data = [list(r) for r in data]
    for j, f in enumerate(fmts):
        if f == "pct" and heads[j].lstrip().startswith("%") and all((r[j] in (0, None)) for r in data):
            k = next((c for c in range(j - 1, 0, -1) if fmts[c] == "num" and any(_is_num(r[c]) and r[c] > 0 for r in data)), None)
            if k is None:
                continue
            den = sum(r[k] for r in data if _is_num(r[k]) and not _is_agg(r))
            if den > 0:
                for r in data:
                    if _is_num(r[k]):
                        r[j] = (r[k] / den) if not _is_agg(r) else (r[k] / den if re.search(r"subtotal", str(r[:2]), re.I) else 1.0)
                recomputed = True
    if recomputed:
        notes = notes + ["% columns were recalculated from the counts in this table (the source file held no calculated values)."]
    out_rows = [[_val(v, fmts[j]) for j, v in enumerate(r)] for r in data]
    # "TOTAL" rows whose numbers were all formula errors carry no information
    out_rows = [r for r in out_rows
                if not (isinstance(r[0], str) and re.match(r"^total", r[0], re.I) and not any(_is_num(v) for v in r[1:]))]
    if not out_rows:
        return None
    blk = {
        "t": "table", "title": title, "sub": sub,
        "groups": groups,
        "cols": [{"h": h, "f": f} for h, f in zip(heads, fmts)],
        "rows": out_rows, "notes": notes,
    }
    if not blk["title"]:
        h0 = heads[0] if heads and heads[0] else "Data"
        blk["title"] = h0 if (h0.lower().startswith("by ") or h0.endswith("s")) else f"By {h0}"
    return blk


def parse_sheet(name, rows):
    """Parse one worksheet into {name, slug, title, sub, blocks[]}."""
    cleaned = [[_clean(v) for v in r] for r in rows]
    raw_blocks = _split_blocks(cleaned)

    sheet = {"name": name, "slug": _slug(name), "title": name, "sub": "", "blocks": []}
    pending = None  # (title, sub) lines awaiting the next data block
    first = True

    for brows in raw_blocks:
        W = max(len(r) for r in brows)
        brows = [list(r) + [None] * (W - len(r)) for r in brows]

        # title-only blocks
        if _single_str_rows(brows):
            lines = [r[_nn(r)[0]] for r in brows]
            if first:
                sheet["title"] = lines[0]
                sheet["sub"] = " · ".join(lines[1:])
                first = False
                continue
            first = False
            if len(lines) == 1 and (len(lines[0]) > 100 or lines[0].startswith(("•", "Note", "Highest / Lowest"))):
                # stray footnote: attach to the previous table if there is one
                prev = next((b for b in reversed(sheet["blocks"]) if b["t"] == "table"), None)
                if prev is not None:
                    prev.setdefault("notes", []).append(lines[0])
                else:
                    sheet["blocks"].append({"t": "note", "title": "", "lines": lines})
                continue
            if _titleish(lines):
                pending = (lines[0], " · ".join(lines[1:]))
                continue
            if SKIP_TITLE_RE.search(lines[0]):
                pending = None
                continue
            sheet["blocks"].append({"t": "note", "title": lines[0], "lines": lines[1:]})
            pending = None
            continue

        first = False
        pending_had_title = bool(pending)
        prev_pending = pending
        blks = _parse_data_block(brows, pending)
        pending = None
        for b in blks:
            if b["t"] == "pairs" and prev_pending and b["title"] and b["title"] != prev_pending[0]:
                b["items"].insert(0, {"h": b["title"]})      # keep the section heading above the commentary list
                b["title"], b["sub"] = prev_pending[0], (prev_pending[1] if len(prev_pending) > 1 else "")
            if b["t"] == "table" and sheet["blocks"]:
                prev = sheet["blocks"][-1]
                if prev["t"] == "table" and " — " in prev["title"] and b["title"] and " — " not in b["title"]                         and not b["title"].startswith("By ") and b["title"] == b["cols"][0]["h"] and not pending_had_title:
                    b["title"] = prev["title"].split(" — ")[0] + " — " + b["title"]
            if b["t"] == "title":
                pending = (b["title"], b["sub"])
                continue
            # a one-card KPI block that only identifies what follows (e.g. "Institute: …")
            if b["t"] == "kpi" and len(b["items"]) == 1 and b["title"]:
                it = b["items"][0]
                pending = (b["title"], f'{it["l"]}: {it["v"]}')
                continue
            sheet["blocks"].append(b)

    return sheet


def parse_extra_sheets(extra):
    out = []
    for name, rows in extra:
        sh = parse_sheet(name, rows)
        if sh["blocks"]:
            out.append(sh)
    return out
