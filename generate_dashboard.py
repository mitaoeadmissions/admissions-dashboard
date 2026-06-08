"""
Reads Masterdata File.xlsx and injects all data into the HTML dashboard template.
Section boundaries are found dynamically so adding rows never breaks offsets.
Uses a temp-file copy so Excel is never blocked from saving.
"""
import json
import re
import shutil
import tempfile
import openpyxl
from pathlib import Path
from datetime import datetime

import os as _os
_IN_CLOUD = _os.environ.get("GITHUB_ACTIONS") == "true"

if _IN_CLOUD:
    # Running in GitHub Actions — paths are relative to repo root
    EXCEL_FILE    = Path("Masterdata File.xlsx")          # downloaded by dropbox_download.py
    HTML_TEMPLATE = Path("Complete admissions_dashboard.html")  # committed to repo
    OUTPUT_HTML   = Path("dashboard.html")
else:
    # Running locally
    EXCEL_FILE    = Path(r"C:\Users\guruv\Dropbox\Admission Dashboard\Masterdata File.xlsx")
    HTML_TEMPLATE = Path(r"C:\Users\guruv\Desktop\Office\Admission Dashboard\05.06.26\Complete admissions_dashboard.html")
    OUTPUT_HTML   = Path(r"C:\Users\guruv\Desktop\Office\Admission Dashboard\dashboard.html")

# ── helpers ───────────────────────────────────────────────────────────────────

def fmt_date(val):
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str):
        s = val.strip().split(" to ")[0].strip().split(" ")[0].strip()
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
    return None

def safe_num(v, fallback=0):
    if v is None:
        return fallback
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s in ("-", "—", "NA", "N/A", ""):
        return fallback
    try:
        return float(s) if "." in s else int(s)
    except Exception:
        return fallback

def safe_str(v, fallback=""):
    return str(v).strip() if v is not None else fallback

def null_or_num(v):
    """Return None for blank/NA, else numeric."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s in ("-", "—", "NA", "N/A", ""):
        return None
    try:
        return float(s) if "." in s else int(s)
    except Exception:
        return None

def title_from_url(url):
    if not url or not url.startswith("http"):
        return url or ""
    name = url.rstrip("/").split("/")[-1]
    name = name.replace("-", " ").replace("_", " ")
    if "." in name:
        name = name.rsplit(".", 1)[0]
    return name[:80] if name else url[:80]

def to_js(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))

def read_rows():
    """Copy Excel to temp, read all rows, delete temp. Keeps Excel unlocked."""
    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    try:
        shutil.copy2(str(EXCEL_FILE), str(tmp))
        wb = openpyxl.load_workbook(str(tmp), read_only=True, data_only=True)
        ws = wb["Sheet1"]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    finally:
        if tmp.exists():
            tmp.unlink()
    return rows

def find_section(rows, keyword):
    kw = keyword.lower().strip()
    for i, r in enumerate(rows):
        if r[0] and isinstance(r[0], str) and r[0].strip().lower().startswith(kw):
            return i
    raise ValueError(f"Section not found: '{keyword}'")

# ── section parsers ───────────────────────────────────────────────────────────

def parse_main(rows, start=0):
    """Daily admissions.
    Cols: Date | Eng Leads(daily) | Des Leads(daily) | Total(cumul) |
          Eng Paid | Des Paid | Total Paid(cumul) |
          Eng Sign | Des Sign | Total Sign(cumul) |
          Eng Prov | Des Prov | Total Prov(cumul) |
          Eng RevPaid | Des RevPaid | Total RevPaid |
          Eng RevProv | Des RevProv | Total RevProv | Grand Total
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], datetime):
            break
        data.append({
            "d":        fmt_date(r[0]),
            "eLeads":   safe_num(r[1]),
            "dLeads":   safe_num(r[2]),
            "ePaid":    safe_num(r[4]),
            "dPaid":    safe_num(r[5]),
            "eSign":    safe_num(r[7]),
            "dSign":    safe_num(r[8]),
            "eProv":    safe_num(r[10]),
            "dProv":    safe_num(r[11]),
            "eRevPaid": safe_num(r[13]),
            "dRevPaid": safe_num(r[14]),
            "eRevProv": safe_num(r[16]),
            "dRevProv": safe_num(r[17]),
        })
    return data


def parse_eng_status(rows, start):
    """Engineering lead status bifurcation.
    Excel cols: Date | CET Reg | DSY | Follow-ups | Future Prospects |
                Interested | Junk | Not Interested | Not Reachable | Untouched
    Template fields: pa=CET Reg, su=DSY, cet=Follow-ups, d2y=FutureProspects,
                     fu=Interested, fp=Junk, in=NotInterested, jk=NotReachable
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], datetime):
            break
        data.append({
            "d":   fmt_date(r[0]),
            "pa":  safe_num(r[1]),   # CET Registered
            "su":  safe_num(r[2]),   # Direct Second Year
            "cet": safe_num(r[3]),   # Follow-ups
            "d2y": safe_num(r[4]),   # Future Prospects
            "fu":  safe_num(r[5]),   # Interested
            "fp":  safe_num(r[6]),   # Junk
            "in":  safe_num(r[7]),   # Not Interested
            "jk":  safe_num(r[8]) if len(r) > 8 and r[8] is not None else 0,  # Not Reachable
            "ni":  0,
            "nr":  0,
        })
    return data


def parse_des_status(rows, start):
    """Design lead status bifurcation (same column layout as Eng)."""
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], datetime):
            break
        data.append({
            "d":   fmt_date(r[0]),
            "pa":  safe_num(r[1]),
            "su":  safe_num(r[2]),
            "cet": safe_num(r[3]),
            "d2y": safe_num(r[4]),
            "fu":  safe_num(r[5]),
            "fp":  safe_num(r[6]),
            "in":  safe_num(r[7]),
            "jk":  safe_num(r[8]) if len(r) > 8 and r[8] is not None else 0,
            "ni":  0,
            "nr":  0,
        })
    return data


def parse_ads(rows, start):
    """Google Ads (Eng or Des).
    Excel cols: Date | CTR% | Impressions | Avg CPC | Cost | Google Leads
    Template stores ctr * 100 (e.g. 16.15 → 1615.0).
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], datetime):
            break
        raw_ctr = safe_num(r[1], 0.0)
        data.append({
            "d":      fmt_date(r[0]),
            "ctr":    round(raw_ctr * 100, 2),   # template expects ctr×100
            "imp":    safe_num(r[2]),
            "cpc":    round(safe_num(r[3], 0.0), 2),
            "cost":   round(safe_num(r[4], 0.0), 2),
            "gLeads": safe_num(r[5]),
        })
    return data


def parse_walkins(rows, start):
    """Walk-ins. Cols: Date | Eng Walkins | Design Walkins | Total."""
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], datetime):
            break
        data.append({
            "d":   fmt_date(r[0]),
            "eng": safe_num(r[1]),
            "des": safe_num(r[2]),
        })
    return data


def parse_social(rows, start):
    """Social media weekly.
    Cols: Date-range string | Facebook | Instagram | Youtube | Linkedin |
          Posts/Reels | Paid Campaign | Paid Campaign Amount
    Template fields: period (full range string), fb, ig, yt, li, posts, paid, paidAmt
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None:
            break
        period = safe_str(r[0])
        if not period:
            break
        data.append({
            "period":  period,
            "fb":      safe_num(r[1]),
            "ig":      safe_num(r[2]),
            "yt":      safe_num(r[3]),
            "li":      safe_num(r[4]),
            "posts":   safe_num(r[5]),
            "paid":    safe_num(r[6]),
            "paidAmt": safe_num(r[7]),
        })
    return data


def parse_branding(rows, start):
    """Branding activities. Skip rows with no type/location/link."""
    data = []
    for r in rows[start + 2:]:
        if r[0] is None:
            break
        btype    = safe_str(r[1])
        location = safe_str(r[2])
        link     = safe_str(r[3])
        if not btype and not location and not link:
            continue
        ds = fmt_date(r[0]) or safe_str(r[0])
        data.append({"d": ds, "type": btype, "location": location, "link": link})
    return data


def parse_counselors(rows, start):
    """Counselor report — template only needs name, pa, prov.
    Excel cols: Name | Untouched | NI | Follow-Up | Interested | NR |
                CET | FP | DSY | New Leads | Junk | UCEED | Paid Apps | Prov
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], str):
            break
        name = safe_str(r[0])
        if not name:
            break
        data.append({
            "name": name,
            "pa":   safe_num(r[12]),   # Paid Applications
            "prov": safe_num(r[13]),   # Provisional Admissions
        })
    return data


def parse_notices(rows, start):
    """Notices/Circulars. Cols: Date | Category | Programme | Link."""
    data = []
    for r in rows[start + 2:]:
        if r[0] is None:
            break
        ds = fmt_date(r[0]) if isinstance(r[0], datetime) else (fmt_date(r[0]) or safe_str(r[0]))
        link  = safe_str(r[3])
        title = title_from_url(link) if link.startswith("http") else link
        data.append({
            "d":         ds,
            "category":  safe_str(r[1]),
            "programme": safe_str(r[2]),
            "title":     title,
            "link":      link,
        })
    return data


def parse_states(rows, start):
    """State-wise leads.
    Excel cols: State | Total | Design | Engineering
    Template fields: state, total, eng=Design(r[2]), des=Engineering(r[3])
    (template uses eng/des labels but maps to Design/Engineering columns respectively)
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], str):
            break
        state = safe_str(r[0])
        if not state or state.lower() in ("grand total", "state", "total"):
            break
        data.append({
            "state": state,
            "total": safe_num(r[1]),
            "eng":   safe_num(r[2]),   # Design column → template 'eng'
            "des":   safe_num(r[3]),   # Engineering column → template 'des'
        })
    return data


def parse_budget(rows, start):
    """Budget analysis. Extra blank row after header so data starts at start+4.
    Excel cols: Sr | School | Particulars | Date | Vendor | Budget Head |
                PO Number | Invoice Number | PO Amount | Advance | This Expenditure | ...
    Template fields: sr, school, desc, d, vendor, head, poNum, invNum, poAmt, advance, thisExp
    """
    data = []
    for r in rows[start + 4:]:
        if r[0] is None or not isinstance(r[0], (int, float)):
            break
        date_val = r[3]
        if isinstance(date_val, datetime):
            ds = fmt_date(date_val)
        else:
            ds = fmt_date(date_val) or safe_str(date_val)
        data.append({
            "sr":      int(safe_num(r[0])),
            "school":  safe_str(r[1], "MITAOE"),
            "desc":    safe_str(r[2]),
            "d":       ds,
            "vendor":  safe_str(r[4]),
            "head":    safe_str(r[5]),
            "poNum":   safe_str(r[6], ""),
            "invNum":  safe_str(r[7], ""),
            "poAmt":   null_or_num(r[8]),
            "advance": null_or_num(r[9]),
            "thisExp": null_or_num(r[10]),
        })
    return data


def parse_transactions(rows, start):
    """Transaction updates.
    Excel cols: Sr | Name | UG/PG | Program | Branch | Contact | Parent Contact |
                Email | Gender | Category | HSC | JEE/UCEED | CET/UCEED/JEE |
                Caution Fees | Counsellor | Payment Date | Payment Amount |
                Transaction ID | Mode | Verified Accounts | Admission Status | Refund Status
    Template fields: sr, name, ugpg, program, branch, gender, category,
                     cet(via), csl(counsellor), d, amt, txid, mode,
                     admStatus, refundStatus, hsc, jee
    """
    data = []
    for r in rows[start + 2:]:
        if r[0] is None or not isinstance(r[0], (int, float)):
            break
        date_val = r[15]
        if isinstance(date_val, datetime):
            ds = fmt_date(date_val)
        else:
            ds = fmt_date(date_val) or safe_str(date_val)
        data.append({
            "sr":           safe_str(r[0]),
            "name":         safe_str(r[1]),
            "ugpg":         safe_str(r[2]),
            "program":      safe_str(r[3]),
            "branch":       safe_str(r[4]),
            "gender":       safe_str(r[8]),
            "category":     safe_str(r[9]),
            "cet":          safe_str(r[12]),   # via: CET / UCEED / JEE
            "csl":          safe_str(r[14]),   # counsellor
            "d":            ds,
            "amt":          safe_num(r[16]),
            "txid":         safe_str(r[17]),
            "mode":         safe_str(r[18]),
            "admStatus":    safe_str(r[20]),
            "refundStatus": safe_str(r[21]) if len(r) > 21 and r[21] else "",
            "hsc":          null_or_num(r[10]),
            "jee":          null_or_num(r[11]),
        })
    return data


def parse_score_analysis(rows, start):
    """Score Analysis of Provisional Admissions (Design).
    Excel layout:
      start+0 : section header
      start+1 : col headers  (Scores | CET & UCEED | Only UCEED | Only CET)
      start+2 : Above 100
      start+3 : 75-100
      start+4 : 50-75
      start+5 : Below 50
      start+6 : Total Provisional Admissions
    """
    band_rows = rows[start + 2 : start + 6]   # 4 score bands
    total_row = rows[start + 6] if len(rows) > start + 6 else None

    bands = []
    for r in band_rows:
        if r[0] is None:
            break
        cu  = safe_num(r[1]) if r[1] not in (None, 'NA', 'N/A') else None
        uu  = safe_num(r[2]) if r[2] not in (None, 'NA', 'N/A') else None
        oc  = safe_num(r[3]) if r[3] not in (None, 'NA', 'N/A') else None
        bands.append({"label": safe_str(r[0]), "cu": cu, "uu": uu, "oc": oc})

    if total_row:
        kpi_cu    = safe_num(total_row[1])
        kpi_uu    = safe_num(total_row[2])
        kpi_oc    = safe_num(total_row[3])
    else:
        kpi_cu = sum(b["cu"] or 0 for b in bands)
        kpi_uu = sum(b["uu"] or 0 for b in bands)
        kpi_oc = sum(b["oc"] or 0 for b in bands)

    kpi_total = kpi_cu + kpi_uu + kpi_oc
    return {"kpi": {"cu": kpi_cu, "uu": kpi_uu, "oc": kpi_oc, "total": kpi_total},
            "bands": bands}


def patch_score_analysis(html, data):
    """Replace hardcoded Score Analysis numbers in the static HTML section."""
    MARKER = "<!-- ============ SCORE ANALYSIS ============ -->"
    sec_start = html.find(MARKER)
    if sec_start == -1:
        print("  [Score Analysis] Section marker not found — skipping patch")
        return html

    # Find end of section (next HTML comment block or </section>)
    sec_end = html.find("</section>", sec_start)
    if sec_end == -1:
        sec_end = len(html)
    else:
        sec_end += len("</section>")

    section = html[sec_start:sec_end]
    kpi   = data["kpi"]
    bands = data["bands"]  # [Above100, 75-100, 50-75, Below50]

    # ── KPI cards: each has a unique color on line-height:1; ─────────────────
    def repl_kpi(color, val, s):
        pat = rf'(color:{re.escape(color)};line-height:1;">)\d+(<)'
        return re.sub(pat, lambda m: f"{m.group(1)}{val}{m.group(2)}", s, count=1)

    section = repl_kpi("#1a56db", kpi["cu"],    section)   # CET & UCEED
    section = repl_kpi("#7c3aed", kpi["uu"],    section)   # Only UCEED
    section = repl_kpi("#059669", kpi["oc"],    section)   # Only CET
    section = repl_kpi("#1b2a5c", kpi["total"], section)   # Total

    # ── Table data rows: replace each band in order ──────────────────────────
    def fmt_val(v):
        return str(int(v)) if v is not None else "N/A"

    for band in bands:
        label = band["label"]
        cu_s  = fmt_val(band["cu"])
        uu_s  = fmt_val(band["uu"])
        oc_s  = fmt_val(band["oc"])

        # Find this band's label in the section
        lbl_idx = section.find(f'>{label}<')
        if lbl_idx == -1:
            continue

        # From label position, find the next 3 table cells and replace numbers
        after = section[lbl_idx:]

        def replace_cell(text, color_hint, new_val, occurrence=1):
            # Match td with font-size:22px (data cells) or italic N/A cell
            pat = r'(<td[^>]*>)(<span[^>]*>[^<]*</span>|[^<]+)(</td>)'
            count = [0]
            def replacer(m):
                count[0] += 1
                if count[0] == occurrence:
                    inner = m.group(2).strip()
                    if inner.lstrip('-').isdigit() or inner == 'N/A':
                        return m.group(1) + new_val + m.group(3)
                return m.group(0)
            return re.sub(pat, replacer, text, count=10)

        # Simpler: replace the 3 numeric td's right after the label row
        # Find the </tr> after the label, then the next <tr>
        row_end = after.find('</tr>')
        if row_end == -1:
            continue
        next_row_start = after.find('<tr', row_end)
        if next_row_start == -1:
            continue
        next_row_end   = after.find('</tr>', next_row_start) + len('</tr>')
        row_html = after[next_row_start:next_row_end]

        # Replace values: first numeric td → cu, second → uu, third → oc
        vals = [cu_s, uu_s, oc_s]
        val_idx = [0]
        def cell_replacer(m):
            inner = m.group(2).strip()
            if (inner.lstrip('-').isdigit() or inner == 'N/A') and val_idx[0] < 3:
                new = vals[val_idx[0]]
                val_idx[0] += 1
                return m.group(1) + new + m.group(3)
            return m.group(0)

        new_row = re.sub(r'(<td[^>]*>)([^<]+)(</td>)', cell_replacer, row_html)
        after = after[:next_row_start] + new_row + after[next_row_end:]
        section = section[:lbl_idx] + after

    # ── Total row: replace the 3 totals in the last summary row ──────────────
    total_marker = ">Total Provisional Admissions<"
    t_idx = section.find(total_marker)
    if t_idx != -1:
        after_total = section[t_idx:]
        t_row_end   = after_total.find('</tr>')
        if t_row_end != -1:
            t_next_start = after_total.find('<tr', t_row_end)
            if t_next_start != -1:
                t_next_end = after_total.find('</tr>', t_next_start) + len('</tr>')
                t_row_html = after_total[t_next_start:t_next_end]
                vals2 = [str(kpi["cu"]), str(kpi["uu"]), str(kpi["oc"])]
                vi = [0]
                def total_replacer(m):
                    inner = m.group(2).strip()
                    if inner.lstrip('-').isdigit() and vi[0] < 3:
                        new = vals2[vi[0]]; vi[0] += 1
                        return m.group(1) + new + m.group(3)
                    return m.group(0)
                new_t_row = re.sub(r'(<td[^>]*>)([^<]+)(</td>)', total_replacer, t_row_html)
                after_total = after_total[:t_next_start] + new_t_row + after_total[t_next_end:]
                section = section[:t_idx] + after_total

    print(f"  Score Analysis patched — CET+UCEED:{kpi['cu']}  OnlyUCEED:{kpi['uu']}  OnlyCET:{kpi['oc']}  Total:{kpi['total']}")
    return html[:sec_start] + section + html[sec_end:]


# ── inject helpers ────────────────────────────────────────────────────────────

def js_array(name, records):
    lines = [f"const {name} = ["]
    for rec in records:
        lines.append("  " + to_js(rec) + ",")
    lines.append("];")
    return "\n".join(lines)


def replace_raw(html, name, records):
    new_block = js_array(name, records)
    pattern   = rf"const {re.escape(name)}\s*=\s*\[.*?\];"
    result, n = re.subn(pattern, new_block, html, count=1, flags=re.DOTALL)
    status = f"{len(records)} rows" if n else "NOT FOUND IN TEMPLATE"
    print(f"  {name:<22} {status}")
    return result


# ── main ──────────────────────────────────────────────────────────────────────

def generate():
    print(f"Reading {EXCEL_FILE.name} (via temp copy)...")
    rows = read_rows()

    # ── find section boundaries dynamically ──────────────────────────────────
    print("Locating section boundaries...")
    s_main    = find_section(rows, "ADMISSION REPORT")
    s_eng_st  = find_section(rows, "ENGINEERING LEAD STATUS")
    s_des_st  = find_section(rows, "DESIGN LEADS STATUS")
    s_eng_ads = find_section(rows, "ENGINEERING GOOGLE ADS")
    s_des_ads = find_section(rows, "DESIGN GOOGLE ADS")
    s_walkins = find_section(rows, "Walkins Record")
    s_social  = find_section(rows, "SOCIAL MEDIA REPORT")
    s_branding= find_section(rows, "Branding")
    s_counsel = find_section(rows, "Counselor Report")
    s_notices = find_section(rows, "Notices/Circulars")
    s_states  = find_section(rows, "State-wise Leads")
    s_budget  = find_section(rows, "Budget Analysis")
    s_txn     = find_section(rows, "Transaction updates")
    s_score   = find_section(rows, "Score Analysis of Provisional")

    print(f"  Main={s_main}  EngSt={s_eng_st}  DesSt={s_des_st}  EngAds={s_eng_ads}")
    print(f"  DesAds={s_des_ads}  Walkins={s_walkins}  Social={s_social}  Branding={s_branding}")
    print(f"  Counselors={s_counsel}  Notices={s_notices}  States={s_states}  Budget={s_budget}  TXN={s_txn}")

    # ── parse ─────────────────────────────────────────────────────────────────
    print("Parsing sections...")
    main_data    = parse_main(rows, s_main)
    eng_status   = parse_eng_status(rows, s_eng_st)
    des_status   = parse_des_status(rows, s_des_st)
    eng_ads      = parse_ads(rows, s_eng_ads)
    des_ads      = parse_ads(rows, s_des_ads)
    walkins      = parse_walkins(rows, s_walkins)
    social       = parse_social(rows, s_social)
    branding     = parse_branding(rows, s_branding)
    counselors   = parse_counselors(rows, s_counsel)
    notices      = parse_notices(rows, s_notices)
    states       = parse_states(rows, s_states)
    budget       = parse_budget(rows, s_budget)
    transactions  = parse_transactions(rows, s_txn)
    score_data    = parse_score_analysis(rows, s_score)

    dates        = sorted([r["d"] for r in main_data if r["d"]])
    data_latest  = dates[-1] if dates else datetime.today().strftime("%Y-%m-%d")
    data_earliest = dates[0]  if dates else "2025-01-01"

    total_leads = sum(r["eLeads"] + r["dLeads"] for r in main_data)
    total_paid  = sum(r["ePaid"]  + r["dPaid"]  for r in main_data)
    total_sign  = sum(r["eSign"]  + r["dSign"]  for r in main_data)
    total_prov  = sum(r["eProv"]  + r["dProv"]  for r in main_data)
    print(f"  Totals — Leads:{total_leads}  PaidApps:{total_paid}  SignUps:{total_sign}  Prov:{total_prov}")
    print(f"  Latest date: {data_latest}  |  Main rows: {len(main_data)}")

    # ── read template & inject ────────────────────────────────────────────────
    print("Reading HTML template...")
    html = HTML_TEMPLATE.read_text(encoding="utf-8")

    print("Injecting data arrays...")
    html = replace_raw(html, "RAW_MAIN",       main_data)
    html = replace_raw(html, "RAW_ENG_STATUS", eng_status)
    html = replace_raw(html, "RAW_DES_STATUS", des_status)
    html = replace_raw(html, "RAW_ENG_ADS",    eng_ads)
    html = replace_raw(html, "RAW_DES_ADS",    des_ads)
    html = replace_raw(html, "RAW_WALKINS",    walkins)
    html = replace_raw(html, "RAW_SOCIAL",     social)
    html = replace_raw(html, "RAW_BRANDING",   branding)
    html = replace_raw(html, "RAW_COUNSELORS", counselors)
    html = replace_raw(html, "RAW_NOTICES",    notices)
    html = replace_raw(html, "RAW_STATES",     states)
    html = replace_raw(html, "RAW_BUDGET",     budget)
    html = replace_raw(html, "RAW_TXN",        transactions)

    # ── Patch static Score Analysis section with live Excel data ─────────────
    print("Patching Score Analysis...")
    html = patch_score_analysis(html, score_data)

    # Update DATA_LATEST everywhere
    html = re.sub(
        r"const DATA_LATEST\s*=\s*'[^']*'",
        f"const DATA_LATEST = '{data_latest}'",
        html
    )

    # ── Fix 1: filters const line (js object on one line) ────────────────────
    def fix_filters_line(line):
        line = re.sub(r"to:\s*'[^']*'",   f"to:'{data_latest}'",   line)
        line = re.sub(r"from:\s*'[^']*'", f"from:'{data_earliest}'", line)
        return line

    html_lines = html.splitlines()
    html_lines = [fix_filters_line(l) if l.strip().startswith('const filters') else l
                  for l in html_lines]
    html = "\n".join(html_lines)

    # ── Fix 2: DOMContentLoaded block — replaces ALL hardcoded to/from dates ──
    # e.g.  filters.eng.to = '2026-06-04';  or  to: '2026-06-04'
    html = re.sub(r"(\.to\s*=\s*)'[^']*'",   lambda m: m.group(1) + f"'{data_latest}'",   html)
    html = re.sub(r"(\.from\s*=\s*)'[^']*'", lambda m: m.group(1) + f"'{data_earliest}'", html)
    # Also object-literal form inside DOMContentLoaded: { from: '...', to: '...' }
    html = re.sub(r"(to:\s*)'([0-9]{4}-[0-9]{2}-[0-9]{2})'",
                  lambda m: m.group(1) + f"'{data_latest}'", html)
    html = re.sub(r"(from:\s*)'([0-9]{4}-[0-9]{2}-[0-9]{2})'",
                  lambda m: m.group(1) + f"'{data_earliest}'", html)

    # ── Fix 3: allFrom constant used in setPreset ─────────────────────────────
    html = re.sub(r"const allFrom\s*=\s*'[^']*'",
                  f"const allFrom = '{data_earliest}'", html)

    # ── Remove hard meta-refresh ──────────────────────────────────────────────
    html = re.sub(r'<meta\s+http-equiv=["\']refresh["\'][^>]*>', '', html)

    # ── Mobile-responsive CSS ─────────────────────────────────────────────────
    mobile_css = """<style id="mobile-responsive">
/* ═══════════════════════════════════════════════════
   MOBILE RESPONSIVE OVERRIDES
   Breakpoints: 768px (tablet), 480px (phone)
═══════════════════════════════════════════════════ */

/* ── Tablet (≤900px) ── already partially handled, reinforce ── */
@media (max-width: 900px) {
  .page-body { padding: 16px 16px; }
  .header { padding: 0 16px; }
  .header-inner { height: auto; min-height: 64px; padding: 10px 0; flex-wrap: wrap; gap: 8px; }
  .main-nav { padding: 0 12px; top: auto; position: relative; }
  .main-nav-inner { overflow-x: auto; -webkit-overflow-scrolling: touch; flex-wrap: nowrap; padding-bottom: 2px; }
  .main-nav-inner::-webkit-scrollbar { height: 3px; }
  .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 12px; }
  .chart-row { grid-template-columns: 1fr; gap: 14px; }
  .insights-grid { grid-template-columns: 1fr 1fr; gap: 12px; }
  .io-grid { grid-template-columns: 1fr; gap: 14px; }
  .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  .table-wrap table { min-width: 600px; }
}

/* ── Phone (≤600px) ── */
@media (max-width: 600px) {
  /* Header */
  .header-inner { flex-direction: row; flex-wrap: wrap; height: auto;
    padding: 10px 0; gap: 6px; justify-content: space-between; }
  .logo-block { gap: 10px; }
  .logo-block img { height: 38px; }
  .report-title { font-size: 15px; }
  .report-subtitle { font-size: 10px; }
  .header-divider { display: none; }
  .header-actions { gap: 6px; }
  .header-actions .btn { padding: 6px 10px; font-size: 11px; }
  .badge-live { display: none; }

  /* Nav tabs — horizontal scroll */
  .main-nav { padding: 0 8px; overflow: hidden; }
  .main-nav-inner { display: flex; overflow-x: auto; -webkit-overflow-scrolling: touch;
    flex-wrap: nowrap; gap: 0; padding-bottom: 2px; scrollbar-width: none; }
  .main-nav-inner::-webkit-scrollbar { display: none; }
  .main-tab { white-space: nowrap; padding: 12px 14px; font-size: 12px; gap: 5px; }
  .main-tab svg { width: 13px; height: 13px; }

  /* Sub-tabs */
  .sub-tabs { display: flex; overflow-x: auto; flex-wrap: nowrap;
    -webkit-overflow-scrolling: touch; gap: 6px; padding-bottom: 4px;
    scrollbar-width: none; }
  .sub-tabs::-webkit-scrollbar { display: none; }
  .sub-tab { white-space: nowrap; padding: 6px 12px; font-size: 12px; }

  /* Preset buttons */
  .preset-btns { display: flex; overflow-x: auto; flex-wrap: nowrap; gap: 6px;
    -webkit-overflow-scrolling: touch; padding-bottom: 4px; scrollbar-width: none; }
  .preset-btns::-webkit-scrollbar { display: none; }
  .preset-btn { white-space: nowrap; font-size: 11px; padding: 5px 10px; }

  /* Filter bar */
  .filter-bar { flex-direction: column; gap: 10px; align-items: stretch; }
  .filter-group { flex-direction: column; gap: 6px; align-items: stretch; }
  .filter-group label { font-size: 11px; }
  .filter-group input[type="date"] { width: 100%; font-size: 13px; padding: 6px 8px; }
  .filter-group select { width: 100%; font-size: 13px; padding: 6px 8px; }
  .btn-apply { width: 100%; justify-content: center; }

  /* KPI cards — 2 columns on phone */
  .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 16px; }
  .kpi-card { padding: 14px 12px; }
  .kpi-value { font-size: 22px; }
  .kpi-label { font-size: 11px; }
  .kpi-change { font-size: 10px; }

  /* Charts — single column, constrained height */
  .chart-row { grid-template-columns: 1fr; gap: 12px; }
  .chart-card { padding: 14px 12px; overflow: hidden; }
  .chart-card canvas { max-width: 100% !important; height: auto !important; }

  /* Insights / io grids */
  .insights-grid { grid-template-columns: 1fr; gap: 10px; }
  .io-grid { grid-template-columns: 1fr; gap: 12px; }

  /* Section header */
  .section-header { flex-direction: column; align-items: flex-start; gap: 8px; }
  .section-header h2 { font-size: 16px; }

  /* Tables — horizontal scroll */
  .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch;
    border-radius: 8px; }
  .table-wrap table { min-width: 560px; font-size: 12px; }
  .table-wrap th, .table-wrap td { padding: 8px 10px; white-space: nowrap; }

  /* Page body */
  .page-body { padding: 12px 10px; }

  /* Sect tabs (inside sections) */
  .sect-tabs { display: flex; overflow-x: auto; flex-wrap: nowrap; gap: 6px;
    -webkit-overflow-scrolling: touch; padding-bottom: 4px;
    scrollbar-width: none; }
  .sect-tabs::-webkit-scrollbar { display: none; }
  .sect-tab { white-space: nowrap; font-size: 12px; padding: 6px 12px; }

  /* Download button repositioned on small screens */
  #dl-btn { bottom: 14px; right: 14px; padding: 9px 14px; font-size: 12px; }
}

/* ── Very small phones (≤380px) ── */
@media (max-width: 380px) {
  .kpi-grid { grid-template-columns: repeat(2, 1fr); gap: 8px; }
  .kpi-value { font-size: 19px; }
  .report-title { font-size: 13px; }
  .main-tab { padding: 10px 10px; font-size: 11px; }
}
</style>"""

    # Inject mobile CSS into both live and share HTML
    html = html.replace("</head>", mobile_css + "\n</head>", 1)

    # ── Build the shareable version first (clean HTML, no poller, no dl btn) ──
    share_html = html  # html already has all data injected + dates fixed

    # ── Smart poller for live version ────────────────────────────────────────
    poller = """<script>
(function(){
  var _ver=null;
  function check(){
    fetch('/version.json?_='+Date.now())
      .then(function(r){return r.json();})
      .then(function(d){
        if(_ver===null){_ver=d.v;}
        else if(d.v!==_ver){location.reload();}
      }).catch(function(){});
  }
  setInterval(check,10000);
  check();
})();
</script>"""

    # ── Download button — fetches dashboard_share.html from server ───────────
    dl_filename = f"MIT_Admissions_Dashboard_{data_latest}.html"
    download_btn_css = """<style>
#dl-btn{position:fixed;bottom:24px;right:24px;z-index:9999;
  display:flex;align-items:center;gap:8px;
  background:#1b2a5c;color:#fff;border:none;border-radius:10px;
  padding:11px 20px;font-size:13px;font-weight:600;cursor:pointer;
  box-shadow:0 4px 16px rgba(27,42,92,0.35);transition:background 0.18s;}
#dl-btn:hover{background:#2d3f7c;}
#dl-btn svg{width:16px;height:16px;flex-shrink:0;}
</style>"""

    download_btn_html = f"""<button id="dl-btn" title="Download shareable snapshot" onclick="downloadDash()">
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"
       stroke-linecap="round" stroke-linejoin="round">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
    <polyline points="7 10 12 15 17 10"/>
    <line x1="12" y1="15" x2="12" y2="3"/>
  </svg>
  Download Dashboard
</button>
<script>
function downloadDash(){{
  var btn=document.getElementById('dl-btn');
  btn.textContent='Preparing...';btn.disabled=true;
  fetch('/dashboard_share.html?_='+Date.now())
    .then(function(r){{return r.blob();}})
    .then(function(blob){{
      var a=document.createElement('a');
      a.href=URL.createObjectURL(blob);
      a.download='{dl_filename}';
      document.body.appendChild(a);a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(a.href);
      btn.innerHTML='&#10003; Downloaded';
      setTimeout(function(){{
        btn.disabled=false;
        btn.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download Dashboard';
      }},2500);
    }})
    .catch(function(){{
      btn.textContent='Error – try again';btn.disabled=false;
    }});
}}
</script>"""

    # ── Assemble live dashboard (with poller + download button) ──────────────
    live_html = html
    live_html = live_html.replace("</head>", download_btn_css + "\n</head>", 1)
    live_html = live_html.replace("</body>", download_btn_html + "\n" + poller + "\n</body>", 1)

    # ── Write both files ──────────────────────────────────────────────────────
    print("Writing dashboard files...")
    OUTPUT_HTML.write_text(live_html, encoding="utf-8")
    share_path = OUTPUT_HTML.parent / "dashboard_share.html"
    share_path.write_text(share_html, encoding="utf-8")
    print(f"  Live   -> {OUTPUT_HTML.name}")
    print(f"  Share  -> {share_path.name}  (clean standalone, no poller)")
    print(f"Done -> {OUTPUT_HTML.parent}")


if __name__ == "__main__":
    generate()
