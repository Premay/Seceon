#!/usr/bin/env python3
"""
soc_pdf_to_excel.py
-------------------
Weekly Seceon "SOC Report" PDF  ->  filled Excel executive dashboard.

What it does
  1. Reads the PDF text layer and vector chart geometry (no OCR, no guessing
     from pictures), finds every dashboard widget by its card + title.
  2. Fills the "Data" sheet of your dashboard template (same 25 sections),
     resizes the chart ranges to the number of rows found, and refreshes the
     Dashboard header / reading notes.
  3. Adds two audit sheets:
       "PDF Extract"  every widget row as printed + cleaned value (copy-friendly)
       "Checks"       cross-checks (totals, percentages, peaks, truncation...)

Usage
  python soc_pdf_to_excel.py SOCReport_xxx.pdf
  python soc_pdf_to_excel.py SOCReport_xxx.pdf --template SOC_Dashboard_Template.xlsx --out result.xlsx

Needs:  pip install pdfplumber openpyxl
"""
import argparse
import copy
import os
import re
import sys
from collections import OrderedDict

import pdfplumber
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TEAL = (0.051, 0.5804, 0.5333)
DEFAULT_TEMPLATE = "SOC_Dashboard_Template.xlsx"
FALLBACK_EMAIL_DOMAIN = "unitedcapitalplcgroup.com"
FALLBACK_HOST_DOMAINS = ["ubacapitalgroup.com"]
MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()

# --------------------------------------------------------------------------
# Widget catalogue: regex on the printed title  ->  Data-sheet section
# (first match wins; order matters)
# --------------------------------------------------------------------------
WIDGETS = [
    (3,  r"threat indicator origin",                "donut"),
    (4,  r"user by failed login",                   "bars"),
    (5,  r"remote login success",                   "bars"),
    (6,  r"remote login failure",                   "bars"),
    (7,  r"unusual incoming connection",            "donut"),
    (8,  r"unusual outgoing connection",            "donut"),
    (9,  r"unlocked accounts",                      "bars"),
    (10, r"locked accounts",                        "bars"),
    (11, r"hosts with unusual activity duration",   "bars"),
    (111, r"hosts with unusual duration",           "bars"),   # duplicate of 11
    (12, r"unusual data volume\(outgoing",          "cloud"),
    (13, r"unusual data volume\(incoming",          "bars"),
    (14, r"websites by download",                   "bars"),
    (15, r"websites by upload",                     "bars"),
    (16, r"host by download",                       "bars"),
    (17, r"unusual duration behavior",              "bars"),
    (18, r"source pairs",                           "donut"),
    (19, r"destination pairs",                      "donut"),
    (20, r"alert severity distribution",            "donut"),
    (21, r"critical and major alert by status",     "donut"),
    (22, r"pattern by severity\s*-\s*major",        "line"),
    (23, r"pattern by severity\s*-\s*critical",     "line"),
    (24, r"critical alerts by alert type",          "bars"),
    (25, r"major alerts by alert type",             "bars"),
]

KPI_MAP = OrderedDict([  # printed label -> (template label, unit)
    ("threat indicators detected",   ("Threat Indicators Detected", "indicators")),
    ("total applications monitored", ("Total Applications Monitored", "applications")),
    ("total private hosts monitored", ("Total Private Hosts Monitored", "hosts")),
    ("total external hosts detected", ("Total External Hosts Detected", "hosts")),
    ("traffic monitored",            ("Traffic Monitored", "TB")),
    ("new hosts detected",           ("New Hosts Detected", "hosts")),
    ("total events analyzed",        ("Total Events Analyzed", "events")),
])

# Data sheet layout: section -> (first data row, unit, number format override)
SECTIONS = OrderedDict([
    (1,  dict(first=8,   unit="",                 hdr=6)),
    (2,  dict(first=19,  unit="",                 hdr=17)),
    (3,  dict(first=26,  unit="indicators",       hdr=24, share=True)),
    (4,  dict(first=34,  unit="attempts",         hdr=32)),
    (5,  dict(first=48,  unit="logins",           hdr=46)),
    (6,  dict(first=62,  unit="attempts",         hdr=60)),
    (7,  dict(first=76,  unit="duration (raw)",   hdr=74, share=True, fmt="#,##0")),
    (8,  dict(first=85,  unit="duration (raw)",   hdr=83, share=True, fmt="#,##0")),
    (9,  dict(first=93,  unit="count",            hdr=91)),
    (10, dict(first=102, unit="lockouts",         hdr=100)),
    (11, dict(first=116, unit="duration",         hdr=114, fmt="#,##0.##")),
    (12, dict(first=130, unit="n/a (word cloud)", hdr=128)),
    (13, dict(first=139, unit="GB",               hdr=137, fmt="#,##0.0###")),
    (14, dict(first=153, unit="GB",               hdr=151, fmt="#,##0.0")),
    (15, dict(first=167, unit="GB",               hdr=165, fmt="#,##0.0")),
    (16, dict(first=181, unit="GB",               hdr=179, fmt="#,##0.0")),
    (17, dict(first=195, unit="duration",         hdr=193, fmt="#,##0.##")),
    (18, dict(first=209, unit="indicators",       hdr=207, share=True)),
    (19, dict(first=223, unit="indicators",       hdr=221, share=True)),
    (20, dict(first=237, unit="alerts",           hdr=235, share=True)),
    (21, dict(first=244, unit="alerts",           hdr=242, share=True)),
    (22, dict(first=250, unit="count",            hdr=248)),
    (23, dict(first=257, unit="count",            hdr=255)),
    (24, dict(first=265, unit="alerts",           hdr=263)),
    (25, dict(first=271, unit="alerts",           hdr=269)),
])


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------
def close(a, b, tol=0.02):
    return a is not None and len(a) == len(b) and all(abs(x - y) <= tol for x, y in zip(a, b))


def norm(s):
    return re.sub(r"\s+", " ", s.strip().lower())


NUM_RE = r"(?:\d[\d,]*\.?\d*|\.\d+)"
ROW_RE = re.compile(
    rf"^(?P<label>.+?)\s+(?P<val>{NUM_RE}(?:\s?(?:[KMB]|GB|MB|KB|TB))?)(?:\s+(?P<pct>\d+(?:\.\d+)?)%)?$")


def to_number(text):
    """'5.3B' -> (5.3e9, None, True) ; '15.7 GB' -> (15.7, 'GB', False)"""
    t = text.strip().replace(",", "")
    m = re.match(r"^(\d*\.?\d+)\s?(K|M|B|GB|MB|KB|TB)?$", t)
    if not m:
        return None, None, False
    v, suf = float(m.group(1)), m.group(2)
    if suf in ("K", "M", "B"):
        return v * {"K": 1e3, "M": 1e6, "B": 1e9}[suf], None, True
    return v, suf, False


def to_gb(value, unit):
    f = {"KB": 1e-6, "MB": 1e-3, "GB": 1.0, "TB": 1e3}.get(unit)
    return None if f is None else value * f


def clean_int(v):
    return int(round(v)) if abs(v - round(v)) < 1e-6 else v


def cluster_rows(words, tol=3.0):
    rows = []
    for w in sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2, w["x0"])):
        yc = (w["top"] + w["bottom"]) / 2
        if rows and abs(rows[-1]["y"] - yc) <= tol:
            rows[-1]["words"].append(w)
            rows[-1]["y"] = (rows[-1]["y"] + yc) / 2
        else:
            rows.append({"y": yc, "words": [w]})
    for r in rows:
        r["words"].sort(key=lambda w: w["x0"])
        r["text"] = " ".join(w["text"] for w in r["words"])
    return rows


def is_bold(w):
    return "Bold" in w["fontname"]


# --------------------------------------------------------------------------
# PDF parsing
# --------------------------------------------------------------------------
class Widget:
    def __init__(self, page, title, subtitle, kind):
        self.page, self.title, self.subtitle, self.kind = page, title, subtitle, kind
        self.section = None
        self.rows = []          # dicts: label, shown, value, unit, pct, rounded, note
        self.total_shown = None
        self.empty = False
        self.notes = []


def find_cards(page):
    cards = {}
    for o in page.curves + page.rects:
        if not o.get("fill"):
            continue
        col = o.get("non_stroking_color")
        if not (isinstance(col, (tuple, list)) and len(col) == 3 and all(abs(c - 1) < 0.01 for c in col)):
            continue
        w, h = o["x1"] - o["x0"], o["bottom"] - o["top"]
        if w < 100 or h < 40:
            continue
        key = tuple(round(v) for v in (o["x0"], o["top"], o["x1"], o["bottom"]))
        cards[key] = key
    boxes = list(cards)
    keep = []
    for b in boxes:   # drop containers that hold other cards
        inner = [c for c in boxes if c != b and c[0] >= b[0] - 1 and c[1] >= b[1] - 1
                 and c[2] <= b[2] + 1 and c[3] <= b[3] + 1]
        if not inner:
            keep.append(b)
    return sorted(keep, key=lambda b: (b[1], b[0]))


def words_in(words, box, pad=1):
    x0, t, x1, b = box
    return [w for w in words
            if x0 - pad <= (w["x0"] + w["x1"]) / 2 <= x1 + pad and t - pad <= (w["top"] + w["bottom"]) / 2 <= b + pad]


def parse_row_text(text):
    m = ROW_RE.match(text)
    if not m:
        return None
    return m.group("label").strip(), m.group("val").strip(), m.group("pct")


def parse_bars_or_donut(widget, body_words):
    # donut centre: biggest bold number + "Total"
    big = [w for w in body_words if is_bold(w) and w["size"] >= 10]
    for w in big:
        widget.total_shown = w["text"]
    body = [w for w in body_words if w not in big and w["text"] != "Total"]
    if any("No data in this period" in " ".join(x["text"] for x in body_words[i:i + 5])
           for i in range(len(body_words))):
        widget.empty = True
        return
    for r in cluster_rows(body):
        p = parse_row_text(r["text"])
        if not p:
            widget.notes.append(f"unparsed row: {r['text']!r}")
            continue
        label, val_txt, pct = p
        value, unit, rounded = to_number(val_txt)
        widget.rows.append(dict(label=label, shown=val_txt, value=value, unit=unit, pct=pct,
                                rounded=rounded, note=""))


def parse_cloud(widget, body_words):
    if any(w["text"].lower() == "no" for w in body_words) and not any(
            w["size"] > 8 for w in body_words):
        pass
    terms = sorted(body_words, key=lambda w: (-w["size"], w["x0"]))
    for w in terms:
        widget.rows.append(dict(label=w["text"], shown="", value=None, unit=None, pct=None,
                                rounded=False, note=f"font size {w['size']:.1f}"))
    widget.notes.append("Word cloud: the PDF prints no numbers; terms listed largest-first.")


def parse_line(widget, page, box, body_words):
    objs = []
    for o in page.lines + page.curves:
        if o.get("fill"):
            continue
        col = o.get("stroking_color")
        if not (isinstance(col, (tuple, list)) and len(col) == 3 and close(col, TEAL, 0.02)):
            continue
        if not (box[0] <= o["x0"] and o["x1"] <= box[2] + 1 and box[1] <= o["top"] and o["bottom"] <= box[3] + 1):
            continue
        pts = o.get("pts") or []
        if len(pts) >= 2:
            objs.append(o)
    if not objs:
        widget.notes.append("no line geometry found")
        return
    line = max(objs, key=lambda o: (len(o["pts"]), o["x1"] - o["x0"]))
    pts = [(float(x), float(y)) for x, y in line["pts"]]
    pts = [p for i, p in enumerate(pts) if i == 0 or p != pts[i - 1]]

    # x labels (dates)
    labels = []
    for i, w in enumerate(body_words):
        if w["text"] in MONTHS and 0 < i < len(body_words) - 1:
            d, y = body_words[i - 1], body_words[i + 1]
            if re.fullmatch(r"\d{1,2}", d["text"]) and re.fullmatch(r"\d{4}", y["text"]):
                labels.append(dict(text=f"{d['text']} {w['text']} {y['text']}",
                                   xc=(d["x0"] + y["x1"]) / 2,
                                   words={id(d), id(w), id(y)}))
    used = set().union(*[l["words"] for l in labels]) if labels else set()

    # peak
    peak = None
    for i, w in enumerate(body_words):
        if w["text"] == "Peak" and i + 1 < len(body_words):
            peak = body_words[i + 1]["text"]
    peak_ids = {id(w) for i, w in enumerate(body_words)
                if w["text"] == "Peak" or (i > 0 and body_words[i - 1]["text"] == "Peak")}

    # y ticks
    left = min(p[0] for p in pts)
    ticks = [((w["top"] + w["bottom"]) / 2, float(w["text"])) for w in body_words
             if re.fullmatch(r"\d+", w["text"]) and not is_bold(w) and id(w) not in used
             and id(w) not in peak_ids and w["x1"] <= left + 2]
    if len(ticks) < 2:
        widget.notes.append("could not calibrate y axis")
        return
    n = len(ticks)
    my, mv = sum(t[0] for t in ticks) / n, sum(t[1] for t in ticks) / n
    slope = sum((y - my) * (v - mv) for y, v in ticks) / sum((y - my) ** 2 for y, v in ticks)
    for i, (x, y) in enumerate(pts):
        val = mv + slope * (y - my)
        val = int(round(val)) if abs(val - round(val)) < 0.45 else round(val, 1)
        lab = min(labels, key=lambda l: abs(l["xc"] - x))["text"] if labels else f"point {i + 1}"
        widget.rows.append(dict(label=lab, shown=str(val), value=val, unit=None, pct=None,
                                rounded=False, note="read from chart geometry"))
    widget.peak_shown = peak


def parse_pdf(path):
    pdf = pdfplumber.open(path)
    out = dict(header={}, kpis=[], widgets=[], stats=[], pages=len(pdf.pages), footer_count=None)
    for pno, page in enumerate(pdf.pages, 1):
        words = page.extract_words(extra_attrs=["fontname", "size"])
        text = page.extract_text() or ""
        m = re.search(r"from (\d+) dashboard widgets", text)
        if m:
            out["footer_count"] = int(m.group(1))

        if pno == 1:   # header band
            band = [w for w in words if 125 <= w["top"] <= 165]
            labs = [w for w in band if w["text"] in ("PREPARED", "REPORTING", "GENERATED")]
            for lw in labs:
                out["header"][lw["text"]] = []
            starts = sorted((lw["x0"], lw["text"]) for lw in labs)
            for w in band:
                if re.fullmatch(r"[A-Z]+", w["text"]) and w["size"] < 7 and w["text"] in ("PREPARED", "FOR", "REPORTING", "PERIOD", "GENERATED"):
                    continue
                col = [n for x, n in starts if x <= w["x0"] + 1]
                if col:
                    out["header"][col[-1]].append(w)
            out["header"] = {k: " ".join(w["text"] for w in sorted(v, key=lambda w: (round(w["top"]), w["x0"])))
                             for k, v in out["header"].items()}

        for box in find_cards(page):
            cw = words_in(words, box)
            titles = [w for w in cw if is_bold(w) and 8.3 <= w["size"] <= 9.3 and w["top"] <= box[1] + 30]
            if not titles:      # KPI / stat card (big number, caps label)
                big = [w for w in cw if is_bold(w) and w["size"] >= 12]
                if not big:
                    continue
                value = " ".join(w["text"] for w in sorted(big, key=lambda w: w["x0"]))
                label = " ".join(w["text"] for w in sorted(
                    [w for w in cw if w not in big and re.fullmatch(r"[A-Z#]+", w["text"])],
                    key=lambda w: (round(w["top"]), w["x0"])))
                exact = [w["text"] for w in cw if w not in big and re.fullmatch(r"[\d,]+", w["text"])]
                rec = dict(page=pno, label=label, shown=value, exact=exact[0] if exact else None)
                (out["kpis"] if pno == 1 else out["stats"]).append(rec)
                continue
            title = " ".join(w["text"] for w in sorted(titles, key=lambda w: (round(w["top"]), w["x0"])))
            sub_words = [w for w in cw if not is_bold(w) and 6.8 <= w["size"] <= 7.3 and w["top"] <= box[1] + 36
                         and w not in titles and not re.fullmatch(r"[\d,.]+[KMB]?", w["text"])]
            # subtitle sits in the upper-right corner of the card
            sub_words = [w for w in sub_words if w["x0"] > box[0] + (box[2] - box[0]) * 0.45]
            subtitle = " ".join(w["text"] for w in sorted(sub_words, key=lambda w: (round(w["top"]), w["x0"])))
            body = [w for w in cw if w not in titles and w not in sub_words]

            sec, kind = None, "bars"
            for s, rx, k in WIDGETS:
                if re.search(rx, title.lower()):
                    sec, kind = s, k
                    break
            wd = Widget(pno, title, subtitle, kind)
            wd.section = sec
            if kind == "line":
                parse_line(wd, page, box, body)
            elif kind == "cloud":
                parse_cloud(wd, body)
            else:
                parse_bars_or_donut(wd, body)
            out["widgets"].append(wd)
    pdf.close()
    return out


# --------------------------------------------------------------------------
# label cleaning (the PDF truncates long names with "…")
# --------------------------------------------------------------------------
def build_vocab(parsed, template_labels):
    vocab = set(template_labels)
    for w in parsed["widgets"]:
        for r in w.rows:
            if not r["label"].endswith("…"):
                vocab.add(r["label"])
    return vocab


def resolve_label(label, vocab, email_domain, host_domains):
    if not label.endswith("…"):
        return label, ""
    stem = label[:-1]
    if "@" in stem:
        local, part = stem.split("@", 1)
        if email_domain.startswith(part):
            return f"{local}@{email_domain}", f"truncated in PDF; domain restored as @{email_domain}"
    if "." in stem and not re.fullmatch(r"[\d.]+", stem):
        host, part = stem.split(".", 1)
        cands = [d for d in host_domains if d.startswith(part)]
        if len(cands) == 1:
            return f"{host}.{cands[0]}", f"truncated in PDF; domain restored as .{cands[0]}"
    cands = sorted(v for v in vocab if v.lower().startswith(stem.lower()) and v != label)
    if len(cands) == 1:
        return cands[0], "truncated in PDF; completed from another widget / earlier week"
    return label, "TRUNCATED in PDF and could not be completed"


# --------------------------------------------------------------------------
# build cleaned dataset + checks
# --------------------------------------------------------------------------
def clean_dataset(parsed, vocab, email_domain, host_domains):
    secs = {}
    for w in parsed["widgets"]:
        for r in w.rows:
            r["resolved"], r["label_note"] = resolve_label(r["label"], vocab, email_domain, host_domains)
        if w.section is not None:
            secs[w.section] = w
    return secs


def run_checks(parsed, secs, kpi_exact):
    checks = []

    def add(status, name, detail):
        checks.append((status, name, detail))

    n_titled = len(parsed["widgets"])
    n_all = n_titled + len(parsed["stats"])
    fc = parsed["footer_count"]
    if fc is None:
        add("WARN", "Widget count", f"Found {n_all}; the PDF footer does not state a widget count.")
    elif fc == n_all:
        add("PASS", "Widget count", f"{n_all} widgets found = {fc} stated in the PDF footer.")
    else:
        add("FAIL", "Widget count", f"Found {n_all} widgets but the PDF footer says {fc}. A widget may have moved or been renamed.")

    unmapped = [w.title for w in parsed["widgets"] if w.section is None]
    add("PASS" if not unmapped else "WARN", "All widgets recognised",
        "Every widget title matched the catalogue." if not unmapped else
        "Not mapped to the Data sheet (still in 'PDF Extract'): " + "; ".join(unmapped))

    for w in parsed["widgets"]:
        if w.kind == "donut" and w.rows:
            vals = [r["value"] for r in w.rows if r["value"] is not None]
            tot, _, _ = to_number(w.total_shown or "")
            if tot:
                s = sum(vals)
                diff = abs(s - tot) / tot
                add("PASS" if diff <= 0.02 else "WARN", f"Donut total - {w.title}",
                    f"rows sum {s:,.0f} vs centre total {w.total_shown} ({diff:.1%} apart; headline totals are rounded).")
            pcts = [float(r["pct"]) for r in w.rows if r["pct"]]
            if pcts:
                add("PASS" if abs(sum(pcts) - 100) <= 1.5 else "WARN", f"Percent total - {w.title}",
                    f"printed percentages add up to {sum(pcts):.1f}%.")
        if w.kind == "bars" and w.rows:
            vals = [r["value"] if r["unit"] not in ("KB", "MB", "GB", "TB") else to_gb(r["value"], r["unit"]) for r in w.rows]
            if any(v is None for v in vals):
                continue
            ok = all(a >= b - 1e-9 for a, b in zip(vals, vals[1:]))
            if not ok:
                add("WARN", f"Sort order - {w.title}", "Rows are not in descending order; check for a misread value.")
            if len(w.rows) > 10:
                add("WARN", f"Row count - {w.title}", f"{len(w.rows)} rows (expected at most 10).")
        if w.kind == "line":
            vals = [r["value"] for r in w.rows if r["value"] is not None]
            pk = getattr(w, "peak_shown", None)
            if vals and pk and re.fullmatch(r"\d+(\.\d+)?", pk):
                add("PASS" if abs(max(vals) - float(pk)) < 0.51 else "FAIL", f"Peak - {w.title}",
                    f"max of extracted points {max(vals)} vs printed peak {pk}.")
            elif not vals:
                add("FAIL", f"Line chart - {w.title}", "No points could be read.")
        for n in w.notes:
            if n.startswith(("unparsed", "no line", "could not")):
                add("FAIL", f"Parse - {w.title}", n)

    # cross-widget identities
    sev, sta = secs.get(20), secs.get(21)
    if sev and sta and sev.rows and sta.rows:
        crit_major = sum(r["value"] for r in sev.rows if r["resolved"].lower() in ("critical", "major"))
        st_total = sum(r["value"] for r in sta.rows)
        add("PASS" if crit_major == st_total else "WARN", "Critical+Major = status total",
            f"Critical + Major = {crit_major:g}; Open + Closed = {st_total:g}.")
        tot, _, _ = to_number(sev.total_shown or "")
        if tot:
            s = sum(r["value"] for r in sev.rows)
            add("PASS" if s == tot else "FAIL", "Severity total", f"Minor+Major+Critical = {s:g} vs centre total {tot:g}.")
    for sec_id, sev_name in ((24, "critical"), (25, "major")):
        w = secs.get(sec_id)
        if w and sev:
            tot = next((r["value"] for r in sev.rows if r["resolved"].lower() == sev_name), None)
            s = sum(r["value"] for r in w.rows)
            if tot is not None:
                st = "PASS" if s == tot else ("WARN" if s < tot else "FAIL")
                add(st, f"{sev_name.title()} alerts by type = {sev_name} total",
                    f"type counts sum to {s:g}; severity widget says {tot:g}."
                    + ("  (Top-10 list can omit small types.)" if s < tot else ""))
    o = secs.get(3)
    if o and kpi_exact.get("Threat Indicators Detected"):
        s = sum(r["value"] for r in o.rows)
        ex = kpi_exact["Threat Indicators Detected"]
        add("PASS" if s == ex else "WARN", "Origin split = Threat Indicators KPI", f"origin rows sum {s:,.0f} vs KPI {ex:,.0f}.")
    a, b = secs.get(11), secs.get(111)
    if a and b:
        same = [(r["resolved"], r["shown"]) for r in a.rows] == [(r["resolved"], r["shown"]) for r in b.rows]
        add("PASS" if same else "WARN", "Duplicate widgets agree", "The two 'Hosts with Unusual Duration' widgets " + ("match." if same else "DIFFER."))
    a, b = secs.get(4), secs.get(6)
    if a and b:
        same = [(r["resolved"], r["shown"]) for r in a.rows] == [(r["resolved"], r["shown"]) for r in b.rows]
        add("INFO", "Remote Login Failure vs Failed Login Attempts",
            "identical series in this PDF." if same else "different series this week.")
    for w in parsed["widgets"]:
        for r in w.rows:
            if r["label_note"].startswith("TRUNCATED"):
                add("WARN", "Truncated label", f"{w.title}: {r['label']}")
        if "download" in w.title.lower() and "upload" in w.subtitle.lower() or \
           "upload" in w.title.lower() and "download" in w.subtitle.lower():
            add("INFO", "Title/subtitle conflict", f"{w.title!r} is labelled {w.subtitle!r} in the PDF - confirm direction in Seceon.")
    for w in parsed["widgets"]:
        if w.empty:
            add("INFO", "Empty widget", f"{w.title}: 'No data in this period.'")
    return checks


# --------------------------------------------------------------------------
# Excel writing
# --------------------------------------------------------------------------
def copy_style(src, dst):
    dst._style = copy.copy(src._style)


def set_title_text(title, fn):
    try:
        for p in title.tx.rich.p:
            for r in (p.r or []):
                r.t = fn(r.t)
    except AttributeError:
        pass


def retarget_chart(chart, first, n):
    last = first + max(n, 1) - 1
    for ser in chart.series:
        refs = []
        if ser.cat is not None:
            refs += [ser.cat.strRef, ser.cat.numRef]
        if ser.val is not None:
            refs += [ser.val.numRef]
        for ref in refs:
            if ref is None:
                continue
            ref.f = re.sub(r"\$(\d+):\$([A-Z]+)\$(\d+)",
                           lambda m: f"${first}:${m.group(2)}${last}", ref.f)
            for cache in ("strCache", "numCache"):
                if hasattr(ref, cache):
                    setattr(ref, cache, None)


def fill_section(ws, sec_id, spec, rows, next_hdr, page_no):
    first = spec["first"]
    cap = (next_hdr - first - 1) if next_hdr else max(len(rows), 10)
    if len(rows) > cap:
        print(f"  ! section {sec_id}: {len(rows)} rows but template has room for {cap}; extra rows dropped")
        rows = rows[:cap]
    ref_row = first
    # remember pattern formulas from the first row
    h_formula = ws.cell(ref_row, 8).value
    i_formula_needed = spec.get("share")
    for k in range(max(cap, len(rows))):
        r = first + k
        for col in range(2, 10):
            copy_style(ws.cell(ref_row, col), ws.cell(r, col))
        if k < len(rows):
            d = rows[k]
            ws.cell(r, 2, d["label"])
            ws.cell(r, 3, d["value"])
            ws.cell(r, 4, d["shown"])
            ws.cell(r, 5, d["unit"])
            ws.cell(r, 6, page_no)
            ws.cell(r, 7, d["note"])
        else:
            for col in range(2, 8):
                ws.cell(r, col).value = None
        if isinstance(h_formula, str) and h_formula.startswith("="):
            ws.cell(r, 8).value = re.sub(rf"\bB{ref_row}\b", f"B{r}", h_formula) if k else h_formula
        if i_formula_needed:
            last = first + max(len(rows), 1) - 1
            ws.cell(r, 9).value = (f"=IFERROR(C{r}/SUM($C${first}:$C${last}),0)" if k < len(rows) else None)
        if spec.get("fmt") and k < len(rows):
            ws.cell(r, 3).number_format = spec["fmt"]
    return len(rows)


def prepare_rows(sec_id, widget, spec, drop_other=True):
    """Convert parsed widget rows into rows for the Data sheet."""
    out = []
    if widget is None:
        return out
    for r in widget.rows:
        label = r["resolved"]
        note_bits = []
        if drop_other and re.match(r"^Other \(\d+\)$", label):
            continue
        if r["label_note"]:
            note_bits.append(r["label_note"])
        value, unit, shown = r["value"], spec["unit"], r["shown"]
        if sec_id in (13, 14, 15, 16, 12) and r["unit"] in ("KB", "MB", "GB", "TB"):
            value = to_gb(r["value"], r["unit"])
            shown = f"{r['shown']} {r['unit']}" if r["unit"] not in r["shown"] else r["shown"]
            if r["unit"] != "GB":
                note_bits.append(f"converted from {r['unit']} to GB (1 GB = 1,000 MB)")
        elif sec_id in (13, 14, 15, 16):
            shown = r["shown"]
        if r["rounded"]:
            note_bits.append("rounded as printed")
        if r["pct"] is not None:
            shown = f"{r['shown']} - {r['pct']}%"
        if sec_id in (22, 23):
            m = re.match(r"(\d{1,2}) (\w{3}) (\d{4})", r["resolved"])
            if m:
                label = f"{int(m.group(1)):02d}/{MONTHS.index(m.group(2)) + 1:02d}/{m.group(3)}"
                if m.group(3) == "1970":
                    note_bits.append("PDF plots a 1970 epoch timestamp (Seceon data artefact); count kept")
            note_bits.insert(0, "read from chart geometry")
            shown = f"{value}"
        if sec_id == 21:
            label = label.title()
        if sec_id == 12:
            value, shown = None, ""
            note_bits = ["word cloud in PDF - no values printed; ordered largest font first"]
        if value is not None:
            value = clean_int(value)
        if r["unit"] and sec_id in (7, 8, 11, 17, 18, 19):
            pass
        out.append(dict(label=label, value=value, shown=shown, unit=unit,
                        note="; ".join(note_bits) if note_bits else "Clear"))
    return out


def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="1F3A53")
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def write_extract_sheet(wb, parsed, secs):
    if "PDF Extract" in wb.sheetnames:
        del wb["PDF Extract"]
    ws = wb.create_sheet("PDF Extract")
    ws.append(["PDF page", "Widget (as printed)", "Data section", "#", "Label (as printed)",
               "Label (cleaned)", "Value (as printed)", "Value (number)", "Unit", "% (as printed)", "Notes"])
    style_header_row(ws, 1, 11)
    for k in parsed["kpis"]:
        label = k["label"].title()
        ws.append([k["page"], "Overview KPI", 1, None, k["label"], label, k["shown"],
                   to_number(k["exact"] or k["shown"].replace(" TB", ""))[0], "", None,
                   f"exact figure printed under headline: {k['exact']}" if k["exact"] else ""])
    for k in parsed["stats"]:
        ws.append([k["page"], "Stat card", None, None, k["label"], k["label"].title(), k["shown"],
                   to_number(k["exact"] or k["shown"])[0], "", None,
                   f"exact figure: {k['exact']}" if k["exact"] else ""])
    for w in parsed["widgets"]:
        if w.empty:
            ws.append([w.page, w.title, w.section, None, "(no data in this period)", "", "", None, "", None, ""])
        for i, r in enumerate(w.rows, 1):
            num = r["value"]
            ws.append([w.page, w.title, w.section if w.section != 111 else 11, i, r["label"], r["resolved"],
                       r["shown"], num, r["unit"] or "", float(r["pct"]) if r["pct"] else None,
                       "; ".join(x for x in (r["label_note"], r["note"]) if x)])
    widths = [9, 44, 12, 5, 44, 52, 16, 18, 8, 14, 60]
    for i, w_ in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w_
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=10)
        row[7].number_format = "#,##0.##"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_checks_sheet(wb, checks):
    if "Checks" in wb.sheetnames:
        del wb["Checks"]
    ws = wb.create_sheet("Checks")
    ws.append(["Status", "Check", "Detail"])
    style_header_row(ws, 1, 3)
    colors = {"PASS": "C6EFCE", "WARN": "FFEB9C", "FAIL": "FFC7CE", "INFO": "DDEBF7"}
    for st, name, detail in checks:
        ws.append([st, name, detail])
        c = ws.cell(ws.max_row, 1)
        c.fill = PatternFill("solid", fgColor=colors[st])
        c.font = Font(name="Arial", bold=True, size=10)
        for col in (2, 3):
            ws.cell(ws.max_row, col).font = Font(name="Arial", size=10)
            ws.cell(ws.max_row, col).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 110
    ws.freeze_panes = "A2"


def build_workbook(parsed, template, out_path, pdf_name):
    wb = openpyxl.load_workbook(template)
    required_sheets = {"Data", "Executive Dashboard"}
    missing_sheets = required_sheets.difference(wb.sheetnames)
    if missing_sheets:
        raise ValueError(
            "The Excel template is missing required sheet(s): "
            + ", ".join(sorted(missing_sheets))
        )
    data, dash = wb["Data"], wb["Executive Dashboard"]

    # vocabulary from the template + domains
    tpl_labels = [str(data.cell(r, 2).value) for r in range(1, data.max_row + 1)
                  if isinstance(data.cell(r, 2).value, str) and not str(data.cell(r, 2).value).startswith("=")]
    email_domain = next((l.split("@", 1)[1] for l in tpl_labels if "@" in l and " " not in l), FALLBACK_EMAIL_DOMAIN)
    host_domains = sorted(set(FALLBACK_HOST_DOMAINS + [m.group(1).lower() for l in tpl_labels
                          for m in [re.search(r"\.([a-z0-9-]+\.[a-z]{2,})$", l, re.I)] if m and not re.fullmatch(r"[\d.]+", l)]))
    vocab = build_vocab(parsed, tpl_labels)
    secs = clean_dataset(parsed, vocab, email_domain, host_domains)

    # ----- header / KPIs -----
    hdr = parsed["header"]
    org, rid = hdr.get("PREPARED", ""), ""
    m = re.match(r"^(.*?)\s*\((\w+)\)$", org)
    if m:
        org, rid = m.group(1), m.group(2)
    period = hdr.get("REPORTING", "").replace("–", "-")
    kpi_exact = {}
    kpi_rows = []
    for k in parsed["kpis"]:
        key = norm(k["label"])
        if key in KPI_MAP:
            label, unit = KPI_MAP[key]
            if k["exact"]:
                val = to_number(k["exact"])[0]
            else:
                val, u2, _ = to_number(k["shown"].replace(" TB", ""))
            if unit == "TB":
                val = to_number(k["shown"].replace(" TB", ""))[0]
            kpi_rows.append((label, clean_int(val), k["shown"], unit,
                             "Exact figure printed under headline" if k["exact"] else "Source text (as printed)"))
            kpi_exact[label] = val
    order = [v[0] for v in KPI_MAP.values()]
    for i, label in enumerate(order):
        r = 8 + i
        match = next((x for x in kpi_rows if x[0] == label), None)
        if match:
            data.cell(r, 2, match[0]); data.cell(r, 3, match[1]); data.cell(r, 4, match[2])
            data.cell(r, 5, match[3]); data.cell(r, 6, 1); data.cell(r, 7, match[4])
        else:
            print(f"  ! KPI '{label}' not found in PDF")
    data["C19"], data["C20"], data["C21"] = org, rid, period
    for r in (19, 20, 21):
        data.cell(r, 6, 1)
    # generated-at row (spare row between sections 2 and 3)
    copy_style(data["B21"], data["B22"]); copy_style(data["C21"], data["C22"]); copy_style(data["F21"], data["F22"])
    data["B22"], data["C22"], data["F22"] = "Generated", hdr.get("GENERATED", ""), 1

    # ----- sections -----
    pages_found = {}
    hdr_rows = sorted(s["hdr"] for s in SECTIONS.values())
    for sec_id, spec in SECTIONS.items():
        if sec_id in (1, 2):
            continue
        widget = secs.get(sec_id)
        if widget is None and sec_id == 11:
            widget = secs.get(111)
        rows = prepare_rows(sec_id, widget, spec)
        nxt = next((h for h in hdr_rows if h > spec["hdr"]), None)
        page_no = widget.page if widget else ""
        pages_found[sec_id] = page_no
        n = fill_section(data, sec_id, spec, rows, nxt, page_no)
        spec["n"] = n
        if sec_id == 6:   # source panel can be identical to table 4
            a, b = secs.get(4), secs.get(6)
            if a and b and [(r["resolved"], r["shown"]) for r in a.rows] == [(r["resolved"], r["shown"]) for r in b.rows]:
                for k in range(n):
                    data.cell(spec["first"] + k, 7, "Clear (PDF panel repeats the failed-login series this week)")
        if sec_id in (14, 15, 16) and widget and (
                "download" in widget.title.lower() and "upload" in widget.subtitle.lower()
                or "upload" in widget.title.lower() and "download" in widget.subtitle.lower()):
            for k in range(n):
                c = data.cell(spec["first"] + k, 7)
                c.value = f"{c.value}; title/subtitle conflict in PDF ('{widget.title}' vs '{widget.subtitle}') - verify direction"
        # section header page reference
        hc = data.cell(spec["hdr"], 2)
        if isinstance(hc.value, str) and widget:
            hc.value = re.sub(r"\(PDF pages? [^)]*\)", f"(PDF page {widget.page})", hc.value)
    for sec_id in (1, 2):
        hc = data.cell(SECTIONS[sec_id]["hdr"], 2)
        hc.value = re.sub(r"\(PDF pages? [^)]*\)", "(PDF page 1)", hc.value)
    data["B4"] = (f"Source: {pdf_name} ({parsed['pages']} pages). Parsed automatically from the PDF text layer and chart "
                  f"geometry by soc_pdf_to_excel.py. Reading confidence and caveats are in column G; see the 'Checks' sheet.")

    # ----- charts: resize to row counts, refresh units/titles -----
    first_to_sec = {s["first"]: sid for sid, s in SECTIONS.items()}
    for ch in dash._charts:
        try:
            f = ch.series[0].val.numRef.f
        except AttributeError:
            continue
        m = re.search(r"\$C\$(\d+)", f)
        sid = first_to_sec.get(int(m.group(1))) if m else None
        if sid is None or "n" not in SECTIONS[sid]:
            continue
        retarget_chart(ch, SECTIONS[sid]["first"], SECTIONS[sid]["n"])
        if sid in (13, 14, 15, 16):
            for ax in (ch.x_axis, ch.y_axis):
                if ax.title is not None:
                    set_title_text(ax.title, lambda t: "GB" if "illions of bytes" in t or t.strip() in ("B", "GB") else t)
        if sid == 12 and ch.title is not None:
            set_title_text(ch.title, lambda t: "Top Applications with Unusual Outgoing Data Volume (word cloud in PDF: no values)")

    # ----- dashboard notes -----
    def other_bucket(sec):
        w = secs.get(sec)
        o = [r for r in (w.rows if w else []) if re.match(r"^Other \(\d+\)$", r["resolved"])]
        return f"'{o[0]['resolved']}' = {o[0]['shown']} ({o[0]['pct']}%)" if o else None

    n1 = ["Reading notes", "",
          "Source and destination pair percentages in the PDF are shares of all indicators; the Data sheet shows each pair's share "
          "of the plotted top-10 total, which recalculates when you edit values.", "",
          "Counts such as 43.8K are rounded exactly as printed in the report."]
    ob = [x for x in (other_bucket(18), other_bucket(19)) if x]
    if ob:
        n1.append("Not charted (PDF 'Other' buckets): " + "; ".join(ob) + ".")
    dash["M35"] = "\n".join(n1)

    n2 = ["Reading notes", ""]
    unl = secs.get(9)
    n2.append("Top 10 Unlocked Accounts is blank in the source PDF (\"No data in this period\"), so this chart is empty until you add rows in table 9 on the Data sheet."
              if (unl is None or unl.empty or not unl.rows) else "Top 10 Unlocked Accounts is populated this week.")
    a, b = secs.get(4), secs.get(6)
    if a and b and [(r["resolved"], r["shown"]) for r in a.rows] == [(r["resolved"], r["shown"]) for r in b.rows]:
        n2.append("")
        n2.append("Remote Login Failure repeats the failed-login series in the source (identical values).")
    n2.append("")
    n2.append(f"Usernames are cut off with an ellipsis in the PDF; the domain was restored as @{email_domain}.")
    dash["M153"] = "\n".join(n2)

    n3 = ["Reading notes", "",
          "Data-volume values are converted to GB from the printed GB/MB labels (1 GB = 1,000 MB).", "",
          "Top Applications with Unusual Outgoing Data Volume is a word cloud in the PDF with no numbers, so its chart has labels but no bars.",
          "", "Seceon prints some bandwidth widgets with titles that contradict their subtitles (Download vs Upload); verify direction in Seceon before quoting."]
    dash["M252"] = "\n".join(n3)

    write_extract_sheet(wb, parsed, secs)
    checks = run_checks(parsed, secs, kpi_exact)
    write_checks_sheet(wb, checks)
    wb.active = wb.sheetnames.index("Executive Dashboard")
    wb.save(out_path)
    return checks, org, rid, period


def main():
    ap = argparse.ArgumentParser(description="Seceon SOC report PDF -> Excel executive dashboard")
    ap.add_argument("pdf")
    ap.add_argument("--template", default=None)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    here = os.path.dirname(os.path.abspath(__file__))
    template = a.template or (DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else os.path.join(here, DEFAULT_TEMPLATE))
    if not os.path.exists(template):
        sys.exit(f"Template not found: {template}\nPut {DEFAULT_TEMPLATE} next to this script or pass --template.")
    print(f"Reading {a.pdf} ...")
    parsed = parse_pdf(a.pdf)
    print(f"  {len(parsed['widgets'])} widgets, {len(parsed['kpis'])} KPI cards, {len(parsed['stats'])} stat cards")
    out = a.out
    if not out:
        end = re.findall(r"(\d{1,2} \w{3} \d{4})", parsed["header"].get("REPORTING", ""))
        stamp = end[-1].replace(" ", "") if end else "output"
        out = os.path.join(os.path.dirname(os.path.abspath(a.pdf)), f"SOC_Executive_Dashboard_{stamp}.xlsx")
    checks, org, rid, period = build_workbook(parsed, template, out, os.path.basename(a.pdf))
    print(f"\n{org} ({rid})  |  {period}")
    for st, name, detail in checks:
        if st != "PASS":
            print(f"  [{st}] {name}: {detail}")
    print(f"  {sum(1 for c in checks if c[0] == 'PASS')} checks passed, "
          f"{sum(1 for c in checks if c[0] in ('WARN', 'FAIL'))} need a look")
    print(f"\nSaved: {out}")


if __name__ == "__main__":
    main()
