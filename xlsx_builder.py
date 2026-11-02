"""Builds the report as a real Excel workbook with native PivotTables.

Starts from template.xlsx (the original DXC report: styles, 3 PivotTables, pivot cache) and regenerates the
Main File sheet, Summary sheet, pivot cache and pivot definitions from live rows. The pivot cache and the
Summary cells are written consistently, so Excel shows them as-is (no refresh on open, which would drop the
unused stage items and reset the More-than-four filter).
"""
import io
import os
import re
import zipfile
from datetime import datetime, timezone
from xml.sax.saxutils import escape

from report import FOLLOWUP_THRESHOLD

MAX_STAGE = 30  # stages kept available in the PivotTable stage filters

TEMPLATE = os.path.join(os.path.dirname(__file__), "template.xlsx")

COLUMNS = [
    "CASE_CHECK_ID", "CASE_ARS_NO", "company_name", "CANDIDATE_NAME", "PROCESS_NAME", "location", "RECEIVED_DATE",
    "INSUFF_RAISED_DATE", "insuff_date", "insuff_remarks", "Case_status", "check_status", "CHECK_NAME",
    "verification_source", *[f"CASE_FLEX_FIELD{i}" for i in range(1, 18)],
    "Candidate contact no", "Candidate email id", "check_reopen_date", "stage", "ADDITIONAL_REMARKS",
    "FOLLOWUP_DATE", "NEXT_FOLLOWUP_DATE", "remark", "check_reopendate",
]
DATE_COLS = {"RECEIVED_DATE", "INSUFF_RAISED_DATE", "insuff_date", "check_reopen_date", "FOLLOWUP_DATE",
             "NEXT_FOLLOWUP_DATE", "check_reopendate"}
STAGE = COLUMNS.index("stage")
REMARK = COLUMNS.index("remark")
NS = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
_BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_EPOCH = datetime(1899, 12, 30)


def col_name(n):  # 1-based -> A, B, ... AA
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def esc(v):
    return escape(_BAD.sub("", str(v)), {'"': "&quot;"})


def stage_no(s):
    try:
        return int(str(s).split("-", 1)[1])
    except (IndexError, ValueError):
        return None


def typed(col, raw):
    """raw DB value -> (kind, value, iso) where kind is s|n|d|m."""
    if raw is None or raw == "":
        return ("m", None, None)
    name = COLUMNS[col]
    if name in DATE_COLS:
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}):(\d{2}))?", str(raw))
        if not m or m.group(1) == "0000":
            return ("m", None, None)
        try:
            dt = datetime(*(int(g or 0) for g in m.groups()))
        except ValueError:
            return ("m", None, None)
        return ("d", (dt - _EPOCH).total_seconds() / 86400, dt.strftime("%Y-%m-%dT%H:%M:%S"))
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return ("n", raw, None)
    if (col == 0 or name == "Candidate contact no") and re.fullmatch(r"[1-9]\d{0,14}", str(raw)):
        return ("n", int(raw), None)
    return ("s", str(raw)[:32000], None)


def shared_items_attrs(cells):
    blank = any(c[0] == "m" for c in cells)
    strs = [c[1] for c in cells if c[0] == "s"]
    nums = [c[1] for c in cells if c[0] == "n"]
    dates = [c for c in cells if c[0] == "d"]
    long_ = any(len(s) > 255 for s in strs)
    a = []
    iso = lambda serial: (datetime.fromtimestamp((serial - 25569) * 86400, timezone.utc)).strftime("%Y-%m-%dT%H:%M:%S")
    if dates:
        v = [c[1] for c in dates]
        a += ['containsNonDate="0"', 'containsDate="1"', 'containsString="0"']
        if blank:
            a.append('containsBlank="1"')
        a += [f'minDate="{iso(min(v))}"', f'maxDate="{iso(max(v))}"']
    elif nums and not strs:
        a.append('containsString="0"')
        if blank:
            a.append('containsBlank="1"')
        a.append('containsNumber="1"')
        if all(float(n).is_integer() for n in nums):
            a.append('containsInteger="1"')
        a += [f'minValue="{min(nums)}"', f'maxValue="{max(nums)}"']
    elif nums and strs:
        if blank:
            a.append('containsBlank="1"')
        a += ['containsMixedTypes="1"', 'containsNumber="1"']
        if all(float(n).is_integer() for n in nums):
            a.append('containsInteger="1"')
        a += [f'minValue="{min(nums)}"', f'maxValue="{max(nums)}"']
        if long_:
            a.append('longText="1"')
    elif strs:
        if blank:
            a.append('containsBlank="1"')
        if long_:
            a.append('longText="1"')
    else:
        a += ['containsNonDate="0"', 'containsString="0"', 'containsBlank="1"']
    return " ".join(a)


def _inline(ref, s, text):
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc(text)}</t></is></c>'


def _num(ref, s, v):
    return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'


def _empty(ref, s):
    return f'<c r="{ref}" s="{s}"/>'


def build_summary_sheet(tpl, sections):
    rows: dict[int, list] = {}

    def put(r, c, xml):
        rows.setdefault(r, []).append((c, xml))

    max_row, merges = 1, []
    for sec in sections:
        row, col = sec["row"], sec["col"]
        L, R = col_name(col), col_name(col + 1)
        put(row, col, _inline(f"{L}{row}", 10, sec["title"])); put(row, col + 1, _empty(f"{R}{row}", 10))
        put(row + 1, col, _empty(f"{L}{row + 1}", 10)); put(row + 1, col + 1, _empty(f"{R}{row + 1}", 10))
        merges.append(f"{L}{row}:{R}{row + 1}")
        r = row + 2
        if sec.get("stage_label"):
            put(r, col, _inline(f"{L}{r}", 1, "stage")); put(r, col + 1, _inline(f"{R}{r}", 2, sec["stage_label"]))
            r += 2
        put(r, col, _inline(f"{L}{r}", 1, "Row Labels")); put(r, col + 1, _inline(f"{R}{r}", 2, "Count of remark"))
        total = 0
        for name, n in sec["labels"]:
            r += 1
            put(r, col, _inline(f"{L}{r}", 3, name)); put(r, col + 1, _num(f"{R}{r}", 2, n))
            total += n
        if sec["labels"]:  # a pivot with no matching rows stays blank (no Grand Total)
            r += 1
            put(r, col, _inline(f"{L}{r}", 3, "Grand Total")); put(r, col + 1, _num(f"{R}{r}", 2, total))
        max_row = max(max_row, r)
    body = "".join(
        f'<row r="{r}">{"".join(x for _, x in sorted(rows[r], key=lambda t: t[0]))}</row>' for r in sorted(rows))
    out = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:F{max_row}"/>', tpl)
    out = re.sub(r"<sheetData>[\s\S]*</sheetData>", lambda m: f"<sheetData>{body}</sheetData>", out)
    merge_xml = "".join(f'<mergeCell ref="{m}"/>' for m in merges)
    return re.sub(r"<mergeCells[\s\S]*</mergeCells>", lambda m: f'<mergeCells count="{len(merges)}">{merge_xml}</mergeCells>', out)


def build_main_sheet(tpl, data):
    top = max([d["stage_no"] or 0 for d in data] + [0])
    xml = "<row r=\"1\">" + "".join(
        _inline(f"{col_name(i + 1)}1", 9 if c == "remark" else 8 if c in DATE_COLS else 7, c)
        for i, c in enumerate(COLUMNS)) + f'<c r="AP1"><f>MAX(AO:AO)</f><v>{top}</v></c></row>'
    for k, d in enumerate(data):
        r = k + 2
        parts = []
        for i, (kind, v, _) in enumerate(d["cells"]):
            ref = f"{col_name(i + 1)}{r}"
            s = 6 if kind == "d" else None if i == 0 else 5
            sa = "" if s is None else f' s="{s}"'
            if kind == "m":
                if i:
                    parts.append(f'<c r="{ref}" s="5"/>')
            elif kind == "s":
                parts.append(f'<c r="{ref}"{sa} t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>')
            else:
                parts.append(f'<c r="{ref}"{sa}><v>{v}</v></c>')
        parts.append(f'<c r="AO{r}"><f>VALUE(RIGHT(AI{r},LEN(AI{r})-FIND("-",AI{r})))</f><v>{d["stage_no"] or 0}</v></c>')
        xml += f'<row r="{r}">{"".join(parts)}</row>'
    out = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:AP{len(data) + 1}"/>', tpl)
    return re.sub(r"<sheetData>[\s\S]*</sheetData>", lambda m: f"<sheetData>{xml}</sheetData>", out)


def build_workbook(rows) -> bytes:
    zin = zipfile.ZipFile(TEMPLATE)
    files = {n: zin.read(n) for n in zin.namelist()}
    read = lambda n: files[n].decode("utf-8")
    write = lambda n, s: files.__setitem__(n, s.encode("utf-8"))

    data = [{
        "cells": [typed(i, r.get(c)) for i, c in enumerate(COLUMNS)],
        "stage_no": stage_no(r.get("stage")),
        "stage": r.get("stage"),
        "remark": r.get("remark"),
    } for r in rows]

    # shared item lists for the two pivot fields
    # Stage items: those present in the data, then every other stage up to MAX_STAGE as "retained"
    # (unused) items - exactly how the template keeps followup-5, 6, 7... selectable. Without them Excel
    # cannot hide every stage in the "More" filter and would silently reset it to (All).
    present = sorted({d["stage_no"] for d in data if d["stage_no"] is not None})
    universe = range(1, max(MAX_STAGE, *(present or [0])) + 1)
    missing = [n for n in universe if n not in present]
    stage_items = [f"followup-{n}" for n in present] + [f"followup-{n}" for n in missing]
    unused = {f"followup-{n}" for n in missing}
    less_stages = [n for n in universe if n <= FOLLOWUP_THRESHOLD]  # followup-1 .. followup-4
    more_stages = [n for n in universe if n > FOLLOWUP_THRESHOLD]   # greater than followup-4
    remark_items = sorted({d["remark"] for d in data if d["remark"] is not None}, key=str.lower)
    remark_blank = any(d["remark"] is None for d in data)
    if remark_blank:
        remark_items.append(None)
    stage_idx = {s: i for i, s in enumerate(stage_items)}
    remark_idx = {s: i for i, s in enumerate(remark_items)}

    # pivot cache definition + records
    fields = []
    for i, name in enumerate(COLUMNS):
        head = f'<cacheField name="{esc(name)}" numFmtId="{15 if name in DATE_COLS else 0}">'
        if i == STAGE:
            items = "".join(f'<s v="{esc(s)}"{" u=\"1\"" if s in unused else ""}/>' for s in stage_items)
            fields.append(f'{head}<sharedItems count="{len(stage_items)}">{items}</sharedItems></cacheField>')
        elif i == REMARK:
            items = "".join("<m/>" if s is None else f'<s v="{esc(s)}"/>' for s in remark_items)
            attr = ' containsBlank="1"' if remark_blank else ""
            fields.append(f'{head}<sharedItems{attr} count="{len(remark_items)}">{items}</sharedItems></cacheField>')
        else:
            fields.append(f'{head}<sharedItems {shared_items_attrs([d["cells"][i] for d in data])}/></cacheField>')
    refreshed = (datetime.now(timezone.utc) - datetime(1899, 12, 30, tzinfo=timezone.utc)).total_seconds() / 86400
    write("xl/pivotCache/pivotCacheDefinition1.xml",
          f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition {NS} r:id="rId1" '
          f'refreshedBy="DXC Insuff Report" refreshedDate="{refreshed:.6f}" createdVersion="8" missingItemsLimit="1048576" '
          f'refreshedVersion="8" minRefreshableVersion="3" recordCount="{len(data)}"><cacheSource type="worksheet">'
          f'<worksheetSource ref="A1:AN{len(data) + 1}" sheet="Main File"/></cacheSource>'
          f'<cacheFields count="{len(COLUMNS)}">{"".join(fields)}</cacheFields></pivotCacheDefinition>')

    def rec(d):
        out = []
        for i, (kind, v, iso) in enumerate(d["cells"]):
            if i == STAGE:
                out.append(f'<x v="{stage_idx[d["stage"]]}"/>')
            elif i == REMARK:
                out.append(f'<x v="{remark_idx[d["remark"]]}"/>')
            elif kind == "m":
                out.append("<m/>")
            elif kind == "s":
                out.append(f'<s v="{esc(v)}"/>')
            elif kind == "n":
                out.append(f'<n v="{v}"/>')
            else:
                out.append(f'<d v="{iso}"/>')
        return f'<r>{"".join(out)}</r>'

    write("xl/pivotCache/pivotCacheRecords1.xml",
          f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords {NS} count="{len(data)}">'
          f'{"".join(rec(d) for d in data)}</pivotCacheRecords>')

    # pivots
    def label_of(sel):
        return f"followup-{sel[0]}" if len(sel) == 1 else "(Multiple Items)"

    def visible(lst):
        counts: dict = {}
        for d in lst:
            counts[d["remark"]] = counts.get(d["remark"], 0) + 1
        return [{"name": s, "idx": i, "n": counts[s]} for i, s in enumerate(remark_items) if s in counts]

    overall_vis = visible([d for d in data if d["remark"] is not None])  # Overall pivot hides (blank)
    less_vis = visible([d for d in data if d["stage_no"] in less_stages])
    more_vis = visible([d for d in data if d["stage_no"] in more_stages])
    base = max(16, len(overall_vis) + 8)

    def row_items(vis):
        items = "".join("<i><x/></i>" if v["idx"] == 0 else f'<i><x v="{v["idx"]}"/></i>' for v in vis)
        return f'<rowItems count="{len(vis) + 1}">{items}<i t="grand"><x/></i></rowItems>'

    def remark_field(hide_blank):
        items = "".join(f'<item{" h=\"1\"" if s is None and hide_blank else ""} x="{i}"/>' for i, s in enumerate(remark_items))
        return (f'<pivotField axis="axisRow" dataField="1" showAll="0"><items count="{len(remark_items) + 1}">'
                f'{items}<item t="default"/></items></pivotField>')

    def stage_field(sel):
        items = "".join(
            f'<item{"" if stage_no(s) in sel else " h=\"1\""}{" m=\"1\"" if s in unused else ""} x="{i}"/>'
            for i, s in enumerate(stage_items))
        return (f'<pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0">'
                f'<items count="{len(stage_items) + 1}">{items}<item t="default"/></items></pivotField>')

    def pivot_file(name, ref, vis, hide_blank, sel=None):
        x = read(name)
        x = re.sub(r'<location ref="[^"]*"', f'<location ref="{ref}"', x)
        x = re.sub(r'<pivotField axis="axisRow"[\s\S]*?</pivotField>', lambda m: remark_field(hide_blank), x)
        if sel is not None:
            x = re.sub(r'<pivotField axis="axisPage"[\s\S]*?</pivotField>', lambda m: stage_field(sel), x)
        x = re.sub(r"<rowItems[\s\S]*?</rowItems>", lambda m: row_items(vis), x)
        write(name, x)

    pivot_file("xl/pivotTables/pivotTable2.xml", f"A3:B{3 + len(overall_vis) + 1}", overall_vis, True)
    pivot_file("xl/pivotTables/pivotTable1.xml", f"A{base + 4}:B{base + 4 + len(less_vis) + 1}", less_vis, False, less_stages)
    pivot_file("xl/pivotTables/pivotTable3.xml", f"E{base + 4}:F{base + 4 + len(more_vis) + 1}", more_vis, False, more_stages)

    labels = lambda vis: [(v["name"], v["n"]) for v in vis]
    write("xl/worksheets/sheet1.xml", build_summary_sheet(read("xl/worksheets/sheet1.xml"), [
        {"row": 1, "col": 1, "title": "Overall Status", "labels": labels(overall_vis)},
        {"row": base, "col": 1, "title": "Less Then Four Follow-Up", "stage_label": label_of(less_stages), "labels": labels(less_vis)},
        {"row": base, "col": 5, "title": "More Then Four Follow-Up", "stage_label": label_of(more_stages), "labels": labels(more_vis)},
    ]))
    write("xl/worksheets/sheet2.xml", build_main_sheet(read("xl/worksheets/sheet2.xml"), data))

    # package plumbing: drop the stale calcChain / filter name, recalc on open
    files.pop("xl/calcChain.xml", None)
    write("xl/_rels/workbook.xml.rels", re.sub(r'<Relationship Id="rId7"[^>]*/>', "", read("xl/_rels/workbook.xml.rels")))
    write("[Content_Types].xml", re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", read("[Content_Types].xml")))
    wbx = re.sub(r"<definedNames>[\s\S]*?</definedNames>", "", read("xl/workbook.xml"))
    write("xl/workbook.xml", wbx.replace('<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>'))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        # [Content_Types].xml first, as Excel expects
        for n in sorted(files, key=lambda n: n != "[Content_Types].xml"):
            zout.writestr(n, files[n])
    return buf.getvalue()
