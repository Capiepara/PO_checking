from __future__ import annotations
import os
import re
from rapidfuzz.fuzz import ratio


def norm(s):
    s = str(s or "").upper()
    s = re.sub(r"\b(M|W)-", "", s)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def filename_po(filename):
    nums = re.findall(r"(?<!\d)(\d{10})(?!\d)", os.path.basename(filename or ""))
    return nums[0] if nums else ""


def text_compatible(expected, pdf_desc, threshold=58):
    """PO descriptions are often truncated; accept containment or fuzzy similarity."""
    a, b = norm(expected), norm(pdf_desc)
    if not a:
        return True, 100
    if a in b or b in a:
        return True, 100
    # Compare meaningful tokens so truncated PO colors don't become false failures.
    at = [x for x in a.split() if len(x) >= 3]
    hits = sum(1 for x in at if x in b)
    token_score = 100 * hits / max(1, len(at))
    fuzzy = ratio(a, b)
    score = max(token_score, fuzzy)
    return score >= threshold, round(score)


def compare_po(excel_po, pdf_po):
    checks = []
    if excel_po is None:
        return {"status":"MISMATCH", "checks":["PO not found in Excel"], "item_issues":["PO not found in Excel"]}

    def ck(label, ok, detail=""):
        checks.append({"field": label, "ok": bool(ok), "detail": detail})

    ck("PO Number", excel_po["po_no"] == pdf_po["po_no"], f'Excel {excel_po["po_no"]} / PDF {pdf_po["po_no"]}')
    fpo = filename_po(pdf_po["filename"])
    ck("Filename", bool(fpo) and fpo == pdf_po["po_no"], f'Filename PO {fpo or "not found"} / PDF {pdf_po["po_no"]}')
    ck("Vendor", norm(excel_po["vendor"]) == norm(pdf_po["vendor"]), f'Excel {excel_po["vendor"]} / PDF {pdf_po["vendor"]}')
    ck("Ship-To", str(excel_po["ship_to"]) == str(pdf_po["ship_to"]), f'Excel {excel_po["ship_to"]} / PDF {pdf_po["ship_to"]}')
    ck("Total Qty", abs(excel_po["qty"] - pdf_po["qty"]) < 1e-9, f'Excel {excel_po["qty"]:g} / PDF {pdf_po["qty"]:g}')

    pdf_items = {x["material"]: x for x in pdf_po["items"]}
    item_issues = []
    warnings = []
    for x in excel_po["items"]:
        p = pdf_items.get(x["material"])
        if not p:
            item_issues.append(f'{x["material"]}: missing in PDF')
            continue
        if abs(x["qty"] - p["qty"]) > 1e-9:
            item_issues.append(f'{x["material"]}: qty Excel {x["qty"]:g} / PDF {p["qty"]:g}')
        for side, q in x["sides"].items():
            pq = p["sides"].get(side)
            if pq is not None and abs(q - pq) > 1e-9:
                item_issues.append(f'{x["material"]} {side}: Excel {q:g} / PDF {pq:g}')
        style_ok, ss = text_compatible(x["style"], p["description"])
        color_ok, cs = text_compatible(x["color"], p["description"])
        if not style_ok:
            warnings.append(f'{x["material"]}: style text is abbreviated/truncated ({ss}% confidence) — Excel "{x["style"]}" / PDF "{p["description"]}"')
        if not color_ok:
            warnings.append(f'{x["material"]}: color text is abbreviated/truncated ({cs}% confidence) — Excel "{x["color"]}" / PDF "{p["description"]}"')

    extra = sorted(set(pdf_items) - {x["material"] for x in excel_po["items"]})
    for m in extra:
        item_issues.append(f"{m}: present in PDF but not Excel")

    core_ok = all(c["ok"] for c in checks)
    status = "PASS" if core_ok and not item_issues else "MISMATCH"
    return {"status": status, "checks": checks, "item_issues": item_issues, "warnings": warnings}
