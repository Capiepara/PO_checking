from __future__ import annotations
import io
import os
import re
import fitz


def _text_from_pdf(source) -> str:
    if isinstance(source, (str, os.PathLike)):
        doc = fitz.open(str(source))
    else:
        data = source.read() if hasattr(source, "read") else bytes(source)
        if hasattr(source, "seek"):
            source.seek(0)
        doc = fitz.open(stream=data, filetype="pdf")
    try:
        return "\n".join(page.get_text("text") for page in doc)
    finally:
        doc.close()


def _one(pattern, text, flags=re.I):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def _norm_lines(text):
    return [re.sub(r"\s+", " ", x).strip() for x in text.splitlines() if x.strip()]


def parse_po_pdf(source, filename: str | None = None) -> dict:
    text = _text_from_pdf(source)
    lines = _norm_lines(text)
    po_no = _one(r"Purchase\s+Order\s+No\.?\s*[:]?\s*(\d{7,12})", text)
    vendor_no = _one(r"Vendor\s+No\.?\s*[:]?\s*(\d+)", text)
    ship_to = _one(r"Ship-To\s*:\s*(\d+)", text)

    # Vendor block: content after Vendor Address and before Information.
    vendor_block = _one(r"Vendor\s+Address\s*(.*?)\s*Information", text, re.I | re.S)
    vendor_lines = _norm_lines(vendor_block)
    vendor = vendor_lines[0] if vendor_lines else ""
    vendor_address = " | ".join(vendor_lines[1:])

    # Shipping block: content after Shipping Address and before Ship-To.
    shipping_block = _one(r"Shipping\s+Address\s*:\s*(.*?)\s*Ship-To\s*:", text, re.I | re.S)
    shipping_address = " | ".join(_norm_lines(shipping_block))

    # Parse material items. Anchor on item number + S-material; capture until next item/total/terms/page header.
    item_re = re.compile(
        r"(?ms)^\s*(\d{1,4})\s+(S\d{6,})\s*\n(.*?)(?=^\s*\d{1,4}\s+S\d{6,}\s*$|^\s*Total Value|^\s*Terms and Conditions|^\s*Purchase Order No\.)"
    )
    items = []
    for m in item_re.finditer(text):
        item_no, material, block = m.group(1), m.group(2).upper(), m.group(3)
        blines = _norm_lines(block)
        # Main quantity is on a line containing EA followed by quantity and net amount.
        qty = None
        # PDF extractors may place EA / qty / amount on separate lines, so parse the whole block.
        qm = re.search(r"\bEA\s+(\d+(?:\.\d+)?)\s+\d[\d,]*\.\d{2}\b", block, re.I | re.S)
        if qm:
            qty = float(qm.group(1))
        sides = {}
        for side, q in re.findall(r"(?i)\b(LF|RT)\s+EA\s+(\d+(?:\.\d+)?)", block):
            sides[side.upper()] = sides.get(side.upper(), 0.0) + float(q)

        # Text before pricing is description/color. Remove obvious price/UOM lines.
        desc_parts = []
        for ln in blines:
            if re.search(r"\bEA\b", ln) or re.match(r"^(LF|RT)\b", ln, re.I):
                continue
            if re.match(r"^\d+(?:\.\d+)?\s+1\b", ln):
                continue
            desc_parts.append(ln)
        desc = " ".join(desc_parts)
        items.append({"item": item_no, "material": material, "description": desc, "qty": qty or 0.0, "sides": sides})

    pdf_qty = sum(i["qty"] for i in items)
    return {
        "filename": filename or getattr(source, "name", ""),
        "po_no": po_no,
        "vendor_no": vendor_no,
        "vendor": vendor,
        "vendor_address": vendor_address,
        "ship_to": ship_to,
        "shipping_address": shipping_address,
        "qty": pdf_qty,
        "items": items,
        "raw_text": text,
    }
