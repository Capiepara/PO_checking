from __future__ import annotations
import io
import pandas as pd


def result_rows(results):
    rows=[]
    for r in results:
        checks={c["field"]:c for c in r["comparison"]["checks"]} if isinstance(r["comparison"].get("checks"), list) and r["comparison"]["checks"] and isinstance(r["comparison"]["checks"][0], dict) else {}
        rows.append({
            "PO": r["pdf"]["po_no"],
            "Vendor Name": r["pdf"]["vendor"],
            "Vendor Address (PDF)": r["pdf"]["vendor_address"],
            "Ship-To": r["pdf"]["ship_to"],
            "Shipping Address (PDF)": r["pdf"]["shipping_address"],
            "Excel Qty": r["excel"]["qty"] if r["excel"] else None,
            "PDF Qty": r["pdf"]["qty"],
            "PO/Filename": "PASS" if checks.get("Filename",{}).get("ok") else "MISMATCH",
            "Duplicate": "YES" if r.get("duplicate") else "NO",
            "Result": r["comparison"]["status"],
            "Issues": " | ".join(r["comparison"].get("item_issues",[]) + [c["detail"] for c in checks.values() if not c["ok"]]),
            "Warnings": " | ".join(r["comparison"].get("warnings",[])),
            "Filename": r["pdf"]["filename"],
        })
    return rows


def to_excel_bytes(results):
    summary = pd.DataFrame(result_rows(results))
    detail=[]
    for r in results:
        em={x["material"]:x for x in (r["excel"]["items"] if r["excel"] else [])}
        pm={x["material"]:x for x in r["pdf"]["items"]}
        for mat in sorted(set(em)|set(pm)):
            e,p=em.get(mat),pm.get(mat)
            detail.append({
                "PO":r["pdf"]["po_no"],"Material":mat,
                "Style (Excel)":e["style"] if e else "","Color (Excel)":e["color"] if e else "",
                "Excel Qty":e["qty"] if e else None,"PDF Qty":p["qty"] if p else None,
                "Excel LF":(e["sides"].get("LF") if e else None),"PDF LF":(p["sides"].get("LF") if p else None),
                "Excel RT":(e["sides"].get("RT") if e else None),"PDF RT":(p["sides"].get("RT") if p else None),
                "PDF Description":p["description"] if p else "",
            })
    bio=io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        summary.to_excel(w,index=False,sheet_name="Summary")
        pd.DataFrame(detail).to_excel(w,index=False,sheet_name="Item Detail")
        for ws in w.book.worksheets:
            ws.freeze_panes="A2"
            ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=cell.font.copy(bold=True)
            for col in ws.columns:
                width=min(max(len(str(c.value or "")) for c in col)+2,45)
                ws.column_dimensions[col[0].column_letter].width=max(10,width)
    bio.seek(0)
    return bio.getvalue()
