import streamlit as st
import pandas as pd
from modules.excel_reader import read_excel, po_summary
from modules.pdf_reader import parse_po_pdf
from modules.po_matcher import compare_po
from modules.report import result_rows, to_excel_bytes

st.set_page_config(page_title="PO Checker", page_icon="✅", layout="wide")
st.title("PO Checker — Excel vs Purchase Order PDF")
st.caption("Checks PO number/filename, Vendor, Ship-To, Material, Style/Color, LF/RT and Open Qty. PDF Qty counts the main line Quantity only (LF/RT are not added twice).")

excel_file = st.file_uploader("1. Upload PO source Excel", type=["xlsx", "xls"])
pdf_files = st.file_uploader("2. Upload PO PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Check PO", type="primary", disabled=not(excel_file and pdf_files)):
    try:
        df = read_excel(excel_file)
        results=[]
        seen={}
        for f in pdf_files:
            pdf=parse_po_pdf(f, f.name)
            po=pdf["po_no"]
            duplicate=po in seen
            seen[po]=seen.get(po,0)+1
            excel=po_summary(df,po) if po else None
            comp=compare_po(excel,pdf)
            results.append({"pdf":pdf,"excel":excel,"comparison":comp,"duplicate":duplicate})
        st.session_state["results"] = results
    except Exception as e:
        st.exception(e)

results=st.session_state.get("results")
if results:
    rows=result_rows(results)
    rdf=pd.DataFrame(rows)
    c1,c2,c3,c4=st.columns(4)
    c1.metric("PDF files",len(results))
    c2.metric("Unique PO",len(set(r["pdf"]["po_no"] for r in results)))
    c3.metric("PASS",sum(r["comparison"]["status"]=="PASS" for r in results))
    c4.metric("MISMATCH",sum(r["comparison"]["status"]!="PASS" for r in results))
    st.dataframe(rdf, use_container_width=True, hide_index=True)
    st.download_button("Download result Excel", to_excel_bytes(results), "PO_Check_Result.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Details")
    for r in results:
        po=r["pdf"]["po_no"] or "PO number not detected"
        label=f'{po} — {r["pdf"]["vendor"]} — {r["comparison"]["status"]}' + (" — DUPLICATE" if r["duplicate"] else "")
        with st.expander(label):
            st.write(f'**Filename:** {r["pdf"]["filename"]}')
            st.write(f'**Vendor address:** {r["pdf"]["vendor_address"]}')
            st.write(f'**Shipping address:** {r["pdf"]["shipping_address"]}')
            if r["comparison"].get("checks") and isinstance(r["comparison"]["checks"][0],dict):
                st.dataframe(pd.DataFrame(r["comparison"]["checks"]), hide_index=True, use_container_width=True)
            if r["comparison"].get("item_issues"):
                st.error("\n".join("• "+x for x in r["comparison"]["item_issues"]))
            else:
                st.success("All item-level checks passed.")
            if r["comparison"].get("warnings"):
                st.warning("PDF style/color text may be abbreviated. Material + Qty still match.\n\n" + "\n".join("• "+x for x in r["comparison"]["warnings"]))
