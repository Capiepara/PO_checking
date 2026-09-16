# PO Checker — GitHub Pages Edition

This version is a static browser app. No Python server and no Streamlit are required.

## Checks
- Purchasing Doc in Excel vs Purchase Order No. in PDF
- Purchase Order No. inside PDF vs PDF filename
- Vendor Name
- Displays Vendor Address from PDF
- Ship-To and Shipping Address
- Material / Style / Color
- Open Qty by material and LF/RT
- Total PDF Qty (main PO line quantity; LF/RT are not counted twice)
- Duplicate PO numbers among uploaded PDFs
- Exports Summary + Item Detail to Excel

## Required Excel columns
`Vendor Name`, `Purchasing Doc`, `Material`, `Style name`, `Color`, `Open Qty.`

Optional: `Size`, `Ship to ID`.

## Deploy with GitHub Pages
1. Replace the old Streamlit project contents with these files.
2. Commit and push to `main`.
3. GitHub repo → Settings → Pages.
4. Source: **Deploy from a branch**
5. Branch: **main**, folder: **/(root)** → Save.
6. Open the GitHub Pages URL.

## Privacy
Excel/PDF processing is performed in the browser. The app itself does not upload the selected files to an application server.

## Important
The parser is tailored to the KEEN PO layout used to build V1. If the PO PDF layout changes in a future season, `app.js` may need adjustment.
