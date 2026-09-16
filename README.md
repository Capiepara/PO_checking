# PO Checker

Local Streamlit tool for checking KEEN PO PDFs against a PO source Excel file.

## Checks
- Purchasing Doc in Excel vs Purchase Order No. inside PDF
- Purchase Order No. inside PDF vs PDF filename
- Vendor Name
- Displays Vendor Address from PDF
- Ship-To ID and displays Shipping Address from PDF
- Material / Style / Color
- Open Qty by material and LF/RT
- Total PDF Qty (uses the main PO line Quantity; LF/RT are **not** added a second time)
- Duplicate PO numbers among uploaded PDFs
- Export Summary + Item Detail to Excel

## Expected Excel columns
The current version is based on the source format used for FW27 and expects at least:
`Vendor Name`, `Purchasing Doc`, `Material`, `Style name`, `Color`, `Open Qty.`

It also uses `Size` (LF/RT) and `Ship to ID` when available.

## Run on Windows
1. Install Python 3.11+.
2. Open Command Prompt in this folder.
3. Create a virtual environment:
   `python -m venv .venv`
4. Activate it:
   `.venv\Scripts\activate`
5. Install dependencies:
   `pip install -r requirements.txt`
6. Start:
   `streamlit run app.py`
7. Browser opens the PO Checker. Upload the Excel source and all PO PDFs, then click **Check PO**.

## Privacy
This version runs locally. The files are processed by Python on the computer running Streamlit; the code does not contain an API call that uploads PO data to an AI service.

## Matching note
PO text sometimes abbreviates/truncates colors (for example Excel may contain `COFFEE BEAN/BLACK` while the PDF displays `M-COFFEE`). The checker anchors item matching on Material code and Quantity, then uses tolerant text matching for Style/Color to reduce false mismatches. Review any low-confidence Style/Color issue shown in the report.

## GitHub
Create an empty GitHub repository named `po-checker`, then from this folder run:
`git init`
`git add .`
`git commit -m "Initial PO checker"`
`git branch -M main`
`git remote add origin YOUR_GITHUB_REPO_URL`
`git push -u origin main`
