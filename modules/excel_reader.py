from __future__ import annotations
import io
import re
import pandas as pd

REQUIRED = ["Vendor Name", "Purchasing Doc", "Material", "Style name", "Color", "Open Qty."]


def _clean_col(x):
    return re.sub(r"\s+", " ", str(x).strip())


def read_excel(source) -> pd.DataFrame:
    """Read the PO source Excel. `source` can be a path or uploaded bytes/file."""
    df = pd.read_excel(source, sheet_name=0, dtype={"Purchasing Doc": str, "Material": str, "Ship to ID": str})
    df.columns = [_clean_col(c) for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required Excel columns: {', '.join(missing)}")

    df["Purchasing Doc"] = df["Purchasing Doc"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["Material"] = df["Material"].astype(str).str.strip().str.upper()
    df["Vendor Name"] = df["Vendor Name"].fillna("").astype(str).str.strip()
    df["Style name"] = df["Style name"].fillna("").astype(str).str.strip()
    df["Color"] = df["Color"].fillna("").astype(str).str.strip()
    df["Open Qty."] = pd.to_numeric(df["Open Qty."], errors="coerce").fillna(0)
    if "Size" in df.columns:
        df["Size"] = df["Size"].fillna("").astype(str).str.strip().str.upper()
    else:
        df["Size"] = ""
    if "Ship to ID" in df.columns:
        df["Ship to ID"] = df["Ship to ID"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.lstrip("0")
    else:
        df["Ship to ID"] = ""
    if "Ordering Address" not in df.columns:
        df["Ordering Address"] = ""
    return df


def po_summary(df: pd.DataFrame, po_no: str) -> dict | None:
    rows = df[df["Purchasing Doc"] == str(po_no)].copy()
    if rows.empty:
        return None

    # Excel is usually one row per LF/RT. Sum Open Qty by material and size side.
    details = []
    for material, g in rows.groupby("Material", sort=False):
        sides = {}
        for _, r in g.iterrows():
            side = str(r.get("Size", "")).strip().upper() or "UNSPECIFIED"
            sides[side] = sides.get(side, 0.0) + float(r["Open Qty."])
        details.append({
            "material": material,
            "style": str(g.iloc[0]["Style name"]),
            "color": str(g.iloc[0]["Color"]),
            "qty": float(g["Open Qty."].sum()),
            "sides": sides,
        })

    vendors = [v for v in rows["Vendor Name"].dropna().unique().tolist() if str(v).strip()]
    ship_tos = [v for v in rows["Ship to ID"].dropna().unique().tolist() if str(v).strip()]
    return {
        "po_no": str(po_no),
        "vendor": vendors[0] if vendors else "",
        "ordering_address": str(rows.iloc[0].get("Ordering Address", "")),
        "ship_to": ship_tos[0] if ship_tos else "",
        "qty": float(rows["Open Qty."].sum()),
        "items": details,
    }
