import io
import pandas as pd
from openpyxl import Workbook

# -------------------------
# 1) Yeni SR Listesi (eski akış)
# -------------------------
SHEET_FALLBACK = "Ανατεθειμένες αυτοψίες"

OUT_COLS_CANON = [
    "SR",
    "ADDRESS",
    "A/K",
    "BUILDING ID",
    "ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ",
    "ΚΙΝΗΤΟ ΠΕΛΑΤΗ",
]

ALIASES = {
    "ADRESS": "ADDRESS",
    "Adress": "ADDRESS",
    "Address": "ADDRESS",
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    df = df.rename(columns={c: ALIASES.get(c, c) for c in df.columns})
    return df


def compare_excels(
    old_bytes: bytes,
    new_bytes: bytes,
    sheet_name: str | None = None
) -> tuple[bytes, int]:
    sheet = sheet_name or SHEET_FALLBACK

    old_df = pd.read_excel(io.BytesIO(old_bytes), sheet_name=sheet)
    new_df = pd.read_excel(io.BytesIO(new_bytes), sheet_name=sheet)

    old_df = normalize_columns(old_df)
    new_df = normalize_columns(new_df)

    if "SR" not in old_df.columns or "SR" not in new_df.columns:
        raise ValueError("SR kolonu bulunamadı (seçilen sheet yanlış olabilir).")

    diff = new_df[~new_df["SR"].isin(old_df["SR"])].copy()

    missing = [c for c in OUT_COLS_CANON if c not in diff.columns]
    if missing:
        raise ValueError(f"Beklenen kolon(lar) eksik: {missing}")

    out_df = diff[OUT_COLS_CANON].copy()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="New_SR")
    output.seek(0)

    return output.read(), len(out_df)


# -------------------------
# 2) New Flow Şablonu (SRID + mapping akışı)
# -------------------------
NEWFLOW_SHEET_FALLBACK = "New flow"

def _norm_srid(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        head = s[:-2]
        if head.isdigit():
            s = head
    return s

def compare_newflow_bytes(old_bytes: bytes, new_bytes: bytes, sheet_name: str | None = None) -> tuple[bytes, int]:
    sheet = sheet_name or NEWFLOW_SHEET_FALLBACK

    old_df = pd.read_excel(io.BytesIO(old_bytes), sheet_name=sheet)
    new_df = pd.read_excel(io.BytesIO(new_bytes), sheet_name=sheet)

    # Kolonları temizle (SRID gibi kolonlarda boşluk olabiliyor)
    old_df.columns = [str(c).strip() for c in old_df.columns]
    new_df.columns = [str(c).strip() for c in new_df.columns]

    if "SRID" not in old_df.columns or "SRID" not in new_df.columns:
        raise ValueError("New flow sekmesinde 'SRID' kolonu bulunamadı.")

    old_srid = old_df["SRID"].map(_norm_srid)
    new_srid = new_df["SRID"].map(_norm_srid)

    old_sr = set(old_srid[old_srid != ""])
    new_df = new_df.copy()
    new_df["_SRID_NORM"] = new_srid

    new_only = new_df[(new_df["_SRID_NORM"] != "") & (~new_df["_SRID_NORM"].isin(old_sr))].copy()

    required = ["SRID", "full_adr", "a/k", "building Id", "customer", "mobile"]
    missing = [c for c in required if c not in new_only.columns]
    if missing:
        raise ValueError(f"New flow sekmesinde eksik kolon(lar): {missing}")

    result = pd.DataFrame({
        "SR": new_only["_SRID_NORM"],
        "ADDRESS": new_only["full_adr"],
        "A/K": new_only["a/k"],
        "BUILDING ID": new_only["building Id"],
        "ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ": new_only["customer"],
        "ΚΙΝΗΤΟ ΠΕΛΑΤΗ": new_only["mobile"],
    })

    wb = Workbook()
    ws = wb.active
    ws.title = "New Flow"
    ws.append(list(result.columns))
    for row in result.itertuples(index=False):
        ws.append(list(row))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read(), len(result)



def compare_newflow_bytes(
    old_bytes: bytes,
    new_bytes: bytes,
    sheet_name: str | None = None
) -> tuple[bytes, int]:
    """
    Senin verdiğin script mantığı:
    - Sheet: "New flow" (ya da sheet_name)
    - Anahtar: SRID
    - Mapping:
        SR <- SRID
        ADDRESS <- full_adr
        A/K <- a/k
        BUILDING ID <- building Id
        ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ <- customer
        ΚΙΝΗΤΟ ΠΕΛΑΤΗ <- mobile
    - Çıktı: tek sheet "New Flow"
    """
    sheet = sheet_name or NEWFLOW_SHEET_FALLBACK

    old_df = pd.read_excel(io.BytesIO(old_bytes), sheet_name=sheet)
    new_df = pd.read_excel(io.BytesIO(new_bytes), sheet_name=sheet)

    if "SRID" not in old_df.columns or "SRID" not in new_df.columns:
        raise ValueError("New flow sekmesinde 'SRID' kolonu bulunamadı.")

    old_sr = set(old_df["SRID"].astype(str))
    new_df = new_df.copy()
    new_df["SRID"] = new_df["SRID"].astype(str)

    new_only = new_df[~new_df["SRID"].isin(old_sr)].copy()

    required = ["SRID", "full_adr", "a/k", "building Id", "customer", "mobile"]
    missing = [c for c in required if c not in new_only.columns]
    if missing:
        raise ValueError(f"New flow sekmesinde eksik kolon(lar): {missing}")

    result = pd.DataFrame({
        "SR": new_only["SRID"],
        "ADDRESS": new_only["full_adr"],
        "A/K": new_only["a/k"],
        "BUILDING ID": new_only["building Id"],
        "ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ": new_only["customer"],
        "ΚΙΝΗΤΟ ΠΕΛΑΤΗ": new_only["mobile"],
    })

    wb = Workbook()
    ws = wb.active
    ws.title = "New Flow"
    ws.append(list(result.columns))
    for row in result.itertuples(index=False):
        ws.append(list(row))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return buf.read(), len(result)