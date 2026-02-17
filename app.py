import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="SECUNET TPBE – SR Karşılaştırma", page_icon="📄", layout="wide")

SHEET_FALLBACK = "Ανατεθειμένες αυτοψίες"
CIKTI_KOLONLARI = [
    "SR",
    "ADDRESS",
    "A/K",
    "BUILDING ID",
    "ΌΝΟΜΑΤΕΠΩΝΥΜΟ ΠΕΛΑΤΗ",
    "ΚΙΝΗΤΟ ΠΕΛΑΤΗ",
]

# Yazım farklarına tolerans (bazı dosyalarda ADRESS gelebiliyor)
ALIAS = {
    "ADRESS": "ADDRESS",
    "Adress": "ADDRESS",
    "Address": "ADDRESS",
}

def kolonlari_normalize_et(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    df = df.rename(columns={c: ALIAS.get(c, c) for c in df.columns})
    return df

def sheet_listesi(dosya) -> list[str]:
    xls = pd.ExcelFile(dosya)
    return list(xls.sheet_names)

def sheet_oku(dosya, sheet_adi: str) -> pd.DataFrame:
    return pd.read_excel(dosya, sheet_name=sheet_adi)

# ---- UI ----
st.title("📄 SECUNET TPBE – Yeni SR Bulucu")
st.caption("İki Excel yükle → aynı sayfayı (sheet) seç → yeni gelen SR’ları Excel olarak indir.")

st.markdown("---")

sol, sag = st.columns([1, 2], gap="large")

with sol:
    st.subheader("1) Dosyaları yükle")
    eski_dosya = st.file_uploader("ESKİ TARİH (.xlsx)", type=["xlsx"], key="old")
    yeni_dosya = st.file_uploader("YENİ TARİH (.xlsx)", type=["xlsx"], key="new")

    if yeni_dosya:
        try:
            sheetler = sheet_listesi(yeni_dosya)
            varsayilan = sheetler.index(SHEET_FALLBACK) if SHEET_FALLBACK in sheetler else 0
            sheet_adi = st.selectbox("Sayfa (Sheet) seç", sheetler, index=varsayilan)
        except Exception:
            sheet_adi = st.text_input("Sayfa (Sheet) adı (manuel)", value=SHEET_FALLBACK)
    else:
        sheet_adi = st.text_input("Sayfa (Sheet) adı", value=SHEET_FALLBACK)

    st.markdown("**Çıktı kolonları:**")
    st.write(", ".join(CIKTI_KOLONLARI))

    karsilastir = st.button("🔎 Karşılaştır", type="primary", use_container_width=True)

with sag:
    st.subheader("2) Sonuç")

    if eski_dosya and yeni_dosya and karsilastir:
        try:
            with st.spinner("Dosyalar okunuyor ve karşılaştırılıyor..."):
                eski_df = kolonlari_normalize_et(sheet_oku(eski_dosya, sheet_adi))
                yeni_df = kolonlari_normalize_et(sheet_oku(yeni_dosya, sheet_adi))

                if "SR" not in eski_df.columns or "SR" not in yeni_df.columns:
                    st.error("SR kolonu bulunamadı. Doğru sayfayı (sheet) seçtiğine emin misin?")
                    st.stop()

                yeni_gelenler = yeni_df[~yeni_df["SR"].isin(eski_df["SR"])].copy()

                eksikler = [c for c in CIKTI_KOLONLARI if c not in yeni_gelenler.columns]
                if eksikler:
                    st.error(f"Beklenen kolon(lar) eksik: {eksikler}")
                    with st.expander("Yeni dosya kolonlarını göster"):
                        st.write(list(yeni_df.columns))
                    st.stop()

                cikti_df = yeni_gelenler[CIKTI_KOLONLARI].copy()

            c1, c2, c3 = st.columns(3)
            c1.metric("Yeni gelen SR", len(cikti_df))
            c2.metric("Eski dosya SR (toplam)", eski_df["SR"].nunique())
            c3.metric("Yeni dosya SR (toplam)", yeni_df["SR"].nunique())

            st.markdown("### Yeni gelen SR listesi")
            st.dataframe(cikti_df, use_container_width=True, height=520)

            # Excel çıktısı
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                cikti_df.to_excel(writer, index=False, sheet_name="Yeni_SR")
            buf.seek(0)

            st.download_button(
                "📥 Excel’i indir (YENI_GELEN_SR.xlsx)",
                data=buf,
                file_name="YENI_GELEN_SR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Hata: {e}")
    else:
        st.info("Soldan dosyaları yükleyip **Karşılaştır**’a basınca sonuç burada görünecek.")

# Hafif görünüm iyileştirme (opsiyonel)
st.markdown(
    """
    <style>
      .stMetric { background: rgba(255,255,255,0.04); padding: 14px; border-radius: 14px; }
      [data-testid="stFileUploader"] { padding: 10px; border-radius: 12px; }
      .stDownloadButton button { padding: 10px 14px; border-radius: 12px; }
    </style>
    """,
    unsafe_allow_html=True,
)
