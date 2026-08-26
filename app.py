"""
Dashboard Evaluasi Penyelenggaraan Pelatihan
=============================================
Aplikasi Streamlit untuk memfilter & menganalisis data hasil Google Form
evaluasi pelatihan, terutama berdasarkan "Program pelatihan yang di ikuti".

Cara menjalankan:
    streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st
import plotly.express as px

# --------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Evaluasi Pelatihan",
    page_icon="📊",
    layout="wide",
)

PROGRAM_COL = "Program pelatihan yang di ikuti"
TIMESTAMP_COL = "Timestamp"

# Kolom-kolom skor (Likert 1-5). Kita deteksi otomatis kolom numerik nanti,
# tapi daftar ini dipakai untuk urutan tampilan & label ringkas di grafik.
SCORE_LABELS = {
    "Apakah informasi pelatihan mudah\nuntuk didapatkan?": "Info mudah didapat",
    "Apakah pendaftaran dan tahapannya mudah untuk dilakukan?": "Pendaftaran mudah",
    "Apakah petunjuk tata cara\npendaftaran jelas dan mudah\ndipahami?": "Petunjuk pendaftaran jelas",
    "Apakah program pelatihan jelas\ndan mudah dipahami?": "Program jelas",
    "Apakah program pelatihan\nmenarik?": "Program menarik",
    "Apakah program pelatihan\nbermanfaat?": "Program bermanfaat",
    "Apakah program pelatihan berhasil\nmeningkatkan kompetensi Anda?": "Kompetensi meningkat",
    "Apakah durasi untuk menyelesaikan pelatihan sudah sesuai?": "Durasi sesuai",
    "Apakah Instrukur menguasai program pelatihan yang disampaikan?": "Instruktur menguasai materi",
    "Bagaimana kemampuan Instruktur dalam menyampaikan program pelatihan?": "Kemampuan penyampaian instruktur",
    "Bagaimana kemampuan Instruktur\ndalam mengelola Peserta Pelatihan?": "Kemampuan mengelola peserta",
    "Bagaimana sikap, disiplin, penampilan, dan teladan Instruktur selama pelatihan?": "Sikap & teladan instruktur",
    "Bagaimana pelayanan petugas\nterhadap Peserta Pelatihan?": "Pelayanan petugas",
    "Apakah pelaksanaan jadwal\npelatihan sudah sesuai dengan\nrencana?": "Jadwal sesuai rencana",
    "Apakah perlengkapan Peserta Pelatihan (training material) diberikan tepat waktu? (contoh: atribut pelatihan/seragam\n/Alat Pelindung Diri/modul/materi\n/ATK/bahan/ konten/dll)": "Perlengkapan tepat waktu",
    "Apakah sarana/prasarana/fasilitas\npelatihan sudah memadai?\n(contoh: kelas/workshop/mesin/\nalat/website sistem manajemen\npembelajaran/dll)": "Sarana pelatihan memadai",
    "Apakah sarana/prasarana/fasilitas\npenunjang pelatihan sudah\nmemadai?\n(contoh: asrama/tempat ibadah/\nkantin/toilet/perpustakaan/\nwebsite lembaga pelatihan/dll)": "Sarana penunjang memadai",
}

COMMENT_COLS = [
    "Komentar dan saran Anda terhadap\nprogram pelatihan",
    "Komentar dan saran Anda terhadap\nInstruktur",
    "Komentar dan saran Anda terhadap penyelenggaraan pelatihan ",
    "Keluhan terhadap penyelenggaraan pelatihan ",
]


# --------------------------------------------------------------------------
# LOAD DATA
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="Membaca file Excel...")
def load_data(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    df.columns = [str(c).strip("\n").strip() if c not in df.columns else c for c in df.columns]
    # samakan whitespace di nama kolom tapi tetap simpan aslinya untuk mapping
    return df


def clean_label(col: str) -> str:
    """Ubah nama kolom panjang jadi label ringkas untuk grafik."""
    return SCORE_LABELS.get(col, col.replace("\n", " ").strip())


# --------------------------------------------------------------------------
# SIDEBAR - UPLOAD & FILTER
# --------------------------------------------------------------------------
st.sidebar.title("⚙️ Data & Filter")

uploaded_file = st.sidebar.file_uploader(
    "Upload file Excel hasil Google Form (.xlsx)",
    type=["xlsx"],
)

if uploaded_file is None:
    st.title("📊 Dashboard Evaluasi Penyelenggaraan Pelatihan")
    st.info(
        "👈 Silakan upload file **Excel hasil Google Form** (.xlsx) di sidebar "
        "kiri untuk mulai memfilter dan menganalisis data."
    )
    st.stop()

df_raw = load_data(uploaded_file)

if PROGRAM_COL not in df_raw.columns:
    st.error(
        f"Kolom **'{PROGRAM_COL}'** tidak ditemukan di file yang diupload. "
        f"Kolom yang tersedia: {list(df_raw.columns)}"
    )
    st.stop()

# --- Filter utama: Program Pelatihan ---
program_options = sorted(df_raw[PROGRAM_COL].dropna().unique().tolist())
selected_programs = st.sidebar.multiselect(
    "🎓 Program pelatihan yang diikuti",
    options=program_options,
    default=program_options,
    help="Kosongkan pilihan lalu pilih ulang untuk memfilter satu/lebih program tertentu.",
)

# --- Filter tanggal (kalau ada kolom Timestamp) ---
df_work = df_raw.copy()
if TIMESTAMP_COL in df_work.columns:
    df_work[TIMESTAMP_COL] = pd.to_datetime(df_work[TIMESTAMP_COL], errors="coerce")
    min_date = df_work[TIMESTAMP_COL].min()
    max_date = df_work[TIMESTAMP_COL].max()
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "📅 Rentang tanggal pengisian",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
        )
    else:
        date_range = None
else:
    date_range = None

# --- Filter tambahan: sumber informasi (opsional, kalau ada) ---
info_col = "Dari mana anda mendapatkan informasi tentang pelatihan?"
selected_sources = None
if info_col in df_work.columns:
    source_options = sorted(df_work[info_col].dropna().unique().tolist())
    with st.sidebar.expander("Filter tambahan: sumber informasi"):
        selected_sources = st.multiselect(
            "Sumber informasi pelatihan",
            options=source_options,
            default=source_options,
        )

# --------------------------------------------------------------------------
# APPLY FILTERS
# --------------------------------------------------------------------------
df_filtered = df_work[df_work[PROGRAM_COL].isin(selected_programs)]

if date_range and len(date_range) == 2:
    start, end = date_range
    df_filtered = df_filtered[
        (df_filtered[TIMESTAMP_COL].dt.date >= start)
        & (df_filtered[TIMESTAMP_COL].dt.date <= end)
    ]

if selected_sources is not None:
    df_filtered = df_filtered[df_filtered[info_col].isin(selected_sources)]

# --------------------------------------------------------------------------
# HEADER & METRIK RINGKAS
# --------------------------------------------------------------------------
st.title("📊 Dashboard Evaluasi Penyelenggaraan Pelatihan")
st.caption(
    f"Menampilkan **{len(df_filtered)}** dari **{len(df_raw)}** total responden "
    f"berdasarkan filter yang dipilih."
)

score_cols = [c for c in SCORE_LABELS if c in df_filtered.columns]

col1, col2, col3 = st.columns(3)
col1.metric("Jumlah Responden (terfilter)", len(df_filtered))
col2.metric("Jumlah Program Terpilih", len(selected_programs))
if score_cols:
    overall_avg = df_filtered[score_cols].mean(numeric_only=True).mean()
    col3.metric("Rata-rata Skor Keseluruhan", f"{overall_avg:.2f} / 5" if pd.notna(overall_avg) else "-")

st.divider()

# --------------------------------------------------------------------------
# TABS
# --------------------------------------------------------------------------
tab_data, tab_ringkasan, tab_komentar = st.tabs(
    ["📄 Data Terfilter", "📈 Ringkasan & Grafik", "💬 Komentar & Keluhan"]
)

# --- TAB 1: DATA TABEL ---
with tab_data:
    st.subheader("Data Responden (sesuai filter)")
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)

    # tombol download hasil filter
    csv_bytes = df_filtered.to_csv(index=False).encode("utf-8-sig")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_filtered.to_excel(writer, index=False, sheet_name="Data Terfilter")
    excel_bytes = buf.getvalue()

    dcol1, dcol2 = st.columns(2)
    dcol1.download_button(
        "⬇️ Download CSV",
        data=csv_bytes,
        file_name="evaluasi_pelatihan_terfilter.csv",
        mime="text/csv",
        use_container_width=True,
    )
    dcol2.download_button(
        "⬇️ Download Excel",
        data=excel_bytes,
        file_name="evaluasi_pelatihan_terfilter.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# --- TAB 2: RINGKASAN & GRAFIK ---
with tab_ringkasan:
    if not score_cols:
        st.warning("Tidak ada kolom skor numerik yang terdeteksi pada data ini.")
    else:
        st.subheader("Jumlah Responden per Program")
        count_df = (
            df_filtered[PROGRAM_COL]
            .value_counts()
            .rename_axis("Program")
            .reset_index(name="Jumlah")
        )
        fig_count = px.bar(
            count_df, x="Jumlah", y="Program", orientation="h", text="Jumlah"
        )
        fig_count.update_layout(yaxis_title="", xaxis_title="Jumlah Responden")
        st.plotly_chart(fig_count, use_container_width=True)

        st.subheader("Rata-rata Skor per Aspek Pelatihan")
        avg_scores = (
            df_filtered[score_cols]
            .mean(numeric_only=True)
            .reset_index()
        )
        avg_scores.columns = ["Aspek", "Rata-rata"]
        avg_scores["Aspek"] = avg_scores["Aspek"].apply(clean_label)
        avg_scores = avg_scores.sort_values("Rata-rata", ascending=True)
        fig_avg = px.bar(
            avg_scores, x="Rata-rata", y="Aspek", orientation="h", range_x=[0, 5],
            text=avg_scores["Rata-rata"].round(2),
        )
        fig_avg.update_layout(yaxis_title="", xaxis_title="Rata-rata Skor (1-5)")
        st.plotly_chart(fig_avg, use_container_width=True, height=600)

        if len(selected_programs) > 1:
            st.subheader("Perbandingan Rata-rata Skor Antar Program")
            melt_df = df_filtered.melt(
                id_vars=[PROGRAM_COL], value_vars=score_cols,
                var_name="Aspek", value_name="Skor"
            )
            melt_df["Aspek"] = melt_df["Aspek"].apply(clean_label)
            comp_df = (
                melt_df.groupby([PROGRAM_COL, "Aspek"])["Skor"]
                .mean()
                .reset_index()
            )
            fig_comp = px.bar(
                comp_df, x="Skor", y="Aspek", color=PROGRAM_COL,
                orientation="h", barmode="group", range_x=[0, 5],
            )
            fig_comp.update_layout(yaxis_title="", xaxis_title="Rata-rata Skor (1-5)", legend_title="Program")
            st.plotly_chart(fig_comp, use_container_width=True)

        if info_col in df_filtered.columns:
            st.subheader("Sumber Informasi Pelatihan")
            src_df = (
                df_filtered[info_col]
                .value_counts()
                .rename_axis("Sumber")
                .reset_index(name="Jumlah")
            )
            fig_src = px.pie(src_df, names="Sumber", values="Jumlah", hole=0.4)
            st.plotly_chart(fig_src, use_container_width=True)

# --- TAB 3: KOMENTAR & KELUHAN ---
with tab_komentar:
    st.subheader("Komentar, Saran, dan Keluhan Peserta")
    available_comment_cols = [c for c in COMMENT_COLS if c in df_filtered.columns]

    if not available_comment_cols:
        st.info("Tidak ada kolom komentar/keluhan pada data ini.")
    else:
        picked_col = st.selectbox(
            "Pilih jenis komentar yang ingin dilihat",
            options=available_comment_cols,
            format_func=lambda c: c.replace("\n", " ").strip(),
        )
        comments = df_filtered[[PROGRAM_COL, picked_col]].dropna(subset=[picked_col])
        comments = comments[comments[picked_col].astype(str).str.strip() != ""]

        if comments.empty:
            st.info("Tidak ada komentar untuk kombinasi filter ini.")
        else:
            for program, group in comments.groupby(PROGRAM_COL):
                with st.expander(f"**{program}** ({len(group)} komentar)"):
                    for _, row in group.iterrows():
                        st.markdown(f"- {row[picked_col]}")
