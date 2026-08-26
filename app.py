"""
Dashboard Evaluasi Penyelenggaraan Pelatihan
=============================================
Aplikasi Streamlit untuk memfilter & menganalisis data hasil Google Form
evaluasi pelatihan, terutama berdasarkan program pelatihan yang diikuti.

Didesain tahan terhadap variasi kecil nama kolom antar sheet/file
(beda huruf besar-kecil, spasi ekstra, newline, dsb) dengan cara mencocokkan
kolom secara otomatis (fuzzy matching). Kalau tetap tidak ketemu, ada
dropdown manual di sidebar untuk memilih kolom yang benar.

Cara menjalankan:
    streamlit run app.py
"""

import io
import re
import difflib

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


def normalize(text: str) -> str:
    """Normalisasi teks kolom: lowercase, hapus spasi berlebih/newline,
    supaya 'Program Pelatihan Yang Diikuti ' == 'Program pelatihan yang di ikuti'."""
    text = str(text).lower()
    text = text.replace("\n", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)  # buang tanda baca, gabung jadi spasi
    text = re.sub(r"\s+", " ", text).strip()
    return text


def best_match(canonical: str, columns, cutoff: float = 0.55):
    """Cari kolom di `columns` yang paling mirip dengan teks `canonical`."""
    norm_canonical = normalize(canonical)
    norm_cols = {c: normalize(c) for c in columns}

    # 1) exact match setelah normalisasi
    for original, norm in norm_cols.items():
        if norm == norm_canonical:
            return original

    # 2) substring match (misal canonical adalah bagian dari nama kolom atau sebaliknya)
    for original, norm in norm_cols.items():
        if norm_canonical in norm or norm in norm_canonical:
            return original

    # 3) fuzzy match pakai difflib
    matches = difflib.get_close_matches(
        norm_canonical, list(norm_cols.values()), n=1, cutoff=cutoff
    )
    if matches:
        for original, norm in norm_cols.items():
            if norm == matches[0]:
                return original
    return None


# --------------------------------------------------------------------------
# DEFINISI "PERTANYAAN KANONIK" — teks acuan untuk mencocokkan kolom
# --------------------------------------------------------------------------
PROGRAM_CANONICAL = "Program pelatihan yang diikuti"
TIMESTAMP_CANONICAL = "Timestamp"
INFO_SOURCE_CANONICAL = "Dari mana anda mendapatkan informasi tentang pelatihan"

# label ringkas -> teks pertanyaan acuan (dipakai untuk matching & label grafik)
SCORE_QUESTIONS = {
    "Info mudah didapat": "Apakah informasi pelatihan mudah untuk didapatkan?",
    "Pendaftaran mudah": "Apakah pendaftaran dan tahapannya mudah untuk dilakukan?",
    "Petunjuk pendaftaran jelas": "Apakah petunjuk tata cara pendaftaran jelas dan mudah dipahami?",
    "Program jelas": "Apakah program pelatihan jelas dan mudah dipahami?",
    "Program menarik": "Apakah program pelatihan menarik?",
    "Program bermanfaat": "Apakah program pelatihan bermanfaat?",
    "Kompetensi meningkat": "Apakah program pelatihan berhasil meningkatkan kompetensi Anda?",
    "Durasi sesuai": "Apakah durasi untuk menyelesaikan pelatihan sudah sesuai?",
    "Instruktur menguasai materi": "Apakah Instrukur menguasai program pelatihan yang disampaikan?",
    "Kemampuan penyampaian instruktur": "Bagaimana kemampuan Instruktur dalam menyampaikan program pelatihan?",
    "Kemampuan mengelola peserta": "Bagaimana kemampuan Instruktur dalam mengelola Peserta Pelatihan?",
    "Sikap & teladan instruktur": "Bagaimana sikap, disiplin, penampilan, dan teladan Instruktur selama pelatihan?",
    "Pelayanan petugas": "Bagaimana pelayanan petugas terhadap Peserta Pelatihan?",
    "Jadwal sesuai rencana": "Apakah pelaksanaan jadwal pelatihan sudah sesuai dengan rencana?",
    "Perlengkapan tepat waktu": "Apakah perlengkapan Peserta Pelatihan training material diberikan tepat waktu",
    "Sarana pelatihan memadai": "Apakah sarana prasarana fasilitas pelatihan sudah memadai",
    "Sarana penunjang memadai": "Apakah sarana prasarana fasilitas penunjang pelatihan sudah memadai",
}

COMMENT_LABELS = {
    "Komentar/saran program pelatihan": "Komentar dan saran Anda terhadap program pelatihan",
    "Komentar/saran instruktur": "Komentar dan saran Anda terhadap Instruktur",
    "Komentar/saran penyelenggaraan": "Komentar dan saran Anda terhadap penyelenggaraan pelatihan",
    "Keluhan penyelenggaraan": "Keluhan terhadap penyelenggaraan pelatihan",
}


# --------------------------------------------------------------------------
# LOAD DATA
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="Membaca file Excel...")
def get_sheet_names(file):
    return pd.ExcelFile(file).sheet_names


@st.cache_data(show_spinner="Membaca data sheet...")
def load_data(file, sheet_name) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=sheet_name)
    df.columns = [str(c) for c in df.columns]
    return df


# --------------------------------------------------------------------------
# SIDEBAR - UPLOAD
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

sheet_names = get_sheet_names(uploaded_file)
if len(sheet_names) > 1:
    sheet_choice = st.sidebar.selectbox("Pilih sheet", options=sheet_names, index=0)
else:
    sheet_choice = sheet_names[0]

df_raw = load_data(uploaded_file, sheet_choice)
all_columns = list(df_raw.columns)

# --------------------------------------------------------------------------
# AUTO-MATCH KOLOM PENTING (dengan fallback manual)
# --------------------------------------------------------------------------
auto_program_col = best_match(PROGRAM_CANONICAL, all_columns)
auto_timestamp_col = best_match(TIMESTAMP_CANONICAL, all_columns)
auto_info_col = best_match(INFO_SOURCE_CANONICAL, all_columns)

with st.sidebar.expander(
    "🔧 Pemetaan kolom" + ("" if auto_program_col else " ⚠️ perlu dicek"),
    expanded=(auto_program_col is None),
):
    st.caption(
        "Sistem mencocokkan kolom secara otomatis. Kalau salah / tidak ketemu, "
        "pilih manual di sini."
    )
    options_with_none = ["(tidak ada)"] + all_columns

    program_col = st.selectbox(
        "Kolom: Program pelatihan yang diikuti",
        options=all_columns,
        index=all_columns.index(auto_program_col) if auto_program_col in all_columns else 0,
    )

    ts_index = (
        options_with_none.index(auto_timestamp_col)
        if auto_timestamp_col in options_with_none
        else 0
    )
    timestamp_col_choice = st.selectbox(
        "Kolom: Timestamp (opsional)", options=options_with_none, index=ts_index
    )
    timestamp_col = None if timestamp_col_choice == "(tidak ada)" else timestamp_col_choice

    info_index = (
        options_with_none.index(auto_info_col) if auto_info_col in options_with_none else 0
    )
    info_col_choice = st.selectbox(
        "Kolom: Sumber informasi pelatihan (opsional)",
        options=options_with_none,
        index=info_index,
    )
    info_col = None if info_col_choice == "(tidak ada)" else info_col_choice

if not program_col:
    st.error(
        "Tidak berhasil menemukan kolom 'Program pelatihan yang diikuti'. "
        "Silakan pilih manual di panel **🔧 Pemetaan kolom** pada sidebar."
    )
    st.stop()

# cocokkan kolom-kolom skor & komentar
score_col_map = {}  # label ringkas -> nama kolom asli di df
for label, question in SCORE_QUESTIONS.items():
    match = best_match(question, all_columns)
    if match:
        score_col_map[label] = match

comment_col_map = {}  # label ringkas -> nama kolom asli di df
for label, question in COMMENT_LABELS.items():
    match = best_match(question, all_columns)
    if match:
        comment_col_map[label] = match

# pastikan kolom skor benar-benar numerik (buang kalau ternyata bukan)
for label, col in list(score_col_map.items()):
    numeric_series = pd.to_numeric(df_raw[col], errors="coerce")
    if numeric_series.notna().sum() == 0:
        del score_col_map[label]
    else:
        df_raw[col] = numeric_series

# --------------------------------------------------------------------------
# SIDEBAR - FILTER
# --------------------------------------------------------------------------
program_options = sorted(df_raw[program_col].dropna().astype(str).unique().tolist())
selected_programs = st.sidebar.multiselect(
    "🎓 Program pelatihan yang diikuti",
    options=program_options,
    default=program_options,
    help="Kosongkan lalu pilih ulang untuk memfilter satu/lebih program tertentu.",
)

df_work = df_raw.copy()
date_range = None
if timestamp_col:
    df_work[timestamp_col] = pd.to_datetime(df_work[timestamp_col], errors="coerce")
    min_date = df_work[timestamp_col].min()
    max_date = df_work[timestamp_col].max()
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "📅 Rentang tanggal pengisian",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
        )

selected_sources = None
if info_col:
    source_options = sorted(df_work[info_col].dropna().astype(str).unique().tolist())
    with st.sidebar.expander("Filter tambahan: sumber informasi"):
        selected_sources = st.multiselect(
            "Sumber informasi pelatihan", options=source_options, default=source_options
        )

# --------------------------------------------------------------------------
# APPLY FILTERS
# --------------------------------------------------------------------------
df_filtered = df_work[df_work[program_col].astype(str).isin(selected_programs)]

if timestamp_col and date_range and len(date_range) == 2:
    start, end = date_range
    df_filtered = df_filtered[
        (df_filtered[timestamp_col].dt.date >= start)
        & (df_filtered[timestamp_col].dt.date <= end)
    ]

if info_col and selected_sources is not None:
    df_filtered = df_filtered[df_filtered[info_col].astype(str).isin(selected_sources)]

# --------------------------------------------------------------------------
# HEADER & METRIK RINGKAS
# --------------------------------------------------------------------------
st.title("📊 Dashboard Evaluasi Penyelenggaraan Pelatihan")
st.caption(
    f"Sheet **'{sheet_choice}'** — Menampilkan **{len(df_filtered)}** dari "
    f"**{len(df_raw)}** total responden berdasarkan filter yang dipilih."
)

score_cols = list(score_col_map.values())

col1, col2, col3 = st.columns(3)
col1.metric("Jumlah Responden (terfilter)", len(df_filtered))
col2.metric("Jumlah Program Terpilih", len(selected_programs))
if score_cols:
    overall_avg = df_filtered[score_cols].mean(numeric_only=True).mean()
    col3.metric(
        "Rata-rata Skor Keseluruhan",
        f"{overall_avg:.2f} / 5" if pd.notna(overall_avg) else "-",
    )

if len(score_col_map) < len(SCORE_QUESTIONS):
    missing = len(SCORE_QUESTIONS) - len(score_col_map)
    st.caption(f"ℹ️ {missing} pertanyaan skor tidak terdeteksi di sheet ini dan dilewati.")

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
    if not score_col_map:
        st.warning("Tidak ada kolom skor numerik yang terdeteksi pada sheet ini.")
    else:
        st.subheader("Jumlah Responden per Program")
        count_df = (
            df_filtered[program_col]
            .astype(str)
            .value_counts()
            .rename_axis("Program")
            .reset_index(name="Jumlah")
        )
        fig_count = px.bar(count_df, x="Jumlah", y="Program", orientation="h", text="Jumlah")
        fig_count.update_layout(yaxis_title="", xaxis_title="Jumlah Responden")
        st.plotly_chart(fig_count, use_container_width=True)

        st.subheader("Rata-rata Skor per Aspek Pelatihan")
        avg_scores = df_filtered[score_cols].mean(numeric_only=True).reset_index()
        avg_scores.columns = ["col", "Rata-rata"]
        inv_map = {v: k for k, v in score_col_map.items()}
        avg_scores["Aspek"] = avg_scores["col"].map(inv_map)
        avg_scores = avg_scores.sort_values("Rata-rata", ascending=True)
        fig_avg = px.bar(
            avg_scores,
            x="Rata-rata",
            y="Aspek",
            orientation="h",
            range_x=[0, 5],
            text=avg_scores["Rata-rata"].round(2),
        )
        fig_avg.update_layout(yaxis_title="", xaxis_title="Rata-rata Skor (1-5)")
        st.plotly_chart(fig_avg, use_container_width=True)

        if len(selected_programs) > 1:
            st.subheader("Perbandingan Rata-rata Skor Antar Program")
            melt_df = df_filtered.melt(
                id_vars=[program_col], value_vars=score_cols, var_name="col", value_name="Skor"
            )
            melt_df["Aspek"] = melt_df["col"].map(inv_map)
            comp_df = melt_df.groupby([program_col, "Aspek"])["Skor"].mean().reset_index()
            fig_comp = px.bar(
                comp_df,
                x="Skor",
                y="Aspek",
                color=program_col,
                orientation="h",
                barmode="group",
                range_x=[0, 5],
            )
            fig_comp.update_layout(
                yaxis_title="", xaxis_title="Rata-rata Skor (1-5)", legend_title="Program"
            )
            st.plotly_chart(fig_comp, use_container_width=True)

        if info_col:
            st.subheader("Sumber Informasi Pelatihan")
            src_df = (
                df_filtered[info_col]
                .astype(str)
                .value_counts()
                .rename_axis("Sumber")
                .reset_index(name="Jumlah")
            )
            fig_src = px.pie(src_df, names="Sumber", values="Jumlah", hole=0.4)
            st.plotly_chart(fig_src, use_container_width=True)

# --- TAB 3: KOMENTAR & KELUHAN ---
with tab_komentar:
    st.subheader("Komentar, Saran, dan Keluhan Peserta")

    if not comment_col_map:
        st.info("Tidak ada kolom komentar/keluhan yang terdeteksi pada sheet ini.")
    else:
        picked_label = st.selectbox(
            "Pilih jenis komentar yang ingin dilihat", options=list(comment_col_map.keys())
        )
        picked_col = comment_col_map[picked_label]

        comments = df_filtered[[program_col, picked_col]].dropna(subset=[picked_col])
        comments = comments[comments[picked_col].astype(str).str.strip() != ""]

        if comments.empty:
            st.info("Tidak ada komentar untuk kombinasi filter ini.")
        else:
            for program, group in comments.groupby(program_col):
                with st.expander(f"**{program}** ({len(group)} komentar)"):
                    for _, row in group.iterrows():
                        st.markdown(f"- {row[picked_col]}")