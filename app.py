import io
import re
import difflib
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px

from report_generator import generate_official_report, convert_xlsx_to_pdf, MAX_RESPONDENTS, DEFAULT_CAPACITY

# --------------------------------------------------------------------------
# KONFIGURASI HALAMAN & GAYA VISUAL
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Evaluasi Pelatihan",
    page_icon=":bar_chart:",
    layout="wide",
)

APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "assets" / "template_laporan_resmi.xlsx"

PRIMARY = "#1F3B57"      # navy - warna utama
ACCENT = "#2E6F95"       # biru kalem - aksen
MUTED = "#6B7A8F"        # abu kebiruan - teks sekunder
BG_SOFT = "#F4F6F8"

st.markdown(
    f"""
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css">
    <style>
        .block-container {{ padding-top: 2rem; }}

        h1, h2, h3 {{ color: {PRIMARY}; font-weight: 700; }}

        .app-header {{
            display: flex; align-items: center; gap: 0.6rem;
            margin-bottom: 0.1rem;
        }}
        .app-header i {{ font-size: 1.6rem; color: {ACCENT}; }}
        .app-subtitle {{ color: {MUTED}; font-size: 0.95rem; margin-top: -0.4rem; }}

        .section-title {{
            display: flex; align-items: center; gap: 0.5rem;
            color: {PRIMARY}; font-weight: 600; font-size: 1.15rem;
            margin-top: 0.5rem;
        }}
        .section-title i {{ color: {ACCENT}; }}

        div[data-testid="stMetric"] {{
            background: {BG_SOFT};
            border: 1px solid #E2E8EF;
            border-radius: 10px;
            padding: 0.9rem 1rem;
        }}

        .stTabs [data-baseweb="tab"] {{
            font-weight: 600;
        }}

        .info-note {{
            background: {BG_SOFT};
            border-left: 3px solid {ACCENT};
            padding: 0.6rem 0.9rem;
            border-radius: 4px;
            color: {MUTED};
            font-size: 0.9rem;
        }}
    </style>
    <div class="app-header">
        <i class="bi bi-bar-chart-line-fill"></i>
        <h1 style="margin:0;">Dashboard Evaluasi Penyelenggaraan Pelatihan</h1>
    </div>
    <div class="app-subtitle">Filter, analisis, dan ekspor data evaluasi peserta pelatihan</div>
    """,
    unsafe_allow_html=True,
)


def section_title(icon: str, text: str):
    st.markdown(
        f'<div class="section-title"><i class="bi {icon}"></i>{text}</div>',
        unsafe_allow_html=True,
    )


# --------------------------------------------------------------------------
# UTILITAS PENCOCOKAN KOLOM (tahan variasi nama kolom antar file/sheet)
# --------------------------------------------------------------------------
def normalize(text: str) -> str:
    text = str(text).lower().replace("\n", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def best_match(canonical: str, columns, cutoff: float = 0.55):
    norm_canonical = normalize(canonical)
    norm_cols = {c: normalize(c) for c in columns}
    for original, norm in norm_cols.items():
        if norm == norm_canonical:
            return original
    for original, norm in norm_cols.items():
        if norm_canonical in norm or norm in norm_canonical:
            return original
    matches = difflib.get_close_matches(norm_canonical, list(norm_cols.values()), n=1, cutoff=cutoff)
    if matches:
        for original, norm in norm_cols.items():
            if norm == matches[0]:
                return original
    return None


# --------------------------------------------------------------------------
# DEFINISI PERTANYAAN KANONIK
# --------------------------------------------------------------------------
PROGRAM_CANONICAL = "Program pelatihan yang diikuti"
TIMESTAMP_CANONICAL = "Timestamp"
INFO_SOURCE_CANONICAL = "Dari mana anda mendapatkan informasi tentang pelatihan"

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
with st.sidebar:
    st.markdown(
        '<div class="section-title"><i class="bi bi-sliders"></i>Data &amp; Filter</div>',
        unsafe_allow_html=True,
    )

uploaded_file = st.sidebar.file_uploader(
    "Upload file Excel hasil Google Form (.xlsx)",
    type=["xlsx"],
)

if uploaded_file is None:
    st.markdown(
        '<div class="info-note"><i class="bi bi-info-circle"></i>&nbsp; '
        "Silakan upload file Excel hasil Google Form (.xlsx) di sidebar kiri "
        "untuk mulai memfilter dan menganalisis data.</div>",
        unsafe_allow_html=True,
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
    "Pemetaan kolom" + ("" if auto_program_col else " -- perlu dicek"),
    expanded=(auto_program_col is None),
):
    st.caption(
        "Kolom dicocokkan otomatis. Kalau salah / tidak ketemu, pilih manual di sini."
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
        "Silakan pilih manual di panel Pemetaan kolom pada sidebar."
    )
    st.stop()

# cocokkan kolom-kolom skor & komentar
score_col_map = {}
for label, question in SCORE_QUESTIONS.items():
    match = best_match(question, all_columns)
    if match:
        score_col_map[label] = match

comment_col_map = {}
for label, question in COMMENT_LABELS.items():
    match = best_match(question, all_columns)
    if match:
        comment_col_map[label] = match

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
    "Program pelatihan yang diikuti",
    options=program_options,
    default=program_options,
    help="Kosongkan lalu pilih ulang untuk memfilter satu/lebih program tertentu.",
)

df_work = df_raw.copy()
if timestamp_col:
    df_work[timestamp_col] = pd.to_datetime(df_work[timestamp_col], errors="coerce")

# --------------------------------------------------------------------------
# APPLY FILTERS
# --------------------------------------------------------------------------
df_filtered = df_work[df_work[program_col].astype(str).isin(selected_programs)]

# --------------------------------------------------------------------------
# METRIK RINGKAS
# --------------------------------------------------------------------------
st.caption(
    f"Sheet '{sheet_choice}' -- Menampilkan {len(df_filtered)} dari "
    f"{len(df_raw)} total responden berdasarkan filter yang dipilih."
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
    st.caption(f"Catatan: {missing} pertanyaan skor tidak terdeteksi di sheet ini dan dilewati.")

st.divider()

# --------------------------------------------------------------------------
# TABS
# --------------------------------------------------------------------------
tab_data, tab_ringkasan, tab_komentar, tab_resmi = st.tabs(
    ["Data Terfilter", "Ringkasan & Grafik", "Komentar & Keluhan", "Laporan Resmi"]
)

# --- TAB 1: DATA TABEL ---
with tab_data:
    section_title("bi-table", "Data Responden (sesuai filter)")
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)

    csv_bytes = df_filtered.to_csv(index=False).encode("utf-8-sig")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_filtered.to_excel(writer, index=False, sheet_name="Data Terfilter")
    excel_bytes = buf.getvalue()

    dcol1, dcol2 = st.columns(2)
    dcol1.download_button(
        "Download CSV",
        data=csv_bytes,
        file_name="evaluasi_pelatihan_terfilter.csv",
        mime="text/csv",
        use_container_width=True,
        icon=":material/download:",
    )
    dcol2.download_button(
        "Download Excel (data mentah)",
        data=excel_bytes,
        file_name="evaluasi_pelatihan_terfilter.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        icon=":material/download:",
    )

# --- TAB 2: RINGKASAN & GRAFIK ---
with tab_ringkasan:
    if not score_col_map:
        st.warning("Tidak ada kolom skor numerik yang terdeteksi pada sheet ini.")
    else:
        section_title("bi-people", "Jumlah Responden per Program")
        count_df = (
            df_filtered[program_col]
            .astype(str)
            .value_counts()
            .rename_axis("Program")
            .reset_index(name="Jumlah")
        )
        fig_count = px.bar(count_df, x="Jumlah", y="Program", orientation="h", text="Jumlah")
        fig_count.update_traces(marker_color=ACCENT)
        fig_count.update_layout(yaxis_title="", xaxis_title="Jumlah Responden", plot_bgcolor="white")
        st.plotly_chart(fig_count, use_container_width=True)

        section_title("bi-bar-chart-steps", "Rata-rata Skor per Aspek Pelatihan")
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
        fig_avg.update_traces(marker_color=PRIMARY)
        fig_avg.update_layout(yaxis_title="", xaxis_title="Rata-rata Skor (1-5)", plot_bgcolor="white")
        st.plotly_chart(fig_avg, use_container_width=True)

        if len(selected_programs) > 1:
            section_title("bi-bar-chart-line", "Perbandingan Rata-rata Skor Antar Program")
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
                color_discrete_sequence=[PRIMARY, ACCENT, "#8FA6BD", "#C9D4DE", "#4C7A9C"],
            )
            fig_comp.update_layout(
                yaxis_title="", xaxis_title="Rata-rata Skor (1-5)", legend_title="Program",
                plot_bgcolor="white",
            )
            st.plotly_chart(fig_comp, use_container_width=True)

        if info_col:
            section_title("bi-pie-chart", "Sumber Informasi Pelatihan")
            src_df = (
                df_filtered[info_col]
                .astype(str)
                .value_counts()
                .rename_axis("Sumber")
                .reset_index(name="Jumlah")
            )
            fig_src = px.pie(
                src_df, names="Sumber", values="Jumlah", hole=0.45,
                color_discrete_sequence=[PRIMARY, ACCENT, "#8FA6BD", "#C9D4DE", "#4C7A9C", "#A9B8C7"],
            )
            st.plotly_chart(fig_src, use_container_width=True)

# --- TAB 3: KOMENTAR & KELUHAN ---
with tab_komentar:
    section_title("bi-chat-left-text", "Komentar, Saran, dan Keluhan Peserta")

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
                with st.expander(f"{program} ({len(group)} komentar)"):
                    for _, row in group.iterrows():
                        st.markdown(f"- {row[picked_col]}")

# --- TAB 4: LAPORAN RESMI (format Kemnaker, sama persis dengan template) ---
with tab_resmi:
    section_title("bi-file-earmark-spreadsheet", "Ekspor ke Format Laporan Resmi")
    st.markdown(
        '<div class="info-note"><i class="bi bi-info-circle"></i>&nbsp; '
        "Laporan ini mengikuti format resmi (kop surat, tabel nilai per peserta, "
        "kategori BAIK/SANGAT BAIK) persis seperti file contoh yang diberikan. "
        "Pilih <b>satu program</b> yang mewakili satu angkatan/kelas pelatihan, "
        "supaya jumlah peserta pada laporan sesuai.</div>",
        unsafe_allow_html=True,
    )
    st.write("")

    report_program = st.selectbox(
        "Program pelatihan untuk laporan ini",
        options=program_options,
        index=0,
    )
    df_report = df_work[df_work[program_col].astype(str) == report_program]

    n_report = len(df_report)
    if n_report == 0:
        st.warning("Tidak ada responden pada program ini.")
    else:
        info_cols = st.columns(2)
        info_cols[0].metric("Jumlah Peserta", n_report)
        info_cols[1].metric("Pertanyaan Terisi", f"{len([l for l in SCORE_QUESTIONS if l in score_col_map])}/{len(SCORE_QUESTIONS)}")

        if n_report > DEFAULT_CAPACITY:
            st.info(
                f"Jumlah peserta ({n_report}) lebih dari {DEFAULT_CAPACITY}, jadi kolom cadangan "
                f"pada template akan otomatis dipakai (kapasitas total {MAX_RESPONDENTS} peserta)."
            )

        if n_report > MAX_RESPONDENTS:
            st.warning(
                f"Jumlah peserta ({n_report}) melebihi kapasitas maksimum template ({MAX_RESPONDENTS}). "
                f"Hanya data {MAX_RESPONDENTS} peserta pertama yang akan dimasukkan. "
                "Persempit dulu datanya (mis. per angkatan) untuk hasil yang akurat."
            )

        st.write("")
        st.markdown("**Informasi kop laporan** (wajib diisi)")
        c1, c2 = st.columns(2)
        training_title = c1.text_input("Nama pelatihan / kompetensi *", value="")
        instructor_name = c2.text_input("Nama instruktur *", value="")
        date_text = st.text_input(
            "Teks tanggal pelaksanaan * (mis. 'TGL Senin 3 Agustus S.D Jumat 28 Agustus 2026')",
            value="",
        )

        missing_fields = [
            label
            for label, value in [
                ("Nama pelatihan / kompetensi", training_title),
                ("Nama instruktur", instructor_name),
                ("Teks tanggal pelaksanaan", date_text),
            ]
            if not value.strip()
        ]

        if missing_fields:
            st.caption(
                "Lengkapi dulu: " + ", ".join(missing_fields) + " -- laporan belum bisa dibuat."
            )

        if st.button(
            "Buat Laporan Resmi (.xlsx)",
            type="primary",
            icon=":material/description:",
            disabled=bool(missing_fields),
        ):
            try:
                report_bytes, meta = generate_official_report(
                    str(TEMPLATE_PATH),
                    df_report,
                    score_col_map,
                    SCORE_QUESTIONS,
                    program_name=report_program,
                    training_title=training_title,
                    date_text=date_text,
                    instructor_name=instructor_name,
                )
                st.session_state["report_bytes"] = report_bytes
                st.session_state["report_meta"] = meta
                st.session_state["report_fname_base"] = report_program[:40].strip().replace(" ", "_")
                st.session_state.pop("report_pdf_bytes", None)  # laporan baru -> PDF lama tidak berlaku lagi
                st.success(
                    f"Laporan berhasil dibuat -- {meta['respondents_used']} peserta, "
                    f"{meta['questions_matched']}/{meta['questions_total']} pertanyaan terisi."
                )
            except Exception as e:
                st.error(f"Gagal membuat laporan: {e}")

        if st.session_state.get("report_bytes"):
            fname_base = st.session_state["report_fname_base"]
            col_xlsx, col_pdf = st.columns(2)
            with col_xlsx:
                st.download_button(
                    "Unduh Laporan Resmi (.xlsx)",
                    data=st.session_state["report_bytes"],
                    file_name=f"Laporan_Evaluasi_{fname_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                    key="dl_xlsx_report",
                )
            with col_pdf:
                if not st.session_state.get("report_pdf_bytes"):
                    if st.button(
                        "Konversi ke PDF",
                        icon=":material/picture_as_pdf:",
                        key="btn_convert_pdf",
                    ):
                        try:
                            with st.spinner("Mengonversi laporan ke PDF..."):
                                st.session_state["report_pdf_bytes"] = convert_xlsx_to_pdf(
                                    st.session_state["report_bytes"]
                                )
                            st.rerun()
                        except Exception as e:
                            st.error(f"Gagal mengonversi ke PDF: {e}")
                else:
                    st.download_button(
                        "Unduh Laporan Resmi (.pdf)",
                        data=st.session_state["report_pdf_bytes"],
                        file_name=f"Laporan_Evaluasi_{fname_base}.pdf",
                        mime="application/pdf",
                        icon=":material/download:",
                        key="dl_pdf_report",
                    )