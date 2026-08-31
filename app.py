import io
import re
import difflib
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px

from report_generator import (
    generate_official_report, convert_xlsx_to_pdf, generate_comment_recap,
    MAX_RESPONDENTS, DEFAULT_CAPACITY,
)

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
BORDER = "#E2E8EF"
SUCCESS = "#1E7A4C"

st.markdown(
    f"""
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css">
    <style>
        .block-container {{ padding-top: 2rem; max-width: 1200px; }}

        h1, h2, h3 {{ color: {PRIMARY}; font-weight: 700; }}

        .app-header {{
            display: flex; align-items: center; gap: 0.6rem;
            margin-bottom: 0.1rem;
        }}
        .app-header i {{ font-size: 1.6rem; color: {ACCENT}; }}
        .app-subtitle {{ color: {MUTED}; font-size: 0.95rem; margin-top: -0.4rem; margin-bottom: 1.5rem; }}

        .section-title {{
            display: flex; align-items: center; gap: 0.5rem;
            color: {PRIMARY}; font-weight: 600; font-size: 1.15rem;
            margin-top: 0.5rem;
        }}
        .section-title i {{ color: {ACCENT}; }}

        div[data-testid="stMetric"] {{
            background: {BG_SOFT};
            border: 1px solid {BORDER};
            border-radius: 10px;
            padding: 0.9rem 1rem;
        }}

        .stTabs [data-baseweb="tab"] {{ font-weight: 600; }}

        .info-note {{
            background: {BG_SOFT};
            border-left: 3px solid {ACCENT};
            padding: 0.65rem 0.9rem;
            border-radius: 4px;
            color: {MUTED};
            font-size: 0.9rem;
        }}

        /* --- Kartu upload (langkah 1) --- */
        .upload-card {{
            background: white;
            border: 1px solid {BORDER};
            border-radius: 14px;
            padding: 2rem 2.2rem;
            margin-top: 0.5rem;
        }}
        .upload-card h3 {{ margin-top: 0; }}
        .step-badge {{
            display: inline-flex; align-items: center; justify-content: center;
            width: 1.6rem; height: 1.6rem; border-radius: 50%;
            background: {ACCENT}; color: white; font-weight: 700; font-size: 0.85rem;
            margin-right: 0.5rem;
        }}
        .file-summary {{
            display: flex; align-items: center; gap: 0.6rem;
            background: {BG_SOFT}; border: 1px solid {BORDER}; border-radius: 8px;
            padding: 0.7rem 1rem; margin: 0.8rem 0;
        }}
        .file-summary i {{ color: {SUCCESS}; font-size: 1.2rem; }}

        .status-pill {{
            display: inline-flex; align-items: center; gap: 0.4rem;
            background: #EAF4EE; color: {SUCCESS}; border: 1px solid #CFE8D9;
            border-radius: 999px; padding: 0.25rem 0.8rem; font-size: 0.85rem; font-weight: 600;
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
# LOAD DATA (helper murni, dipanggil saat tombol submit ditekan)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    df.columns = [str(c) for c in df.columns]
    return df


def process_file(file_bytes: bytes, sheet_name: str):
    """Baca data, cocokkan kolom otomatis, siapkan semua yang dibutuhkan
    dashboard. Mengembalikan dict siap simpan ke session_state, atau
    melempar ValueError dengan pesan yang ramah kalau kolom program tidak
    ketemu sama sekali."""
    df_raw = load_data(file_bytes, sheet_name)
    all_columns = list(df_raw.columns)

    program_col = best_match(PROGRAM_CANONICAL, all_columns)
    if not program_col:
        raise ValueError(
            "Kolom 'Program pelatihan yang diikuti' tidak ditemukan di sheet ini. "
            "Pastikan file yang diupload adalah hasil Google Form evaluasi pelatihan, "
            "lalu coba lagi."
        )
    timestamp_col = best_match(TIMESTAMP_CANONICAL, all_columns)
    info_col = best_match(INFO_SOURCE_CANONICAL, all_columns)

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

    if timestamp_col:
        df_raw[timestamp_col] = pd.to_datetime(df_raw[timestamp_col], errors="coerce")

    return {
        "df_raw": df_raw,
        "sheet_choice": sheet_name,
        "program_col": program_col,
        "timestamp_col": timestamp_col,
        "info_col": info_col,
        "score_col_map": score_col_map,
        "comment_col_map": comment_col_map,
    }


# --------------------------------------------------------------------------
# STATE
# --------------------------------------------------------------------------
st.session_state.setdefault("data_ready", False)
st.session_state.setdefault("pending_file", None)


def reset_data():
    for key in ["data_ready", "df_raw", "sheet_choice", "program_col", "timestamp_col",
                "info_col", "score_col_map", "comment_col_map", "pending_file",
                "report_bytes", "report_meta", "report_fname_base", "report_pdf_bytes"]:
        st.session_state.pop(key, None)
    st.session_state["data_ready"] = False


# ==========================================================================
# LANGKAH 1: UPLOAD & PROSES (ditampilkan sampai data siap)
# ==========================================================================
if not st.session_state["data_ready"]:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown(
        '<h3><span class="step-badge">1</span>Upload Data Evaluasi</h3>',
        unsafe_allow_html=True,
    )
    st.caption("Upload file Excel (.xlsx) hasil Google Form evaluasi pelatihan, lalu klik Proses Data.")

    uploaded_file = st.file_uploader(
        "File Excel (.xlsx)",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    sheet_choice = None
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        st.markdown(
            f'<div class="file-summary"><i class="bi bi-file-earmark-check-fill"></i>'
            f"<div><b>{uploaded_file.name}</b><br>"
            f'<span style="color:{MUTED}; font-size:0.85rem;">{len(file_bytes)/1024:.0f} KB</span></div></div>',
            unsafe_allow_html=True,
        )

        try:
            sheet_names = get_sheet_names(file_bytes)
        except Exception:
            st.error("File tidak bisa dibaca. Pastikan formatnya .xlsx dan tidak rusak.")
            sheet_names = []

        if sheet_names:
            if len(sheet_names) > 1:
                sheet_choice = st.selectbox("Pilih sheet yang berisi data responden", options=sheet_names, index=0)
            else:
                sheet_choice = sheet_names[0]

            st.write("")
            submit = st.button(
                "Proses Data",
                type="primary",
                icon=":material/play_arrow:",
                use_container_width=True,
            )
            if submit:
                with st.spinner("Memproses data..."):
                    try:
                        result = process_file(file_bytes, sheet_choice)
                        for k, v in result.items():
                            st.session_state[k] = v
                        st.session_state["data_ready"] = True
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
    else:
        st.markdown(
            '<div class="info-note"><i class="bi bi-info-circle"></i>&nbsp; '
            "Belum ada file yang dipilih. Klik area di atas untuk memilih file dari komputer Anda.</div>",
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


# ==========================================================================
# LANGKAH 2: DASHBOARD (tampil setelah data diproses)
# ==========================================================================
df_raw = st.session_state["df_raw"]
sheet_choice = st.session_state["sheet_choice"]
program_col = st.session_state["program_col"]
timestamp_col = st.session_state["timestamp_col"]
info_col = st.session_state["info_col"]
score_col_map = st.session_state["score_col_map"]
comment_col_map = st.session_state["comment_col_map"]

with st.sidebar:
    st.markdown(
        '<span class="status-pill"><i class="bi bi-check-circle-fill"></i>Data siap</span>',
        unsafe_allow_html=True,
    )
    st.caption(f"Sheet: {sheet_choice} • {len(df_raw)} responden")
    if st.button("Upload Data Baru", icon=":material/upload_file:", use_container_width=True):
        reset_data()
        st.rerun()

    st.markdown(
        '<div class="section-title"><i class="bi bi-sliders"></i>Filter</div>',
        unsafe_allow_html=True,
    )

program_options = sorted(df_raw[program_col].dropna().astype(str).unique().tolist())
selected_programs = st.sidebar.multiselect(
    "Program pelatihan yang diikuti",
    options=program_options,
    default=program_options,
    help="Kosongkan lalu pilih ulang untuk memfilter satu/lebih program tertentu.",
)

df_work = df_raw

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
tab_data, tab_ringkasan, tab_komentar, tab_rekap_komentar, tab_resmi = st.tabs(
    ["Data Terfilter", "Ringkasan & Grafik", "Komentar & Keluhan", "Rekap Komentar", "Laporan Resmi"]
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

# --- TAB 4: REKAP KOMENTAR (1 baris per responden, format resmi) ---
with tab_rekap_komentar:
    section_title("bi-file-earmark-spreadsheet", "Rekap Komentar per Responden")
    st.markdown(
        '<div class="info-note"><i class="bi bi-info-circle"></i>&nbsp; '
        "Rekap ini menghasilkan satu baris per responden (jumlah baris selalu "
        "sama dengan jumlah peserta) dengan format kolom: <b>NO. | Sumber Informasi | "
        "Komentar Pelatihan | Komentar Instruktur | Komentar Penyelenggaraan | "
        "Keluhan Penyelenggaraan</b> -- responden dengan komentar kosong tetap "
        "disertakan, tidak dilewati.</div>",
        unsafe_allow_html=True,
    )
    st.write("")

    recap_program = st.selectbox(
        "Program pelatihan untuk rekap ini",
        options=program_options,
        index=0,
        key="recap_program_select",
    )
    df_recap = df_work[df_work[program_col].astype(str) == recap_program]
    n_recap = len(df_recap)

    if n_recap == 0:
        st.warning("Tidak ada responden pada program ini.")
    elif not comment_col_map:
        st.warning("Tidak ada kolom komentar yang terdeteksi pada sheet ini.")
    else:
        st.metric("Jumlah Responden (= jumlah baris rekap)", n_recap)

        recap_bytes = generate_comment_recap(df_recap, info_col, comment_col_map)

        # pratinjau tabel dengan header format resmi
        from report_generator import COMMENT_RECAP_HEADERS
        preview_rows = []
        for i, (_, r) in enumerate(df_recap.iterrows(), start=1):
            preview_rows.append({
                COMMENT_RECAP_HEADERS[0]: i,
                COMMENT_RECAP_HEADERS[1]: r.get(info_col, "") if info_col else "",
                COMMENT_RECAP_HEADERS[2]: r.get(comment_col_map.get("Komentar/saran program pelatihan", ""), ""),
                COMMENT_RECAP_HEADERS[3]: r.get(comment_col_map.get("Komentar/saran instruktur", ""), ""),
                COMMENT_RECAP_HEADERS[4]: r.get(comment_col_map.get("Komentar/saran penyelenggaraan", ""), ""),
                COMMENT_RECAP_HEADERS[5]: r.get(comment_col_map.get("Keluhan penyelenggaraan", ""), ""),
            })
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

        st.download_button(
            "Unduh Rekap Komentar (.xlsx)",
            data=recap_bytes,
            file_name=f"Komentar_Peserta_{recap_program[:40].strip().replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
            use_container_width=True,
        )

# --- TAB 5: LAPORAN RESMI (format Kemnaker, sama persis dengan template) ---
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