import re
import difflib
from io import BytesIO

import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

PRIMARY_COLS = [get_column_letter(c) for c in range(3, 19)]   # C..R (16)
SECONDARY_COLS = [get_column_letter(c) for c in range(20, 28)]  # T..AA (8)

# Kapasitas "default" formulir adalah 16 responden (kolom C..R).
# Jika jumlah responden lebih dari itu, kolom cadangan T..AA (8 kolom) otomatis
# dipakai sehingga total kapasitas template menjadi 24 responden.
DEFAULT_CAPACITY = len(PRIMARY_COLS)                     # 16
MAX_RESPONDENTS = DEFAULT_CAPACITY + len(SECONDARY_COLS)  # 24

SCORE_COL = "S"
KET_COL = "AC"

MASTER_SHEET = "KOLOM ISIAN"

# Judul laporan tiap sheet ("HASIL ANALISIS EVALUASI") + posisi baris judul,
# supaya bisa dipasang garis pemisah (kop) dan teks judul yang konsisten
# dengan sheet master di SEMUA sheet, bukan cuma KOLOM ISIAN.
# value: (baris_judul, kolom_terakhir_untuk_garis)
SHEET_TITLE_ROWS = {
    "KOLOM ISIAN": (6, 29),      # A..AC
    "Histogram": (6, 4),         # A..D (mengikuti 'Materi Pelatihan' via formula)
    "Materi Pelatihan": (6, 29),
    "Program Pelatihan": (6, 29),
    "INSTRUKTUR": (7, 29),
    "Penyelenggaraan": (7, 29),
    "NILAI INST": (6, 6),        # A..F (mengikuti 'Materi Pelatihan' via formula)
}

# Sheet yang judulnya ("HASIL EVALUASI") ditulis literal (bukan formula ke
# master) sehingga perlu disamakan manual ke "HASIL ANALISIS EVALUASI".
TITLE_TEXT_CELLS = {
    "Materi Pelatihan": "A6",
    "Program Pelatihan": "A6",
    "INSTRUKTUR": "A7",
    "Penyelenggaraan": "A7",
}

REPORT_TITLE = "HASIL ANALISIS EVALUASI"

# Baris-baris header tabel isian (nomor urut responden) di sheet master.
# Kolom T..AA di baris-baris ini secara bawaan berisi nomor yang salah (22-29)
# padahal seharusnya melanjutkan nomor primer (17-24).
HEADER_ROWS = [13, 22, 30, 38]


def _normalize(text: str) -> str:
    text = str(text).lower().replace("\n", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _find_question_row(ws, question_text: str, cutoff: float = 0.55):
    """Cari baris di kolom B yang cocok dengan teks pertanyaan (fuzzy)."""
    target = _normalize(question_text)
    best_row, best_ratio = None, 0.0
    for row in ws.iter_rows(min_col=2, max_col=2):
        cell = row[0]
        if not cell.value:
            continue
        norm = _normalize(cell.value)
        if norm == target:
            return cell.row
        ratio = difflib.SequenceMatcher(None, norm, target).ratio()
        if norm in target or target in norm:
            ratio = max(ratio, 0.9)
        if ratio > best_ratio:
            best_ratio, best_row = ratio, cell.row
    return best_row if best_ratio >= cutoff else None


def kategori_nilai(score: float) -> str:
    if score >= 5:
        return "SANGAT BAIK"
    if score >= 4:
        return "BAIK"
    if score >= 3:
        return "CUKUP"
    if score >= 2:
        return "KURANG"
    return "TIDAK BAIK"


def _format_training_title(training_title: str) -> str:
    """'Industri dan Jasa' -> 'PELATIHAN KOMPETENSI BERBASIS INDUSTRI DAN JASA'."""
    name = re.sub(r"\s+", " ", str(training_title).strip()).upper()
    if not name:
        return ""
    # Hindari duplikasi kalau user sudah mengetik prefiksnya sendiri.
    name = re.sub(r"^PELATIHAN KOMPETENSI BERBASIS\s+", "", name)
    return f"PELATIHAN KOMPETENSI BERBASIS {name}"


def _format_program_title(program_name: str) -> str:
    """'Daily Make Up 3' -> 'PROGRAM PELATIHAN DASAR DAILY MAKE UP PAKET 3'.

    - Selalu CAPSLOCK.
    - Kalau nama program diakhiri angka batch (mis. '... 3') dan belum ada
      kata 'PAKET', kata 'PAKET' otomatis disisipkan sebelum angka tsb.
    """
    name = re.sub(r"\s+", " ", str(program_name).strip()).upper()
    if not name:
        return ""
    name = re.sub(r"^PROGRAM PELATIHAN DASAR\s+", "", name)
    if "PAKET" not in name:
        match = re.match(r"^(.*\S)\s+(\d+)$", name)
        if match:
            name = f"{match.group(1)} PAKET {match.group(2)}"
    return f"PROGRAM PELATIHAN DASAR {name}"


def _add_title_separator_line(ws, row: int = 6, min_col: int = 1, max_col: int = 29):
    """Tambahkan garis horizontal (double line) tepat di atas 'HASIL ANALISIS
    EVALUASI', memisahkan blok kop surat dari blok judul laporan."""
    thin_double = Side(style="double", color="000000")
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=row, column=col)
        existing = cell.border
        cell.border = Border(
            top=thin_double,
            bottom=existing.bottom,
            left=existing.left,
            right=existing.right,
        )


def _sync_titles_and_separator_lines(wb):
    """Pasang garis pemisah di atas judul laporan pada SEMUA sheet, dan
    samakan teks judul ("HASIL ANALISIS EVALUASI") supaya semua sheet
    konsisten dengan sheet master -- beberapa sheet turunan sebelumnya
    punya teks literal "HASIL EVALUASI" yang tidak sinkron."""
    for sheet_name, (row, max_col) in SHEET_TITLE_ROWS.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        _add_title_separator_line(ws, row=row, max_col=max_col)

    for sheet_name, cell_coord in TITLE_TEXT_CELLS.items():
        if sheet_name not in wb.sheetnames:
            continue
        wb[sheet_name][cell_coord] = REPORT_TITLE


def _fix_secondary_header_numbers(ws):
    """Perbaiki penomoran responden 17-24 pada kolom cadangan T..AA yang di
    template bawaan salah tertulis 22-29."""
    for row in HEADER_ROWS:
        if ws.cell(row=row, column=19).value != "JUMLAH":  # kolom S
            continue
        for i, col in enumerate(SECONDARY_COLS):
            ws[f"{col}{row}"] = DEFAULT_CAPACITY + i + 1  # 17, 18, ..., 24


def _update_divisor_everywhere(wb, n_respondents: int, use_secondary: bool):
    """Perbaiki pembagi '/12' di rumus SUM pada kolom S, di SEMUA sheet
    (karena tiap sheet punya rumus SUM sendiri, bukan rujukan ke sheet lain)."""
    sum_range = (
        f"C{{r}}:R{{r}},T{{r}}:AA{{r}}" if use_secondary else "C{r}:R{r}"
    )
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_col=19, max_col=19):  # kolom S
            cell = row[0]
            if isinstance(cell.value, str) and cell.value.startswith("=SUM(") and "/12" in cell.value:
                r = cell.row
                new_range = sum_range.format(r=r)
                cell.value = f"=SUM({new_range})/{n_respondents}"


def generate_official_report(
    template_path: str,
    df_filtered,
    score_col_map: dict,          # label_ringkas -> nama kolom asli di dataframe
    canonical_questions: dict,     # label_ringkas -> teks pertanyaan kanonik
    program_name: str,
    training_title: str = "",
    date_text: str = "",
    instructor_name: str = "",
) -> bytes:
    n = len(df_filtered)
    if n == 0:
        raise ValueError("Tidak ada responden pada data terfilter.")

    capped = min(n, MAX_RESPONDENTS)
    use_secondary = capped > DEFAULT_CAPACITY

    wb = openpyxl.load_workbook(template_path)
    ws_master = wb[MASTER_SHEET]

    # ---- Perbaiki penomoran kolom cadangan (17-24) yang salah di template ----
    _fix_secondary_header_numbers(ws_master)

    # ---- Garis pemisah di atas judul + samakan teks judul di semua sheet ----
    _sync_titles_and_separator_lines(wb)

    # ---- Header ----
    ws_master["A6"] = REPORT_TITLE
    if training_title:
        ws_master["A7"] = _format_training_title(training_title)
    if program_name:
        ws_master["A8"] = _format_program_title(program_name)
    if date_text:
        ws_master["A9"] = date_text
    if instructor_name:
        ws_master["C31"] = f"Nama Instruktur : {instructor_name}"
        if "INSTRUKTUR" in wb.sheetnames:
            wb["INSTRUKTUR"]["C12"] = f"INSTRUKTUR : {instructor_name}"

    # ---- Cocokkan tiap pertanyaan skor dengan baris di sheet master ----
    row_map = {}  # label -> row index di KOLOM ISIAN
    for label, question in canonical_questions.items():
        if label not in score_col_map:
            continue
        row = _find_question_row(ws_master, question)
        if row:
            row_map[label] = row

    # ---- Tulis nilai peserta ----
    all_cols = PRIMARY_COLS + SECONDARY_COLS
    values_by_label = {}
    for label, row in row_map.items():
        col_name = score_col_map[label]
        values = (
            df_filtered[col_name]
            .dropna()
            .astype(float)
            .round()
            .astype(int)
            .tolist()[:capped]
        )
        values_by_label[label] = values

        # bersihkan slot lama
        for col in all_cols:
            ws_master[f"{col}{row}"] = None

        for i, val in enumerate(values):
            col = all_cols[i]
            ws_master[f"{col}{row}"] = val

    # ---- Perbaiki pembagi (divisor) di semua sheet ----
    actual_n = max((len(v) for v in values_by_label.values()), default=capped)
    _update_divisor_everywhere(wb, actual_n, use_secondary)

    # ---- Paksa Excel menghitung ulang rumus saat file dibuka ----
    wb.calculation.fullCalcOnLoad = True

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue(), {
        "respondents_used": actual_n,
        "respondents_total_filtered": n,
        "capped": n > MAX_RESPONDENTS,
        "used_secondary_columns": use_secondary,
        "questions_matched": len(row_map),
        "questions_total": len(canonical_questions),
    }