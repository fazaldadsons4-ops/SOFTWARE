"""
PDF Cleaner & Column Mapper
Flask app — extracts tables from distributor PDFs, lets the user map
columns via plain-text commands, and exports a clean Excel file.
"""
import os
import re
import uuid
import json
import io
import logging
from pathlib import Path

import pdfplumber
import pandas as pd
from flask import (
    Flask, Blueprint, render_template, request,
    redirect, url_for, session, flash, send_file,
)
from werkzeug.utils import secure_filename

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
BASE_PATH  = os.environ.get("BASE_PATH", "/pdf-cleaner").rstrip("/")
PORT       = int(os.environ.get("PORT", 5001))

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf"}

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Flask app + Blueprint ─────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-pdf-cleaner-2025")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024   # 50 MB

bp = Blueprint("main", __name__)


@app.context_processor
def inject_base_path():
    return {"base_path": BASE_PATH}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── Column-command parser ─────────────────────────────────────────────────────

COLUMN_PATTERNS = [
    (r"product\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "product"),
    (r"(?:col(?:umn)?\s+(\d+)\s+(?:is\s+)?(?:the\s+)?product)", "product"),
    (r"item\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "product"),
    (r"rate\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "rate"),
    (r"price\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "rate"),
    (r"(?:col(?:umn)?\s+(\d+)\s+(?:is\s+)?(?:the\s+)?rate)", "rate"),
    (r"net\s*sales?\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "net_sales"),
    (r"(?:col(?:umn)?\s+(\d+)\s+(?:is\s+)?(?:the\s+)?net\s*sales?)", "net_sales"),
    (r"sales?\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "net_sales"),
    (r"amount\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "net_sales"),
    (r"net\s+amt\s+(?:col(?:umn)?)?\s*(?:is\s+)?(\d+)", "net_sales"),
]


def parse_column_commands(text: str) -> dict:
    """
    Parse free-text column mapping instructions.
    Returns a dict like {"product": 0, "rate": 1, "net_sales": 4} (0-indexed).
    """
    mapping: dict = {}
    text_lower = text.lower()
    for pattern, col_type in COLUMN_PATTERNS:
        m = re.search(pattern, text_lower)
        if m and col_type not in mapping:
            col_num = int(m.group(1))
            mapping[col_type] = col_num - 1   # convert 1-based → 0-based
    return mapping


# ── PDF extraction ────────────────────────────────────────────────────────────

def _is_empty_row(row: list) -> bool:
    return not any(str(c or "").strip() for c in row)


def _normalize_row(row: list) -> tuple:
    """Return a hashable, normalised version of a row for header dedup."""
    return tuple(re.sub(r"\s+", " ", str(c or "").strip().lower()) for c in row)


def _looks_like_header(row: list, first_header: tuple) -> bool:
    """True if this row is a repeat of the first detected header."""
    return _normalize_row(row) == first_header


def _extract_pdfplumber(pdf_path: str) -> tuple:
    """
    Extract ALL tables from ALL pages using pdfplumber.
    Returns (rows, meta_dict, warning_or_None).
    meta_dict has: total_pages, tables_per_page, engine
    """
    all_rows: list = []
    first_header: tuple | None = None
    meta = {"engine": "pdfplumber", "total_pages": 0, "tables_per_page": {}}
    warning = None

    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            meta["total_pages"] = total_pages
            logger.info("[pdfplumber] Total pages: %d", total_pages)

            pages_with_tables = 0

            for page_num, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                page_rows = 0

                if tables:
                    pages_with_tables += 1
                    # Use ALL tables on the page (not just the largest)
                    for table in tables:
                        for row in table:
                            if _is_empty_row(row):
                                continue
                            cleaned = [str(c or "").strip() for c in row]
                            norm = _normalize_row(cleaned)

                            # Identify first header row
                            if first_header is None:
                                first_header = norm
                                all_rows.append(cleaned)
                                page_rows += 1
                            elif norm == first_header:
                                # Duplicate header from a subsequent page — skip
                                logger.debug("[pdfplumber] Skipped repeat header on page %d", page_num)
                            else:
                                all_rows.append(cleaned)
                                page_rows += 1

                meta["tables_per_page"][page_num] = len(tables)
                if page_rows:
                    logger.info("[pdfplumber] Page %d: %d table(s), %d rows kept", page_num, len(tables), page_rows)

            logger.info("[pdfplumber] Pages with tables: %d / %d", pages_with_tables, total_pages)
            logger.info("[pdfplumber] Total rows extracted: %d", len(all_rows))

            if pages_with_tables == 0:
                warning = "pdfplumber: no structured tables found — will try fallback."

    except Exception as exc:
        logger.exception("[pdfplumber] extraction failed")
        warning = f"pdfplumber error: {exc}"

    return all_rows, meta, warning


def _extract_tabula(pdf_path: str) -> tuple:
    """
    Extract ALL tables from ALL pages using tabula-py (Java-based).
    Returns (rows, meta_dict, warning_or_None).
    """
    try:
        import tabula
    except ImportError:
        return [], {"engine": "tabula", "total_pages": 0, "tables_per_page": {}}, "tabula-py not installed"

    all_rows: list = []
    first_header: tuple | None = None
    meta = {"engine": "tabula", "total_pages": 0, "tables_per_page": {}}
    warning = None

    try:
        # Read all pages at once
        dfs = tabula.read_pdf(
            pdf_path,
            pages="all",
            multiple_tables=True,
            silent=True,
            lattice=False,
            stream=True,
        )
        logger.info("[tabula] Found %d table(s) across all pages", len(dfs))
        meta["total_pages"] = len(dfs)  # approximate

        for t_idx, df in enumerate(dfs):
            if df.empty:
                continue
            # Convert header to a row too
            header_row = [str(c).strip() for c in df.columns.tolist()]
            norm_h = _normalize_row(header_row)

            if first_header is None:
                first_header = norm_h
                all_rows.append(header_row)
            elif norm_h != first_header:
                # A differently-structured table — include its header
                all_rows.append(header_row)

            for _, data_row in df.iterrows():
                row = [str(v).strip() if str(v) not in ("nan", "None", "") else "" for v in data_row.tolist()]
                if _is_empty_row(row):
                    continue
                norm = _normalize_row(row)
                if norm == first_header:
                    continue   # repeated header as data row
                all_rows.append(row)

            logger.info("[tabula] Table %d: %d rows", t_idx + 1, len(df))

        logger.info("[tabula] Total rows extracted: %d", len(all_rows))

    except Exception as exc:
        logger.warning("[tabula] extraction failed: %s", exc)
        warning = f"tabula error: {exc}"

    return all_rows, meta, warning


def _extract_text_fallback(pdf_path: str) -> tuple:
    """Last-resort: extract text lines from all pages."""
    all_rows = []
    meta = {"engine": "text-fallback", "total_pages": 0, "tables_per_page": {}}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            meta["total_pages"] = len(pdf.pages)
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                page_rows = 0
                for line in text.splitlines():
                    parts = [p.strip() for p in line.split("  ") if p.strip()]
                    if not parts:
                        parts = line.split()
                    if parts:
                        all_rows.append(parts)
                        page_rows += 1
                logger.info("[text-fallback] Page %d: %d lines", page_num, page_rows)
    except Exception as exc:
        logger.exception("[text-fallback] failed")
    logger.info("[text-fallback] Total rows: %d", len(all_rows))
    return all_rows, meta, "Text-only extraction — no table structure detected. Column mapping may be approximate."


def extract_tables_from_pdf(pdf_path: str) -> tuple:
    """
    Multi-engine, multi-page PDF table extractor.
    1. pdfplumber — scans ALL pages, collects ALL tables, deduplicates headers.
    2. tabula-py   — automatic fallback if pdfplumber yields < 5 data rows.
    3. text lines  — last resort if both table engines fail.

    Returns (rows, extraction_meta, warning_or_None).
    extraction_meta = {engine, total_pages, tables_per_page, row_count}
    """
    # ── Engine 1: pdfplumber ──────────────────────────────────────────────────
    rows, meta, warning = _extract_pdfplumber(pdf_path)
    data_rows = rows[1:] if rows else []   # exclude header for count check

    if len(data_rows) >= 5:
        logger.info("Using pdfplumber results (%d data rows)", len(data_rows))
        meta["row_count"] = len(rows)
        return rows, meta, warning

    # ── Engine 2: tabula-py fallback ─────────────────────────────────────────
    logger.info("pdfplumber returned %d data rows — trying tabula fallback", len(data_rows))
    rows2, meta2, warning2 = _extract_tabula(pdf_path)
    data_rows2 = rows2[1:] if rows2 else []

    if len(data_rows2) > len(data_rows):
        logger.info("tabula returned more rows (%d) — using tabula results", len(data_rows2))
        meta2["row_count"] = len(rows2)
        note = "pdfplumber returned few rows; tabula-py was used as fallback."
        return rows2, meta2, note

    # Use whichever engine got more rows
    best_rows = rows if len(rows) >= len(rows2) else rows2
    best_meta = meta if len(rows) >= len(rows2) else meta2

    if best_rows:
        best_meta["row_count"] = len(best_rows)
        return best_rows, best_meta, warning

    # ── Engine 3: text-line fallback ─────────────────────────────────────────
    logger.info("Both table engines failed — using text-line fallback")
    rows3, meta3, warning3 = _extract_text_fallback(pdf_path)
    meta3["row_count"] = len(rows3)
    return rows3, meta3, warning3


def apply_column_mapping(rows, mapping, skip_header=True):
    """Apply mapping to raw rows; return (DataFrame, list_of_warnings)."""
    warnings = []
    if not mapping:
        warnings.append("No column mapping provided.")
        return pd.DataFrame(), warnings

    data_rows = rows[1:] if skip_header and len(rows) > 1 else rows
    records = []
    for row in data_rows:
        record = {}
        for field, col_idx in mapping.items():
            record[field] = row[col_idx] if col_idx < len(row) else ""
        records.append(record)

    rename_map = {"product": "Product", "rate": "Rate", "net_sales": "Net Sales"}
    df = pd.DataFrame(records)
    df.rename(columns=rename_map, inplace=True)

    keep_cols = [rename_map[k] for k in mapping if k in rename_map]
    df = df[[c for c in keep_cols if c in df.columns]]
    df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]

    return df, warnings


def generate_excel(df: pd.DataFrame, output_path: str) -> None:
    """Write a styled Excel file to output_path."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Clean Data")
        ws = writer.sheets["Clean Data"]

        header_fill = PatternFill("solid", fgColor="1E3A5F")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        thin = Side(style="thin", color="C9D8E8")
        center = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = Border(
                top=Side(style="medium", color="1E3A5F"),
                bottom=Side(style="medium", color="1E3A5F"),
                left=thin, right=thin,
            )

        alt_fill = PatternFill("solid", fgColor="EBF3FB")
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                if row_idx % 2 == 0:
                    cell.fill = alt_fill

        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            max_w = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[letter].width = min(max_w + 4, 50)

        ws.freeze_panes = "A2"

    with open(output_path, "wb") as f:
        f.write(buf.getvalue())


# ── Routes ────────────────────────────────────────────────────────────────────

@bp.route("/")
def index():
    return render_template("index.html")


@bp.route("/upload", methods=["POST"])
def upload():
    if "pdf_file" not in request.files:
        flash("No file selected.", "danger")
        return redirect(url_for("main.index"))

    f = request.files["pdf_file"]
    if not f or not f.filename:
        flash("No file selected.", "danger")
        return redirect(url_for("main.index"))

    if not allowed_file(f.filename):
        flash("Only PDF files are supported.", "danger")
        return redirect(url_for("main.index"))

    filename = secure_filename(f.filename)
    upload_id = str(uuid.uuid4())[:8]
    save_name = f"{upload_id}_{filename}"
    save_path = str(UPLOAD_DIR / save_name)
    f.save(save_path)

    rows, meta, warning = extract_tables_from_pdf(save_path)

    if not rows:
        flash(warning or "Could not extract any data from this PDF.", "danger")
        return redirect(url_for("main.index"))

    data_file = str(UPLOAD_DIR / f"{upload_id}_data.json")
    with open(data_file, "w") as fh:
        json.dump({"rows": rows, "filename": filename, "meta": meta}, fh)

    session["upload_id"]    = upload_id
    session["data_file"]    = data_file
    session["pdf_filename"] = filename
    session["pdf_path"]     = save_path
    session["extract_meta"] = meta

    if warning:
        flash(warning, "warning")

    return redirect(url_for("main.preview"))


@bp.route("/preview")
def preview():
    if "data_file" not in session:
        flash("Please upload a PDF first.", "info")
        return redirect(url_for("main.index"))

    data_file = session["data_file"]
    if not os.path.exists(data_file):
        flash("Session expired. Please re-upload.", "warning")
        return redirect(url_for("main.index"))

    with open(data_file) as fh:
        payload = json.load(fh)

    rows     = payload["rows"]
    meta     = payload.get("meta", session.get("extract_meta", {}))
    filename = session.get("pdf_filename", "file.pdf")

    # Show up to 200 rows in preview; indicate if truncated
    PREVIEW_LIMIT = 200
    preview_rows  = rows[:PREVIEW_LIMIT]
    n_cols = max((len(r) for r in preview_rows), default=0)

    # Build per-page summary string for the debug panel
    tpp = meta.get("tables_per_page", {})
    page_summary = [
        f"Page {p}: {t} table(s)"
        for p, t in sorted(tpp.items(), key=lambda x: int(x[0]))
        if t > 0
    ]

    return render_template(
        "preview.html",
        filename=filename,
        rows=preview_rows,
        total_rows=len(rows),
        preview_limit=PREVIEW_LIMIT,
        n_cols=n_cols,
        col_range=list(range(1, n_cols + 1)),
        meta=meta,
        page_summary=page_summary,
    )


@bp.route("/extract", methods=["POST"])
def extract():
    if "data_file" not in session:
        flash("Session expired. Please re-upload.", "warning")
        return redirect(url_for("main.index"))

    data_file = session["data_file"]
    commands  = request.form.get("commands", "").strip()
    skip_hdr  = request.form.get("skip_header", "1") == "1"

    if not commands:
        flash("Please enter at least one column mapping command.", "warning")
        return redirect(url_for("main.preview"))

    with open(data_file) as fh:
        payload = json.load(fh)
    rows = payload["rows"]

    mapping = parse_column_commands(commands)
    if not mapping:
        flash(
            "Could not understand commands. "
            "Try: 'Product column is 1, Rate column is 4, Net Sales column is 7'",
            "warning",
        )
        return redirect(url_for("main.preview"))

    df, warnings = apply_column_mapping(rows, mapping, skip_header=skip_hdr)
    for w in warnings:
        flash(w, "warning")

    if df.empty:
        flash("No data extracted with the provided column mapping.", "danger")
        return redirect(url_for("main.preview"))

    upload_id    = session.get("upload_id", "out")
    out_filename = f"clean_{upload_id}.xlsx"
    out_path     = str(OUTPUT_DIR / out_filename)
    generate_excel(df, out_path)

    session["result_filename"] = out_filename
    session["result_preview"]  = df.head(20).to_html(
        classes="table table-sm table-hover table-bordered mb-0",
        border=0,
        index=False,
    )
    session["result_mapping"] = {k: v + 1 for k, v in mapping.items()}
    session["result_rows"]    = len(df)

    return redirect(url_for("main.result"))


@bp.route("/result")
def result():
    if "result_filename" not in session:
        return redirect(url_for("main.index"))
    return render_template(
        "result.html",
        filename=session.get("result_filename"),
        preview_html=session.get("result_preview", ""),
        mapping=session.get("result_mapping", {}),
        row_count=session.get("result_rows", 0),
        pdf_filename=session.get("pdf_filename", ""),
    )


@bp.route("/download/<filename>")
def download(filename: str):
    safe = secure_filename(filename)
    path = OUTPUT_DIR / safe
    if not path.exists():
        flash("File not found or already deleted.", "danger")
        return redirect(url_for("main.index"))
    return send_file(str(path), as_attachment=True, download_name=safe)


@bp.route("/reset")
def reset():
    for key in ("pdf_path", "data_file"):
        p = session.pop(key, None)
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass
    session.clear()
    return redirect(url_for("main.index"))


# ── Register blueprint & run ──────────────────────────────────────────────────

app.register_blueprint(bp, url_prefix=BASE_PATH)


@app.route("/healthz")
def healthz():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=False)
