import os
import io
import uuid
import logging
import time
import subprocess
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(
    __name__,
    template_folder="../frontend/templates",
    static_folder="../frontend/static",
)
CORS(app)

UPLOAD_FOLDER = Path(__file__).parent.parent / "tmp_uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "doc", "docx"}
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── Helper: konversi .doc → .docx via LibreOffice/antiword ──────────────────
def _convert_doc_to_docx(save_path: Path) -> tuple[Path, bool, str]:
    """
    Coba konversi file .doc ke .docx.
    Return: (path_hasil, berhasil, metode)
      - path_hasil: Path ke .docx jika berhasil, tetap save_path jika gagal
      - berhasil: True jika konversi sukses
      - metode: 'libreoffice' | 'antiword_text' | 'failed'
    """
    docx_path = save_path.with_suffix(".docx")

    # Kandidat path LibreOffice
    custom_path = os.environ.get("SOFFICE_PATH", "")
    soffice_candidates = list(filter(None, [
        custom_path,
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"C:\Program Files\LibreOffice 7\program\soffice.exe",
        r"C:\Program Files\LibreOffice 24\program\soffice.exe",
        r"C:\Program Files\LibreOffice 25\program\soffice.exe",
        r"D:\Program Files\LibreOffice\program\soffice.exe",
        r"D:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"D:\Program Files\LibreOffice 7\program\soffice.exe",
        r"D:\Program Files\LibreOffice 24\program\soffice.exe",
        r"D:\Program Files\LibreOffice 25\program\soffice.exe",
    ]))

    for soffice in soffice_candidates:
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "docx",
                 "--outdir", str(save_path.parent), str(save_path)],
                capture_output=True, timeout=90
            )
            if result.returncode == 0 and docx_path.exists():
                logger.info(f".doc dikonversi ke .docx via {soffice}: {docx_path.name}")
                return docx_path, True, "libreoffice"
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            continue

    # Fallback: antiword → teks mentah
    try:
        result = subprocess.run(
            ["antiword", str(save_path)],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and result.stdout.strip():
            logger.info(f".doc dibaca via antiword: {len(result.stdout)} chars")
            txt_path = save_path.with_suffix(".txt")
            txt_path.write_text(result.stdout, encoding="utf-8")
            return txt_path, True, "antiword_text"
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    return save_path, False, "failed"


def _extract_raw_text_from_docx(doc_path: str) -> str:
    """
    Ekstrak teks dari .docx semirip mungkin dengan tampilan aslinya.
    Paragraf kosong tetap dipertahankan sebagai baris kosong (whitespace asli dijaga).
    Tabel diekstrak dengan format grid sederhana agar tetap terbaca.
    """
    import docx as _docx
    doc = _docx.Document(doc_path)
    lines = []

    # Kumpulkan semua block (paragraf + tabel) dalam urutan dokumen
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    body = doc.element.body
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            # Paragraf — ambil teks apa adanya, termasuk yang kosong
            para_text = child.text_content() if hasattr(child, 'text_content') else ''
            # Gunakan python-docx Paragraph untuk ambil teks dengan benar
            from docx.text.paragraph import Paragraph
            para = Paragraph(child, doc)
            lines.append(para.text)  # bisa berupa string kosong — itu disengaja

        elif tag == 'tbl':
            # Tabel — render sebagai grid teks sederhana
            from docx.table import Table
            tbl = Table(child, doc)
            # Tentukan lebar kolom maksimal per kolom
            col_widths = []
            all_rows_data = []
            for row in tbl.rows:
                row_data = [cell.text.replace('\n', ' ') for cell in row.cells]
                all_rows_data.append(row_data)
                for ci, val in enumerate(row_data):
                    if ci >= len(col_widths):
                        col_widths.append(0)
                    col_widths[ci] = max(col_widths[ci], len(val))

            # Batas lebar kolom agar tidak terlalu lebar
            col_widths = [min(w, 40) for w in col_widths]

            sep = '+' + '+'.join('-' * (w + 2) for w in col_widths) + '+'
            lines.append(sep)
            for ri, row_data in enumerate(all_rows_data):
                row_str = '|'
                for ci, val in enumerate(row_data):
                    w = col_widths[ci] if ci < len(col_widths) else 10
                    cell_val = val[:w].ljust(w)
                    row_str += f' {cell_val} |'
                lines.append(row_str)
                if ri == 0:  # Separator setelah header
                    lines.append(sep)
            lines.append(sep)
            lines.append('')  # Baris kosong setelah tabel

    return '\n'.join(lines)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "2.0.0"})


@app.route("/api/preview-upload", methods=["POST"])
def preview_upload():
    """Ekstrak teks mentah dari file yang diupload untuk ditampilkan ke user — semirip mungkin dengan isi aslinya."""
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400

    file = request.files["file"]
    if not file.filename or not allowed_file(file.filename):
        return jsonify({"error": "Format tidak didukung"}), 400

    file.seek(0, os.SEEK_END)
    if file.tell() > MAX_FILE_SIZE:
        return jsonify({"error": "File terlalu besar"}), 400
    file.seek(0)

    filename  = secure_filename(file.filename)
    uid       = str(uuid.uuid4())[:8]
    save_path = UPLOAD_FOLDER / f"{uid}_{filename}"

    temp_files: list[Path] = [save_path]

    try:
        file.save(str(save_path))
        ext = filename.rsplit(".", 1)[1].lower()

        if ext == "pdf":
            # PDF: gunakan parser yang ada — hasilnya sudah cukup raw
            from .file_parser import parse_uploaded_file
            text = parse_uploaded_file(str(save_path))

        elif ext == "doc":
            converted_path, ok, method = _convert_doc_to_docx(save_path)

            if not ok:
                return jsonify({
                    "error": (
                        "File .doc (format Word lama) tidak dapat dibaca langsung. "
                        "Pastikan LibreOffice terinstall di server, atau simpan ulang "
                        "file sebagai .docx lalu upload kembali."
                    )
                }), 400

            if converted_path != save_path:
                temp_files.append(converted_path)

            if method == "antiword_text":
                # antiword sudah menghasilkan teks mentah yang cukup baik
                text = converted_path.read_text(encoding="utf-8")
            else:
                # LibreOffice → .docx → ekstrak raw
                text = _extract_raw_text_from_docx(str(converted_path))

        elif ext == "docx":
            # Ekstrak teks semirip mungkin dengan dokumen asli
            text = _extract_raw_text_from_docx(str(save_path))

        else:
            return jsonify({"error": "Format tidak didukung"}), 400

        preview   = text[:8000]
        truncated = len(text) > 8000
        return jsonify({
            "preview":     preview,
            "truncated":   truncated,
            "total_chars": len(text),
            "filename":    filename,
        })

    except Exception as e:
        logger.error(f"Preview upload gagal: {e}", exc_info=True)
        return jsonify({"error": f"Gagal membaca file: {str(e)[:200]}"}), 500
    finally:
        for f in temp_files:
            if f.exists():
                f.unlink(missing_ok=True)


@app.route("/api/preview-xlsform/<uid>")
def preview_xlsform(uid: str):
    """Baca file xlsx hasil konversi dan kembalikan isinya sebagai JSON untuk preview."""
    matches = list(UPLOAD_FOLDER.glob(f"{uid}_*"))
    if not matches:
        return jsonify({"error": "File tidak ditemukan atau sudah kadaluarsa"}), 404

    out_path = matches[0]
    try:
        import openpyxl
        wb = openpyxl.load_workbook(str(out_path), read_only=True)

        def _sheet_to_rows(ws):
            rows = []
            headers = []
            for ri, row in enumerate(ws.iter_rows(values_only=True)):
                vals = [str(v) if v is not None else "" for v in row]
                if ri == 0:
                    headers = vals
                else:
                    if any(v.strip() for v in vals):
                        rows.append(dict(zip(headers, vals)))
            return {"headers": headers, "rows": rows}

        survey_data  = _sheet_to_rows(wb["survey"])  if "survey"  in wb.sheetnames else {"headers": [], "rows": []}
        choices_data = _sheet_to_rows(wb["choices"]) if "choices" in wb.sheetnames else {"headers": [], "rows": []}
        wb.close()

        return jsonify({"survey": survey_data, "choices": choices_data})
    except Exception as e:
        logger.error(f"Preview xlsform gagal: {e}", exc_info=True)
        return jsonify({"error": f"Gagal membaca hasil konversi: {str(e)[:200]}"}), 500


@app.route("/api/download/<uid>")
def download_file(uid: str):
    """Endpoint untuk mengunduh file hasil konversi berdasarkan UID."""
    matches = list(UPLOAD_FOLDER.glob(f"{uid}_*"))
    if not matches:
        return jsonify({"error": "File tidak ditemukan atau sudah kadaluarsa"}), 404
    out_path    = matches[0]
    output_name = out_path.name[len(uid) + 1:]
    return send_file(
        str(out_path),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=output_name,
    )


@app.route("/api/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file di request"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Tidak ada file yang dipilih"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Format tidak didukung. Gunakan PDF, DOC, atau DOCX."}), 400

    file.seek(0, os.SEEK_END)
    file_size = file.tell()
    file.seek(0)
    if file_size > MAX_FILE_SIZE:
        return jsonify({"error": "File terlalu besar. Maksimal 10MB."}), 400

    filename  = secure_filename(file.filename)
    uid       = str(uuid.uuid4())[:8]
    save_path = UPLOAD_FOLDER / f"{uid}_{filename}"

    try:
        file.save(str(save_path))
        logger.info(f"File disimpan: {save_path} ({file_size:,} bytes)")
        t0  = time.time()
        ext = filename.rsplit(".", 1)[1].lower()

        # ── Konversi .doc legacy ke .docx ────────────────────────────────────
        if ext == "doc":
            converted_path, ok, method = _convert_doc_to_docx(save_path)

            if method == "antiword_text" and ok:
                logger.warning(".doc → antiword text — fallback ke LLM pipeline")
                xlsx_bytes = _llm_pipeline(str(converted_path))
                converted_path.unlink(missing_ok=True)
                elapsed     = time.time() - t0
                logger.info(f"Selesai dalam {elapsed:.1f}s | output {len(xlsx_bytes):,} bytes")
                output_name = filename.rsplit(".", 1)[0] + "_xlsform.xlsx"
                out_uid     = str(uuid.uuid4())[:8]
                out_path    = UPLOAD_FOLDER / f"{out_uid}_{output_name}"
                out_path.write_bytes(xlsx_bytes)
                return jsonify({
                    "download_uid":  out_uid,
                    "download_name": output_name,
                    "notes": {"fallback_questions": [], "placeholder_choices": []},
                })

            if not ok:
                return jsonify({
                    "error": (
                        "File .doc (format Word lama) tidak bisa dikonversi otomatis. "
                        "Silakan buka file di Microsoft Word atau LibreOffice, lalu simpan ulang sebagai .docx, "
                        "kemudian upload file .docx tersebut."
                    )
                }), 400

            save_path = converted_path
            ext       = "docx"

        if ext in ("doc", "docx"):
            logger.info("Mode: Deka Research DOCX parser")
            try:
                from .docx_parser import parse_questionnaire
                from .json_to_xlsform import convert_json_to_xlsform

                questions = parse_questionnaire(str(save_path))
                parsed    = {"_meta": {"source": filename}, "questions": questions}
                logger.info(f"Parse selesai: {len(questions)} baris")

                xlsx_bytes, conversion_notes = convert_json_to_xlsform(parsed)

            except ValueError as e:
                logger.warning(f"Bukan format Deka: {e} — fallback ke LLM pipeline")
                xlsx_bytes        = _llm_pipeline(str(save_path))
                conversion_notes  = {"fallback_questions": [], "placeholder_choices": []}

        elif ext == "pdf":
            logger.info("Mode: LLM pipeline (PDF)")
            xlsx_bytes       = _llm_pipeline(str(save_path))
            conversion_notes = {"fallback_questions": [], "placeholder_choices": []}

        else:
            return jsonify({"error": "Format tidak didukung"}), 400

        elapsed = time.time() - t0
        logger.info(f"Selesai dalam {elapsed:.1f}s | output {len(xlsx_bytes):,} bytes")

        output_name = filename.rsplit(".", 1)[0] + "_xlsform.xlsx"
        out_uid     = str(uuid.uuid4())[:8]
        out_path    = UPLOAD_FOLDER / f"{out_uid}_{output_name}"
        out_path.write_bytes(xlsx_bytes)

        return jsonify({
            "download_uid":  out_uid,
            "download_name": output_name,
            "notes":         conversion_notes,
        })

    except ValueError as e:
        logger.error(f"Validation error: {e}")
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Konversi gagal: {e}", exc_info=True)
        return jsonify({"error": f"Konversi gagal: {str(e)[:300]}"}), 500
    finally:
        if save_path.exists():
            save_path.unlink(missing_ok=True)


def _llm_pipeline(file_path: str) -> bytes:
    """Fallback pipeline lama via LLM untuk PDF atau format non-Deka."""
    from .file_parser import parse_uploaded_file
    from .llm_client import call_llm_for_xlsform
    from .xlsform_builder import build_xlsform_from_json

    text = parse_uploaded_file(file_path)
    if not text or len(text.strip()) < 50:
        raise ValueError("Tidak dapat membaca konten file.")

    xlsform_data = call_llm_for_xlsform(text)
    return build_xlsform_from_json(xlsform_data)


if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV", "production") == "development"
    app.run(host="0.0.0.0", port=port, debug=debug)
