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


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok", "version": "2.0.0"})


@app.route("/api/preview-upload", methods=["POST"])
def preview_upload():
    """Ekstrak teks mentah dari file yang diupload untuk ditampilkan ke user."""
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400

    file = request.files["file"]
    if not file.filename or not allowed_file(file.filename):
        return jsonify({"error": "Format tidak didukung"}), 400

    file.seek(0, os.SEEK_END)
    if file.tell() > MAX_FILE_SIZE:
        return jsonify({"error": "File terlalu besar"}), 400
    file.seek(0)

    filename = secure_filename(file.filename)
    uid      = str(uuid.uuid4())[:8]
    save_path = UPLOAD_FOLDER / f"{uid}_{filename}"

    try:
        file.save(str(save_path))
        ext = filename.rsplit(".", 1)[1].lower()

        if ext == "pdf":
            from .file_parser import parse_uploaded_file
            text = parse_uploaded_file(str(save_path))
        elif ext in ("doc", "docx"):
            import docx as _docx
            doc  = _docx.Document(str(save_path))
            lines = []
            for para in doc.paragraphs:
                if para.text.strip():
                    lines.append(para.text.strip())
            for tbl in doc.tables:
                for row in tbl.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        lines.append(" | ".join(cells))
            text = "\n".join(lines)
        else:
            return jsonify({"error": "Format tidak didukung"}), 400

        # Batasi ke 8000 karakter untuk preview
        preview = text[:8000]
        truncated = len(text) > 8000
        return jsonify({
            "preview": preview,
            "truncated": truncated,
            "total_chars": len(text),
            "filename": filename,
        })
    except Exception as e:
        logger.error(f"Preview upload gagal: {e}", exc_info=True)
        return jsonify({"error": f"Gagal membaca file: {str(e)[:200]}"}), 500
    finally:
        if save_path.exists():
            save_path.unlink()


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
                    # Skip baris kosong
                    if any(v.strip() for v in vals):
                        rows.append(dict(zip(headers, vals)))
            return {"headers": headers, "rows": rows}

        survey_data  = _sheet_to_rows(wb["survey"])  if "survey"  in wb.sheetnames else {"headers":[],"rows":[]}
        choices_data = _sheet_to_rows(wb["choices"]) if "choices" in wb.sheetnames else {"headers":[],"rows":[]}
        wb.close()

        return jsonify({
            "survey":  survey_data,
            "choices": choices_data,
        })
    except Exception as e:
        logger.error(f"Preview xlsform gagal: {e}", exc_info=True)
        return jsonify({"error": f"Gagal membaca hasil konversi: {str(e)[:200]}"}), 500


@app.route("/api/download/<uid>")
def download_file(uid: str):
    """Endpoint untuk mengunduh file hasil konversi berdasarkan UID."""
    # Cari file dengan prefix uid di UPLOAD_FOLDER
    matches = list(UPLOAD_FOLDER.glob(f"{uid}_*"))
    if not matches:
        return jsonify({"error": "File tidak ditemukan atau sudah kadaluarsa"}), 404
    out_path = matches[0]
    output_name = out_path.name[len(uid) + 1:]  # strip "uid_" prefix
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
        t0 = time.time()

        ext = filename.rsplit(".", 1)[1].lower()

        # ── Konversi .doc legacy ke .docx ─────────────────────────────────────
        if ext == "doc":
            docx_path = save_path.with_suffix(".docx")
            converted = False

            # Kandidat path LibreOffice — Linux + Windows C: dan D:
            # Bisa di-override via .env: SOFFICE_PATH=D:\MyApps\LibreOffice\program\soffice.exe
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
                        save_path = docx_path
                        ext = "docx"
                        converted = True
                        break
                except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
                    continue

            if not converted:
                # Coba antiword (Windows: perlu install manual)
                try:
                    result = subprocess.run(
                        ["antiword", str(save_path)],
                        capture_output=True, text=True, timeout=30
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        logger.warning(".doc → antiword text — fallback ke LLM pipeline")
                        # Simpan sebagai txt sementara lalu jalankan LLM pipeline
                        txt_path = save_path.with_suffix(".txt")
                        txt_path.write_text(result.stdout, encoding="utf-8")
                        xlsx_bytes = _llm_pipeline(str(txt_path))
                        txt_path.unlink(missing_ok=True)
                        elapsed = time.time() - t0
                        logger.info(f"Selesai dalam {elapsed:.1f}s | output {len(xlsx_bytes):,} bytes")
                        output_name = filename.rsplit(".", 1)[0] + "_xlsform.xlsx"
                        out_uid  = str(uuid.uuid4())[:8]
                        out_path = UPLOAD_FOLDER / f"{out_uid}_{output_name}"
                        out_path.write_bytes(xlsx_bytes)
                        return jsonify({
                            "download_uid": out_uid,
                            "download_name": output_name,
                            "notes": {"fallback_questions": [], "placeholder_choices": []},
                        })
                except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
                    pass

            if not converted:
                return jsonify({
                    "error": (
                        "File .doc (format Word lama) tidak bisa dikonversi otomatis. "
                        "Silakan buka file di Microsoft Word atau LibreOffice, lalu simpan ulang sebagai .docx, "
                        "kemudian upload file .docx tersebut."
                    )
                }), 400

        if ext in ("doc", "docx"):
            # ── Pipeline baru: DOCX → parse → deterministik + LLM minimal ──
            logger.info("Mode: Deka Research DOCX parser")
            try:
                from .docx_parser import parse_questionnaire
                from .json_to_xlsform import convert_json_to_xlsform

                questions = parse_questionnaire(str(save_path))
                parsed = {
                    "_meta": {"source": filename},
                    "questions": questions
                }
                logger.info(f"Parse selesai: {len(questions)} baris")

                xlsx_bytes, conversion_notes = convert_json_to_xlsform(parsed)

            except ValueError as e:
                # Bukan format Deka Research → fallback ke LLM pipeline lama
                logger.warning(f"Bukan format Deka: {e} — fallback ke LLM pipeline")
                xlsx_bytes = _llm_pipeline(str(save_path))
                conversion_notes = {"fallback_questions": [], "placeholder_choices": []}

        elif ext == "pdf":
            # PDF tetap pakai LLM pipeline
            logger.info("Mode: LLM pipeline (PDF)")
            xlsx_bytes = _llm_pipeline(str(save_path))
            conversion_notes = {"fallback_questions": [], "placeholder_choices": []}

        else:
            return jsonify({"error": "Format tidak didukung"}), 400

        elapsed = time.time() - t0
        logger.info(f"Selesai dalam {elapsed:.1f}s | output {len(xlsx_bytes):,} bytes")

        output_name = filename.rsplit(".", 1)[0] + "_xlsform.xlsx"

        # Simpan file sementara agar bisa di-fetch via /api/download/<uid>
        out_uid  = str(uuid.uuid4())[:8]
        out_path = UPLOAD_FOLDER / f"{out_uid}_{output_name}"
        out_path.write_bytes(xlsx_bytes)

        return jsonify({
            "download_uid": out_uid,
            "download_name": output_name,
            "notes": conversion_notes,
        })

    except ValueError as e:
        logger.error(f"Validation error: {e}")
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.error(f"Konversi gagal: {e}", exc_info=True)
        return jsonify({"error": f"Konversi gagal: {str(e)[:300]}"}), 500
    finally:
        if save_path.exists():
            save_path.unlink()


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
