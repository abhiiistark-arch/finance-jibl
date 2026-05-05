"""
Flask application: upload PDF → process in memory → download Excel (.xlsx).
"""

from __future__ import annotations

import io
import logging
from pathlib import Path

import pandas as pd
import pymupdf
from flask import Flask, Request, jsonify, make_response, render_template_string, request, send_file
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from app.extractor.mapper import rows_to_output_dataframe
from app.extractor.pdf_parser import extract_upload_bytes

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"


def download_excel_filename(original_name: str) -> str:
    """Human-friendly Excel name from uploaded PDF filename."""
    safe = secure_filename(original_name or "upload.pdf")
    stem = Path(safe).stem or "commission"
    return f"{stem}_commission.xlsx"


def merged_excel_download_name(parts: list[str]) -> str:
    """Stable name when several PDFs are merged into one workbook."""
    if len(parts) == 1:
        return download_excel_filename(parts[0])
    stem0 = Path(secure_filename(parts[0]) or "upload.pdf").stem or "commission"
    return f"{stem0}_plus_{len(parts) - 1}_more_commission.xlsx"


def collect_uploaded_pdf_files(req: Request) -> list[FileStorage]:
    """
    Normalize single- or multi-file uploads (field name pdf or file).
    """
    uploads = [f for f in req.files.getlist("pdf") if getattr(f, "filename", "")]
    if not uploads:
        uploads = [f for f in req.files.getlist("file") if getattr(f, "filename", "")]
    if uploads:
        return uploads
    one = req.files.get("pdf") or req.files.get("file")
    return [one] if one and one.filename else []


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    )


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Commission PDF → Excel</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0; min-height: 100vh; font-family: "Segoe UI", system-ui, sans-serif;
      background: linear-gradient(160deg, #0f766e 0%, #134e4a 45%, #0c4a6e 100%);
      color: #0f172a; display: flex; align-items: center; justify-content: center; padding: 24px;
    }
    .card {
      background: #fff; border-radius: 16px; box-shadow: 0 24px 50px rgba(0,0,0,.28);
      max-width: 520px; width: 100%; padding: 28px 32px;
    }
    h1 { margin: 0 0 8px; font-size: 1.45rem; color: #0f172a; }
    .sub { margin: 0 0 20px; font-size: 0.93rem; color: #475569; line-height: 1.5; }
    .highlight { background: #ecfeff; border: 1px solid #99f6e4; border-radius: 10px;
      padding: 12px 14px; font-size: 0.88rem; color: #0e7490; margin-bottom: 20px; line-height: 1.45; }
    label.file {
      display: block; border: 2px dashed #94a3b8; border-radius: 12px; padding: 28px;
      text-align: center; cursor: pointer; color: #64748b; transition: border-color .2s, background .2s;
    }
    label.file:hover { border-color: #0d9488; background: #f0fdfa; }
    label.file strong { display: block; color: #0f172a; margin-top: 8px; }
    input[type=file] { display: none; }
    .chk { margin: 18px 0 4px; font-size: 0.9rem; color: #475569; display: flex; gap: 8px; align-items: flex-start; }
    .chk input { margin-top: 3px; }
    button.primary {
      width: 100%; margin-top: 16px; padding: 14px 18px; font-size: 1rem; font-weight: 600;
      color: #fff; background: linear-gradient(135deg, #0d9488, #0891b2); border: none;
      border-radius: 10px; cursor: pointer; box-shadow: 0 10px 24px rgba(13,148,136,.35);
    }
    button.primary:hover { filter: brightness(1.05); }
    .foot { margin-top: 18px; font-size: 0.8rem; color: #94a3b8; text-align: center; }
    code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 0.85em; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Commission bill → Excel</h1>
    <p class="sub">Upload one or more PDFs—nothing is saved under <code>input/</code> for this.
      Each PDF is parsed in memory; all matching commission rows are combined into one sheet with sequential
      <strong>Sr No</strong>. Your browser downloads a single <strong>.xlsx</strong> file.</p>
    <div class="highlight">
      <strong>Default:</strong> immediate Excel download (one workbook for everything you selected).<br>
      Optional: preview the first merged rows in the browser instead of downloading.
    </div>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <label class="file">
        <span>Select PDF(s)—hold Ctrl/⌘ for multiple…</span>
        <strong id="fn">No files chosen</strong>
        <input type="file" name="pdf" id="pdf" accept="application/pdf,.pdf" multiple required>
      </label>
      <div class="chk">
        <input type="checkbox" name="preview" value="1" id="preview">
        <label for="preview">Preview in browser instead of downloading (first 40 rows).</label>
      </div>
      <button class="primary" type="submit">Convert & download Excel</button>
    </form>
    <p class="foot">CLI batch (<code>input/</code> → <code>output/</code>): <code>python process_all_input_pdfs.py</code></p>
  </div>
  <script>
    document.getElementById('pdf').addEventListener('change', function(e) {
      var files = e.target.files;
      var el = document.getElementById('fn');
      if (!files || !files.length) { el.textContent = 'No files chosen'; return; }
      if (files.length === 1) { el.textContent = files[0].name; return; }
      el.textContent = files.length + ' PDFs: ' + Array.from(files).map(function(f){ return f.name; }).join(', ');
    });
  </script>
</body>
</html>
"""

    PREVIEW_TMPL = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Preview — Commission extract</title>
  <style>
    body { font-family: system-ui, sans-serif; margin: 0; padding: 24px; background: #f8fafc;
      max-width: 1100px; margin-left: auto; margin-right: auto; color: #0f172a; }
    table.preview { border-collapse: collapse; font-size: 13px; width: 100%; background: #fff; border-radius: 8px;
      overflow: hidden; box-shadow: 0 8px 24px rgba(0,0,0,.08); }
    table.preview th, table.preview td { border: 1px solid #e2e8f0; padding: 8px 10px; }
    table.preview th { background: #0d9488; color: #fff; font-weight: 600; }
    .muted { color: #64748b; font-size: 0.92rem; }
    .banner { padding: 12px 16px; border-radius: 10px; margin-bottom: 16px;
      {% if rows == 0 %}background:#fff7ed;border:1px solid #fdba74;{% else %}background:#ecfdf5;border:1px solid #6ee7b7;{% endif %} }
    a.download { display: inline-block; margin-top: 12px; padding: 10px 18px; background: #0d9488; color: #fff;
      text-decoration: none; border-radius: 8px; font-weight: 600; }
  </style>
</head>
<body>
  <div class="banner">
    <strong>{{ pdf_count }}</strong> PDF(s): {{ filenames_note }}<br>
    <strong>Combined data rows extracted:</strong> {{ rows }}
    {% if rows == 0 %}
    <p class="muted" style="margin:8px 0 0">No table matched the expected commission columns in any of these PDFs. The workbook
       would contain headers only.</p>
    {% endif %}
  </div>
  <p><a class="download" href="#" onclick="window.history.back();return false;">← Back to upload</a></p>
  <p class="muted">Showing up to <strong>40</strong> rows. For full data, submit again without preview.</p>
  <div style="overflow-x:auto">{{ table_html|safe }}</div>
</body>
</html>"""

    @app.route("/", methods=["GET"])
    def index():
        return render_template_string(INDEX_HTML)

    @app.route("/upload", methods=["POST"])
    def upload():
        file_list = collect_uploaded_pdf_files(request)
        if not file_list:
            logger.warning("Upload missing PDF(s).")
            return jsonify({"error": "No PDF file uploaded."}), 400

        preview = request.form.get("preview") == "1" or request.args.get("preview") == "1"
        combined_rows: list[dict[str, object]] = []
        source_names: list[str] = []

        for fh in file_list:
            fname = fh.filename or "upload.pdf"
            raw = fh.read()
            if not raw:
                logger.warning("Empty upload: %s", fname)
                return jsonify({"error": "Empty file.", "file": fname}), 400
            if not raw.startswith(b"%PDF"):
                logger.error("Invalid PDF signature: %s", fname)
                return jsonify({"error": "Invalid PDF upload (missing %PDF header).", "file": fname}), 400

            try:
                preview_doc = pymupdf.open(stream=raw, filetype="pdf")
                preview_doc.close()
            except Exception as e:
                logger.error("Unreadable PDF %s: %s", fname, e)
                return jsonify({"error": "Unreadable or corrupt PDF.", "file": fname, "detail": str(e)}), 400

            parsed = extract_upload_bytes(raw, filename=fname)
            logger.info(
                "Upload PDF=%s valid_table_found=%s row_count=%d",
                fname,
                bool(parsed.matched_headers),
                len(parsed.rows),
            )
            combined_rows.extend(parsed.rows)
            source_names.append(fname)

        df = rows_to_output_dataframe(combined_rows)
        excel_buf = io.BytesIO()

        try:
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Commission")
            excel_buf.seek(0)
        except Exception as e:
            logger.exception(
                "Excel build failed (%d PDF(s)): %s",
                len(source_names),
                e,
            )
            return jsonify({"error": "Failed to generate Excel.", "detail": str(e)}), 500

        out_name = merged_excel_download_name(source_names)
        filenames_note = "; ".join(source_names[:14]) + ("; …" if len(source_names) > 14 else "")
        pdf_count = len(source_names)

        if preview:
            table_html = df.head(40).to_html(classes="preview", border=0, na_rep="")
            page = render_template_string(
                PREVIEW_TMPL,
                filenames_note=filenames_note,
                pdf_count=pdf_count,
                rows=len(df),
                table_html=table_html,
            )
            resp = make_response(page)
            resp.headers["X-Excel-Rows"] = str(len(df))
            resp.headers["X-Download-Filename"] = out_name
            return resp

        resp = send_file(
            excel_buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=out_name,
        )
        resp.headers["X-Processed-Rows"] = str(len(df))
        resp.headers["X-Source-Pdf-Count"] = str(pdf_count)
        return resp

    @app.route("/health", methods=["GET"])
    def health():
        return jsonify({"status": "ok"})

    return app


def main() -> None:
    configure_logging()
    app.run(host="0.0.0.0", port=5000, debug=False)


# WSGI / platform entry: gunicorn "app.app:app" or `app` script in pyproject.toml
app = create_app()


if __name__ == "__main__":
    main()
