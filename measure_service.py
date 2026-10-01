"""
measure_service.py - tiny web service that measures exam PDFs (indents, columns, tables,
borders, shapes, font sizes) and returns a plain-text report the AI grader can use.

    POST /measure   (multipart/form-data)
        pdf      the PDF file (one file per request)
        name     optional file name to show in the report
      header  X-API-Key: <your secret>
    -> {"name": ..., "pages": [{"page": 1, "scanned": false}, ...], "report": "..."}

    GET /health  -> {"ok": true}

This is a separate service from the stamper (different folder, container and port),
so it can be started, stopped or updated without touching stamping.

Set the environment variable MEASURE_API_KEY to a long random secret on the host.
"""
import hmac
import os

from flask import Flask, jsonify, request

from measure_pdf import measure_pdf

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024  # 40 MB per request


def _authorised():
    expected = os.environ.get("MEASURE_API_KEY", "")
    if not expected:
        return False  # refuse to run open: the key must be configured on the server
    given = request.headers.get("X-API-Key", "")
    return hmac.compare_digest(given.encode(), expected.encode())


@app.get("/health")
def health():
    return jsonify(ok=True, service="ict-measure")


@app.post("/measure")
def measure():
    if not _authorised():
        return jsonify(error="unauthorised"), 401
    f = request.files.get("pdf") or (next(iter(request.files.values())) if request.files else None)
    if f is None:
        return jsonify(error="missing 'pdf' file field"), 400
    name = request.form.get("name") or f.filename or "document.pdf"
    data = f.read()
    if not data[:1024].lstrip().startswith(b"%PDF"):
        return jsonify(error="the uploaded file is not a PDF"), 422
    try:
        result = measure_pdf(data, name)
    except Exception as e:  # damaged / encrypted PDF etc.
        return jsonify(error=f"could not measure: {e}", name=name,
                       report=f"Measurements for '{name}' are not available ({e})."), 200
    return jsonify(result)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8001)))
