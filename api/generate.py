from __future__ import annotations

import base64
import json
import mimetypes
import re
import sys
import tempfile
import uuid
from datetime import date
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import unquote, urlparse


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from bulletin_generator import generate_bulletin_artifacts  # noqa: E402


MAX_UPLOAD_BYTES = 4 * 1024 * 1024
MAX_RESPONSE_BYTES = 4 * 1024 * 1024


def build_response_payload(
    excel_bytes: bytes,
    filename: str,
    report_date: str | None = None,
    prepared_for: str = "SBP Management Forum",
) -> dict:
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("Please upload an .xlsx file created from the template.")
    if len(excel_bytes) > MAX_UPLOAD_BYTES:
        raise ValueError("This Vercel preview accepts Excel files up to 4 MB.")

    run_id = uuid.uuid4().hex[:12]
    safe_name = Path(filename).name

    with tempfile.TemporaryDirectory() as temp_root:
        temp_dir = Path(temp_root)
        upload_path = temp_dir / f"{run_id}_{safe_name}"
        report_dir = temp_dir / "reports"
        preview_dir = temp_dir / "previews"
        upload_path.write_bytes(excel_bytes)

        report_path, preview_path, metadata = generate_bulletin_artifacts(
            upload_path,
            report_dir,
            preview_dir,
            report_date,
            prepared_for,
        )

        final_name = f"Daily_News_Bulletin_{date.today().strftime('%Y%m%d')}_{run_id}.docx"
        docx_bytes = report_path.read_bytes()
        preview_html = preview_path.read_text(encoding="utf-8")
        docx_base64 = base64.b64encode(docx_bytes).decode("ascii")

    estimated_response_bytes = len(docx_base64.encode("ascii")) + len(preview_html.encode("utf-8"))
    if estimated_response_bytes > MAX_RESPONSE_BYTES:
        raise ValueError(
            "The generated bulletin is too large for this Vercel preview. "
            "Use a smaller test file or move the preview build to Blob storage."
        )

    return {
        "ok": True,
        "filename": final_name,
        "previewHtml": preview_html,
        "docxBase64": docx_base64,
        "articleCount": metadata["article_count"],
        "storyCount": metadata["story_count"],
        "sectionCount": metadata["section_count"],
        "warnings": metadata["warnings"],
    }


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/generate":
            return self.send_json({"ok": False, "error": "Not found."}, status=404)

        try:
            fields, files = self.parse_multipart()
            uploaded = files.get("excel_file")
            if not uploaded:
                raise ValueError("Please choose a completed Excel file first.")

            filename, content = uploaded
            payload = build_response_payload(
                content,
                filename,
                fields.get("bulletin_date", "").strip() or None,
                fields.get("prepared_for", "").strip() or "SBP Management Forum",
            )
            self.send_json(payload)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, status=400)

    def do_GET(self) -> None:
        path = urlparse(self.path).path

        if path in {"/", "/index.html"}:
            return self.send_file(ROOT / "index.html", "text/html; charset=utf-8")
        if path.startswith("/static/"):
            safe_name = Path(unquote(path.removeprefix("/static/"))).name
            return self.send_file(ROOT / "static" / safe_name)
        if path == "/data/Daily_Bulletin_Empty_Template.xlsx":
            return self.send_file(
                ROOT / "data" / "Daily_Bulletin_Empty_Template.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                download_name="Daily_Bulletin_Empty_Template.xlsx",
            )
        if path == "/api/generate":
            return self.send_json({"ok": False, "error": "Use POST /api/generate."}, status=405)
        return self.send_json({"ok": False, "error": "Not found."}, status=404)

    def parse_multipart(self) -> tuple[dict[str, str], dict[str, tuple[str, bytes]]]:
        content_type = self.headers.get("Content-Type", "")
        match = re.search(r"boundary=(.+)", content_type)
        if not match:
            raise ValueError("The upload request was not formed correctly.")

        boundary = match.group(1).strip().strip('"').encode("utf-8")
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            raise ValueError("The uploaded file was empty.")
        if length > MAX_UPLOAD_BYTES:
            raise ValueError("This Vercel preview accepts upload requests up to 4 MB.")

        body = self.rfile.read(length)
        delimiter = b"--" + boundary
        fields: dict[str, str] = {}
        files: dict[str, tuple[str, bytes]] = {}

        for raw_part in body.split(delimiter):
            part = raw_part.strip()
            if not part or part == b"--":
                continue
            if part.endswith(b"--"):
                part = part[:-2].strip()
            if b"\r\n\r\n" not in part:
                continue

            header_blob, content = part.split(b"\r\n\r\n", 1)
            content = content.rstrip(b"\r\n")
            headers = header_blob.decode("utf-8", errors="replace")
            disposition = next(
                (line for line in headers.split("\r\n") if line.lower().startswith("content-disposition:")),
                "",
            )
            name_match = re.search(r'name="([^"]+)"', disposition)
            if not name_match:
                continue

            field_name = name_match.group(1)
            filename_match = re.search(r'filename="([^"]*)"', disposition)
            if filename_match:
                files[field_name] = (Path(filename_match.group(1)).name, content)
            else:
                fields[field_name] = content.decode("utf-8", errors="replace")
        return fields, files

    def send_json(self, payload: dict, status: int = 200) -> None:
        encoded = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def send_file(self, path: Path, content_type: str | None = None, download_name: str | None = None) -> None:
        resolved = path.resolve()
        allowed_roots = [(ROOT / "static").resolve(), (ROOT / "data").resolve()]
        allowed_files = {(ROOT / "index.html").resolve()}
        if resolved not in allowed_files and not any(root == resolved or root in resolved.parents for root in allowed_roots):
            return self.send_json({"ok": False, "error": "Not found."}, status=404)
        if not resolved.exists() or not resolved.is_file():
            return self.send_json({"ok": False, "error": "Not found."}, status=404)

        payload = resolved.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type or mimetypes.guess_type(resolved.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "public, max-age=300")
        if download_name:
            self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.end_headers()
        self.wfile.write(payload)
