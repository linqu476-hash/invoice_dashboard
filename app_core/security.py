"""Security helpers shared by the pages. No Streamlit in here, so it can be tested on its own.

Layer 1 (sign-in)  -> app_core/auth.py        hashed password, lockout that survives a refresh, session timeouts
Layer 2 (uploads)  -> this file               every upload is checked before openpyxl / xlrd / sqlite / json read it
Layer 3 (data)     -> this file + Dockerfile  audit log, private file permissions, no secrets in the image, CI scans
"""
import io
import json
import logging
import os
import stat
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath

# ---------------------------------------------------------------- limits
MAX_ZIP_MEMBERS = 2000                    # entries in one .xlsx / .zip
MAX_UNCOMPRESSED = 300 * 1024 * 1024      # all entries of one workbook together (zip bomb guard)
MAX_RATIO = 200                           # uncompressed / compressed size of one entry (zip bomb guard)
MAX_JSON_BYTES = 20 * 1024 * 1024
MAX_SQLITE_BYTES = 200 * 1024 * 1024

XLSX_MAGIC = b"PK\x03\x04"
XLS_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
SQLITE_MAGIC = b"SQLite format 3\x00"
FORMULA_STARTS = ("=", "+", "-", "@", "\t", "\r")


class UploadRejected(ValueError):
    """The uploaded file is not safe to open. The message is safe to show to the person."""


# ---------------------------------------------------------------- layer 2: uploads
def inspect_zip(data, *, label="file", expect_xlsx=False):
    """Refuse zip bombs, path traversal and encrypted members BEFORE anything decompresses the file."""
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise UploadRejected(f"{label}: this is not a valid Excel / zip file.") from None
    with z:
        infos = z.infolist()
        if len(infos) > MAX_ZIP_MEMBERS:
            raise UploadRejected(f"{label}: too many entries inside the file ({len(infos)}).")
        total = 0
        for i in infos:
            name = i.filename.replace("\\", "/")
            if name.startswith("/") or ".." in PurePosixPath(name).parts or (len(name) > 1 and name[1] == ":"):
                raise UploadRejected(f"{label}: the file contains an unsafe path ({name[:60]}).")
            if i.flag_bits & 0x1:
                raise UploadRejected(f"{label}: encrypted files are not supported.")
            total += i.file_size
            if total > MAX_UNCOMPRESSED:
                raise UploadRejected(f"{label}: the file expands to more than {MAX_UNCOMPRESSED // (1024 * 1024)} MB (possible zip bomb).")
            if i.compress_size and i.file_size / i.compress_size > MAX_RATIO and i.file_size > 1024 * 1024:
                raise UploadRejected(f"{label}: suspicious compression ratio (possible zip bomb).")
        if expect_xlsx and "[Content_Types].xml" not in z.namelist():
            raise UploadRejected(f"{label}: this is a zip file, not an Excel .xlsx workbook.")


def check_upload(name, data, *, kinds=("xlsx",), max_bytes=20 * 1024 * 1024, ext=None):
    """Check one uploaded file by its CONTENT, not only by its extension.
    kinds: any of "xlsx", "xls", "zip", "json", "sqlite". ext overrides the extension taken from name. Raises UploadRejected."""
    label = Path(str(name)).name[:80] or "file"
    if not data:
        raise UploadRejected(f"{label}: the file is empty.")
    if len(data) > max_bytes:
        raise UploadRejected(f"{label}: the file is larger than {max_bytes // (1024 * 1024)} MB.")
    ext = (ext or Path(str(name)).suffix).lower().lstrip(".")
    if ext in ("xlsx", "xlsm", "xltx") and "xlsx" in kinds:
        if not data.startswith(XLSX_MAGIC):
            raise UploadRejected(f"{label}: the content is not an .xlsx workbook (the file extension does not match).")
        inspect_zip(data, label=label, expect_xlsx=True)
    elif ext == "xls" and "xls" in kinds:
        if not data.startswith(XLS_MAGIC):
            raise UploadRejected(f"{label}: the content is not an .xls workbook (the file extension does not match).")
    elif ext == "zip" and "zip" in kinds:
        if not data.startswith(XLSX_MAGIC):
            raise UploadRejected(f"{label}: the content is not a zip file.")
        inspect_zip(data, label=label)
    elif ext == "json" and "json" in kinds:
        if len(data) > MAX_JSON_BYTES:
            raise UploadRejected(f"{label}: the backup is larger than {MAX_JSON_BYTES // (1024 * 1024)} MB.")
    elif ext in ("sqlite3", "sqlite", "db") and "sqlite" in kinds:
        if len(data) > MAX_SQLITE_BYTES:
            raise UploadRejected(f"{label}: the database is larger than {MAX_SQLITE_BYTES // (1024 * 1024)} MB.")
        if not data.startswith(SQLITE_MAGIC):
            raise UploadRejected(f"{label}: the content is not a SQLite database.")
    else:
        raise UploadRejected(f"{label}: this file type is not allowed here.")
    return data


def load_json_backup(data):
    """bytes -> parsed backup. Size is limited and the top level must be an object."""
    check_upload("backup.json", data, kinds=("json",), max_bytes=MAX_JSON_BYTES)
    try:
        obj = json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        raise UploadRejected(f"backup.json: not valid JSON ({e}).") from None
    if not isinstance(obj, dict):
        raise UploadRejected("backup.json: expected a JSON object.")
    return obj


def neutralize_formula(value, strict=False):
    """openpyxl writes a text that starts with "=" into the workbook as a FORMULA, so customer text (names, addresses, marks)
    typed or imported by a user could run code in Excel. A leading quote keeps it text.
    strict=True also covers + - @ tab CR (for CSV, which spreadsheets re-interpret; see invoicing/csvout.py).
    Numbers and empty values are returned unchanged."""
    if not isinstance(value, str) or not value:
        return value
    return "'" + value if value.startswith(FORMULA_STARTS if strict else ("=",)) else value


# ---------------------------------------------------------------- layer 3: audit log and file permissions
_log = logging.getLogger("invoice_dashboard.audit")


def _audit_path():
    p = os.environ.get("INVOICE_AUDIT_LOG")
    if p:
        return Path(p)
    db = os.environ.get("INVOICE_DB_PATH")
    base = Path(db).parent if db else Path(__file__).resolve().parent.parent / "data"
    return base / "audit.log"


def audit(event, **fields):
    """Append one line (UTC time, event, key=value ...) to the audit log. Never records passwords or file contents.
    Logging must never break the app, so every error is swallowed."""
    try:
        clean = " ".join(f"{k}={str(v).replace(chr(10), ' ')[:120]}" for k, v in fields.items())
        line = f"{datetime.now(timezone.utc).isoformat(timespec='seconds')} {event} {clean}".rstrip()
        p = _audit_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        new = not p.exists()
        with open(p, "a", encoding="utf-8") as f:
            f.write(line + "\n")
        if new:
            private(p)
        _log.info(line)
    except Exception:       # noqa: BLE001
        pass


def private(path):
    """Owner-only permissions (0600) on a file that holds customer data. No effect on Windows."""
    try:
        os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass
