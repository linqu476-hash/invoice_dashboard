"""Shipping Instruction uploads: many Excel files at once, or one or more .zip files of them.

Works on (name, bytes) pairs so it does not depend on the web framework."""
import io
import zipfile
from pathlib import PurePosixPath

from invoicing import shipmarks

from . import security, store

MAX_SI_FILES = 500                 # Excel files in one upload (zip members count)
MAX_SI_BYTES = 20 * 1024 * 1024    # one Excel file
MAX_ZIP_BYTES = 200 * 1024 * 1024  # one .zip upload


def _checked(display, base, content):
    """(display, base, content, None) when the bytes really are an Excel workbook, else the same with the reason."""
    try:
        security.check_upload(base, content, kinds=("xls", "xlsx"), max_bytes=MAX_SI_BYTES)
    except security.UploadRejected as e:
        return display, base, None, str(e)
    return display, base, content, None


def iter_si_uploads(files):
    """files: iterable of (name, bytes) -> yields (display name, file name, bytes or None, error or None).
    A zip is opened and only .xls / .xlsx members are read."""
    n = 0
    for name, data in files:
        low = name.lower()
        if low.endswith(".zip"):
            try:
                security.check_upload(name, data, kinds=("zip",), max_bytes=MAX_ZIP_BYTES)
                z = zipfile.ZipFile(io.BytesIO(data))
            except security.UploadRejected as e:
                yield name, name, None, str(e)
                continue
            except zipfile.BadZipFile:
                yield name, name, None, "Cannot read the zip file."
                continue
            with z:
                members = [i for i in z.infolist()
                           if not i.is_dir() and i.filename.lower().endswith((".xls", ".xlsx"))
                           and "__MACOSX" not in i.filename and not PurePosixPath(i.filename).name.startswith(("~$", "."))]
                if not members:
                    yield name, name, None, "The zip file has no .xls or .xlsx files."
                for i in members:
                    n += 1
                    base = PurePosixPath(i.filename).name
                    if n > MAX_SI_FILES:
                        yield f"{name} / {base}", base, None, f"Skipped: more than {MAX_SI_FILES} files in one upload."
                    elif i.file_size > MAX_SI_BYTES:
                        yield f"{name} / {base}", base, None, "Skipped: file is larger than 20 MB."
                    else:
                        yield _checked(f"{name} / {base}", base, z.read(i))
        elif low.endswith((".xls", ".xlsx")):
            n += 1
            if n > MAX_SI_FILES:
                yield name, name, None, f"Skipped: more than {MAX_SI_FILES} files in one upload."
            elif len(data) > MAX_SI_BYTES:
                yield name, name, None, "Skipped: file is larger than 20 MB."
            else:
                yield _checked(name, name, data)
        else:
            yield name, name, None, "Not an Excel file: use .xls, .xlsx or a .zip of them."


def save_si_uploads(files):
    """Read every Shipping Instruction and save its consignee. One failing file does not stop the others.
    -> dict(rows=[...], total, added, replaced, unchanged, failed)"""
    rows = []
    for display, fname, data, err in iter_si_uploads(files):
        row = dict(file=display, status="failed", country="", customer_no="", name="", error=err or "")
        if not err:
            try:
                status, c = store.save_si_consignee(store.read_si_bytes(fname, data))
                row.update(status=status, country=c["country"], customer_no=c["customer_no"], name=c["name"])
            except shipmarks.MarksError as e:
                row["error"] = str(e)
            except Exception as e:                      # bad workbook ...
                row["error"] = f"Could not read the shipping instruction: {e}"
        rows.append(row)
    count = lambda st: sum(1 for r in rows if r["status"] == st)
    return dict(rows=rows, total=len(rows), added=count("added"), replaced=count("replaced"), unchanged=count("unchanged"), failed=count("failed"))
