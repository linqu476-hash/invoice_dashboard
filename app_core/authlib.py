"""Sign-in logic without Streamlit (so it can be unit tested): password hashing and a lockout that survives a page refresh.

Password hash format:  scrypt$<N>$<r>$<p>$<salt hex>$<hash hex>      make one with:  python scripts/make_password_hash.py
"""
import hashlib
import hmac
import json
import os
import secrets
import tempfile
import time
from pathlib import Path

SCRYPT_N, SCRYPT_R, SCRYPT_P = 2 ** 14, 8, 1

MAX_FAILS = 5                     # wrong passwords in a row before the lock starts
BASE_LOCK_SECONDS = 60            # first lock; doubles at every further failure, up to MAX_LOCK_SECONDS
MAX_LOCK_SECONDS = 15 * 60
FAIL_WINDOW_SECONDS = 15 * 60     # failures older than this are forgotten


def hash_password(password, *, n=SCRYPT_N, r=SCRYPT_R, p=SCRYPT_P):
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=n, r=r, p=p, dklen=32)
    return f"scrypt${n}${r}${p}${salt.hex()}${digest.hex()}"


def verify_password(password, stored):
    """stored is a scrypt hash (see above) or, for old setups, the plain password. Constant-time comparison."""
    if stored.startswith("scrypt$"):
        try:
            _, n, r, p, salt, digest = stored.split("$")
            calc = hashlib.scrypt(password.encode("utf-8"), salt=bytes.fromhex(salt), n=int(n), r=int(r), p=int(p), dklen=len(digest) // 2)
            return hmac.compare_digest(calc.hex(), digest)
        except (ValueError, TypeError):
            return False
    return hmac.compare_digest(password.encode("utf-8"), stored.encode("utf-8"))


# ---------------------------------------------------------------- lockout shared by ALL sessions
# Streamlit does not give the app the visitor's IP address, so the lock is global: after MAX_FAILS wrong passwords in a row
# NOBODY can sign in until the lock ends. That stops password guessing from many tabs / browsers at the cost of a short
# lock-out of the real users while someone is guessing (the audit log shows when it happened).
def _state_path():
    p = os.environ.get("INVOICE_LOCKOUT_FILE")
    if p:
        return Path(p)
    db = os.environ.get("INVOICE_DB_PATH")
    base = Path(db).parent if db else Path(__file__).resolve().parent.parent / "data"
    return base / "lockout.json"


def _read():
    try:
        return json.loads(_state_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _write(state):
    p = _state_path()
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=p.parent, prefix=".lockout-")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(state, f)
        os.replace(tmp, p)
        os.chmod(p, 0o600)
    except OSError:
        pass


def seconds_locked(now=None):
    """0 when sign-in is allowed, else how many seconds are left."""
    now = time.time() if now is None else now
    return max(0, int(_read().get("locked_until", 0) - now + 0.999))


def record_failure(now=None):
    """Count one wrong password -> seconds the sign-in is now locked (0 = not locked yet)."""
    now = time.time() if now is None else now
    s = _read()
    fails = s.get("fails", 0) if now - s.get("last_fail", 0) < FAIL_WINDOW_SECONDS else 0
    fails += 1
    s.update(fails=fails, last_fail=now)
    lock = 0
    if fails >= MAX_FAILS:
        lock = min(BASE_LOCK_SECONDS * 2 ** (fails - MAX_FAILS), MAX_LOCK_SECONDS)
        s["locked_until"] = now + lock
    _write(s)
    return lock


def record_success():
    _write({})
