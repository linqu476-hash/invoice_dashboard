"""Make the APP_PASSWORD_HASH value:   python scripts/make_password_hash.py
Paste the printed line into .streamlit/secrets.toml (or the host's Secrets box) and delete APP_PASSWORD."""
import getpass
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app_core.authlib import hash_password  # noqa: E402

pw = getpass.getpass("New password (at least 12 characters): ")
if len(pw) < 12:
    sys.exit("Too short: use a passphrase of at least 12 characters.")
if pw != getpass.getpass("Repeat the password: "):
    sys.exit("The two passwords are different.")
print(f'\nAPP_PASSWORD_HASH = "{hash_password(pw)}"')
