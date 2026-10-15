# Security policy

This app handles customer addresses and invoice data. If you find a vulnerability, do not open a public issue: use GitHub's
"Report a vulnerability" (Security tab) or contact the repository owner directly.

What protects the app is described in the Security section of the README. Never commit `.streamlit/secrets.toml`, `data/*.sqlite3`
or `data/seed.json`; they are in `.gitignore` and the secret scan in CI checks the history.
