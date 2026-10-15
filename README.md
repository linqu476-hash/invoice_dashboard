# Invoice dashboard (Streamlit)

Packing plan → `MASTER_INVOICE.xlsx` → commercial invoices + packing lists, with the same rules as the Django version.
The invoice logic is the **unchanged `invoicing/` package**; this project replaces the Django layer (views, templates, models, auth) with Streamlit pages.

## Run it locally

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml     # then set APP_PASSWORD
streamlit run streamlit_app.py
```

The app refuses to start without `APP_PASSWORD` (it handles customer addresses and invoice data).

## Pages

| Page | What it does |
|---|---|
| **Build invoices** | Upload `MASTER_INVOICE.xlsx` → validation (stops with every problem + Excel row) → invoices. Result: totals, merge summary, downloads (zip, combined workbook, PDF run report, validation CSV, one workbook per invoice), grouped "Check before sending", searchable invoice list with line gauge, invoice detail, by-country chart, **Update invoices** after adding a missing consignee. |
| **Packing plan to master** | Packing plan or Pre-sell file → master (Pre-sell is converted automatically), skip list, under-$50 / USA / AIR-CC rules, "Build the invoices from this master →". Second tab: only rename the Pre-sell columns. |
| **Consignees** | One entry per customer number; add / edit / delete; Shipping Instruction upload (many files or a zip of up to 500); missing-consignee list of the last run. |
| **Shipping marks** | CONSIGNEE Excel upload (`.xls` / `.xlsx`), add / edit / delete. |
| **Backup and data** | Download / restore a JSON backup; import consignees and marks from the old Django `db.sqlite3`. |

## Not carried over from the Django version

These need a database of runs and user accounts, which a Streamlit app on an ordinary host does not have:

- **Run history, All runs page, Analytics page, CSV export of runs, `purge_jobs`.** A run now lives in the browser session: download what you need before closing the tab or signing out.
- **Roles (Operator / Administrator), per-user run ownership, sign-in lockout, audit trail, Django admin.** Replaced by **one shared password** (`APP_PASSWORD`). For real per-user sign-in use Streamlit's `st.login` (OIDC) or put the app behind your company's SSO / reverse proxy.

If you need those, keep the Django app for them; both can share the same `invoicing/` package.

## Deploying

### Where the data lives (read this first)
Consignees and shipping marks are stored in SQLite at `data/invoice_dashboard.sqlite3` (override with `INVOICE_DB_PATH`).
**Streamlit Community Cloud wipes local files on every restart or redeploy**, so anything saved in the app is lost there. Choose one:

1. **Host with a persistent disk** (VPS, Render, Railway, Fly.io…): set `INVOICE_DB_PATH` to a path on that disk. A `Dockerfile` is included (not built or tested here).
2. **Community Cloud + seed file:** download a backup on *Backup and data*, save it as `data/seed.json` in a **private** repository. An empty database is filled from it at start-up. Changes made in the app after that are lost at the next restart unless you download a new backup and commit it.

### Streamlit Community Cloud
1. Push this folder to a **private** GitHub repository (see "Keep out of git" below).
2. On share.streamlit.io: *Create app* → pick the repo → main file `streamlit_app.py`.
3. *Advanced settings → Secrets*: `APP_PASSWORD = "a-long-passphrase"`.
4. A Community Cloud app URL is reachable by anyone who has it unless you restrict viewers in the app's sharing settings; the password gate is your minimum protection.

### Keep out of git
`.gitignore` already excludes `.streamlit/secrets.toml`, `data/*.sqlite3` and `data/seed.json`. The old Django archive also held `db.sqlite3` (13 consignees, 17 shipping marks, job history and an audit log) and a `sample_data/` folder of real-looking workbooks: **do not commit those anywhere public.** Migrate the consignees and marks with `python scripts/import_from_django.py path/to/db.sqlite3` (or the Import box on *Backup and data*).

## Good to know

- **Build time.** A packing plan of about 340 PO lines becomes 26 invoices in roughly 35 seconds on a normal laptop; shared hosting can be slower. Profiling shows most of it is openpyxl's `merge_cells` called from `engine.cut_rows` and `engine.apply`, so it is an engine cost, not a Streamlit one, and it is the first place to optimise.
- **Old masters.** A master without the `INV DATE` column (for example the old `sample_data/MASTER_INVOICE.xlsx`) is refused, exactly as in the Django app; rebuild it on *Packing plan to master*.
- **Memory.** Uploaded files and built invoices are kept in the session (not on disk) until you sign out or close the tab.
- **Upload limit.** 50 MB in `.streamlit/config.toml`; one workbook is capped at 20 MB by the validation rules.

## Tests

```bash
python -m unittest discover -s tests -v          # service layer, synthetic data only (about 2 s)
python tests/smoke_pages.py [packing_plan.xlsx]  # runs every page script against a fake streamlit module
```
The smoke test checks that the pages handle real built data; it cannot check Streamlit's own widgets or layout. Click through the real app once after deploying.

## Layout

    streamlit_app.py     entry point: password gate, navigation
    views/               build.py  master.py  consignees.py  marks.py  data.py
    app_core/            store.py (SQLite)  runner.py (build + checks)  uploads.py  auth.py  ui.py
    invoicing/           the invoice logic, unchanged from the Django project
    assets/              invoice_template.xlsx
    scripts/             import_from_django.py
    tests/
