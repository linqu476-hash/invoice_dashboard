# Changelog

## 4.0.0
- **Builds run in the background.** Streamlit stops a page's script when the user opens another page, so a build that ran inside *Build invoices*
  died silently on any click elsewhere. The build (and "Update invoices") now runs in a worker thread (`app_core/jobs.py`); the pages only start it,
  watch it and collect the result. Changing page, re-runs and closing the tab no longer stop it.
- **Failures are reported.** A failed build (with every problem of a bad master and its row) or a cancelled one stays on the **Dashboard** and the Build
  page until dismissed, a pop-up appears on whatever page you are on, a link shows in the sidebar, and each start / end / failure is a line in the audit log.
- Live progress (bar, elapsed, time left, **Cancel**) on the Build page, the Dashboard and the sidebar; new **Recent builds** table on the Dashboard.
- One build at a time per session, at most 2 on the server (others wait), sessions cannot see each other's jobs, sign-out / expiry cancels the session's
  builds, finished jobs are dropped after 2 hours (last 10 per session kept).
- Limits: a server restart ends running builds (jobs are in memory). Invoice rules, stored data, look and security layers are unchanged from 3.0.
- Tests: `tests/test_jobs.py` (success, failure, bad master, cancel, queueing, isolation, clean-up, real builds in a thread).

## 3.0.0
- New typography. **Inter** for text, tables and forms, **Plus Jakarta Sans** for page titles, headings and big numbers, **JetBrains Mono** for code.
  Numbers use tabular figures so columns of amounts line up. Tighter heading spacing, clearer small captions and sidebar section labels.
- Fonts are **self-hosted** (`static/fonts`, SIL Open Font License files included) and declared in `.streamlit/config.toml`: no request goes to
  Google or any CDN, so the app still works offline and leaks nothing.
- Colours, layout, invoice rules, stored data and security layers are unchanged from 2.0.
- Requires Streamlit 1.50 or newer (font settings in the theme).

## 2.0.0
- New look: navy and blue palette on a light canvas, white cards with soft shadows, rounded buttons, tabs, uploaders and tables
  (`.streamlit/config.toml` for the theme, `app_core/theme.py` for the card styles).
- New **Dashboard** home page: key numbers, the 4-step workflow with shortcuts, and the last build with an amount-by-country chart.
- Grouped sidebar navigation: Overview, Invoicing, Master data, System. Version number shown in the sidebar.
- Centred, branded sign-in page. Page headers with icon badges. Metrics shown as cards.
- No change to the invoice rules, the stored data or the security layers of 1.x.

## 1.1.0
- Three security layers: hardened sign-in, upload validation, data and deployment protection (see the Security section of the README).
