# Changelog

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
