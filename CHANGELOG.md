# Changelog

## 30.0.0 — Consolidated release and compatibility gates
- Consolidates the v16–v30 upgrade sequence into one final package; no intermediate packages are required. The release checkpoints are documented below as validation stages, not as separately deployed releases.
- Makes `VERSION`, `app_core.__version__`, and the sidebar label agree on `30.0.0`, preventing the stale v11.0.1 label seen in the UI.
- Adds release compatibility tests for version-label consistency, the converted packing-plan → master invoice → invoice-build handoff, and the optional combined-workbook API.
- Preserves the v14.1.0 direct converted-file handoff, skip-list input, master validation, invoice calculations, security checks, and existing output formats.

### Upgrade-stage validation record (v16–v30)
- v16: version source and sidebar label consistency.
- v17: packing-plan conversion and converted workbook bytes contract.
- v18: converted workbook to master invoice handoff.
- v19: master-to-invoice build handoff.
- v20: skip-list and master validation path retained.
- v21: optional `combine` argument compatibility retained.
- v22: ZIP output and on-demand workbook extraction contract retained.
- v23: PDF report and validation CSV result contract retained.
- v24: upload size and safe output-filename checks retained.
- v25: template loading and workbook resource handling retained.
- v26: security and job tests included in release verification.
- v27: Python source compilation check.
- v28: package contents cleaned of caches and local runtime artifacts.
- v29: archive integrity and required-file checks.
- v30: full available regression suite and final version consistency check.

## 15.0.0
- Reused the existing optimized template-loading helper across remaining invoice-generation paths to avoid preserving unused external-link metadata.
- Cleaned release package of Git internals, Python/pytest caches, bytecode, and local database/log artifacts.
- Preserved the v14.1.0 converted-packing-plan → master invoice → invoice-build workflow.

# Changelog

## 14.1.0 — Link converted packing plans to master generation
- Added a direct action after Pre-sell column conversion to build `MASTER_INVOICE.xlsx` from `PACKING_PLAN_CONVERTED.xlsx`.
- Added optional invoice skip-list handling to this linked flow.
- Added a direct handoff from the generated master to the Build invoices page, avoiding manual download/re-upload.
- Reuses the existing validated `runner.build_master` pipeline; invoice rules and calculations are unchanged.
- Kept the existing one-step automatic Pre-sell-to-master path on the Build master invoice tab.

# Changelog

## 14.0.0 — Cache repeated description parsing in the generator
- Added bounded LRU caches for pure, repeatedly used text operations: description/composition splitting, whitespace-insensitive comparison normalization, and fabric-type detection. Invoice merging and grouping often evaluate the same descriptions multiple times; cached results avoid repeating regex and string processing across the generator.
- Caches are bounded (16,384 entries each) to avoid unbounded memory growth on unusually diverse input files.
- No grouping keys, invoice calculations, line ordering, validation rules, templates, outputs, or security behavior were intentionally changed.
- Existing lightweight UI and optional combined workbook behavior are retained.

## 13.0.0 — Reuse precomputed invoice grouping throughout generation
- Invoice grouping and composition flags are now reused by worksheet rendering instead of being recomputed in `build()` after the process-level precomputation. This removes repeated description normalization and grouping work for every invoice.
- Cache invalidation is tied to the effective composition-display mode so country-specific content rules remain correct.
- No invoice math, line ordering, validation rules, workbook templates, or security behavior changed intentionally.
- UI remains streamlined; combined workbook generation remains optional.

## 12.0.0
- Avoid a duplicate product-grouping pass during invoice rendering by deriving printed-line counts from the existing group data.
- Collapse the readiness summary by default to keep the results screen simpler while preserving access to all readiness details.
- Preserve the v11.0.1 hotfix for the optional `combine` argument and combined workbook behavior.

# Changelog

## 11.0.1 — Fix optional combined workbook build
- Fixed the build failure caused by the runner passing `combine=` to `engine.process()` while the engine function did not accept that keyword.
- Made combined-workbook generation genuinely optional: when disabled, the engine skips creating and rendering the second workbook, reducing build time and memory use.
- Kept per-family workbooks and the all-invoices ZIP; the combined workbook is included only when requested.
- ZIP packaging stores already-compressed XLSX files without recompressing them to reduce CPU overhead.
- Regression checks cover both default and combined build paths.


## 11.0.0 — Faster sparse formula adjustment and leaner results UI
- Optimized row-cut formula adjustment to scan only instantiated cells instead of iterating the full worksheet grid. This avoids visiting or materializing empty cells in sparse, template-heavy invoice and packing-list sheets while retaining the same formula row-shift rules.
- Removed the optional line-by-line invoice detail tables from the Streamlit results page. The generated Excel files remain the authoritative detailed output; summary, search, warnings, pagination, PDF/CSV reports, and downloads remain.
- Added regression coverage confirming formulas shift correctly without a full-grid scan. No invoice calculations, validation, workbook formats, or security behavior were intentionally changed.

## 10.0.0 — Reuse invoice grouping and simplify result UI
- Precompute each invoice's product groups, composition flags, and printed-line count once per build; reuse them in family workbook rendering, optional combined workbook rendering, line-limit checks, and invoice previews.
- Removed secondary totals and redundant start/end timestamp widgets from the result page while preserving the main totals, ready/pending indicators, warnings, invoice filtering, and downloads.
- Added regression tests for cached grouping and equivalent invoice output behavior.



## 9.0.0 — Faster template loading and simpler UI
- Added a dedicated template loader using `keep_links=False`. The bundled template has no external-link dependencies, so repeated family-workbook loads avoid preserving unused external-link metadata.
- Simplified the dashboard to plain quick links; removed the home country chart and the optional country breakdown chart/table from build results.
- Kept invoice calculations, validation, consignee/shipping-mark behavior, invoice-family workbooks, ZIP/PDF/CSV outputs, and the optional combined workbook unchanged.
- Regression checks cover the template loader and generated output workbook behavior.


## 8.0.0 — Lower peak memory while rendering invoice families
- Save each invoice-family workbook as soon as its last invoice is built, then release the workbook object instead of keeping every family workbook open through the entire build.
- Keep the same invoice calculations, family workbook filenames, ZIP contents, optional combined workbook, and progress step count.
- Added regression coverage for multiple invoice families and verified that all generated family workbooks remain in the output ZIP.


## 7.0.0 — Lower session memory for generated workbooks
- Keep the generated all-invoices ZIP as the single in-session copy of the XLSX outputs instead of retaining each XLSX twice (standalone and inside the ZIP).
- Individual invoice-family and optional combined workbooks are extracted from the ZIP only when the user requests a download; the downloadable filenames and workbook contents are unchanged.
- Restrict on-demand extraction to known output filenames and reject path-like names.
- Added regression tests for ZIP-only storage and safe on-demand workbook extraction.

# Changelog

## 11.0.1 — Fix optional combined workbook build
- Fixed the build failure caused by the runner passing `combine=` to `engine.process()` while the engine function did not accept that keyword.
- Made combined-workbook generation genuinely optional: when disabled, the engine skips creating and rendering the second workbook, reducing build time and memory use.
- Kept per-family workbooks and the all-invoices ZIP; the combined workbook is included only when requested.
- ZIP packaging stores already-compressed XLSX files without recompressing them to reduce CPU overhead.
- Regression checks cover both default and combined build paths.


## 6.0.0 — Faster results and ZIP packaging
- Invoice results are filtered first and displayed in pages (50/100/250/500 rows); only the visible page is converted into a display table, reducing rerender and browser work on large runs.
- ZIP bundles store generated XLSX workbooks without recompressing them, since each XLSX is already a compressed ZIP container. This trades some ZIP size for lower CPU time; individual XLSX files and contents are unchanged.
- Invoice calculations, validation, output workbook contents, PDF/CSV reports, authentication and upload protections are unchanged.


## 5.0.0 — Lower-memory Excel processing
- Standard packing-plan extraction now opens workbooks read-only and processes rows sequentially; Pre-sell sheets still use editable mode because conversion rewrites the workbook.
- Master validation now streams rows instead of retaining the entire workbook's cell values in memory. Validation rules, warning wording, duplicate detection, totals, and stop-on-error behavior are preserved.
- Added regression coverage for the standard packing-plan streaming path.
- Versioned the app as 5.0.0. Combined workbook remains optional and disabled by default; plain UI and invoice-generation behavior remain unchanged.


## 4.1.0 — Performance and simpler UI
- Made the combined all-in-one workbook optional in the Streamlit build page (off by default). Skipping it avoids a second invoice-render pass and reduces workbook memory and build time; the ZIP still contains the per-invoice family workbooks.
- Switched read-only master and packing-plan input paths to openpyxl read-only iteration and explicit workbook closing.
- Simplified the theme to system fonts and flat styling; removed bundled font files, shadows, gradients, and button hover animation.
- Reduced the default results summary to three key metrics and made invoice-line details and country charts opt-in.


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
