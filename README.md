# Invoice Dashboard (Django)

Upload the packing plan (or an existing MASTER_INVOICE.xlsx) and get commercial invoices, packing lists, a combined workbook
(sheets INV, PL, INV-1, PL-1 ...; with several invoice numbers: 'INV SC260458', 'PL SC260458-1' ...) and a zip of everything, plus a summary dashboard.

## Run it (Windows or Linux)

1. Install Python 3.10 or newer.
2. Open a terminal in this folder and install the packages:
   `pip install -r requirements.txt`
3. Create the database (this also adds the United Kingdom consignee):
   `python manage.py migrate`
4. Start the server:
   `python manage.py runserver`
5. Open http://127.0.0.1:8000 and upload your file.
   To try it first, use the files in `sample_data/`.

## Using it

1. Dashboard page: upload the packing plan. The master invoice is extracted from it first (each carton
   after the first one of an invoice becomes SC260458-1, -2, ...), then the invoices are generated.
   "Only these invoices" limits the run, e.g. `SC260458`. If you already have a MASTER_INVOICE.xlsx, upload
   it as well and it is used as it is. The invoice date is the CRD of the packing plan (master column INV DATE); there is no date field in the UI.
2. Result page: summary cards, charts, the list of invoices (click one to see its groups and PO lines),
   the "Check before sending" warnings, and the download buttons.
3. Consignees page: add the consignee address and shipping marks for every destination country
   (DENMARK, GERMANY, ...). Until a country is added, its invoice shows yellow placeholders.

## Packing plan -> master (the extract step)

- Columns are found by header text (INV.NO., PO Number, Gps Customer Number, COUNTRY, DESCRIPTION, HTS CODE1:,
  Ordered Quantity, Price Per Unit, Total Adjustments, weight, CARTON WEIGHT, CARTON NO.).
- A carton ends at the orange total row (CARTON NO. 1, 2, 3 ...) or at the green "1 OF 1" row for a single carton.
- N.W = sum of quantity x weight of the carton; G.W = N.W + carton weight; AMOUNT = quantity x price + adjustment.
- It reads the first sheet that has the INV.NO. and PO Number headers. From the command line:
  `python dashboard/extract.py` is the same code as `packing_to_master.py` (`--sheet`, `--invoices` options).

## Master columns and one file per invoice family

- MASTER_INVOICE.xlsx columns: INV NO, **INV DATE**, PO NO, **CUSTOMER PO NO**, **ART NO**, CUS NO, COUNTRY, DESCRIPTION, HTS CODE, AMOUNT, QTY, N.W, G.W.
  CUSTOMER PO NO comes from the packing plan column `Customer Number`, ART NO from `Article Number`, both per PO line. INV DATE comes from `Customer Request Date (CRD)` and is printed as the invoice DATE.
  An older master without these two columns still works (values are then looked up in the packing plan by PO number).
- Invoices are written **one workbook per invoice family**: `SC260460.xlsx` holds SC260460, -1, -2, -3 ... as sheets
  INV, PL, INV-1, PL-1, INV-2, PL-2 ... (`--combine` / the combined file still holds every invoice).

## Rules built in (from README.MD)

- Same description + HTS code + composition = one group (tick "Remove content line from the next groups" on the dashboard, or `--no-content`: the first group keeps its composition line, the next groups lose it and groups are description + HTS code only); groups show 4 lines then the PO lines.
- Invoice file name = invoice number (SC260458, SC260458-1 ...); sheets are INV / PL / INV-1 / PL-1.
- Packing list: N.W(KGS) and G.W(KGS) per PO line, split by quantity from the master totals,
  weights rounded to the nearest 0.1 KG (5.39 -> 5.40, 6.43 -> 6.40).

## The two scripts (command line and UI)

- `packing_to_master.py` (step 1) = `dashboard/extract.py` = UI page **Packing plan -> Master** (`/master/`).
- `generate_invoices.py` (step 2) = `dashboard/engine.py` = UI page **Dashboard** (`/`). With `--packing` and no `--master` it builds the master first.
- One code base: change `extract.py` or `engine.py` and both the UI and the scripts follow.

## Files

- `dashboard/extract.py`: packing plan -> MASTER_INVOICE.xlsx.
- `dashboard/engine.py`: the invoice builder (same logic as generate_invoices.py, also runnable alone).
- `assets/invoice_template.xlsx`: layout copied for every invoice. Replace it, or upload a template per run.
- `jobs/`: one folder per upload with the inputs and generated files.
- Tests: `python manage.py test`
