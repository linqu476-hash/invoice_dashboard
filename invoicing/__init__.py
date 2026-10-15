"""Invoice logic with no Django in it: packing plan -> master -> invoices.

    presell     Packing Plan Pre-sell -> packing plan layout
    extract     packing plan -> MASTER_INVOICE.xlsx
    engine      MASTER_INVOICE.xlsx -> commercial invoices + packing lists
    shipmarks   CONSIGNEE / Shipping Instruction Excel -> shipping marks
    validation  checks of the master before anything is built

The Django app (dashboard/) imports from here; the scripts in the project root call the same modules.
"""
