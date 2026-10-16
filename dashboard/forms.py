from django import forms

from .models import Consignee

XLSX = ".xlsx"


def _xlsx(f):
    if f and not f.name.lower().endswith(XLSX):
        raise forms.ValidationError("Please upload an .xlsx file.")
    return f


class UploadForm(forms.Form):
    packing = forms.FileField(label="Packing plan", required=False,
                              help_text="The master invoice is built from it automatically (cartons become SC260458, SC260458-1, -2 ...). "
                                        "With a master file below it only adds ART.NO. and CUSTOMER PO NO.")
    master = forms.FileField(label="MASTER_INVOICE.xlsx (optional)", required=False,
                             help_text="Only needed if you already have a master. Then it is used as it is.")
    invoices = forms.CharField(label="Only these invoices (optional)", required=False,
                               widget=forms.TextInput(attrs={"placeholder": "SC260458, SC260460"}),
                               help_text="When building the master from the packing plan. Empty = all invoices.")
    template = forms.FileField(label="Invoice template (optional)", required=False,
                               help_text="Leave empty to use the built-in Sample.xlsx layout")

    hide_content = forms.BooleanField(
        label="Remove content (composition) line from the next groups", required=False,
        help_text="The first group keeps its content line (e.g. 100%POLYESTER(100%RECYCLED) WOVEN). Tick to remove it from "
                  "all following groups; products are then combined by gender + garment + HTS code only.")

    clean_master = lambda self: _xlsx(self.cleaned_data.get("master"))
    clean_packing = lambda self: _xlsx(self.cleaned_data.get("packing"))
    clean_template = lambda self: _xlsx(self.cleaned_data.get("template"))


    def clean(self):
        d = super().clean()
        if not d.get("master") and not d.get("packing"):
            raise forms.ValidationError("Upload the packing plan (the master is built from it) or a MASTER_INVOICE.xlsx.")
        return d


class MasterForm(forms.Form):
    """Step 1 on its own: packing plan -> MASTER_INVOICE.xlsx (same as packing_to_master.py)."""
    packing = forms.FileField(label="Packing plan")
    sheet = forms.CharField(label="Sheet name (optional)", required=False,
                            help_text="Empty = first sheet that has the INV.NO. and PO Number headers")
    invoices = forms.CharField(label="Only these invoices (optional)", required=False,
                               widget=forms.TextInput(attrs={"placeholder": "SC260458, SC260460"}),
                               help_text="Empty = all invoices")

    clean_packing = lambda self: _xlsx(self.cleaned_data.get("packing"))


class ConsigneeForm(forms.ModelForm):
    class Meta:
        model = Consignee
        fields = ["country", "name", "address", "marks"]
        widgets = {"address": forms.Textarea(attrs={"rows": 7}), "marks": forms.Textarea(attrs={"rows": 7})}

    def clean(self):
        d = super().clean()
        country = (d.get("country") or "").strip().upper()
        if country != "UNITED KINGDOM" and not (d.get("address") and d.get("marks")):
            raise forms.ValidationError(
                "Address and shipping marks are required (only UNITED KINGDOM may stay empty: it uses the template).")
        return d
