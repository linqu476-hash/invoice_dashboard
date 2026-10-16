from django.db import models


class Consignee(models.Model):
    """Consignee + shipping marks printed on the invoice for one destination country."""
    country = models.CharField(max_length=60, unique=True, help_text="Exactly as written in the COUNTRY column, e.g. DENMARK")
    name = models.CharField(max_length=120, blank=True)
    address = models.TextField(blank=True, help_text="Printed under CONSIGNEE (one line per row)")
    marks = models.TextField("Shipping marks", blank=True)

    class Meta:
        ordering = ["country"]

    def save(self, *a, **kw):
        self.country = self.country.strip().upper()
        super().save(*a, **kw)

    def __str__(self):
        return self.country
