from django.db import migrations, models


def seed_uk(apps, schema_editor):
    # United Kingdom keeps the consignee address and shipping marks already in the Excel template
    apps.get_model("dashboard", "Consignee").objects.get_or_create(
        country="UNITED KINGDOM", defaults=dict(name="ADIDAS INTERNATIONAL TRADING AG"))


class Migration(migrations.Migration):
    initial = True
    dependencies = []
    operations = [
        migrations.CreateModel(
            name="Consignee",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("country", models.CharField(help_text="Exactly as written in the COUNTRY column, e.g. DENMARK", max_length=60, unique=True)),
                ("name", models.CharField(blank=True, max_length=120)),
                ("address", models.TextField(blank=True, help_text="Printed under CONSIGNEE (one line per row)")),
                ("marks", models.TextField(blank=True, verbose_name="Shipping marks")),
            ],
            options={"ordering": ["country"]},
        ),
        migrations.RunPython(seed_uk, migrations.RunPython.noop),
    ]
