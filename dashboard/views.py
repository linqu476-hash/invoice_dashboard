import json
import re
import shutil
import uuid
from datetime import date as _date
from pathlib import Path

from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from . import engine, extract
from .forms import ConsigneeForm, MasterForm, UploadForm
from .models import Consignee

JOB_RE = re.compile(r"^[0-9a-f]{32}$")


def _job_dir(job_id):
    if not JOB_RE.match(job_id):
        raise Http404
    d = Path(settings.JOBS_DIR) / job_id
    if not (d / "result.json").exists():
        raise Http404
    return d


def _consignees():
    return {c.country: dict(name=c.name, address=c.address, marks=c.marks) for c in Consignee.objects.all()}


def _recent(limit=8):
    out = []
    root = Path(settings.JOBS_DIR)
    if root.exists():
        for d in sorted(root.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
            f = d / "result.json"
            if JOB_RE.match(d.name) and f.exists():
                r = json.loads(f.read_text(encoding="utf-8"))
                out.append(dict(id=d.name, master=r["master_name"], created=r["created"], **r["totals"]))
            if len(out) >= limit:
                break
    return out


def index(request):
    error = None
    if request.method == "POST":
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            cd = form.cleaned_data
            job_id = uuid.uuid4().hex
            d = Path(settings.JOBS_DIR) / job_id
            (d / "input").mkdir(parents=True)

            def save(f, name):
                p = d / "input" / name
                with open(p, "wb") as out:
                    for chunk in f.chunks():
                        out.write(chunk)
                return p

            try:
                packing = save(cd["packing"], "packing.xlsx") if cd["packing"] else None
                built = None
                if cd["master"]:
                    master = save(cd["master"], "master.xlsx")
                else:                                  # packing plan -> MASTER_INVOICE.xlsx first, then the generator
                    only = {x.strip() for x in re.split(r"[,\s]+", cd["invoices"]) if x.strip()} or None
                    rows, notes = extract.extract(packing, only=only)
                    (d / "output").mkdir(exist_ok=True)
                    master = extract.write_master(rows, d / "output" / "MASTER_INVOICE.xlsx")
                    built = dict(file="MASTER_INVOICE.xlsx", notes=notes, rows=len(rows),
                                 splits=sum(1 for s in extract.summarize(rows) if "-" in s["inv"]))
                template = save(cd["template"], "template.xlsx") if cd["template"] else Path(settings.TEMPLATE_XLSX)
                result = engine.process(master, template, packing, None, d / "output", _consignees(),
                                        hide_comp=cd["hide_content"])
                result.update(master_name=cd["master"].name if cd["master"] else cd["packing"].name,
                              created=_date.today().isoformat(), built=built)
                (d / "result.json").write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
                return redirect(reverse("job", args=[job_id]))
            except Exception as e:  # bad workbook, missing columns, ...
                shutil.rmtree(d, ignore_errors=True)
                error = f"Could not build the invoices: {e}"
    else:
        form = UploadForm()
    return render(request, "dashboard/index.html", dict(form=form, error=error, recent=_recent(),
                                                        consignees=Consignee.objects.count()))


def job(request, job_id):
    d = _job_dir(job_id)
    r = json.loads((d / "result.json").read_text(encoding="utf-8"))
    chart = dict(
        countries=[c["country"] for c in r["by_country"]], country_qty=[c["qty"] for c in r["by_country"]],
        invoices=[i["no"] for i in r["invoices"]], inv_amount=[i["amount"] for i in r["invoices"]])
    warnings = [dict(inv=i["no"], msg=m) for i in r["invoices"] for m in i["warnings"]]
    return render(request, "dashboard/job.html", dict(job_id=job_id, r=r, chart=chart, warnings=warnings))


def download(request, job_id, name):
    d = _job_dir(job_id)
    r = json.loads((d / "result.json").read_text(encoding="utf-8"))
    allowed = set(r["files"]["items"]) | {r["files"]["combined"], r["files"]["zip"]}
    if r.get("built"):
        allowed.add(r["built"]["file"])
    if name not in allowed:
        raise Http404
    return FileResponse(open(d / "output" / name, "rb"), as_attachment=True, filename=name)


@require_POST
def delete_job(request, job_id):
    d = _job_dir(job_id)
    shutil.rmtree(d, ignore_errors=True)
    return redirect("index")


# ---- step 1 on its own: packing plan -> master (packing_to_master.py) ----------------------
def _master_dir(mid):
    if not JOB_RE.match(mid):
        raise Http404
    d = Path(settings.JOBS_DIR) / "masters" / mid
    if not (d / "MASTER_INVOICE.xlsx").exists():
        raise Http404
    return d


def master_tool(request):
    error, built = None, None
    if request.method == "POST":
        form = MasterForm(request.POST, request.FILES)
        if form.is_valid():
            cd = form.cleaned_data
            mid = uuid.uuid4().hex
            d = Path(settings.JOBS_DIR) / "masters" / mid
            d.mkdir(parents=True)
            try:
                src = d / "packing.xlsx"
                with open(src, "wb") as out:
                    for chunk in cd["packing"].chunks():
                        out.write(chunk)
                only = {x.strip() for x in re.split(r"[,\s]+", cd["invoices"]) if x.strip()} or None
                rows, notes = extract.extract(src, cd["sheet"] or None, only)
                extract.write_master(rows, d / "MASTER_INVOICE.xlsx")
                summary = extract.summarize(rows)
                built = dict(id=mid, rows=len(rows), qty=sum(r["qty"] for r in rows), notes=notes, summary=summary,
                             amount=round(sum(s["amount"] for s in summary), 2),
                             splits=sum(1 for s in summary if "-" in s["inv"]), source=cd["packing"].name)
            except Exception as e:
                shutil.rmtree(d, ignore_errors=True)
                error = f"Could not build the master invoice: {e}"
    else:
        form = MasterForm()
    return render(request, "dashboard/master.html", dict(form=form, error=error, built=built))


def master_download(request, mid):
    d = _master_dir(mid)
    return FileResponse(open(d / "MASTER_INVOICE.xlsx", "rb"), as_attachment=True, filename="MASTER_INVOICE.xlsx")


# ---- consignees --------------------------------------------------------------------------
def consignees(request):
    return render(request, "dashboard/consignees.html", dict(items=Consignee.objects.all()))


def consignee_edit(request, pk=None):
    obj = get_object_or_404(Consignee, pk=pk) if pk else None
    form = ConsigneeForm(request.POST or None, instance=obj)
    if request.method == "POST" and form.is_valid():
        form.save()
        return redirect("consignees")
    return render(request, "dashboard/consignee_form.html", dict(form=form, obj=obj))


@require_POST
def consignee_delete(request, pk):
    get_object_or_404(Consignee, pk=pk).delete()
    return redirect("consignees")
