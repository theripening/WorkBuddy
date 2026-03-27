import io
import os
import contextlib
import tempfile
from datetime import date
from pathlib import Path

from django.shortcuts import render
from django.http import HttpResponse
from django.views.decorators.http import require_POST

from . import pr

_HUB = "scripts/hub.html"


def hub(request):
    return render(request, _HUB)


def _hub_results(request, tool, output, success=True):
    return render(request, _HUB, {"results": {"tool": tool, "output": output, "success": success}})


# ---------------------------------------------------------------------------
# PR Tool
# ---------------------------------------------------------------------------

@require_POST
def pr_run(request):
    card_file = request.FILES.get("card")
    ach_file = request.FILES.get("ach")
    ta_file = request.FILES.get("ta")

    missing = [n for n, f in [("Card.xlsx", card_file), ("ACH.xlsx", ach_file), ("TA.xlsx", ta_file)] if not f]
    if missing:
        return _hub_results(request, "pr", f"Missing files: {', '.join(missing)}", success=False)

    try:
        import pandas  # noqa: F401
    except ImportError:
        return _hub_results(request, "pr", "pandas is not installed. Run: pip install pandas openpyxl", success=False)

    try:
        csv_bytes, filename = pr.run_for_today(card_file, ach_file, ta_file)
    except Exception as exc:
        return _hub_results(request, "pr", str(exc), success=False)

    response = HttpResponse(csv_bytes, content_type="text/csv")
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


# ---------------------------------------------------------------------------
# Find Phrase
# ---------------------------------------------------------------------------

@require_POST
def findphrase_run(request):
    folder = request.POST.get("folder", "").strip()
    term = request.POST.get("term", "").strip()
    filter_str = request.POST.get("filter", "").strip() or None

    if not folder or not term:
        return _hub_results(request, "findphrase", "Folder and search term are required.", success=False)

    buf = io.StringIO()
    try:
        from .findphrase import search_in_files
        with contextlib.redirect_stdout(buf):
            search_in_files(folder, term, filter_str)
        output = buf.getvalue().strip() or "(no matches found)"
    except Exception as exc:
        return _hub_results(request, "findphrase", str(exc), success=False)

    return _hub_results(request, "findphrase", output)


# ---------------------------------------------------------------------------
# Merge CSVs
# ---------------------------------------------------------------------------

@require_POST
def mergecsv_run(request):
    folder = request.POST.get("folder", "").strip()
    filter_str = request.POST.get("filter", "").strip() or None
    output_name = request.POST.get("output", "").strip() or "merged.csv"

    if not folder:
        return _hub_results(request, "mergecsv", "Folder path is required.", success=False)

    buf = io.StringIO()
    try:
        from .merge_csvs import merge_csvs_in_folder
        with contextlib.redirect_stdout(buf):
            merge_csvs_in_folder(folder, output_filename=output_name, filter_str=filter_str)
        output = buf.getvalue().strip() or "Done."
    except Exception as exc:
        return _hub_results(request, "mergecsv", str(exc), success=False)

    return _hub_results(request, "mergecsv", output)


# ---------------------------------------------------------------------------
# SQL Import Folder
# ---------------------------------------------------------------------------

@require_POST
def sqlimport_run(request):
    folder = request.POST.get("folder", "").strip()
    server = request.POST.get("server", "").strip() or "MPOPE-11V\\SQLEXPRESS"
    database = request.POST.get("database", "").strip() or "ebppsupport"
    if_exists = request.POST.get("if_exists", "replace")

    if not folder:
        return _hub_results(request, "sqlimport", "Folder path is required.", success=False)

    buf = io.StringIO()
    try:
        from .sql_import_folder import import_file_to_sqlserver
        lines = []
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path) and filename.lower().endswith((".csv", ".xlsx")):
                table_name = os.path.splitext(filename)[0]
                with contextlib.redirect_stdout(buf):
                    import_file_to_sqlserver(file_path, table_name, server, database, if_exists)
        output = buf.getvalue().strip() or "No CSV/Excel files found in folder."
    except Exception as exc:
        return _hub_results(request, "sqlimport", str(exc), success=False)

    return _hub_results(request, "sqlimport", output)


# ---------------------------------------------------------------------------
# Search Big (find files by partial name)
# ---------------------------------------------------------------------------

@require_POST
def searchbig_run(request):
    folder = request.POST.get("folder", "").strip()
    term = request.POST.get("term", "").strip()

    if not folder or not term:
        return _hub_results(request, "searchbig", "Folder and search term are required.", success=False)

    try:
        matches = []
        for root, _, files in os.walk(folder):
            for file in files:
                if term.lower() in file.lower():
                    matches.append(os.path.join(root, file))
        output = "\n".join(matches) if matches else "(no matching files found)"
    except Exception as exc:
        return _hub_results(request, "searchbig", str(exc), success=False)

    return _hub_results(request, "searchbig", output)


# ---------------------------------------------------------------------------
# SQL Import (Single File) — uses sql_importer3.py
# ---------------------------------------------------------------------------

@require_POST
def sqlimport_file_run(request):
    uploaded = request.FILES.get("file")
    server = request.POST.get("server", "").strip() or "MPOPE-11V\\SQLEXPRESS"
    database = request.POST.get("database", "").strip() or "ebppsupport"
    table_name = request.POST.get("table_name", "").strip() or None
    if_exists = request.POST.get("if_exists", "replace")
    delimiter = "|" if request.POST.get("pipe") else request.POST.get("delimiter", ",").strip() or ","

    if not uploaded:
        return _hub_results(request, "sqlimport_file", "No file uploaded.", success=False)

    buf = io.StringIO()
    with tempfile.TemporaryDirectory() as tmpdir:
        dest = Path(tmpdir) / uploaded.name
        with open(dest, "wb") as f:
            for chunk in uploaded.chunks():
                f.write(chunk)
        try:
            from .sql_importer3 import import_file_to_sqlserver
            with contextlib.redirect_stdout(buf):
                import_file_to_sqlserver(str(dest), table_name, server, database, if_exists, delimiter)
            output = buf.getvalue().strip() or "Done."
        except Exception as exc:
            return _hub_results(request, "sqlimport_file", str(exc), success=False)

    return _hub_results(request, "sqlimport_file", output)
