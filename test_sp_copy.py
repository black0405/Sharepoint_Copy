"""
Self-check for the offline parts of sp_copy.py (no network, no framework).

    python test_sp_copy.py
"""
import base64
import os
import tempfile

import openpyxl

import sp_copy as sp

# --- normalise_link ---------------------------------------------------------
browser = (
    "https://t.sharepoint.com/sites/Fin/Shared%20Documents/Forms/AllItems.aspx"
    "?id=%2Fsites%2FFin%2FShared%20Documents%2FReports%2F2026&viewid=abc"
)
assert sp.normalise_link(browser) == \
    "https://t.sharepoint.com/sites/Fin/Shared%20Documents/Reports/2026"

old_style = "https://t.sharepoint.com/sites/Fin/Forms/AllItems.aspx?RootFolder=%2Fsites%2FFin%2FDocs"
assert sp.normalise_link(old_style) == "https://t.sharepoint.com/sites/Fin/Docs"

sharing = "https://t.sharepoint.com/:f:/s/Fin/EabcDEF123?e=xyz"
assert sp.normalise_link(sharing) == sharing

plain = " https://t.sharepoint.com/sites/Fin/Shared Documents/Reports "
assert sp.normalise_link(plain) == \
    "https://t.sharepoint.com/sites/Fin/Shared%20Documents/Reports"

# --- share_token ------------------------------------------------------------
token = sp.share_token("https://t.sharepoint.com/sites/Fin/Docs")
assert token.startswith("u!") and "=" not in token
padded = token[2:] + "=" * (-len(token[2:]) % 4)
assert base64.urlsafe_b64decode(padded).decode() == "https://t.sharepoint.com/sites/Fin/Docs"

# --- read_jobs --------------------------------------------------------------
with tempfile.TemporaryDirectory() as tmp:
    path = os.path.join(tmp, "jobs.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["File", " Source", "TARGET"])          # case/space tolerant
    ws.append(["a.xlsx", "https://s/1", "https://t/1"])
    ws.append(["b.xlsx", None, "https://t/1"])         # blank -> skipped
    ws.append([None, None, None])                      # empty row -> skipped
    ws.append(["c.xlsx", "https://s/2", "https://t/2"])
    wb.save(path)
    assert sp.read_jobs(path) == [
        ("a.xlsx", "https://s/1", "https://t/1"),
        ("c.xlsx", "https://s/2", "https://t/2"),
    ]

    bad = os.path.join(tmp, "bad.xlsx")
    wb = openpyxl.Workbook()
    wb.active.append(["file", "from", "to"])
    wb.save(bad)
    try:
        sp.read_jobs(bad)
        raise AssertionError("missing headers should raise CopyError")
    except sp.CopyError as exc:
        assert "source" in str(exc) and "target" in str(exc)

print("ok")
