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

# --- resolve_folder / folder_listing (fake Graph, no network) ---------------
class FakeResponse:
    def __init__(self, body, status=200):
        self.status_code, self.ok, self.text = status, status < 400, str(body)
        self.json = lambda: body


class FakeSession:
    """URL fragment -> JSON body. Anything unmatched is a 404."""
    def __init__(self, routes):
        self.routes = routes

    def request(self, method, url, **kwargs):
        for fragment, body in self.routes.items():
            if fragment in url:
                return FakeResponse(body)
        return FakeResponse({"error": "not found"}, 404)


site = ("t.sharepoint.com", "/sites/Fin", "")
fake = FakeSession({
    "/sites/t.sharepoint.com:/sites/Fin": {"id": "site1"},
    "/sites/site1/drives": {"value": [{"name": "Documents", "id": "drv1"}]},
    "/drives/drv1/root:/Reports/2026": {
        "id": "fld1", "folder": {}, "parentReference": {"driveId": "drv1"}},
    "/shares/": {
        "id": "fld2", "folder": {}, "parentReference": {"driveId": "drv2"}},
    "/items/fld2?$expand=children": {
        "webUrl": "https://t/x", "children": [{"name": "b.xlsx"}, {"name": "a.xlsx"}]},
})
cache = {}
# plain path -> configured site, library by name, folder by path
assert sp.resolve_folder(fake, "Documents/Reports/2026", cache, site) == ("drv1", "fld1")
# full link -> /shares, untouched by the site tuple
link = "https://t.sharepoint.com/sites/Fin/Shared Documents/X"
assert sp.resolve_folder(fake, link, cache, site) == ("drv2", "fld2")
assert len(cache) == 2
try:
    sp.resolve_folder(fake, "Documents/Nope", {}, site)
    raise AssertionError("missing plain-path folder should raise CopyError")
except sp.CopyError as exc:
    assert "Nope" in str(exc) and "Documents" in str(exc)

listing = sp.folder_listing(fake, "drv2", "fld2")
assert "https://t/x" in listing
assert listing.index("a.xlsx") < listing.index("b.xlsx")   # sorted

print("ok")
