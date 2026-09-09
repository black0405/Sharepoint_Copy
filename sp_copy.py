"""
Copy files from one SharePoint folder to another.

Runs through Microsoft Graph with an Entra ID app registration (SPN):
tenant id + client id + client secret. Same-tenant copies are done
server-side by SharePoint -- the file's bytes never travel through this
machine.

    pip install msal requests openpyxl

Two ways to run it:

    python sp_copy.py jobs.xlsx     # batch: one copy per Excel row
    python sp_copy.py               # single file, from the constants below

BATCH MODE (jobs.xlsx)
    Row 1 holds the headers (any capitalisation):   file | source | target
      file    name of the file to copy, e.g. report.xlsx
      source  SharePoint link to the folder the file is in now
      target  SharePoint link to the folder it should be copied into
    Links may be copied from the browser address bar, from the "Copy link"
    button, or typed as a plain path -- all three forms work. Rows with a
    blank cell are skipped. A row that fails is reported and the run
    continues with the next one. The SRC_* / DST_* site settings below are
    ignored in batch mode (the links carry that information); credentials
    and the Behaviour settings still apply.

---------------------------------------------------------------------------
CONFIGURATION
---------------------------------------------------------------------------
Everything is configured in this file -- edit the constants in the
CONFIGURATION section below and that is the whole configuration story.
Nothing is read from environment variables.

Optionally, a "config.env" sitting next to this script overrides those
constants, which is how you run against two tenants from one folder. See
config.env.example for the key names. Blank values are ignored.
---------------------------------------------------------------------------
"""

import base64
import os
import sys
import time
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qs, quote, urlsplit

import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ===========================================================================
# CONFIGURATION
# ===========================================================================

# --- Entra ID app registration (SPN) ---------------------------------------
# The three static values your admin gave you when the service principal
# was created. Not to be confused with the Graph ACCESS TOKEN, which this
# script fetches at runtime by exchanging these three -- that one is
# short-lived and is never configured anywhere.
TENANT_ID: str = "PUT_TENANT_ID_HERE"
CLIENT_ID: str = "PUT_CLIENT_ID_HERE"
CLIENT_SECRET: str = "PUT_CLIENT_SECRET_HERE"

# --- Source ----------------------------------------------------------------
# Two ways to identify the site:
#   A) leave SITE_ID blank and fill in hostname + path; the script resolves
#      the ID for you (easier, costs one extra API call)
#   B) paste a known SITE_ID and the hostname/path are ignored
SRC_SITE_HOSTNAME: str = "yourtenant.sharepoint.com"
SRC_SITE_PATH: str = "/sites/SourceSite"
SRC_SITE_ID: str = ""

# Document library ("drive"). Leave SRC_DRIVE_ID blank to look the library
# up by name; "Documents" is the default library on most sites (it shows in
# the UI as "Shared Documents").
SRC_LIBRARY_NAME: str = "Documents"
SRC_DRIVE_ID: str = ""

# Path to the file INSIDE that library. No leading slash, no library name.
SRC_FILE_PATH: str = "Reports/report.xlsx"

# --- Destination -----------------------------------------------------------
DST_SITE_HOSTNAME: str = "yourtenant.sharepoint.com"
DST_SITE_PATH: str = "/sites/DestinationSite"
DST_SITE_ID: str = ""

DST_LIBRARY_NAME: str = "Documents"
DST_DRIVE_ID: str = ""

# Folder inside the destination library. "" copies to the library root.
# Nested paths are fine ("Archive/2026"); missing folders are created.
DST_FOLDER: str = "Archive"

# --- Behaviour -------------------------------------------------------------
# "" keeps the source filename. Set a name to rename on copy.
NEW_NAME: str = ""

# Overwrite a file of the same name if it's already at the destination.
# False makes Graph auto-rename instead (report 1.xlsx).
REPLACE_EXISTING: bool = True

# Seconds to wait for the server-side copy to finish before giving up.
COPY_TIMEOUT: int = 300

# Seconds for any single HTTP request.
REQUEST_TIMEOUT: int = 60

# Extra request/response detail on stdout. Turn on to troubleshoot.
DEBUG: bool = False

# ===========================================================================
# END OF CONFIGURATION
# ===========================================================================

GRAPH_ROOT: str = "https://graph.microsoft.com/v1.0"

# Config files looked for next to this script, in order. First one found
# wins. These are optional -- the constants above work on their own.
CONFIG_FILENAMES: Tuple[str, ...] = ("config.env", "config.ini", "sp_copy.env")


class CopyError(RuntimeError):
    """Raised when the copy can't be completed."""


# ---------------------------------------------------------------------------
# config.env OVERRIDE
# ---------------------------------------------------------------------------
# The constants above are DEFAULTS. If a config file sits next to this
# script, its values override them. Blank values are ignored, so a partly
# filled file only overrides what it actually sets.
# ---------------------------------------------------------------------------

def _parse_bool(value: str) -> bool:
    return value.strip().lower() in ("1", "true", "yes", "on")


def _load_config_file() -> Optional[str]:
    """Read a config file next to this script and override the constants."""
    path = None
    for name in CONFIG_FILENAMES:
        candidate = os.path.join(SCRIPT_DIR, name)
        if os.path.isfile(candidate):
            path = candidate
            break
    if path is None:
        return None

    values: Dict[str, str] = {}
    with open(path, "r", encoding="utf-8-sig") as handle:
        for line in handle:
            line = line.strip()
            # Skip comments, blanks, and .ini section headers.
            if not line or line.startswith("#") or line.startswith("["):
                continue
            if "=" not in line:
                continue
            key, _, raw = line.partition("=")
            values[key.strip().upper()] = raw.strip().strip('"').strip("'")

    globals_ = globals()

    def put(key: str, const: str, cast: Any = str) -> None:
        raw = values.get(key, "")
        if raw == "":            # blank means "leave the constant alone"
            return
        globals_[const] = cast(raw)

    put("SPN_TENANT_ID", "TENANT_ID")
    put("SPN_CLIENT_ID", "CLIENT_ID")
    put("SPN_CLIENT_SECRET", "CLIENT_SECRET")

    put("SRC_SITE_HOSTNAME", "SRC_SITE_HOSTNAME")
    put("SRC_SITE_PATH", "SRC_SITE_PATH")
    put("SRC_SITE_ID", "SRC_SITE_ID")
    put("SRC_LIBRARY_NAME", "SRC_LIBRARY_NAME")
    put("SRC_DRIVE_ID", "SRC_DRIVE_ID")
    put("SRC_FILE_PATH", "SRC_FILE_PATH")

    put("DST_SITE_HOSTNAME", "DST_SITE_HOSTNAME")
    put("DST_SITE_PATH", "DST_SITE_PATH")
    put("DST_SITE_ID", "DST_SITE_ID")
    put("DST_LIBRARY_NAME", "DST_LIBRARY_NAME")
    put("DST_DRIVE_ID", "DST_DRIVE_ID")
    put("DST_FOLDER", "DST_FOLDER")

    put("NEW_NAME", "NEW_NAME")
    put("REPLACE_EXISTING", "REPLACE_EXISTING", _parse_bool)
    put("COPY_TIMEOUT", "COPY_TIMEOUT", int)
    put("REQUEST_TIMEOUT", "REQUEST_TIMEOUT", int)
    put("DEBUG", "DEBUG", _parse_bool)

    return path


# ---------------------------------------------------------------------------
# Graph plumbing
# ---------------------------------------------------------------------------

def log(message: str) -> None:
    print(message, flush=True)


def debug(message: str) -> None:
    if DEBUG:
        print(f"  [debug] {message}", flush=True)


def graph_credentials() -> Tuple[str, str, str]:
    """
    Return the SPN credentials, raising a clear error if they're still the
    placeholder values rather than letting Graph fail with a vague 401.
    """
    missing = [
        name for name, value in
        (("tenant id", TENANT_ID), ("client id", CLIENT_ID),
         ("client secret", CLIENT_SECRET))
        if not value or value.startswith("PUT_")
    ]
    if missing:
        raise CopyError(
            f"Missing credential(s): {', '.join(missing)}. Set TENANT_ID / "
            f"CLIENT_ID / CLIENT_SECRET in the CONFIGURATION section of "
            f"{os.path.basename(__file__)}, or in config.env next to it."
        )
    return TENANT_ID, CLIENT_ID, CLIENT_SECRET


def get_graph_token() -> str:
    """Acquire an app-only Graph token via the client credentials flow."""
    try:
        import msal
    except ImportError:
        raise CopyError(
            "The 'msal' package is required. Install it with:  pip install msal"
        )

    tenant, client, secret = graph_credentials()
    try:
        app = msal.ConfidentialClientApplication(
            client_id=client,
            client_credential=secret,
            authority=f"https://login.microsoftonline.com/{tenant}",
        )

        # .default asks for whatever application permissions were consented.
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
    except ValueError as exc:
        # msal raises this when the authority itself can't be reached or the
        # tenant doesn't exist -- almost always a wrong TENANT_ID.
        raise CopyError(
            f"Could not reach the login authority for tenant '{tenant}'. "
            f"Check TENANT_ID (it is a GUID, or your tenant's domain name).\n"
            f"  msal said: {exc}"
        )
    except requests.RequestException as exc:
        raise CopyError(f"Network error while requesting a token: {exc}")

    if "access_token" not in result:
        raise CopyError(
            f"Token request failed: {result.get('error')} -- "
            f"{result.get('error_description', '')[:300]}"
        )
    return result["access_token"]


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {get_graph_token()}",
        "Accept": "application/json",
    })
    return session


def graph_request(session: requests.Session, method: str, url: str,
                  **kwargs: Any) -> requests.Response:
    """Call a Graph endpoint, turning failures into readable errors."""
    kwargs.setdefault("timeout", REQUEST_TIMEOUT)
    response = session.request(method, url, **kwargs)
    debug(f"{method} {url} -> {response.status_code}")

    # Both of these mean the same thing under Sites.Selected: the app
    # authenticated fine but has no permission on THIS site. SharePoint
    # reports it inconsistently -- sometimes a clean 403, sometimes a 401
    # wrapped as "generalException / spException". Treat them together.
    body = response.text or ""
    is_site_grant_issue = (
        response.status_code == 403
        or (response.status_code == 401
            and ("spException" in body or "generalException" in body))
    )
    if is_site_grant_issue:
        raise CopyError(
            f"Graph returned {response.status_code} for {url}.\n"
            f"  This is the classic Sites.Selected symptom: the app has the "
            f"API permission but has NOT been granted access to this specific "
            f"site.\n"
            f"  The API permission alone grants nothing -- an admin must run "
            f"the site-level grant once, on BOTH the source and destination "
            f"sites.\n"
            f"  Run  python {os.path.basename(sys.argv[0])} --grant-help  "
            f"for the exact calls."
        )
    if not response.ok:
        raise CopyError(
            f"Graph {method} {url} failed ({response.status_code}): "
            f"{response.text[:300]}"
        )
    return response


def graph_get(session: requests.Session, url: str) -> Dict[str, Any]:
    return graph_request(session, "GET", url).json()


def resolve_site_id(session: requests.Session, hostname: str, path: str,
                    preset: str, label: str) -> str:
    """Look up a site's Graph ID from hostname + path, unless preset."""
    if preset:
        return preset
    path = path if path.startswith("/") else "/" + path
    site = graph_get(session, f"{GRAPH_ROOT}/sites/{hostname}:{path}")
    if "id" not in site:
        raise CopyError(
            f"No site id returned for the {label} site {hostname}{path}"
        )
    debug(f"{label} site id: {site['id']}")
    return site["id"]


def resolve_drive_id(session: requests.Session, site_id: str, library: str,
                     preset: str, label: str) -> str:
    """Find a document library's driveId by display name, unless preset."""
    if preset:
        return preset
    drives = graph_get(session, f"{GRAPH_ROOT}/sites/{site_id}/drives")["value"]
    for drive in drives:
        if drive.get("name", "").lower() == library.lower():
            debug(f"{label} drive id: {drive['id']}")
            return drive["id"]
    names = ", ".join(d.get("name", "?") for d in drives) or "(none)"
    raise CopyError(
        f"{label} library '{library}' not found. Libraries on that site: {names}"
    )


def encode_path(path: str) -> str:
    """URL-encode a drive-relative path, keeping the slashes as separators."""
    return quote(path.strip("/"), safe="/")


def get_file_item(session: requests.Session, drive_id: str,
                  path: str) -> Dict[str, Any]:
    url = f"{GRAPH_ROOT}/drives/{drive_id}/root:/{encode_path(path)}"
    item = graph_get(session, url)
    if "file" not in item:
        raise CopyError(f"Source path '{path}' is a folder, not a file.")
    return item


def ensure_folder(session: requests.Session, drive_id: str,
                  folder: str) -> Dict[str, Any]:
    """
    Return the destination folder item, creating it (and any missing parent)
    if it isn't there yet.
    """
    folder = folder.strip("/")
    if not folder:
        return graph_get(session, f"{GRAPH_ROOT}/drives/{drive_id}/root")

    item = graph_get(session, f"{GRAPH_ROOT}/drives/{drive_id}/root")

    for segment in folder.split("/"):
        children = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item['id']}/children"
        existing = graph_get(session, children).get("value", [])
        match = next(
            (c for c in existing
             if c.get("name", "").lower() == segment.lower() and "folder" in c),
            None,
        )
        if match:
            item = match
            continue

        log(f"  Creating folder '{segment}'")
        item = graph_request(
            session, "POST", children,
            json={
                "name": segment,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            },
        ).json()

    return item


def wait_for_copy(monitor_url: str, timeout: int) -> Dict[str, Any]:
    """
    Graph's copy is asynchronous: it hands back a monitor URL that reports
    progress. Poll it until the copy settles.

    The monitor URL is pre-authenticated -- sending our own Authorization
    header at it makes it fail, so this uses a bare request.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        response = requests.get(monitor_url, timeout=REQUEST_TIMEOUT)
        if not response.ok:
            raise CopyError(
                f"Copy monitor failed ({response.status_code}): "
                f"{response.text[:300]}"
            )
        status = response.json()
        state = status.get("status")
        debug(f"copy status: {state} {status.get('percentageComplete', '')}")

        if state == "completed":
            return status
        if state == "failed":
            raise CopyError(f"SharePoint reported the copy failed: {status}")
        time.sleep(2)

    raise CopyError(
        f"Copy did not finish within {timeout}s. It may still complete on "
        f"the server -- check the destination folder."
    )


def copy_item(session: requests.Session, src_drive: str, item: Dict[str, Any],
              dst_drive: str, dst_folder_id: str, name: str) -> None:
    """Ask SharePoint to copy one item server-side and wait for it."""
    body: Dict[str, Any] = {
        "parentReference": {"driveId": dst_drive, "id": dst_folder_id},
        "name": name,
        "@microsoft.graph.conflictBehavior":
            "replace" if REPLACE_EXISTING else "rename",
    }
    response = graph_request(
        session, "POST",
        f"{GRAPH_ROOT}/drives/{src_drive}/items/{item['id']}/copy",
        json=body,
    )
    monitor = response.headers.get("Location")
    if monitor:
        wait_for_copy(monitor, COPY_TIMEOUT)
    else:
        # Small files sometimes complete inline with a 200/201 and no monitor.
        debug("no monitor URL returned; copy completed inline")


# ---------------------------------------------------------------------------
# SharePoint links -> drive items (batch mode)
# ---------------------------------------------------------------------------

def normalise_link(url: str) -> str:
    """
    Turn whatever SharePoint link a user pasted into one that Graph's
    /shares endpoint understands.

    Browser address-bar links (".../Forms/AllItems.aspx?id=...") point at a
    list VIEW page, not the folder; the real folder path is in the "id"
    (older: "RootFolder") query parameter. Sharing links (":f:/s/...") and
    plain paths pass through, with spaces percent-encoded.
    """
    url = url.strip()
    parts = urlsplit(url)
    params = parse_qs(parts.query)
    folder = (params.get("id") or params.get("RootFolder") or [""])[0]
    if folder:
        return f"{parts.scheme}://{parts.netloc}{quote(folder, safe='/')}"
    return url.replace(" ", "%20")


def share_token(url: str) -> str:
    """Encode a URL the way Graph's /shares/{token} endpoint expects."""
    raw = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + raw.rstrip("=")


def resolve_folder(session: requests.Session, url: str,
                   cache: Dict[str, Tuple[str, str]]) -> Tuple[str, str]:
    """
    Resolve any SharePoint folder link to (driveId, itemId). Answers are
    cached so ten rows pointing at the same folder cost one call.
    """
    if url in cache:
        return cache[url]
    token = share_token(normalise_link(url))
    try:
        item = graph_get(session, f"{GRAPH_ROOT}/shares/{token}/driveItem")
    except CopyError as exc:
        if "(404)" in str(exc) or "(400)" in str(exc):
            raise CopyError(
                f"link could not be resolved (does it open in a browser?): {url}"
            )
        raise
    if "folder" not in item:
        raise CopyError(f"link points at a file, not a folder: {url}")
    cache[url] = (item["parentReference"]["driveId"], item["id"])
    debug(f"{url} -> drive {cache[url][0]} item {cache[url][1]}")
    return cache[url]


def read_jobs(path: str) -> List[Tuple[str, str, str]]:
    """Read (file, source, target) rows from an .xlsx. Rows with blanks are skipped."""
    try:
        import openpyxl
    except ImportError:
        raise CopyError(
            "The 'openpyxl' package is required for batch mode. Install it "
            "with:  pip install openpyxl"
        )
    if not os.path.isfile(path):
        raise CopyError(f"Excel file not found: {path}")

    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = book.active.iter_rows(values_only=True)
        header = [str(h or "").strip().lower() for h in next(rows, ())]
        wanted = ("file", "source", "target")
        missing = [c for c in wanted if c not in header]
        if missing:
            found = ", ".join(h for h in header if h) or "nothing"
            raise CopyError(
                f"Excel is missing column(s): {', '.join(missing)}. Row 1 "
                f"must have the headers file, source, target (found: {found})."
            )
        idx = [header.index(c) for c in wanted]
        jobs: List[Tuple[str, str, str]] = []
        for row in rows:
            cells = tuple(
                str(row[i]).strip() if i < len(row) and row[i] is not None else ""
                for i in idx
            )
            if all(cells):
                jobs.append(cells)  # type: ignore[arg-type]
        return jobs
    finally:
        book.close()


def copy_from_excel(path: str) -> int:
    """Copy every row of the workbook. Returns the number of failed rows."""
    jobs = read_jobs(path)
    if not jobs:
        raise CopyError(
            f"No usable rows in {path} (each row needs file, source and target)."
        )
    log(f"Jobs   : {len(jobs)} row(s) from {os.path.basename(path)}")

    session = build_session()
    folders: Dict[str, Tuple[str, str]] = {}
    width = max(len(name) for name, _, _ in jobs)
    failed = 0
    for n, (name, source, target) in enumerate(jobs, 1):
        try:
            src_drive, src_folder = resolve_folder(session, source, folders)
            try:
                item = graph_get(
                    session,
                    f"{GRAPH_ROOT}/drives/{src_drive}/items/{src_folder}:/"
                    f"{encode_path(name)}",
                )
            except CopyError as exc:
                if "(404)" in str(exc):
                    raise CopyError("file not found in source folder")
                raise
            if "file" not in item:
                raise CopyError("that name is a folder, not a file")
            dst_drive, dst_folder = resolve_folder(session, target, folders)
            copy_item(session, src_drive, item, dst_drive, dst_folder, item["name"])
            log(f"[{n}/{len(jobs)}] {name:<{width}} -> OK")
        except (CopyError, requests.RequestException) as exc:
            failed += 1
            log(f"[{n}/{len(jobs)}] {name:<{width}} -> FAILED: {exc}")

    log(f"Done: {len(jobs) - failed} ok, {failed} failed")
    return failed


# ---------------------------------------------------------------------------
# Single-file mode (constants / config.env)
# ---------------------------------------------------------------------------

def copy_file() -> None:
    session = build_session()

    src_site = resolve_site_id(
        session, SRC_SITE_HOSTNAME, SRC_SITE_PATH, SRC_SITE_ID, "source")
    src_drive = resolve_drive_id(
        session, src_site, SRC_LIBRARY_NAME, SRC_DRIVE_ID, "Source")

    dst_site = resolve_site_id(
        session, DST_SITE_HOSTNAME, DST_SITE_PATH, DST_SITE_ID, "destination")
    dst_drive = resolve_drive_id(
        session, dst_site, DST_LIBRARY_NAME, DST_DRIVE_ID, "Destination")

    item = get_file_item(session, src_drive, SRC_FILE_PATH)
    size_mb = item.get("size", 0) / (1024 * 1024)
    log(f"Source : {SRC_LIBRARY_NAME}/{SRC_FILE_PATH}  ({size_mb:.2f} MB)")

    folder = ensure_folder(session, dst_drive, DST_FOLDER)
    name = NEW_NAME or item["name"]
    target = "/".join(
        p for p in (DST_LIBRARY_NAME, DST_FOLDER.strip("/"), name) if p
    )
    log(f"Target : {target}")
    log("Copying (server-side)...")
    copy_item(session, src_drive, item, dst_drive, folder["id"], name)
    log("Done.")


def print_grant_help() -> None:
    """
    Print the one-time site-permission grant an admin needs to run, with the
    app's own IDs filled in. This is what fixes the 401/403 spException.
    """
    sites = [
        ("SOURCE", SRC_SITE_HOSTNAME, SRC_SITE_PATH, "read"),
        ("DESTINATION", DST_SITE_HOSTNAME, DST_SITE_PATH, "write"),
    ]
    print("=" * 70)
    print("ONE-TIME SharePoint site grants (run by a SharePoint / Graph admin)")
    print("=" * 70)
    print()
    print("Sites.Selected grants NO access until the app is granted rights on")
    print("each specific site. The copy touches two sites, so grant both.")
    print("(Skip all of this if the app has Sites.ReadWrite.All consented.)")
    print()
    for label, hostname, path, role in sites:
        path = path if path.startswith("/") else "/" + path
        print(f"--- {label} site ({role}) ---")
        print(f"  STEP 1  GET {GRAPH_ROOT}/sites/{hostname}:{path}")
        print('          -> copy the "id" from the response')
        print(f"  STEP 2  POST {GRAPH_ROOT}/sites/{{site-id}}/permissions")
        print("          Content-Type: application/json")
        print()
        print("  {")
        print(f'    "roles": ["{role}"],')
        print('    "grantedToIdentities": [{')
        print('      "application": {')
        print(f'        "id": "{CLIENT_ID}",')
        print('        "displayName": "SharePoint File Copy"')
        print("      }")
        print("    }]")
        print("  }")
        print()
    print("Both steps can be run from Graph Explorer (aka.ms/ge) signed in as")
    print("an admin, or via PowerShell (Get-MgSite / New-MgSitePermission).")
    print("=" * 70)


def main() -> int:
    config_path = _load_config_file()

    if "--grant-help" in sys.argv:
        print_grant_help()
        return 0

    if config_path:
        log(f"Config : {os.path.basename(config_path)} "
            f"(overriding script defaults)")
    else:
        log(f"Config : {os.path.basename(__file__)} (no config file found)")

    xlsx = next((a for a in sys.argv[1:] if a.lower().endswith(".xlsx")), None)
    try:
        if xlsx:
            return 1 if copy_from_excel(xlsx) else 0
        copy_file()
    except CopyError as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    except requests.RequestException as exc:
        print(f"\nERROR: network problem talking to Graph: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        return 130
    return 0


if __name__ == "__main__":
    sys.exit(main())