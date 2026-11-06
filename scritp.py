#!/usr/bin/env python3
"""
UCMDB 25.4 - Installed Software Authorization & Compliance Report (READ ONLY - no CI is ever modified)

*** LARGE-SCALE EDITION ***

Built for millions of installed-software CIs:
  * Streams every server chunk: processed, written to CSV and released - memory stays flat no matter how many records exist.
  * Fetches chunks in parallel (WORKERS) with retry/backoff and automatic re-authentication if the token expires during a long run.
  * Full detail goes to CSV part files (max CSV_ROWS_PER_FILE rows each, so every part opens in Excel). The Excel workbook holds the summaries.
  * Reports are ALWAYS written, even if the run fails or is interrupted (they are then marked PARTIAL).

Report columns (per computer, per installed software):
    Hostname | IP Address | Serial | Software Name | Version | Publisher | Software Type | Subscription Type | Deployment Environment | Authorization | Compliance | Reason

TQL REQUIREMENTS (software_version_compliance):
    Node  --(Composition)-->  InstalledSoftware      (required)
    Node  --(Containment)-->  IpAddress              (optional, for IPs)

Layout (attributes returned):
    Node              : name, primary_dns_name, serial_number
    InstalledSoftware : name, display_label, version, deployment_env,
                        publisher (optional)
    IpAddress         : name

Pagination root stays "InstalledSoftware". Each chunk must contain the Node/IP CIs and relations of the software in that chunk (normal behaviour).

Requires: requests, urllib3   (optional: openpyxl for the .xlsx summary)
"""

import csv
import json
import os
import re
import sys
import threading
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from getpass import getpass
from pathlib import Path
from urllib.parse import quote

import requests
import urllib3

try:
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    HAVE_OPENPYXL = True
except ImportError:
    HAVE_OPENPYXL = False


# ============================================================
# CONFIGURATION
# ============================================================

SERVER_URL = "https://localhost.localdomain.com:8443"
API_PATH = "/rest-api"

USERNAME = "admin"
PASSWORD_ENV_VAR = "UCMDB_PASSWORD"   # optional: set to run unattended
CLIENT_CONTEXT = 1

TQL_NAME = "software_version_compliance"
PAGINATION_ROOT = "InstalledSoftware"   # TQL node label, not CI type

# ---- performance (tune for your server) ----------------------------------
# 100 means ~44,000 requests for 4.4M records. Start with 2000, and raise to
# 5000 if the server copes (check memory/response time on the UCMDB server).
CHUNK_SIZE = 2000
WORKERS = 4                 # parallel chunk downloads (2-8 is sensible)
TIMEOUT = (15, 600)         # (connect, read) seconds - big chunks are slow
MAX_RETRIES = 4             # retries for timeouts / 429 / 5xx
RETRY_BACKOFF = 5           # seconds, multiplied by attempt number
PROGRESS_EVERY = 1          # print progress every N chunks

# ---- output size controls -------------------------------------------------
CSV_ROWS_PER_FILE = 1_000_000   # detail CSV part size (Excel limit 1,048,576)
EXCEL_MAX_ROWS = 200_000        # row cap per Excel sheet (full data is in CSV)
UNRESOLVED_SAVE_LIMIT = 100_000 # unresolved records kept (all are counted)
SAVE_RAW_CHUNKS = 1             # raw JSON chunks saved: 0=none, N=first N,
                                # None=all (NOT advised for millions)

VERIFY_SSL = False                      # production: True or CA bundle path
HIDE_CERTIFICATE_WARNINGS = True

SOFTWARE_CI_TYPE = "installed_software"
IP_CI_TYPE = "ip_address"

# Attribute names tried (in order) on each CI type.
SOFTWARE_NAME_ATTRS = ("display_label", "name")
SOFTWARE_VERSION_ATTRS = ("version", "software_version")
SOFTWARE_PUBLISHER_ATTRS = ("publisher", "vendor")
DEPLOYMENT_ENV_ATTRS = ("deployment_env",)
ENV_NOT_SET = "Not Set"
HOST_NAME_ATTRS = ("name", "primary_dns_name", "display_label")
SERIAL_ATTRS = ("serial_number", "serial")
IP_ATTRS_ON_IP_CI = ("name", "ip_address", "display_label")
IP_ATTRS_ON_NODE = ("primary_ip_address", "ip_address")

EXCLUDE_IPV6 = True
EXCLUDE_IP_PREFIXES = ("127.", "169.254.")

# Software that is NOT in the catalog:
#   REPORT_UNLISTED_SOFTWARE = False -> ignored completely.
#   REPORT_UNLISTED_SOFTWARE = True  -> counted as Unauthorized/Non-Compliant
#       (always shown in the Software Summary and device counts).
#   WRITE_UNLISTED_DETAIL    = True  -> ALSO written row-by-row to the detail
#       CSV. With millions of records leave this False: the detail then holds
#       only catalog software, and unlisted software is still counted.
REPORT_UNLISTED_SOFTWARE = True
WRITE_UNLISTED_DETAIL = False

# Case-insensitive fragments; matching software is skipped entirely.
IGNORE_NAME_PATTERNS = []

UNCLASSIFIED = "Unclassified"

# ------------------------------------------------------------
# SOFTWARE CATALOG  (edit to match the customer's contracts)
#   name / aliases    : matched case-insensitively, exact or by prefix
#   approved_versions : exact version strings => "Authorized"
#   min_version       : installed >= this => "Compliant"
#   software_type     : Freeware / Commercial / Licensed
#   subscription_type : Freeware / Perpetual / Annual Subscription / ...
# ------------------------------------------------------------
SOFTWARE_CATALOG = [
    {
        "name": "WPS Office",
        "aliases": [],
        "approved_versions": ["10,2,0,7549"],
        "min_version": "10,2,0,7549",
        "software_type": "Commercial",
        "subscription_type": "Annual Subscription",
    },
    {
        "name": "Microsoft Edge",
        "aliases": [],
        "approved_versions": ["153.0.4234.48"],
        "min_version": "153.0.4234.48",
        "software_type": "Freeware",
        "subscription_type": "Freeware",
    },
    {
        "name": "7-Zip",
        "aliases": [],
        "approved_versions": ["26.01"],
        "min_version": "26.01",
        "software_type": "Freeware",
        "subscription_type": "Freeware",
    },
    {
        "name": "Universal Discovery Agent",
        "aliases": [],
        "approved_versions": ["25.4.000.98"],
        "min_version": "25.4.000.98",
        "software_type": "Licensed",
        "subscription_type": "Perpetual",
    },
]

API_BASE = SERVER_URL.rstrip("/") + "/" + API_PATH.strip("/")
OUTPUT_DIR = Path(__file__).resolve().parent
RAW_DIR = OUTPUT_DIR / "raw_chunks"

if VERIFY_SSL is False and HIDE_CERTIFICATE_WARNINGS:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================
# FILE HELPERS
# ============================================================

def save_json(filename, data, folder=None):
    folder = folder or OUTPUT_DIR
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / filename
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return path


def write_csv(path, headers, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        writer.writerows(rows)


class RotatingCsv:
    """Streams rows to CSV part files of at most max_rows rows each."""

    def __init__(self, base_name, headers, max_rows):
        self.base_name = base_name
        self.headers = headers
        self.max_rows = max_rows
        self.files = []
        self.total = 0
        self._rows = 0
        self._fh = None
        self._writer = None

    def _open_next(self):
        self.close_current()
        path = OUTPUT_DIR / f"{self.base_name}_part{len(self.files) + 1:03d}.csv"
        self._fh = open(path, "w", newline="", encoding="utf-8-sig")
        self._writer = csv.writer(self._fh)
        self._writer.writerow(self.headers)
        self.files.append(path)
        self._rows = 0

    def write(self, row):
        if self._fh is None or self._rows >= self.max_rows:
            self._open_next()
        self._writer.writerow(row)
        self._rows += 1
        self.total += 1

    def flush(self):
        if self._fh:
            self._fh.flush()

    def close_current(self):
        if self._fh:
            self._fh.close()
            self._fh = None

    def close(self):
        if not self.files:
            self._open_next()          # header-only file if nothing written
        self.close_current()


# ============================================================
# HTTP HELPERS (retry + automatic re-authentication)
# ============================================================

_credentials = {}
_auth_lock = threading.Lock()


def request_json(session, method, path, quiet=False, _reauth=True, **kwargs):
    endpoint = API_BASE + "/" + path.lstrip("/")
    label = (
        "topology/result/[result-id]/[chunk]"
        if path.startswith("topology/result/")
        else path
    )
    attempt = 0

    while True:
        attempt += 1
        try:
            response = session.request(
                method=method, url=endpoint, timeout=TIMEOUT,
                verify=VERIFY_SSL, allow_redirects=False, **kwargs
            )
        except requests.exceptions.SSLError as exc:
            raise RuntimeError(
                "TLS certificate verification failed. "
                "Configure a trusted CA bundle."
            ) from exc
        except requests.exceptions.RequestException as exc:
            if attempt <= MAX_RETRIES:
                wait = RETRY_BACKOFF * attempt
                print(f"  {label}: {type(exc).__name__} - retry "
                      f"{attempt}/{MAX_RETRIES} in {wait}s")
                time.sleep(wait)
                continue
            raise RuntimeError(f"{method} {label}: {type(exc).__name__}") from exc

        if not quiet:
            print(f"{method} {label} -> HTTP {response.status_code}")

        # Token expired during a long run -> log in again, retry once.
        if response.status_code == 401 and _reauth and path != "authenticate":
            reauthenticate(session)
            _reauth = False
            continue

        if response.status_code in (429, 500, 502, 503, 504) \
                and attempt <= MAX_RETRIES:
            wait = RETRY_BACKOFF * attempt
            print(f"  {label}: HTTP {response.status_code} - retry "
                  f"{attempt}/{MAX_RETRIES} in {wait}s")
            time.sleep(wait)
            continue

        if not 200 <= response.status_code < 300:
            raise RuntimeError(
                f"{method} {label} failed with HTTP {response.status_code}. "
                "Check server logs, API permissions, and the pagination "
                "node label."
            )

        try:
            return json.loads(response.content.decode("utf-8-sig"))
        except (ValueError, UnicodeError) as exc:
            raise RuntimeError(f"{label}: response is not valid JSON.") from exc


def authenticate(session, password=None):
    if password is None:
        password = os.environ.get(PASSWORD_ENV_VAR) or getpass(
            "Enter UCMDB password: ")
    if not password:
        raise RuntimeError("Password cannot be empty.")

    result = request_json(
        session, "POST", "authenticate", _reauth=False,
        json={
            "username": USERNAME,
            "password": password,
            "clientContext": CLIENT_CONTEXT,
        },
    )
    token = result.get("token") if isinstance(result, dict) else None
    if not isinstance(token, str) or not token.strip():
        raise RuntimeError("Authentication returned no valid token.")

    session.headers["Authorization"] = f"Bearer {token.strip()}"
    _credentials["password"] = password     # kept in memory only, for re-login
    print("Authentication successful. Token is not displayed.")


def reauthenticate(session):
    with _auth_lock:
        print("  Token rejected (HTTP 401) - re-authenticating...")
        authenticate(session, _credentials.get("password"))


# ============================================================
# ATTRIBUTE READING
# ============================================================

def extract_attributes(ci):
    """Read explicitly named attributes (properties / attributeDefinitions)."""
    attributes = {}

    properties = ci.get("properties")
    if isinstance(properties, dict):
        attributes.update(properties)

    definitions = ci.get("attributeDefinitions")

    def read_value(entry):
        if isinstance(entry, dict):
            for key in ("rawValue", "value"):
                if key in entry and entry[key] is not None:
                    return True, entry[key]
            return False, None
        if entry is None or isinstance(entry, (str, int, float, bool)):
            return True, entry
        return False, None

    def add_attribute(name, entry):
        if not isinstance(name, str) or not name:
            return
        present, value = read_value(entry)
        if not present:
            return
        if name not in attributes or attributes[name] in (None, ""):
            attributes[name] = value

    if isinstance(definitions, list):
        for entry in definitions:
            if isinstance(entry, dict):
                add_attribute(entry.get("name") or entry.get("attributeName"), entry)
    elif isinstance(definitions, dict):
        for name, entry in definitions.items():
            add_attribute(name, entry)

    return attributes


def first_text(attrs, names):
    for name in names:
        value = attrs.get(name)
        if value is None:
            continue
        value = str(value).strip()
        if value:
            return value
    return ""


# ============================================================
# VERSION + CATALOG LOGIC
# ============================================================

def parse_version(text):
    if not isinstance(text, str):
        return None
    parts = re.findall(r"\d+", text)
    return tuple(int(p) for p in parts) if parts else None


def version_gte(installed, minimum):
    a, b = parse_version(installed), parse_version(minimum)
    if a is None or b is None:
        return None
    size = max(len(a), len(b))
    a += (0,) * (size - len(a))
    b += (0,) * (size - len(b))
    return a >= b


def normalise_version(text):
    return re.sub(r"\s+", "", text or "").lower()


def build_catalog_index():
    index = []
    for entry in SOFTWARE_CATALOG:
        keys = [entry["name"]] + list(entry.get("aliases", []))
        index.append((sorted((k.lower() for k in keys), key=len, reverse=True), entry))
    return index


def find_catalog_entry(name, index):
    lowered = name.lower().strip()
    for keys, entry in index:
        if lowered in keys:
            return entry
    best, best_len = None, 0
    for keys, entry in index:
        for key in keys:
            if lowered.startswith(key) and len(key) > best_len:
                best, best_len = entry, len(key)
    return best


def evaluate(version, entry):
    if entry is None:
        return {
            "authorization": "Unauthorized",
            "compliance": "Non-Compliant",
            "software_type": UNCLASSIFIED,
            "subscription_type": UNCLASSIFIED,
            "reason": "Software not in approved catalog",
        }

    base = {
        "software_type": entry.get("software_type", UNCLASSIFIED),
        "subscription_type": entry.get("subscription_type", UNCLASSIFIED),
    }

    if not version:
        return {**base,
                "authorization": "Unauthorized",
                "compliance": "Non-Compliant",
                "reason": "Version missing - cannot verify"}

    approved = {normalise_version(v) for v in entry.get("approved_versions", [])}
    authorized = normalise_version(version) in approved

    min_version = entry.get("min_version")
    if min_version:
        gte = version_gte(version, min_version)
        compliant = bool(gte)
        comp_note = (
            "version cannot be compared" if gte is None
            else f"version {'>=' if gte else '<'} minimum {min_version}"
        )
    else:
        compliant = authorized
        comp_note = "no min_version defined; follows authorization"

    auth_note = (
        "approved version" if authorized
        else f"version not in approved list "
             f"({', '.join(entry.get('approved_versions', []))})"
    )

    return {
        **base,
        "authorization": "Authorized" if authorized else "Unauthorized",
        "compliance": "Compliant" if compliant else "Non-Compliant",
        "reason": f"{auth_note}; {comp_note}",
    }


def clean_ip(value):
    if not value:
        return None
    value = value.strip()
    if EXCLUDE_IPV6 and ":" in value:
        return None
    if value.startswith(EXCLUDE_IP_PREFIXES):
        return None
    return value


# ============================================================
# REPORT HEADERS
# ============================================================

DETAIL_HEADERS = [
    "Hostname", "IP Address", "Serial", "Software Name", "Version",
    "Publisher", "Software Type", "Subscription Type",
    "Deployment Environment", "Authorization", "Compliance", "Reason",
]
SUMMARY_HEADERS = [
    "Hostname", "IP Address", "Serial", "Deployment Environment",
    "Total Software", "Authorized", "Unauthorized", "Compliant",
    "Non-Compliant", "Computer Status",
]
ENV_SUMMARY_HEADERS = [
    "Deployment Environment", "Devices", "Software Records",
    "Authorized", "Unauthorized", "Compliant", "Non-Compliant",
    "Fully Compliant Devices", "Devices With Issues",
]
SOFTWARE_SUMMARY_HEADERS = [
    "Software Name", "Version", "Software Type", "Subscription Type",
    "Deployment Environment", "Installations", "Authorization",
    "Compliance", "Reason",
]
UNRESOLVED_HEADERS = ["ucmdbId", "Name", "Version",
                      "Deployment Environment", "Reason"]

 
# ============================================================
# STREAMING COLLECTOR
# ============================================================

class Collector:
    """Consumes one server chunk at a time. Only small aggregates are kept in
    memory: one entry per computer and one per distinct software/version."""

    def __init__(self, stamp):
        self.index = build_catalog_index()
        self.ignore = [p.lower() for p in IGNORE_NAME_PATTERNS]
        self.eval_cache = {}                       # (name, version) -> tuple
        self.nodes = {}                            # host id -> attrs
        self.host_env = defaultdict(lambda: defaultdict(lambda: [0] * 5))
        self.sw_summary = {}                       # (name, ver, env) -> [count, result]
        self.counts = Counter()
        self.unresolved = []
        self.chunks_done = 0
        self.records_seen = 0
        self.detail = RotatingCsv(
            f"software_compliance_detail_{stamp}", DETAIL_HEADERS, CSV_ROWS_PER_FILE)

    def _unresolved(self, record):
        self.counts["unresolved"] += 1
        if len(self.unresolved) < UNRESOLVED_SAVE_LIMIT:
            self.unresolved.append(record)

    def _classify(self, name, version):
        key = (name, version)
        hit = self.eval_cache.get(key)
        if hit is None:
            entry = find_catalog_entry(name, self.index)
            hit = (entry, evaluate(version, entry))
            self.eval_cache[key] = hit
        return hit

    def process_chunk(self, data):
        records = data.get("cis")
        if not isinstance(records, list):
            raise RuntimeError(
                "Topology result has no valid cis list. "
                f"Returned keys: {list(data.keys())}"
            )

        # merge CIs of THIS chunk only
        cis = {}
        for ci in records:
            if not isinstance(ci, dict) or not ci.get("ucmdbId"):
                raise RuntimeError("A topology record is missing ucmdbId.")
            ci_id = str(ci["ucmdbId"])
            parsed = extract_attributes(ci)
            if ci_id not in cis:
                cis[ci_id] = {"type": ci.get("type"), "attributes": parsed}
            else:
                existing = cis[ci_id]["attributes"]
                for key, value in parsed.items():
                    if key not in existing or value not in (None, ""):
                        existing[key] = value

        software_ids = [i for i, c in cis.items()
                        if c["type"] == SOFTWARE_CI_TYPE]
        sw_set = set(software_ids)
        ip_set = {i for i, c in cis.items() if c["type"] == IP_CI_TYPE}

        # link software -> computer and computer -> IPs (this chunk)
        sw_to_host = {}
        host_to_ips = defaultdict(set)
        for rel in data.get("relations") or []:
            if not isinstance(rel, dict):
                continue
            e1, e2 = rel.get("end1Id"), rel.get("end2Id")
            if not e1 or not e2:
                continue
            e1, e2 = str(e1), str(e2)
            for a, b in ((e1, e2), (e2, e1)):
                if a in cis and a not in sw_set and a not in ip_set:
                    if b in sw_set:
                        sw_to_host.setdefault(b, a)
                    elif b in ip_set:
                        host_to_ips[a].add(b)

        # update computer cache (non-empty values win)
        for host_id in set(sw_to_host.values()) | set(host_to_ips):
            attrs = cis[host_id]["attributes"]
            node = self.nodes.setdefault(
                host_id, {"hostname": "", "serial": "", "ips": set()})
            node["hostname"] = first_text(attrs, HOST_NAME_ATTRS) or node["hostname"]
            node["serial"] = first_text(attrs, SERIAL_ATTRS) or node["serial"]
            for ip_id in host_to_ips.get(host_id, ()):
                ip = clean_ip(first_text(cis[ip_id]["attributes"], IP_ATTRS_ON_IP_CI))
                if ip:
                    node["ips"].add(ip)
            if not node["ips"]:
                ip = clean_ip(first_text(attrs, IP_ATTRS_ON_NODE))
                if ip:
                    node["ips"].add(ip)

        for sw_id in software_ids:
            self.records_seen += 1
            attrs = cis[sw_id]["attributes"]
            name = first_text(attrs, SOFTWARE_NAME_ATTRS)
            version = first_text(attrs, SOFTWARE_VERSION_ATTRS)
            publisher = first_text(attrs, SOFTWARE_PUBLISHER_ATTRS)
            env = first_text(attrs, DEPLOYMENT_ENV_ATTRS) or ENV_NOT_SET

            if not name:
                self._unresolved({
                    "ucmdbId": sw_id, "name": "", "version": version,
                    "deployment_env": env, "reason": "software name missing"})
                continue

            lowered = name.lower()
            if self.ignore and any(p in lowered for p in self.ignore):
                self.counts["ignored"] += 1
                continue

            host_id = sw_to_host.get(sw_id)
            if not host_id:
                self._unresolved({
                    "ucmdbId": sw_id, "name": name, "version": version,
                    "deployment_env": env,
                    "reason": "no computer relation found in query result"})
                continue

            entry, result = self._classify(name, version)
            if entry is None and not REPORT_UNLISTED_SOFTWARE:
                self.counts["unlisted_skipped"] += 1
                continue

            # aggregates
            stats = self.host_env[host_id][env]
            stats[0] += 1
            is_auth = result["authorization"] == "Authorized"
            is_comp = result["compliance"] == "Compliant"
            stats[1] += is_auth
            stats[2] += not is_auth
            stats[3] += is_comp
            stats[4] += not is_comp

            key = (name, version, env)
            row = self.sw_summary.get(key)
            if row is None:
                self.sw_summary[key] = [1, result]
            else:
                row[0] += 1

            self.counts[result["authorization"].lower()] += 1
            self.counts[result["compliance"].lower().replace("-", "_")] += 1
            if entry is None:
                self.counts["unlisted"] += 1

            # detail row (streamed to disk, nothing kept in memory)
            if entry is not None or WRITE_UNLISTED_DETAIL:
                node = self.nodes[host_id]
                self.detail.write([
                    node["hostname"] or f"<unnamed {host_id[:8]}>",
                    ", ".join(sorted(node["ips"])), node["serial"],
                    name, version, publisher,
                    result["software_type"], result["subscription_type"],
                    env, result["authorization"], result["compliance"],
                    result["reason"],
                ])

        self.chunks_done += 1

    # ---------------- final tables ----------------

    def device_rows(self):
        rows = []
        for host_id, envs in self.host_env.items():
            node = self.nodes.get(host_id, {"hostname": "", "serial": "", "ips": set()})
            hostname = node["hostname"] or f"<unnamed {host_id[:8]}>"
            for env, (total, auth, unauth, comp, noncomp) in envs.items():
                status = "Compliant" if unauth == 0 and noncomp == 0 else "Non-Compliant"
                rows.append([hostname, ", ".join(sorted(node["ips"])),
                             node["serial"], env, total, auth, unauth,
                             comp, noncomp, status, host_id])
        rows.sort(key=lambda r: (r[0].lower(), r[3]))
        return rows

    @staticmethod
    def env_rows(device_rows):
        stats = defaultdict(lambda: [0] * 8)
        for r in device_rows:
            s = stats[r[3]]
            s[0] += 1                        # devices
            s[1] += r[4]                     # software records
            s[2] += r[5]
            s[3] += r[6]
            s[4] += r[7]
            s[5] += r[8]
            ok = r[6] == 0 and r[8] == 0
            s[6] += ok
            s[7] += not ok
        rows = [[env] + s for env, s in sorted(stats.items())]
        if rows:
            distinct = len({r[10] for r in device_rows})
            rows.append(["TOTAL (distinct devices)", distinct]
                        + [sum(r[i] for r in rows) for i in range(2, 9)])
        return rows

    def software_rows(self):
        rows = []
        for (name, version, env), (count, res) in self.sw_summary.items():
            rows.append([name, version, res["software_type"],
                         res["subscription_type"], env, count,
                         res["authorization"], res["compliance"], res["reason"]])
        rows.sort(key=lambda r: (-r[5], r[0].lower(), r[1]))
        return rows

    def unresolved_rows(self):
        return [[u.get("ucmdbId"), u.get("name", ""), u.get("version", ""),
                 u.get("deployment_env", ""), u.get("reason", "")]
                for u in self.unresolved]


# ============================================================
# TOPOLOGY FETCH (parallel chunks)
# ============================================================

def iter_chunks(session):
    """Yields (number, total, data) for each server chunk, in order."""
    result = request_json(
        session, "GET", "topology",
        params={
            "queryName": TQL_NAME,
            "chunkSize": CHUNK_SIZE,
            "rootNodeForPagination": PAGINATION_ROOT,
            "includeEmptyLayoutProperties": "true",
        },
    )
    if not isinstance(result, dict):
        raise RuntimeError("Topology response must be a JSON object.")

    if "numberOfChunks" not in result:
        print(f"\nServer returned a direct result (chunk size {CHUNK_SIZE}).")
        yield 1, 1, result
        return

    try:
        chunk_count = int(result["numberOfChunks"])
    except (TypeError, ValueError) as exc:
        raise RuntimeError("Invalid numberOfChunks.") from exc
    if chunk_count < 0:
        raise RuntimeError("numberOfChunks cannot be negative.")
    if chunk_count == 0:
        if isinstance(result.get("cis"), list):
            yield 1, 1, result
        return

    result_id = result.get("queryResultId")
    if not result_id:
        raise RuntimeError("Chunked response has no queryResultId.")
    encoded_id = quote(str(result_id), safe="")

    print(f"\nServer chunks: {chunk_count} | Chunk size: {CHUNK_SIZE} | "
          f"Parallel workers: {WORKERS}")

    def fetch(number):
        return request_json(
            session, "GET", f"topology/result/{encoded_id}/{number}",
            quiet=True, params={"includeEmptyLayoutProperties": "true"})

    window = max(WORKERS * 2, 1)
    with ThreadPoolExecutor(max_workers=max(WORKERS, 1)) as pool:
        for start in range(1, chunk_count + 1, window):
            numbers = range(start, min(start + window, chunk_count + 1))
            futures = [(n, pool.submit(fetch, n)) for n in numbers]
            for n, future in futures:
                yield n, chunk_count, future.result()


# ============================================================
# EXCEL SUMMARY WORKBOOK
# ============================================================

def write_excel(path, collector, device_rows, env_rows, software_rows,
                unresolved_rows, partial):
    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F3864")
    green = PatternFill("solid", fgColor="C6EFCE", bgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE", bgColor="FFC7CE")
    notes = []

    def sheet(ws, headers, rows, color_headers=(), bold_last=False, label=""):
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center")
        limit = EXCEL_MAX_ROWS
        for i, row in enumerate(rows):
            if i >= limit:
                notes.append(f"{label or ws.title}: showing first {limit:,} of "
                             f"{len(rows):,} rows - full data is in the CSV files.")
                break
            ws.append(row[:len(headers)])
        last = ws.max_row
        for header in color_headers:
            letter = get_column_letter(headers.index(header) + 1)
            rng = f"{letter}2:{letter}{max(last, 2)}"
            for text in ("Authorized", "Compliant"):
                ws.conditional_formatting.add(
                    rng, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=green))
            for text in ("Unauthorized", "Non-Compliant"):
                ws.conditional_formatting.add(
                    rng, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=red))
        if bold_last and last > 1:
            for cell in ws[last]:
                cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        if last > 1:
            ws.auto_filter.ref = ws.dimensions
        for idx, header in enumerate(headers, 1):
            longest = max(
                [len(str(header))] +
                [len(str(ws.cell(row=r, column=idx).value or ""))
                 for r in range(2, min(last, 300) + 1)])
            ws.column_dimensions[get_column_letter(idx)].width = min(longest + 3, 60)

    ws = wb.active
    ws.title = "Environment Summary"
    sheet(ws, ENV_SUMMARY_HEADERS, env_rows, bold_last=True)
    sheet(wb.create_sheet("Computer Summary"), SUMMARY_HEADERS, device_rows,
          color_headers=("Computer Status",), label="Computer Summary")
    sheet(wb.create_sheet("Software Summary"), SOFTWARE_SUMMARY_HEADERS,
          software_rows, color_headers=("Authorization", "Compliance"),
          label="Software Summary")
    sheet(wb.create_sheet("Catalog"),
          ["Name", "Aliases", "Approved Versions", "Min Version",
           "Software Type", "Subscription Type"],
          [[e["name"], ", ".join(e.get("aliases", [])),
            ", ".join(e.get("approved_versions", [])), e.get("min_version", ""),
            e.get("software_type", ""), e.get("subscription_type", "")]
           for e in SOFTWARE_CATALOG])
    sheet(wb.create_sheet("Unresolved"), UNRESOLVED_HEADERS, unresolved_rows,
          label="Unresolved")

    info = wb.create_sheet("Run Info", 0)
    info.append(["Item", "Value"])
    info.append(["Status", "PARTIAL - run did not finish" if partial else "Complete"])
    info.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    info.append(["Server chunks processed", collector.chunks_done])
    info.append(["Installed software records seen", collector.records_seen])
    info.append(["Detail rows written to CSV", collector.detail.total])
    for p in collector.detail.files:
        info.append(["Detail CSV file", p.name])
    for n in notes:
        info.append(["Note", n])
    info.column_dimensions["A"].width = 36
    info.column_dimensions["B"].width = 90
    for cell in info[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    wb.save(path)


# ============================================================
# FINALIZE (always runs - even after a failure)
# ============================================================

def finalize(collector, stamp, partial):
    outputs = []

    collector.detail.close()
    outputs.extend(collector.detail.files)

    device_rows = collector.device_rows()
    env_rows = collector.env_rows(device_rows)
    software_rows = collector.software_rows()
    unresolved_rows = collector.unresolved_rows()
    tag = "_PARTIAL" if partial else ""

    jobs = [
        (f"computer_summary_{stamp}{tag}.csv", SUMMARY_HEADERS,
         [r[:len(SUMMARY_HEADERS)] for r in device_rows]),
        (f"environment_summary_{stamp}{tag}.csv", ENV_SUMMARY_HEADERS, env_rows),
        (f"software_summary_{stamp}{tag}.csv", SOFTWARE_SUMMARY_HEADERS, software_rows),
        (f"software_unresolved_{stamp}{tag}.csv", UNRESOLVED_HEADERS, unresolved_rows),
    ]
    for filename, headers, rows in jobs:
        path = OUTPUT_DIR / filename
        write_csv(path, headers, rows)
        outputs.append(path)

    if HAVE_OPENPYXL:
        path = OUTPUT_DIR / f"software_compliance_summary_{stamp}{tag}.xlsx"
        try:
            write_excel(path, collector, device_rows, env_rows,
                        software_rows, unresolved_rows, partial)
            outputs.append(path)
        except Exception as exc:     # CSV files above are already complete
            print(f"\nWARNING: Excel summary could not be written "
                  f"({type(exc).__name__}: {exc}). CSV reports are complete.")
    else:
        print("\nNOTE: openpyxl not installed - .xlsx skipped. "
              "Install with: pip install openpyxl")

    return outputs, env_rows


# ============================================================
# MAIN: READ ONLY
# ============================================================

def main():
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    started = time.time()

    print("=" * 70)
    print("UCMDB InstalledSoftware Inventory Report Generator (LARGE SCALE)")
    print("=" * 70)
    print(f"TQL: {TQL_NAME} | Pagination root: {PAGINATION_ROOT}")
    print(f"Chunk size: {CHUNK_SIZE} | Workers: {WORKERS}")
    print(f"Write unlisted software rows to detail CSV: {WRITE_UNLISTED_DETAIL}")
    print("Mode: READ ONLY - NO CI UPDATES")

    save_json("software_run_status.json", {"status": "running", "mode": "read_only"})

    collector = Collector(stamp)
    status, error = "completed", None
    interrupted = False

    try:
        with requests.Session() as session:
            pool_size = max(WORKERS, 1) + 2
            adapter = requests.adapters.HTTPAdapter(
                pool_connections=pool_size, pool_maxsize=pool_size)
            session.mount("https://", adapter)
            session.mount("http://", adapter)
            session.headers.update({
                "Accept": "application/json",
                "Content-Type": "application/json",
            })

            authenticate(session)

            for number, total, data in iter_chunks(session):
                if not isinstance(data, dict):
                    raise RuntimeError("Topology result must be a JSON object.")
                if SAVE_RAW_CHUNKS is None or number <= SAVE_RAW_CHUNKS:
                    save_json(f"chunk_{number:05d}.json", data, RAW_DIR)

                collector.process_chunk(data)

                if number % PROGRESS_EVERY == 0 or number == total:
                    elapsed = time.time() - started
                    rate = collector.chunks_done / elapsed if elapsed else 0
                    eta = (total - collector.chunks_done) / rate if rate else 0
                    print(f"Chunk {number}/{total} | software: "
                          f"{collector.records_seen:,} | devices: "
                          f"{len(collector.host_env):,} | elapsed "
                          f"{elapsed / 60:.1f} min | ETA {eta / 60:.1f} min")
                    collector.detail.flush()
    except KeyboardInterrupt:
        status, interrupted = "interrupted", True
        print("\nInterrupted - writing partial reports...")
    except Exception as exc:
        status, error = "failed", str(exc)
        print(f"\nERROR: {exc}\nWriting partial reports from data collected so far...")

    partial = status != "completed"
    outputs, env_rows = finalize(collector, stamp, partial)

    summary = {
        "mode": "read_only",
        "status": status,
        "error": error,
        "generated": stamp,
        "chunks_processed": collector.chunks_done,
        "software_records_seen": collector.records_seen,
        "devices": len(collector.host_env),
        "distinct_software_versions": len(collector.sw_summary),
        "detail_rows_written": collector.detail.total,
        "detail_files": [p.name for p in collector.detail.files],
        "counts": dict(collector.counts),
        "devices_by_environment": {r[0]: r[1] for r in env_rows},
        "duration_minutes": round((time.time() - started) / 60, 1),
    }
    save_json("software_summary.json", summary)
    save_json("software_run_status.json", {
        "status": status, "mode": "read_only", "error": error})

    print("\n" + "=" * 70)
    print("SUMMARY" + ("  (PARTIAL RESULTS)" if partial else ""))
    print("=" * 70)
    print(json.dumps(summary, indent=2))

    print("\nDEVICES BY DEPLOYMENT ENVIRONMENT")
    print("-" * 70)
    for r in env_rows:
        print(f"  {r[0]:<28} devices: {r[1]:<8} software records: {r[2]}")

    print("\nReports written:")
    for p in outputs:
        print(f"  {p}")

    if collector.counts.get("unresolved"):
        print(f"\nATTENTION: {collector.counts['unresolved']:,} software record(s) "
              "could not be attributed (see software_unresolved CSV). They "
              "were NOT marked Unauthorized.")
    if collector.chunks_done and not collector.host_env:
        print("\nWARNING: no software could be linked to a computer. Check that "
              "the TQL returns Node -> InstalledSoftware relations.")

    print("\nNo CIs were modified.")
    if interrupted:
        sys.exit(130)
    if status == "failed":
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:
        save_json("software_run_status.json",
                  {"status": "failed", "mode": "read_only", "error": str(exc)})
        print(f"\nERROR: {exc}")
        sys.exit(1)  
