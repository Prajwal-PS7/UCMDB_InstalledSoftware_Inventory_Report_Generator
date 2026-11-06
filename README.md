# UCMDB InstalledSoftware Inventory Report Generator (REST API)
<img width="3584" height="1184" alt="Gemini_Generated_Image_9dzjza9dzjza9dzj" src="https://github.com/user-attachments/assets/f00169a4-1584-4b8d-b63d-761ef649fa7e" />

## Extract all installed software inventory data from UCMDB REST API
🔹 **Script** : script.py <br/>
🔹 **Target** : Micro Focus / OpenText UCMDB v25.4 & upper (REST API) <br/>
🔹 **Mode**   : READ ONLY - The script never creates, updates or deletes any CI <br/>

## 1. PURPOSE
This is the LARGE-SCALE edition, tested on multi-million record estates (for example, 4.4 million installed-software CIs).
UCMDB is unable to handle large-scale data loading and report extraction due to inherent limitations. To overcome this and successfully retrieve all installed-software reports from UCMDB, this script has been designed.

## 2. Key Highlights & Architecture

* **Tested for Multi-Million CI Estates**: Verified across large-scale enterprise deployments exceeding 4.4+ million installed software records.
* **Flat Memory Footprint**: Rather than loading multi-gigabyte topologies into RAM, each chunk is streamed, parsed, classified, flushed to disk, and freed.
* **Safe & Non-Destructive**: **100% Read-Only**. No CI modifications, deletions, or schema writes are ever sent to UCMDB.
* **Resilient Parallel Processing**:
  * Concurrent chunk fetching using Python's `ThreadPoolExecutor`.
  * Exponential backoff and retry for transient network drops, rate limits (HTTP 429), and gateway errors (HTTP 5xx).
  * Automatic token refresh/re-authentication on HTTP 401 mid-stream without terminating long-running jobs.
* **Fail-Safe & Graceful Degradation**:
  * If interrupted (`Ctrl+C`) or terminated by an unrecoverable network error, the collector flushes all data gathered up to that second into timestamped `_PARTIAL` reports.
* **Spreadsheet-Ready Output**:
  * Detailed output automatically rolls over into split CSV part files (default: 1,000,000 rows/file) to respect native Microsoft Excel boundaries.
  * Formatted summary workbook (`.xlsx`) featuring auto-filters, dynamic column sizing, and conditional formatting.

## 3. Generated Reports

Every execution outputs timestamped files in the script’s directory:

| Artifact | Format | Description |
| :--- | :--- | :--- |
| `software_compliance_detail_*.csv` | CSV (Rotated) | Granular line-item breakdown: Hostname, IP, Serial, Software, Version, Publisher, Authorization, Compliance, Reason. Rotates per 1,000,000 rows. |
| `computer_summary_*.csv` | CSV | Host-level breakdown aggregated by deployment environment with total/authorized/compliant counts. |
| `environment_summary_*.csv` | CSV | High-level summary of devices and records across environments (e.g., Production, Staging, QA). |
| `software_summary_*.csv` | CSV | Grouped list of unique software packages, versions, total installs, and compliance states. |
| `software_unresolved_*.csv` | CSV | Orphaned records (missing host composition relations or labels) for UCMDB data hygiene audits. |
| `software_compliance_summary_*.xlsx` | Excel Workbook | Consolidated management dashboard including color-coded status tabs and run metadata. |
| `software_summary.json` | JSON | Machine-readable metrics summary (total chunks, run duration, failure/partial flags, counts). |


## 4. UCMDB TQL Configuration

Before running the script, verify that your TQL query (default: `software_version_compliance`) exists and contains the necessary schema and layouts:

1. **Hierarchy / Relations**:
   * `Node` $\xrightarrow{\text{Composition}}$ `InstalledSoftware` *(Required)*
   * `Node` $\xrightarrow{\text{Containment}}$ `IpAddress` *(Optional, for multi-homed IP resolution)*
2. **Layout Properties (Attributes to return)**:
   * **Node**: `name`, `primary_dns_name`, `serial_number`, `primary_ip_address`, `ip_address`
   * **InstalledSoftware**: `name`, `display_label`, `version`, `software_version`, `deployment_env`, `publisher`, `vendor`
   * **IpAddress**: `name`, `ip_address`, `display_label`
3. **Pagination Root**: Set to `InstalledSoftware` *(the TQL node label)*.

## 5. Configuration

Modify the parameters at the top of the script according to your infrastructure:

```python
# Connection & Endpoints
SERVER_URL = "https://ucmdb.your-domain.com:8443"
API_PATH = "/rest-api"
USERNAME = "admin"
PASSWORD_ENV_VAR = "UCMDB_PASSWORD"   # Recommended for CI/CD or unattended cron tasks

# TQL Details
TQL_NAME = "software_version_compliance"
PAGINATION_ROOT = "InstalledSoftware"

# Performance Tuning
CHUNK_SIZE = 2000                     # Recommended: 2,000 - 5,000 depending on server heap
WORKERS = 4                            # Parallel chunk download threads (2 to 8)
TIMEOUT = (15, 600)                    # (Connect, Read) timeouts in seconds
MAX_RETRIES = 4                        # Retry attempts on HTTP 429/5xx/network drops

# Export Controls
CSV_ROWS_PER_FILE = 1_000_000          # Part-file row limit (Excel safety threshold)
REPORT_UNLISTED_SOFTWARE = True        # Mark non-catalog items as Unauthorized
WRITE_UNLISTED_DETAIL = False          # False = only approved catalog software goes to line-item CSV
```

## 6. Interactive Authentication (Prompted)

Run the script directly. You will be prompted securely for your password:

```bash
python scritp.py
```

## 7. Unattended / Automation Mode

Set the environment variable `UCMDB_PASSWORD` before starting:

```bash
export UCMDB_PASSWORD="YourStrongPasswordHere"
python scritp.py
```

## 8. Sample Console Output

```text
======================================================================
UCMDB InstalledSoftware Inventory Report Generator (LARGE SCALE)
======================================================================
TQL: software_version_compliance | Pagination root: InstalledSoftware
Chunk size: 2000 | Workers: 4
Write unlisted software rows to detail CSV: False
Mode: READ ONLY - NO CI UPDATES
Authentication successful. Token is not displayed.

Server chunks: 2200 | Chunk size: 2000 | Parallel workers: 4
Chunk 1/2200 | software: 2,000 | devices: 1,410 | elapsed 0.1 min | ETA 35.2 min
Chunk 2/2200 | software: 4,000 | devices: 2,820 | elapsed 0.2 min | ETA 34.8 min
...
Chunk 2200/2200 | software: 4,400,000 | devices: 142,390 | elapsed 36.1 min | ETA 0.0 min

======================================================================
SUMMARY
======================================================================
{
  "mode": "read_only",
  "status": "completed",
  "software_records_seen": 4400000,
  "devices": 142390,
  "distinct_software_versions": 8421,
  "detail_rows_written": 182300,
  "duration_minutes": 36.1
}

Reports written:
  software_compliance_detail_20261107_120000_part001.csv
  computer_summary_20261107_120000.csv
  environment_summary_20261107_120000.csv
  software_summary_20261107_120000.csv
  software_unresolved_20261107_120000.csv
  software_compliance_summary_20261107_120000.xlsx

No CIs were modified.
```

## 9. Operational Troubleshooting

* **SSL Verification Failures**: Set `VERIFY_SSL = True` and point it to your enterprise internal root CA bundle `.pem` path if self-signed cert warnings are blocked.
* **UCMDB Server Heap / Slow Responses**: Decrease `CHUNK_SIZE` (e.g., from `2000` to `1000`) or reduce `WORKERS` from `4` down to `2`.
* **High Number of Unresolved Records**: Check `software_unresolved_*.csv`. If records show `"no computer relation found"`, inspect your TQL query layout to verify that the relationship edge between `Node` and `InstalledSoftware` is included in the query result.
