# Enterprise B2B Lead Extraction & Schema Validation Pipeline ⚡🛡️

[![Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-3776AB.svg?logo=python&logoColor=white)](https://www.python.org/)
[![Pydantic v2](https://img.shields.io/badge/Pydantic-v2-E92063.svg?logo=pydantic&logoColor=white)](https://docs.pydantic.dev/)
[![AsyncIO](https://img.shields.io/badge/Concurrency-AsyncIO-009688.svg)]()
[![OpenPyXL](https://img.shields.io/badge/Excel-OpenPyXL-217346.svg)]()

A fault-tolerant, asynchronous B2B lead extraction and validation pipeline engineered with **Python `asyncio`**, **Pydantic v2 data contracts**, **exponential backoff with jitter**, a **Dead-Letter Quarantine (DLQ)** audit log, and automated **OpenPyXL Excel modeling**.

![Pipeline Overview](portfolio_assets/1_upwork_cover.jpg)

---

## 🏛️ System Architecture

```mermaid
flowchart LR
    A["AsyncIO Worker Pool<br>(Semaphore + HTTP 429 Backoff)"] --> B["BeautifulSoup4<br>DOM Entity Parser"]
    B --> C{"Pydantic v2<br>Data Contract Gate"}
    C -->|"Valid Schema (86.0%)"| D["Styled OpenPyXL Workbook<br>(leads_validated.xlsx)"]
    C -->|"Corrupt DOM / Schema Violation (14.0%)"| E["Dead-Letter Quarantine<br>(quarantine_dead_letter.jsonl)"]
```

---

## 📊 Empirical Execution Telemetry (50 Target Entities)

| Pipeline Metric | Measured Benchmark Output |
| :--- | :--- |
| **Total Targets Scanned** | `50` B2B Directory Entities |
| **Validated Production Leads** | `43` Records (`86.0%` Pass Rate) $\rightarrow$ [`leads_validated.xlsx`](leads_validated.xlsx) |
| **Quarantined Corrupt Records** | `7` Records (`14.0%` Isolated) $\rightarrow$ [`quarantine_dead_letter.jsonl`](quarantine_dead_letter.jsonl) |
| **HTTP 429 Rate Limits Recovered** | `15` Automatic Exponential Backoffs + Random Jitter |
| **Pipeline Crash Rate** | `0.0%` (100% Fault-Tolerant Completion) |

### Rich CLI Live Telemetry & Quarantine Proof
![Rich CLI Telemetry](portfolio_assets/2_rich_cli_telemetry_and_dlq.jpg)

### Generated OpenPyXL Executive Spreadsheet (`leads_validated.xlsx`)
![Excel Output](portfolio_assets/3_excel_validated_leads_output.jpg)

---

## ⚡ Quickstart

```bash
pip install -r requirements.txt
python pipeline.py
```

### Sample Deliverable Files Included in Repository
- **[`leads_validated.xlsx`](leads_validated.xlsx)**: Styled Excel output containing 43 schema-verified B2B leads.
- **[`quarantine_dead_letter.jsonl`](quarantine_dead_letter.jsonl)**: Append-only JSONL quarantine log capturing exact DOM & Pydantic validation failures.
- **[`pipeline.py`](pipeline.py)**: Complete asynchronous ETL implementation.
