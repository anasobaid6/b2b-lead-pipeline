"""
================================================================================
Enterprise B2B Lead Extraction & Validation Pipeline
Author: Senior Backend Automation Engineer
Description: Resilient asynchronous scraping ETL engine featuring Pydantic schema
             contracts, rate-limiting with exponential backoff & jitter, a dead-letter
             quarantine queue, an interactive Rich CLI dashboard, and styled Excel export.
================================================================================
"""

import asyncio
from datetime import datetime, timezone
import json
import logging
from pathlib import Path
import random
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

# Ensure Windows terminal standard streams handle UTF-8 properly
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field, ValidationError, field_validator
from rich import box
from rich.console import Console, Group
from rich.layout import Layout
from rich.live import Live
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskID,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.table import Table
from rich.text import Text

# -----------------------------------------------------------------------------
# 1. LOGGING & PATH CONFIGURATION
# -----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "pipeline_execution.log"
EXCEL_OUTPUT_FILE = BASE_DIR / "leads_validated.xlsx"
DEAD_LETTER_FILE = BASE_DIR / "quarantine_dead_letter.jsonl"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [%(name)s]: %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w"),
    ],
)
logger = logging.getLogger("EnterpriseLeadETL")


# -----------------------------------------------------------------------------
# 2. PYDANTIC SCHEMA & RIGOROUS DATA CONTRACT
# -----------------------------------------------------------------------------
class CompanyLead(BaseModel):
    """
    Strict B2B company lead data contract.
    Enforces regex email validation, normalized domains, cleaned phone numbers,
    and sanitized numeric metrics.
    """
    company_name: str = Field(..., min_length=2, max_length=150, description="Legal or operating company name")
    domain: str = Field(..., description="Sanitized root domain (e.g. acme-tech.com)")
    employee_count: int = Field(..., ge=1, le=5_000_000, description="Full-time employee headcount")
    industry: str = Field(..., min_length=2, max_length=100, description="Primary industry vertical")
    phone: str = Field(..., description="Cleaned corporate telephone number")
    email: str = Field(..., description="Corporate primary contact email")
    annual_revenue: float = Field(..., ge=0.0, description="Annual recurring/estimated revenue in USD")
    confidence_score: float = Field(default=0.95, ge=0.0, le=1.0, description="Extraction confidence score")
    scraped_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc), description="UTC extraction timestamp")

    @field_validator("domain", mode="before")
    @classmethod
    def normalize_domain(cls, v: Any) -> str:
        if not isinstance(v, str):
            raise ValueError(f"Domain must be a string, got {type(v).__name__}")
        cleaned = v.strip().lower()
        cleaned = re.sub(r"^https?://", "", cleaned)
        cleaned = re.sub(r"^www\.", "", cleaned)
        cleaned = cleaned.split("/")[0].split("?")[0].strip()
        # Domain validation: must contain a dot and valid host characters
        if not re.match(r"^[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?(\.[a-z0-9-]{2,})+$", cleaned):
            raise ValueError(f"Invalid domain syntax: '{v}'")
        return cleaned

    @field_validator("email", mode="before")
    @classmethod
    def validate_corporate_email(cls, v: Any) -> str:
        if not isinstance(v, str):
            raise ValueError(f"Email must be a string, got {type(v).__name__}")
        email = v.strip().lower()
        email_regex = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        if not re.match(email_regex, email):
            raise ValueError(f"Malformed or non-compliant email: '{v}'")
        # Reject generic non-actionable placeholder prefixes
        if email.startswith(("null@", "undefined@", "none@", "test@")):
            raise ValueError(f"Disposable or dummy placeholder email rejected: '{v}'")
        return email

    @field_validator("phone", mode="before")
    @classmethod
    def normalize_phone(cls, v: Any) -> str:
        if not isinstance(v, str):
            raise ValueError(f"Phone must be a string, got {type(v).__name__}")
        raw = v.strip()
        digits = re.sub(r"[^\d+]", "", raw)
        if len(re.sub(r"\D", "", digits)) < 10:
            raise ValueError(f"Phone number too short or invalid: '{v}'")
        # Format 10-digit North American numbers or keep clean international E.164
        num_only = re.sub(r"\D", "", digits)
        if len(num_only) == 10:
            return f"+1 ({num_only[:3]}) {num_only[3:6]}-{num_only[6:]}"
        elif len(num_only) == 11 and num_only.startswith("1"):
            return f"+1 ({num_only[1:4]}) {num_only[4:7]}-{num_only[7:]}"
        return digits

    @field_validator("employee_count", mode="before")
    @classmethod
    def parse_employee_count(cls, v: Any) -> int:
        if isinstance(v, int):
            if v <= 0:
                raise ValueError("Employee count must be positive")
            return v
        if isinstance(v, str):
            # Parse formats like "1,500+", "250-500", "750 employees"
            cleaned = v.replace(",", "").replace("+", "").strip()
            # If a range e.g. "100-250", pick upper bound
            range_match = re.search(r"(\d+)\s*-\s*(\d+)", cleaned)
            if range_match:
                return int(range_match.group(2))
            digit_match = re.search(r"\d+", cleaned)
            if digit_match:
                val = int(digit_match.group(0))
                if val <= 0:
                    raise ValueError(f"Extracted zero/negative employee count: {val}")
                return val
        raise ValueError(f"Unable to parse valid employee count from: '{v}'")

    @field_validator("annual_revenue", mode="before")
    @classmethod
    def parse_annual_revenue(cls, v: Any) -> float:
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str):
            s = v.strip().upper().replace("$", "").replace(",", "")
            multiplier = 1.0
            if "B" in s:
                multiplier = 1_000_000_000.0
                s = s.replace("B", "")
            elif "M" in s:
                multiplier = 1_000_000.0
                s = s.replace("M", "")
            elif "K" in s:
                multiplier = 1_000.0
                s = s.replace("K", "")
            try:
                numeric_val = float(re.findall(r"[\d.]+", s)[0])
                return numeric_val * multiplier
            except Exception:
                pass
        raise ValueError(f"Invalid monetary revenue specification: '{v}'")


# -----------------------------------------------------------------------------
# 3. ASYNC EXTRACTION SIMULATION & RESILIENCE ENGINE
# -----------------------------------------------------------------------------
class MockHttpEngine:
    """
    Simulates high-throughput asynchronous scraping against enterprise B2B directories.
    Features synthetic DOM construction, intermittent HTTP 429 rate-limiting,
    and simulated corrupt markup.
    """

    SAMPLE_INDUSTRIES = [
        "Cloud Infrastructure",
        "Cybersecurity",
        "Fintech & Payments",
        "Enterprise SaaS",
        "AI & Autonomous Systems",
        "HealthTech",
        "Supply Chain Logistics",
    ]

    SAMPLE_COMPANIES = [
        ("CloudScale Systems", "cloudscalesystems.com", "500+", "$48.5M", "+1-415-555-0192", "inquiries@cloudscalesystems.com"),
        ("Apex CyberGuard", "https://www.apexcyberguard.io/", "1,200", "$120M", "650-555-0821", "partners@apexcyberguard.io"),
        ("Quantum Ledger", "quantum-ledger.finance", "85", "$14.2M", "212.555.0345", "contact@quantum-ledger.finance"),
        ("BioSynth Analytics", "biosynth-analytics.org", "350", "$65M", "+1 (617) 555-0988", "leads@biosynth-analytics.org"),
        ("OmniRoute Freight", "omniroutefreight.com", "2,400", "$310M", "312-555-4491", "sales@omniroutefreight.com"),
        ("DeepCognition AI", "deepcognition.ai", "140", "$28.0M", "408-555-9124", "hello@deepcognition.ai"),
        ("Vanguard Risk Solutions", "vanguard-risk.com", "900", "$85.5M", "214.555.7723", "bizdev@vanguard-risk.com"),
        ("Veritas Health Systems", "veritas-health.net", "4,100", "$520M", "+1 305 555 6112", "enterprise@veritas-health.net"),
        ("AeroDynamics Logic", "aerodynamics-logic.com", "450", "$72.4M", "206-555-8819", "info@aerodynamics-logic.com"),
        ("HyperVolt Energy", "hypervoltenergy.com", "680", "$94M", "713.555.3340", "growth@hypervoltenergy.com"),
        ("Synergy Cloud AI", "synergycloud.tech", "210", "$32.8M", "+1-800-555-1200", "connect@synergycloud.tech"),
        ("Sentinel Data Labs", "sentinel-datalabs.io", "50", "$8.5M", "415-555-9981", "hello@sentinel-datalabs.io"),
        ("Global Matrix Global", "globalmatrix.co", "180", "$22M", "+1-202-555-0144", "team@globalmatrix.co"),
        ("Stratum FinOps", "stratum-finops.com", "320", "$45M", "646.555.0117", "contact@stratum-finops.com"),
        ("PulsePoint Systems", "pulsepoint-sys.com", "110", "$16.8M", "+1 (858) 555-0177", "inquire@pulsepoint-sys.com"),
    ]

    CORRUPTED_SCENARIOS = [
        # Missing essential company name
        """<div class="company-card"><div class="meta-tag domain">ghost-corp.com</div><div class="contact"><span class="email">info@ghost-corp.com</span></div></div>""",
        # Corrupted email format
        """<div class="company-card"><h1 class="company-name">BrokenData Inc</h1><div class="meta-tag domain">brokendata.io</div><div class="badge employees">150</div><div class="contact"><span class="email">mailto://invalid-address</span><span class="phone">555-0199</span></div><div class="financials revenue">$10M</div></div>""",
        # Negative / corrupt headcount
        """<div class="company-card"><h1 class="company-name">GlitchWorks LLC</h1><div class="meta-tag domain">glitchworks.co</div><div class="badge employees">-45</div><div class="contact"><span class="email">ops@glitchworks.co</span><span class="phone">415-555-9000</span></div><div class="financials revenue">$5M</div></div>""",
        # Corrupted domain format
        """<div class="company-card"><h1 class="company-name">Malformed Web</h1><div class="meta-tag domain">htt p://invalid:domain</div><div class="badge employees">200</div><div class="contact"><span class="email">web@invalid.com</span><span class="phone">212-555-1234</span></div><div class="financials revenue">$12M</div></div>""",
        # Truncated markup
        """<div class="company-card"><h1 class="company-name">Empty Shell Corp</h1><div class="financials revenue">N/A</div>""",
    ]

    @classmethod
    def generate_synthetic_html(cls, target_id: int) -> Tuple[str, bool]:
        """
        Generates realistic HTML DOM string.
        Returns: (html_content, is_corrupted_flag)
        """
        # Inject ~15% corrupt markup to demonstrate resilient dead-letter routing
        if target_id % 7 == 0:
            corrupted_html = random.choice(cls.CORRUPTED_SCENARIOS)
            return corrupted_html, True

        # Generate healthy enterprise HTML markup
        name, domain, employees, revenue, phone, email = random.choice(cls.SAMPLE_COMPANIES)
        industry = random.choice(cls.SAMPLE_INDUSTRIES)
        # Add slight variation to ensure unique records
        unique_name = f"{name} {target_id}" if target_id > 15 else name
        unique_email = f"lead.{target_id}.{email}" if target_id > 15 else email

        html = f"""
        <!DOCTYPE html>
        <html>
        <head><title>{unique_name} Profile</title></head>
        <body>
            <div class="company-card" data-entity-id="ENT-{target_id:04d}">
                <div class="company-header">
                    <h1 class="company-name">{unique_name}</h1>
                    <span class="industry-badge">{industry}</span>
                </div>
                <div class="company-details">
                    <div class="meta-tag domain"><a href="https://{domain}">{domain}</a></div>
                    <div class="badge employees">{employees} Employees</div>
                    <div class="financials revenue">{revenue}</div>
                </div>
                <div class="contact-section">
                    <div class="contact phone">{phone}</div>
                    <div class="contact email"><a href="mailto:{unique_email}">{unique_email}</a></div>
                </div>
            </div>
        </body>
        </html>
        """
        return html, False

    @classmethod
    async def fetch_page(cls, target_id: int, url: str) -> Dict[str, Any]:
        """
        Simulates network request.
        Randomly simulates HTTP 429 Too Many Requests (~18% chance) to showcase
        exponential backoff handling.
        """
        # Simulate network latency (50ms - 180ms)
        await asyncio.sleep(random.uniform(0.05, 0.18))

        # 18% probability of HTTP 429 Rate Limit
        if random.random() < 0.18:
            return {
                "status_code": 429,
                "target_id": target_id,
                "url": url,
                "error": "HTTP 429 Too Many Requests (Rate Limit Triggered)",
                "html": "",
            }

        html, is_corrupt = cls.generate_synthetic_html(target_id)
        return {
            "status_code": 200,
            "target_id": target_id,
            "url": url,
            "html": html,
            "is_corrupt": is_corrupt,
        }


# -----------------------------------------------------------------------------
# 4. DEAD-LETTER QUARANTINE MANAGER
# -----------------------------------------------------------------------------
class DeadLetterQuarantine:
    """
    Safely captures unparsable, malformed, or contract-violating records
    into an append-only JSON-lines quarantine file without terminating the pipeline.
    """

    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.lock = asyncio.Lock()
        # Initialize clean file
        with open(self.file_path, "w", encoding="utf-8") as f:
            pass

    async def log_quarantine(
        self,
        target_id: int,
        url: str,
        stage: str,
        reason: str,
        raw_payload: str,
    ) -> None:
        entry = {
            "quarantine_id": f"DLQ-{target_id:05d}",
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "url": url,
            "failure_stage": stage,
            "error_reason": reason,
            "raw_snippet": (raw_payload[:250] + "...") if len(raw_payload) > 250 else raw_payload,
        }
        async with self.lock:
            with open(self.file_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry) + "\n")

        logger.warning(
            f"Quarantined target #{target_id} [{stage}]: {reason}"
        )


# -----------------------------------------------------------------------------
# 5. EXCEL EXPORT & STYLING ENGINE
# -----------------------------------------------------------------------------
class ExcelOutputHandler:
    """
    Transforms validated Pydantic lead objects into a formatted enterprise
    spreadsheet with stylized headers, borders, number formatting, and auto-width.
    """

    @staticmethod
    def export(leads: List[CompanyLead], output_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Validated B2B Leads"

        # Apply gridlines visibility
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Company Name",
            "Root Domain",
            "Industry",
            "Employees",
            "Corporate Phone",
            "Contact Email",
            "Annual Revenue",
            "Confidence",
            "Extracted At (UTC)",
        ]
        ws.append(headers)

        # Style Header Row
        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="thin", color="CBD5E1"),
        )

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
            cell.border = thin_border
        ws.row_dimensions[1].height = 28

        # Insert Records & Format Cells
        regular_font = Font(name="Calibri", size=10)
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        for row_idx, lead in enumerate(leads, start=2):
            ws.append([
                lead.company_name,
                lead.domain,
                lead.industry,
                lead.employee_count,
                lead.phone,
                lead.email,
                lead.annual_revenue,
                lead.confidence_score,
                lead.scraped_at.strftime("%Y-%m-%d %H:%M:%S"),
            ])
            ws.row_dimensions[row_idx].height = 20

            # Formatting
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = regular_font
                cell.border = thin_border
                if row_idx % 2 == 0:
                    cell.fill = zebra_fill

                # Number formatting
                if col_idx == 4:  # Employees
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx == 7:  # Revenue
                    cell.number_format = '"$"#,##0'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx == 8:  # Confidence
                    cell.number_format = "0.0%"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Auto-adjust column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if cell.column == 7 and isinstance(cell.value, (int, float)):
                    val_str = f"${cell.value:,.0f}"
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        # Freeze header row & add auto-filter
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        wb.save(output_path)
        logger.info(f"Successfully generated structured Excel workbook: {output_path} ({len(leads)} leads)")


# -----------------------------------------------------------------------------
# 6. PIPELINE CONTROLLER & RICH LIVE DASHBOARD
# -----------------------------------------------------------------------------
class PipelineTelemetry:
    """Thread-safe telemetry state for live Rich dashboard rendering."""

    def __init__(self, total_targets: int):
        self.total_targets = total_targets
        self.processed = 0
        self.valid_count = 0
        self.quarantine_count = 0
        self.rate_limit_hits = 0
        self.recent_valid_leads: List[CompanyLead] = []
        self.start_time = time.time()
        self.latest_event = "Initializing async ETL workers..."

    @property
    def throughput(self) -> float:
        elapsed = time.time() - self.start_time
        return self.processed / elapsed if elapsed > 0 else 0.0


class EnterpriseLeadPipeline:
    """
    Main orchestration pipeline coordinating:
    - Asynchronous worker pool
    - Rate limit exponential backoff with jitter
    - BeautifulSoup DOM extraction
    - Pydantic schema contract validation
    - Dead-letter quarantine routing
    - Rich live telemetry terminal display
    """

    def __init__(self, total_targets: int = 50, concurrency_limit: int = 5):
        self.total_targets = total_targets
        self.concurrency_limit = concurrency_limit
        self.telemetry = PipelineTelemetry(total_targets)
        self.quarantine = DeadLetterQuarantine(DEAD_LETTER_FILE)
        self.validated_leads: List[CompanyLead] = []
        self.console = Console(force_terminal=True, legacy_windows=False)

    def parse_dom(self, html: str) -> Dict[str, Any]:
        """Extract raw attributes from HTML using BeautifulSoup."""
        soup = BeautifulSoup(html, "html.parser")
        card = soup.find("div", class_="company-card")
        if not card:
            raise ValueError("Corrupted DOM: Missing '.company-card' root container")

        name_el = card.find("h1", class_="company-name")
        domain_el = card.find("div", class_="domain")
        employees_el = card.find("div", class_="employees")
        industry_el = card.find("span", class_="industry-badge")
        phone_el = card.find("div", class_="phone")
        email_el = card.find("div", class_="email")
        revenue_el = card.find("div", class_="financials")

        if not name_el or not name_el.text.strip():
            raise ValueError("Corrupted DOM: Target company name is absent or empty")

        return {
            "company_name": name_el.text.strip(),
            "domain": domain_el.text.strip() if domain_el else "",
            "employee_count": employees_el.text.strip() if employees_el else 0,
            "industry": industry_el.text.strip() if industry_el else "General Enterprise",
            "phone": phone_el.text.strip() if phone_el else "",
            "email": email_el.text.strip() if email_el else "",
            "annual_revenue": revenue_el.text.strip() if revenue_el else "0",
            "confidence_score": round(random.uniform(0.92, 0.99), 2),
        }

    async def scrape_target(self, target_id: int, semaphore: asyncio.Semaphore) -> None:
        """
        Processes single target with retry loop, backoff, DOM extraction, and Pydantic validation.
        """
        async with semaphore:
            url = f"https://enterprise-directory.internal/company/ENT-{target_id:04d}"
            max_retries = 3
            base_backoff = 0.4
            attempt = 0

            while attempt <= max_retries:
                fetch_res = await MockHttpEngine.fetch_page(target_id, url)

                # Case 1: Handle HTTP 429 Too Many Requests (Exponential Backoff + Jitter)
                if fetch_res.get("status_code") == 429:
                    attempt += 1
                    self.telemetry.rate_limit_hits += 1
                    if attempt > max_retries:
                        await self.quarantine.log_quarantine(
                            target_id=target_id,
                            url=url,
                            stage="RATE_LIMIT_EXHAUSTED",
                            reason=f"Exceeded {max_retries} exponential backoff retry attempts",
                            raw_payload="HTTP 429 Too Many Requests",
                        )
                        self.telemetry.quarantine_count += 1
                        self.telemetry.processed += 1
                        self.telemetry.latest_event = f"[bold red]DLQ:[/] Target #{target_id:02d} dropped (Rate-limit exhausted)"
                        return

                    # Backoff = base * 2^(attempt-1) + jitter
                    jitter = random.uniform(0.05, 0.25)
                    backoff_delay = (base_backoff * (2 ** (attempt - 1))) + jitter
                    self.telemetry.latest_event = (
                        f"[bold yellow]429 BACKOFF:[/] Target #{target_id:02d} hit rate limit "
                        f"(Retrying in {backoff_delay:.2f}s, attempt {attempt}/{max_retries})"
                    )
                    logger.warning(
                        f"Target #{target_id} hit HTTP 429. Backing off for {backoff_delay:.2f}s "
                        f"(Attempt {attempt}/{max_retries})"
                    )
                    await asyncio.sleep(backoff_delay)
                    continue

                # Case 2: Process Successful HTTP 200 Response
                html = fetch_res.get("html", "")

                # DOM Parsing Stage
                try:
                    raw_data = self.parse_dom(html)
                except Exception as dom_err:
                    await self.quarantine.log_quarantine(
                        target_id=target_id,
                        url=url,
                        stage="DOM_PARSING",
                        reason=str(dom_err),
                        raw_payload=html,
                    )
                    self.telemetry.quarantine_count += 1
                    self.telemetry.processed += 1
                    self.telemetry.latest_event = f"[bold red]CORRUPT DOM:[/] Target #{target_id:02d} failed HTML extraction"
                    return

                # Pydantic Schema Validation Stage
                try:
                    lead = CompanyLead(**raw_data)
                    self.validated_leads.append(lead)
                    self.telemetry.valid_count += 1
                    self.telemetry.recent_valid_leads.append(lead)
                    if len(self.telemetry.recent_valid_leads) > 5:
                        self.telemetry.recent_valid_leads.pop(0)

                    self.telemetry.latest_event = (
                        f"[bold green]VALIDATED:[/] {lead.company_name} "
                        f"({lead.industry} • ${lead.annual_revenue:,.0f})"
                    )
                    logger.info(f"Validated lead #{target_id}: {lead.company_name} [{lead.domain}]")
                except ValidationError as val_err:
                    error_summary = "; ".join(
                        f"{e['loc'][0]}: {e['msg']}" for e in val_err.errors()
                    )
                    await self.quarantine.log_quarantine(
                        target_id=target_id,
                        url=url,
                        stage="PYDANTIC_VALIDATION",
                        reason=error_summary,
                        raw_payload=html,
                    )
                    self.telemetry.quarantine_count += 1
                    self.telemetry.latest_event = f"[bold red]INVALID SCHEMA:[/] Target #{target_id:02d} ({error_summary[:45]}...)"
                finally:
                    self.telemetry.processed += 1

                break

    def render_dashboard(self, progress_component: Progress) -> Group:
        """Constructs an integrated Rich layout for the live CLI dashboard."""
        # 1. Header Banner
        header_text = Text()
        header_text.append(">>> ENTERPRISE B2B LEAD EXTRACTION & ETL PIPELINE <<<\n", style="bold cyan")
        header_text.append("Async Scraping Engine | Pydantic Contract Validation | DLQ Quarantine", style="dim white")
        header_panel = Panel(
            header_text,
            box=box.DOUBLE_EDGE,
            style="cyan",
            padding=(0, 2),
        )

        # 2. KPI Metrics Grid Table
        kpi_table = Table(box=box.SIMPLE_HEAD, expand=True, padding=(0, 2))
        kpi_table.add_column("Total Targets", justify="center", style="bold white")
        kpi_table.add_column("Processed", justify="center", style="bold cyan")
        kpi_table.add_column("Valid Leads", justify="center", style="bold green")
        kpi_table.add_column("Quarantined (DLQ)", justify="center", style="bold red")
        kpi_table.add_column("429 Backoffs", justify="center", style="bold yellow")
        kpi_table.add_column("Throughput", justify="center", style="bold magenta")

        valid_pct = (self.telemetry.valid_count / self.telemetry.processed * 100) if self.telemetry.processed > 0 else 0.0
        quarantine_pct = (self.telemetry.quarantine_count / self.telemetry.processed * 100) if self.telemetry.processed > 0 else 0.0

        kpi_table.add_row(
            str(self.telemetry.total_targets),
            f"{self.telemetry.processed}/{self.telemetry.total_targets}",
            f"{self.telemetry.valid_count} ({valid_pct:.1f}%)",
            f"{self.telemetry.quarantine_count} ({quarantine_pct:.1f}%)",
            str(self.telemetry.rate_limit_hits),
            f"{self.telemetry.throughput:.1f} rec/s",
        )
        metrics_panel = Panel(kpi_table, title="[bold]Real-Time Extraction Telemetry[/bold]", border_style="blue")

        # 3. Live Updating Table of Recent Validated Leads
        leads_table = Table(
            box=box.ROUNDED,
            expand=True,
            show_header=True,
            header_style="bold steel_blue1",
            padding=(0, 1),
        )
        leads_table.add_column("Company", style="bold white", width=22)
        leads_table.add_column("Root Domain", style="cyan", width=20)
        leads_table.add_column("Industry", style="magenta", width=18)
        leads_table.add_column("Headcount", justify="right", style="bright_yellow", width=10)
        leads_table.add_column("Annual Revenue", justify="right", style="bold green", width=14)
        leads_table.add_column("Confidence", justify="center", style="bright_cyan", width=10)

        if not self.telemetry.recent_valid_leads:
            leads_table.add_row(
                "[dim]Awaiting records...[/dim]",
                "[dim]--[/dim]",
                "[dim]--[/dim]",
                "[dim]--[/dim]",
                "[dim]--[/dim]",
                "[dim]--[/dim]",
            )
        else:
            for lead in reversed(self.telemetry.recent_valid_leads):
                leads_table.add_row(
                    lead.company_name,
                    lead.domain,
                    lead.industry,
                    f"{lead.employee_count:,}",
                    f"${lead.annual_revenue:,.0f}",
                    f"{lead.confidence_score * 100:.1f}%",
                )

        recent_panel = Panel(
            leads_table,
            title="[bold]Last 5 Validated Enterprise Leads (In-Flight Buffer)[/bold]",
            border_style="green",
        )

        # 4. Status Ticker / Event Stream
        ticker_panel = Panel(
            Text.from_markup(f"> {self.telemetry.latest_event}"),
            title="[bold]Pipeline Stream Ticker[/bold]",
            border_style="dim white",
            padding=(0, 1),
        )

        return Group(
            header_panel,
            metrics_panel,
            Panel(progress_component, title="[bold]Pipeline Ingestion Progress[/bold]", border_style="cyan"),
            recent_panel,
            ticker_panel,
        )

    async def execute(self) -> None:
        """Runs the entire asynchronous scraping ETL lifecycle."""
        logger.info(f"Starting Enterprise B2B Pipeline: {self.total_targets} targets, concurrency={self.concurrency_limit}")

        semaphore = asyncio.Semaphore(self.concurrency_limit)
        progress = Progress(
            SpinnerColumn(spinner_name="dots"),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(bar_width=45, complete_style="cyan", finished_style="green"),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            console=self.console,
        )
        task_id = progress.add_task("[bold cyan]Extracting & Validating Leads...", total=self.total_targets)

        # Spawn asynchronous tasks
        tasks = [
            asyncio.create_task(self.scrape_target(i, semaphore))
            for i in range(1, self.total_targets + 1)
        ]

        # Render Rich Live Dashboard at 8 frames per second
        with Live(self.render_dashboard(progress), console=self.console, refresh_per_second=8, screen=False) as live:
            while not all(t.done() for t in tasks):
                progress.update(task_id, completed=self.telemetry.processed)
                live.update(self.render_dashboard(progress))
                await asyncio.sleep(0.12)

            # Await all tasks to ensure clean closure
            await asyncio.gather(*tasks)
            progress.update(task_id, completed=self.telemetry.processed)
            live.update(self.render_dashboard(progress))

        # Output Generation
        self.console.print("\n[bold cyan]Processing complete! Exporting validated records to Excel...[/bold cyan]")
        ExcelOutputHandler.export(self.validated_leads, EXCEL_OUTPUT_FILE)

        # Final Summary Console Card
        summary_table = Table(title="[bold]ETL Execution Summary[/bold]", box=box.ROUNDED, show_header=False)
        summary_table.add_column("Metric", style="bold white")
        summary_table.add_column("Value", style="cyan")
        summary_table.add_row("Total Targets Scanned", str(self.telemetry.total_targets))
        summary_table.add_row("Successfully Validated", f"[bold green]{len(self.validated_leads)} records[/bold green]")
        summary_table.add_row("Quarantined to Dead-Letter", f"[bold red]{self.telemetry.quarantine_count} records[/bold red]")
        summary_table.add_row("Rate Limits Handled (429)", f"[bold yellow]{self.telemetry.rate_limit_hits} backoffs[/bold yellow]")
        summary_table.add_row("Output Spreadsheet", str(EXCEL_OUTPUT_FILE))
        summary_table.add_row("Quarantine Log", str(DEAD_LETTER_FILE))
        summary_table.add_row("Full Execution Log", str(LOG_FILE))
        self.console.print(summary_table)


# -----------------------------------------------------------------------------
# 7. SCRIPT ENTRYPOINT
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    # Ensure Windows asyncio event loop policy is compatible
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    pipeline = EnterpriseLeadPipeline(total_targets=50, concurrency_limit=6)
    asyncio.run(pipeline.execute())
