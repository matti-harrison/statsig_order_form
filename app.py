from __future__ import annotations

import io
import re
import base64
import os
from datetime import date, datetime, timedelta
from dataclasses import dataclass, asdict, field
from typing import Dict, List, Tuple

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


@dataclass
class OrderData:
    account_name: str = ""
    primary_contact_name: str = ""
    primary_contact_email: str = ""
    billing_email: str = ""
    shipping_address: str = ""
    billing_address: str = "Same as shipping address"
    start_date: str = ""
    subscription_term_months: int = 12
    billing_frequency: str = "Annual"
    payment_terms: str = "Net 30"
    payment_method: str = "Bank Transfer"
    billing_id: str = ""
    po_number: str = ""
    opportunity_type: str = ""
    addendum_effective_date: str = ""
    terms_type: str = "Online"
    msa_execution_date: str = ""
    special_terms: List[str] = field(default_factory=list)
    expiration_enabled: bool = False
    expiration_date: str = ""
    usage_terms: str = ""


DEFAULT_SERVICES: List[Dict[str, str]] = [
    {
        "subscription_period": "",
        "service": "Feature Gates and SDKs",
        "annual_usage_commitment": "N/A",
        "unit": "N/A",
        "annual_service_fee": 0.0,
    },
    {
        "subscription_period": "",
        "service": "Experimentation",
        "annual_usage_commitment": "",
        "unit": "Billable Events",
        "annual_service_fee": 0.0,
    },
    {
        "subscription_period": "",
        "service": "Advanced Product Analytics",
        "annual_usage_commitment": "",
        "unit": "N/A",
        "annual_service_fee": 0.0,
    },
]

WAREHOUSE_TYPES = ["Cloud", "Warehouse Native", "Credit/Usage Based"]

CLOUD_PRODUCTS = [
    "Feature Gates and SDKs",
    "Experimentation",
    "Advanced Product Analytics",
    "Platform Fee",
    "Session Replay",
]

WAREHOUSE_NATIVE_EXPERIMENTATION_OPTIONS = [
    "Experimentation: Analysis Only",
    "Experimentation: Analysis + Assignment",
]

WAREHOUSE_NATIVE_OTHER_PRODUCTS = [
    "Feature Gates and SDKs",
    "Advanced Product Analytics",
    "Platform Fee",
    "Session Replay",
]

USAGE_BASED_REQUIRED_PRODUCTS = ["Warehouse Native", "Platform Fee"]
SUPPORT_TIERS = ["Premium", "Standard", "Community"]


def normalize_text(text: str) -> str:
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())


def extract_text_from_upload(uploaded_file) -> str:
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore")

    if file_name.endswith(".pdf"):
        reader = PdfReader(uploaded_file)
        pages = [page.extract_text() or "" for page in reader.pages]
        return "\n".join(pages)

    if file_name.endswith(".docx"):
        doc = Document(uploaded_file)
        paragraphs = [p.text for p in doc.paragraphs]
        return "\n".join(paragraphs)

    return ""


def find_field(text: str, labels: List[str]) -> str:
    for label in labels:
        pattern = rf"(?im)^\s*{re.escape(label)}\s*[:\-]\s*(.+)$"
        match = re.search(pattern, text)
        if match:
            return match.group(1).strip()
    return ""


def extract_order_fields(text: str) -> OrderData:
    cleaned = normalize_text(text)
    return OrderData(
        account_name=find_field(cleaned, ["customer", "account", "account name", "customer name"]),
        primary_contact_name=find_field(cleaned, ["primary contact", "contact", "contact name"]),
        primary_contact_email=find_field(cleaned, ["contact email", "customer contact", "email"]),
        billing_email=find_field(cleaned, ["billing email", "invoice email"]),
        shipping_address=find_field(cleaned, ["shipping address", "ship to", "address"]),
        billing_address=find_field(cleaned, ["billing address", "bill to"]),
        start_date=find_field(cleaned, ["start date", "subscription start"]),
        subscription_term_months=_to_int(
            find_field(
                cleaned,
                ["subscription term (months)", "subscription term", "term months", "term"],
            )
        ),
        billing_frequency=find_field(cleaned, ["billing frequency", "frequency"]),
        payment_terms=find_field(cleaned, ["payment terms", "terms"]),
        payment_method=find_field(cleaned, ["payment method"]),
        billing_id=find_field(cleaned, ["billing id", "aws billing id", "gcp billing id", "azure billing id"]),
        po_number=find_field(cleaned, ["po", "po number", "purchase order"]),
        opportunity_type=find_field(cleaned, ["opportunity type", "deal type", "deal label"]),
        addendum_effective_date=find_field(
            cleaned,
            ["addendum effective date", "effective date", "upsell effective date"],
        ),
        terms_type=find_field(cleaned, ["terms", "terms type", "agreement type"]),
        msa_execution_date=find_field(
            cleaned,
            ["msa execution date", "msa executed on", "msa date"],
        ),
        expiration_date=find_field(cleaned, ["expiration date", "quote expiration", "expires on"]),
        usage_terms=find_field(cleaned, ["usage terms", "terms details", "notes"]),
    )


def safe_money(value: str) -> float:
    try:
        return float(str(value).replace("$", "").replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _to_int(value: str, default: int = 12) -> int:
    try:
        parsed = int(str(value).strip())
        return max(1, parsed)
    except ValueError:
        return default


def parse_date(value: str) -> date:
    if not value:
        today = date.today()
        return date(today.year, today.month, 1)
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    today = date.today()
    return date(today.year, today.month, 1)


def first_of_month(d: date) -> date:
    return date(d.year, d.month, 1)


def add_months(d: date, months: int) -> date:
    month_index = d.month - 1 + months
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    return date(year, month, 1)


def compute_end_date(start: date, term_months: int) -> date:
    return add_months(first_of_month(start), max(1, term_months)) - timedelta(days=1)


def display_date(d: date) -> str:
    return d.strftime("%m/%d/%Y")


def default_subscription_start_date(today: date | None = None) -> date:
    d = today or date.today()
    # Always default to the first day of the next month.
    return add_months(first_of_month(d), 1)


def build_output_filename(account_name: str) -> str:
    safe_account = (account_name or "").strip() or "Account Name"
    month_year = date.today().strftime("%m.%Y")
    return f"{month_year} Statsig Order Form - {safe_account}.pdf"


def fmt_money(amount: float) -> str:
    return f"${amount:,.2f}"


def product_options_by_warehouse(warehouse_type: str) -> List[str]:
    if warehouse_type == "Warehouse Native":
        return WAREHOUSE_NATIVE_OTHER_PRODUCTS + WAREHOUSE_NATIVE_EXPERIMENTATION_OPTIONS
    if warehouse_type in {"Credit", "Credit/Usage Based", "Usage Based Pricing"}:
        return USAGE_BASED_REQUIRED_PRODUCTS
    return CLOUD_PRODUCTS


def default_unit_for_service(service: str) -> str:
    if service == "Experimentation":
        return "Billable Events"
    if service == "Session Replay":
        return "Sessions"
    return "N/A"


def is_experimentation_service(service: str, warehouse_type: str = "") -> bool:
    cloud_only = {"Experimentation"}
    warehouse_native_only = {
        "Experimentation: Analysis Only",
        "Experimentation: Analysis + Assignment",
    }
    if warehouse_type == "Cloud":
        return service in cloud_only
    if warehouse_type == "Warehouse Native":
        return service in warehouse_native_only
    return service in (cloud_only | warehouse_native_only)


def default_usage_for_service(service: str) -> str:
    if is_experimentation_service(service):
        return ""
    if service in {"Feature Gates and SDKs", "Platform Fee"}:
        return "N/A"
    if service.endswith("Support"):
        return "N/A"
    return "N/A"


def parse_numeric_value(value: str):
    cleaned = str(value).replace("$", "").replace(",", "").strip()
    if cleaned == "":
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def format_usage_commitment_value(value) -> str:
    raw = str(value).strip()
    if raw == "" or raw.upper() == "N/A":
        return "N/A" if raw.upper() == "N/A" else ""
    parsed = parse_whole_number(raw)
    if parsed <= 0 and raw not in {"0", "0.0", "0.00"}:
        return raw
    return f"{parsed:,}"


def build_services_rows(
    selected_products: List[str], support_tier: str, existing_rows: List[Dict[str, str]]
) -> List[Dict[str, str]]:
    existing_rows = normalize_service_rows(existing_rows)
    existing_by_service: Dict[str, Dict[str, str]] = {}
    for row in existing_rows:
        existing_service = str(row.get("service", "")).strip()
        if existing_service:
            existing_by_service[existing_service] = row

    rows: List[Dict[str, str]] = []
    for product in selected_products:
        service_name = f"{support_tier} Support" if product == "Support" else product
        existing = existing_by_service.get(service_name, {})
        rows.append(
            {
                "subscription_period": "",
                "service": service_name,
                "annual_usage_commitment": format_usage_commitment_value(
                    existing.get("annual_usage_commitment", default_usage_for_service(service_name))
                ),
                "unit": str(existing.get("unit", default_unit_for_service(service_name))),
                "annual_service_fee": normalize_fee_value(existing.get("annual_service_fee", 0.0)),
            }
        )
    return rows


def normalize_service_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    normalized: List[Dict[str, str]] = []
    for row in rows:
        normalized.append(
            {
                "subscription_period": str(
                    row.get("subscription_period", row.get("Subscription Period", ""))
                ),
                "service": str(row.get("service", row.get("Service", ""))),
                "annual_usage_commitment": format_usage_commitment_value(
                    row.get(
                        "annual_usage_commitment",
                        row.get("Annual Usage Commitment", ""),
                    )
                ),
                "unit": str(row.get("unit", row.get("Unit", ""))),
                "annual_service_fee": normalize_fee_value(
                    row.get("annual_service_fee", row.get("Annual Service Fee", ""))
                ),
            }
        )
    return normalized


def rows_from_editor(value) -> List[Dict[str, str]]:
    if isinstance(value, list):
        return normalize_service_rows(value)
    if hasattr(value, "to_dict"):
        try:
            records = value.to_dict(orient="records")
            return normalize_service_rows(records)
        except TypeError:
            pass
    return normalize_service_rows([])


def sort_rows_by_fee_desc(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    normalized = normalize_service_rows(rows)
    # Keep support as the final row; sort all other rows by descending fee.
    return sorted(
        normalized,
        key=lambda row: (
            str(row.get("service", "")).strip().endswith("Support"),
            -normalize_fee_value(row.get("annual_service_fee", 0.0)),
        ),
    )


def is_whole_number(value: str) -> bool:
    cleaned = str(value).replace(",", "").strip()
    return bool(re.fullmatch(r"\d+", cleaned))


def is_numeric_amount(value: str) -> bool:
    cleaned = str(value).replace("$", "").replace(",", "").strip()
    if cleaned == "":
        return False
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def normalize_fee_value(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace("$", "").replace(",", "").strip()
    if cleaned == "":
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def format_fee_display(value) -> str:
    return f"${normalize_fee_value(value):,.2f}"


def validate_services_rows(rows: List[Dict[str, str]], warehouse_type: str) -> List[str]:
    errors: List[str] = []

    for idx, row in enumerate(rows, start=1):
        service = str(row.get("service", "")).strip()
        usage = str(row.get("annual_usage_commitment", "")).strip()
        fee = str(row.get("annual_service_fee", "")).strip()

        if not service:
            errors.append(f"Row {idx}: Service is required.")
            continue

        if not is_numeric_amount(fee):
            errors.append(f"Row {idx} ({service}): Annual Service Fee is required and must be numeric.")

        if warehouse_type in {"Credit", "Credit/Usage Based"} and not is_whole_number(usage):
            errors.append(f"Row {idx} ({service}): Credits must be a whole number.")
        elif is_experimentation_service(service, warehouse_type):
            if not usage or usage.upper() == "N/A" or not is_whole_number(usage):
                errors.append(
                    f"Row {idx} ({service}): Annual Usage Commitment must be a whole number."
                )
        elif usage.upper() != "N/A":
            errors.append(
                f"Row {idx} ({service}): Annual Usage Commitment must be N/A for non-experimentation products."
            )

    return errors


def parse_whole_number(value: str) -> int:
    cleaned = str(value).replace(",", "").strip()
    if not re.fullmatch(r"\d+", cleaned):
        return 0
    return int(cleaned)


def find_row_by_service(rows: List[Dict[str, str]], service_name: str) -> Dict[str, str]:
    for row in rows:
        if str(row.get("service", "")).strip() == service_name:
            return row
    return {}


def compute_excess_usage_rate(rows: List[Dict[str, str]], warehouse_type: str) -> str:
    if warehouse_type == "Cloud":
        exp_row = find_row_by_service(rows, "Experimentation")
        if not exp_row:
            return "N/A"
        total_price = safe_money(str(exp_row.get("annual_service_fee", "")))
        usage_commitment = parse_whole_number(str(exp_row.get("annual_usage_commitment", "")))
        if usage_commitment <= 0:
            return "N/A"
        rate = (total_price / usage_commitment) * 1000
        return f"{rate:.4f}"

    if warehouse_type == "Warehouse Native":
        exp_row = find_row_by_service(rows, "Experimentation: Analysis Only")
        if not exp_row:
            exp_row = find_row_by_service(rows, "Experimentation: Analysis + Assignment")
        if not exp_row:
            return "N/A"
        total_price = safe_money(str(exp_row.get("annual_service_fee", "")))
        usage_commitment = parse_whole_number(str(exp_row.get("annual_usage_commitment", "")))
        if usage_commitment <= 0:
            return "N/A"
        rate = total_price / usage_commitment
        return f"{rate:.2f}"

    if warehouse_type in {"Credit", "Credit/Usage Based"}:
        warehouse_native_row = find_row_by_service(rows, "Warehouse Native")
        if not warehouse_native_row:
            return "N/A"
        total_price = safe_money(str(warehouse_native_row.get("annual_service_fee", "")))
        credit_amount = parse_whole_number(
            str(warehouse_native_row.get("annual_usage_commitment", ""))
        )
        if credit_amount <= 0:
            return "N/A"
        rate = total_price / credit_amount
        return f"{rate:.4f}"

    return "N/A"


def build_usage_terms_for_products(
    warehouse_type: str,
    selected_products: List[str],
    excess_usage_rate: str,
) -> str:
    excess_usage_rate_display = fmt_money(safe_money(excess_usage_rate))
    has_feature_gates = "Feature Gates and SDKs" in selected_products
    has_session_replay = "Session Replay" in selected_products
    cloud_experimentation_set = {
        "Experimentation",
        "Experimentation: Analysis Only",
        "Experimentation: Analysis + Assignment",
    }
    has_cloud_experimentation = any(p in selected_products for p in cloud_experimentation_set)
    # Support both possible labels the workflow has used.
    has_wn_analysis_assignment = any(
        p in selected_products
        for p in [
            "Experimentation: Analysis + Assignment",
            "Warehouse Native: Experimentation (Analysis + Assignment)",
        ]
    )
    has_wn_analysis_only = any(
        p in selected_products
        for p in ["Experimentation: Analysis Only", "Warehouse Native: Analysis"]
    )

    session_replay_term = (
        "Customer will have access to 50,000 recorded user sessions on a rolling 30-day basis, for a total "
        "of 600,000 during the Paid Subscription Term. New sessions above 50,000 in a 30-day window will not "
        "be recorded or stored. Customer can control session recording frequency by adjusting sample rate."
    )

    if warehouse_type == "Cloud":
        cloud_short_fg_only = (
            "Customer may use up to 100,000,000,000 non-analytic Feature Gate and config checks through all "
            "server and client-side SDKs during each subscription period."
        )
        cloud_long_fg_plus_exp = (
            "Customer has access to the number of billable events specified in the table above during the "
            "applicable subscription period (\"Annual Usage Commitment\"). Unused billable events expire at the "
            "end of the applicable subscription period and cannot be rolled over to a future subscription period.\n\n"
            "Statsig records a billable event when Customer's application uses Statsig SDKs or APIs to check "
            "the value of an experiment, analytics-enabled gate, or layer. Statsig deduplicates exposure events "
            "for identical users and features or experiments within each hour on a client-side SDK and within "
            "each minute on a server-side SDK.\n\n"
            "Statsig also records a billable event each time Customer logs an event to Statsig via Statsig SDKs, "
            "ingests a metric, or computes a custom metric. Customer can add one event dimension for each logged "
            "event, without incurring an additional billable event. For every additional dimension added, an extra "
            "log event will be recorded.\n\n"
            "Checks for experiments that result in no allocation (i.e., if the experiment hasn't commenced or has "
            "concluded) or Feature Gates that are deactivated (i.e., fully launched or discarded without any rule "
            "evaluation) do not generate billable events.\n\n"
            "Customer may use up to 100,000,000,000 non-analytic Feature Gate and config checks through all server "
            "and client-side SDKs during each subscription period.\n\n"
            "If Customer exceeds the Annual Usage Commitment during the applicable subscription period, Customer "
            f"shall be invoiced monthly in arrears for any excess usage at a rate of {excess_usage_rate_display} per "
            "1,000 billable events."
        )
        base = ""
        if has_feature_gates and has_cloud_experimentation:
            base = cloud_long_fg_plus_exp
        elif has_feature_gates and not has_cloud_experimentation:
            base = cloud_short_fg_only
        if has_session_replay:
            return f"{base}\n\n{session_replay_term}" if base else session_replay_term
        return base

    if warehouse_type == "Warehouse Native":
        wn_analysis_assignment_text = (
            "If Customer has access to the number of experiments specified in the table above during the "
            "applicable subscription period (\"Annual Usage Commitment\"). Unused experiments expire at the end "
            "of the applicable subscription period and cannot be rolled over to a future subscription period.\n\n"
            "Experiment is defined as an experiment or a feature rollout that results in metric lifts being "
            "computed. Feature rollouts configured to not compute metric lifts are not counted as experiments. "
            "The same experiment being restarted is not counted as a new experiment.\n\n"
            "Customer may use up to 100,000,000,000 Feature Gate checks through all server and client-side SDKs. "
            "Customers may also forward up to 100,000,000,000 exposures to their data warehouse using Statsig's SDKs.\n\n"
            "If Customer exceeds the Annual Usage Commitment during the applicable subscription period, Customer "
            f"shall be invoiced monthly in arrears for any excess usage at a rate of {excess_usage_rate_display} per experiment."
        )
        wn_analysis_only_text = (
            "Customer has access to the number of experiments specified in the table above during the applicable "
            "subscription period (\"Annual Usage Commitment\"). Unused experiments expire at the end of the "
            "applicable subscription period and cannot be rolled over to a future subscription period.\n\n"
            "Experiment is defined as an experiment or a feature rollout that results in metric lifts being computed. "
            "Feature rollouts configured to not compute metric lifts are not counted as experiments. The same "
            "experiment being restarted is not counted as a new experiment.\n\n"
            "If Customer exceeds the Annual Usage Commitment during the applicable subscription period, Customer "
            f"shall be invoiced monthly in arrears for any excess usage at a rate of {excess_usage_rate_display} per experiment."
        )
        base = ""
        if has_wn_analysis_assignment:
            base = wn_analysis_assignment_text
        elif has_wn_analysis_only:
            base = wn_analysis_only_text
        if has_session_replay:
            return f"{base}\n\n{session_replay_term}" if base else session_replay_term
        return base

    return ""


def products_from_services_rows(rows: List[Dict[str, str]]) -> List[str]:
    products: List[str] = []
    for row in normalize_service_rows(rows):
        service = str(row.get("service", "")).strip()
        if not service:
            continue
        if service.endswith("Support"):
            products.append("Support")
            continue
        products.append(service)
    return list(dict.fromkeys(products))


def normalize_special_terms(selected: List[str]) -> List[str]:
    if not selected:
        return []
    cleaned = [str(s).strip() for s in selected if str(s).strip()]
    non_none = [s for s in cleaned if s.lower() != "none"]
    return list(dict.fromkeys(non_none))


def render_merged_table_preview(
    rows: List[Dict[str, str]],
    column_spec: List[Tuple[str, str]],
    period_text: str,
    total: str,
) -> None:
    row_count = max(1, len(rows))
    first_key, first_label = column_spec[0]
    other_cols = column_spec[1:]

    header_html = "".join(f"<th>{label}</th>" for _, label in column_spec)
    body_rows = []
    for i in range(row_count):
        row = rows[i] if i < len(rows) else {}
        cells = []
        if i == 0:
            cells.append(f"<td rowspan='{row_count}'>{period_text}</td>")
        for key, _ in other_cols:
            cells.append(f"<td>{row.get(key, '')}</td>")
        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    total_colspan = max(1, len(column_spec) - 1)
    total_row = (
        f"<tr class='total-row'><td colspan='{total_colspan}'>Total:</td><td>{total}</td></tr>"
        if len(column_spec) > 1
        else f"<tr class='total-row'><td>{total}</td></tr>"
    )

    html = f"""
    <style>
      .merged-preview table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 0.9rem;
      }}
      .merged-preview th, .merged-preview td {{
        border: 1px solid #b8b8b8;
        padding: 8px;
        text-align: center;
      }}
      .merged-preview th {{
        background: #1f4675;
        color: #ffffff;
        font-weight: 700;
      }}
      .merged-preview .total-row td {{
        font-weight: 700;
      }}
    </style>
    <div class="merged-preview">
      <table>
        <thead><tr>{header_html}</tr></thead>
        <tbody>
          {''.join(body_rows)}
          {total_row}
        </tbody>
      </table>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


def resolve_pdf_fonts() -> Tuple[str, str]:
    if "OpenSans" in pdfmetrics.getRegisteredFontNames() and "OpenSans-Bold" in pdfmetrics.getRegisteredFontNames():
        return "OpenSans", "OpenSans-Bold"

    font_candidates = [
        (
            "/Users/matti/Desktop/Codex/assets/OpenSans-Regular.ttf",
            "/Users/matti/Desktop/Codex/assets/OpenSans-Bold.ttf",
        ),
        (
            "/Library/Fonts/OpenSans-Regular.ttf",
            "/Library/Fonts/OpenSans-Bold.ttf",
        ),
        (
            "/System/Library/Fonts/Supplemental/OpenSans-Regular.ttf",
            "/System/Library/Fonts/Supplemental/OpenSans-Bold.ttf",
        ),
    ]
    for regular_path, bold_path in font_candidates:
        if os.path.exists(regular_path) and os.path.exists(bold_path):
            pdfmetrics.registerFont(TTFont("OpenSans", regular_path))
            pdfmetrics.registerFont(TTFont("OpenSans-Bold", bold_path))
            return "OpenSans", "OpenSans-Bold"
    return "Helvetica", "Helvetica-Bold"


def resolve_pdf_italic_font() -> str:
    if "OpenSans-Italic" in pdfmetrics.getRegisteredFontNames():
        return "OpenSans-Italic"
    for path in [
        "/Users/matti/Desktop/Codex/assets/OpenSans-Italic.ttf",
        "/Library/Fonts/OpenSans-Italic.ttf",
        "/System/Library/Fonts/Supplemental/OpenSans-Italic.ttf",
    ]:
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont("OpenSans-Italic", path))
            return "OpenSans-Italic"
    return "Helvetica-Oblique"


def find_header_logo_path() -> str:
    for path in [
        "/Users/matti/Desktop/Statsig_Logo_Transparent_Black.png",
        "/Users/matti/Desktop/Logo-min.png",
        "/Users/matti/Desktop/Statsig Logo.jpeg",
        "/Users/matti/Desktop/Codex/assets/statsig-header.png",
        "/Users/matti/Desktop/Codex/assets/statsig_header.png",
        "/Users/matti/Desktop/Codex/assets/statsig-logo.png",
        "/Users/matti/Desktop/Codex/assets/statsig_logo.png",
    ]:
        if os.path.exists(path):
            return path
    return ""


def find_signature_logo_path() -> str:
    for path in [
        "/Users/matti/Desktop/Logo-min.png",
        "/Users/matti/Desktop/Statsig Logo.jpeg",
        "/Users/matti/Desktop/Codex/assets/statsig-mark.png",
        "/Users/matti/Desktop/Codex/assets/statsig-logo.png",
    ]:
        if os.path.exists(path):
            return path
    return ""


def draw_agreement_section(
    c: canvas.Canvas,
    width: float,
    height: float,
    font_regular: str,
    font_bold: str,
    agreement_section: str,
    start_y: float,
) -> float:
    y = start_y
    min_bottom_margin = 40

    def ensure_space(required_height: float) -> None:
        nonlocal y
        if y - required_height < min_bottom_margin:
            c.showPage()
            c.setFillColor(colors.black)
            y = height - 56

    ensure_space(24)
    c.setFillColor(colors.black)
    c.setFont(font_bold, 12)
    c.drawString(36, y, "Agreement")

    y -= 20
    ensure_space(14)
    c.setFont(font_regular, 10)
    if (
        "Enterprise Subscription Agreement" in agreement_section
        and "Data Processing Addendum" in agreement_section
    ):
        segments = [
            ("This Order Form is subject to the ", None),
            ("Enterprise Subscription Agreement", "https://www.statsig.com/enterprise-terms"),
            (" and ", None),
            ("Data Processing Addendum", "https://statsig.com/legal/online-dpa"),
            (
                " (together, the“MSA”) governing Customer’s use of the Services described herein.",
                None,
            ),
        ]
        max_text_width = width - 72
        left_x = 36
        right_x = left_x + max_text_width
        line_height = 12
        cursor_x = left_x
        cursor_y = y

        rich_tokens: List[Tuple[str, str | None]] = []
        for text, url in segments:
            for token in re.findall(r"\S+\s*", text):
                rich_tokens.append((token, url))

        for token, url in rich_tokens:
            token_w = c.stringWidth(token, font_regular, 10)
            if cursor_x + token_w > right_x and cursor_x > left_x:
                cursor_y -= line_height
                if cursor_y <= 40:
                    c.showPage()
                    c.setFillColor(colors.black)
                    c.setFont(font_regular, 10)
                    cursor_y = height - 56
                cursor_x = left_x

            if url:
                c.setFillColor(colors.HexColor("#194b7d"))
                c.drawString(cursor_x, cursor_y, token)
                c.linkURL(
                    url,
                    (cursor_x, cursor_y - 1, cursor_x + token_w, cursor_y + 10),
                    relative=0,
                    thickness=0,
                )
            else:
                c.setFillColor(colors.black)
                c.drawString(cursor_x, cursor_y, token)
            cursor_x += token_w
        c.setFillColor(colors.black)
        y = cursor_y - 24
    else:
        max_text_width = width - 72
        wrapped_agreement = wrap_text_to_width(c, agreement_section, max_text_width, font_regular, 10)
        for line in wrapped_agreement:
            ensure_space(12)
            c.drawString(36, y, line)
            y -= 12
        y -= 12

    paragraph_lines = [
        "The pricing and terms in this Order Form are Statsig, LLC's Proprietary Information. All fees are in U.S. dollars and",
        "exclude all taxes. Customer is responsible for all applicable taxes, including, but not limited to, U.S. sales,",
        "withholding tax, GST, and VAT. Capitalized terms not defined in this Order Form have the meanings assigned in the MSA.",
        "If a direct conflict exists between this Order Form and the MSA, the terms of this Order Form will control.",
        "",
        "The parties through their duly authorized representative agree to the terms of this Order Form, effective as of",
        "last signature date.",
    ]
    c.setFont(font_regular, 10)
    for line in paragraph_lines:
        if line == "":
            y -= 12
        else:
            ensure_space(12)
            c.setFont(font_regular, 10)
            c.drawString(36, y, line)
            y -= 12

    y -= 30
    ensure_space(132)
    c.setStrokeColor(colors.HexColor("#aaaaaa"))
    c.line(36, y + 23, width - 36, y + 23)

    statsig_font_size = 12
    c.setFont(font_bold, statsig_font_size)
    c.drawString(36, y, "Statsig, LLC:")

    c.setFont(font_bold, 12)
    c.drawString(330, y, "Customer:")

    mid_x = width / 2
    left_line_end = mid_x - 16
    right_line_end = width - 36

    y -= 28
    c.setFont(font_regular, 10)
    c.drawString(36, y, "By:")
    c.line(62, y - 2, left_line_end, y - 2)
    c.drawString(330, y, "By:")
    c.line(356, y - 2, right_line_end, y - 2)

    y -= 24
    c.drawString(36, y, "Name:")
    c.line(70, y - 2, left_line_end, y - 2)
    c.drawString(330, y, "Name:")
    c.line(364, y - 2, right_line_end, y - 2)

    y -= 24
    c.drawString(36, y, "Title:")
    c.line(66, y - 2, left_line_end, y - 2)
    c.drawString(330, y, "Title:")
    c.line(360, y - 2, right_line_end, y - 2)

    y -= 24
    c.drawString(36, y, "Date:")
    c.line(66, y - 2, left_line_end, y - 2)
    c.drawString(330, y, "Date:")
    c.line(360, y - 2, right_line_end, y - 2)
    return y


def _draw_text(
    c: canvas.Canvas,
    x: float,
    y: float,
    label: str,
    value: str,
    bold_label: bool = True,
    font_regular: str = "Helvetica",
    font_bold: str = "Helvetica-Bold",
) -> None:
    if bold_label:
        c.setFont(font_bold, 10)
        c.drawString(x, y, label)
        c.setFont(font_regular, 10)
    else:
        c.setFont(font_regular, 10)
        c.drawString(x, y, label)
    c.drawString(x + (c.stringWidth(label, font_bold, 10) if bold_label else 0), y, value)


def wrap_text_to_width(
    c: canvas.Canvas,
    text: str,
    max_width: float,
    font_name: str,
    font_size: int,
) -> List[str]:
    content = (text or "").strip()
    if not content:
        return [""]
    words = content.split()
    lines: List[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if c.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def preserve_input_lines(text: str) -> List[str]:
    if text is None:
        return [""]
    lines = str(text).splitlines()
    return lines if lines else [""]


def create_branded_pdf(order: OrderData, services: List[Dict[str, str]], warehouse_type: str) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=LETTER)
    width, height = LETTER

    blue = colors.HexColor("#1f4675")
    font_regular, font_bold = resolve_pdf_fonts()
    font_italic = resolve_pdf_italic_font()

    start = first_of_month(parse_date(order.start_date))
    end = compute_end_date(start, order.subscription_term_months)
    start_str = display_date(start)
    end_str = display_date(end)

    # Header (first page only)
    logo_path = find_header_logo_path()
    if logo_path:
        c.drawImage(
            logo_path,
            (width - 190) / 2,
            height - 92,
            width=190,
            height=38,
            preserveAspectRatio=True,
            mask="auto",
        )
    else:
        c.setFont(font_bold, 26)
        c.setFillColor(colors.black)
        c.drawCentredString(width / 2, height - 82, "STATSIG")
    c.setFont(font_bold, 12)
    c.setFillColor(colors.black)
    c.drawCentredString(width / 2, height - 112, "Order Form")
    header_line_y = height - 112
    if order.expiration_date:
        exp_date = parse_date(order.expiration_date).strftime("%m.%d.%Y")
        c.setFont(font_italic, 10)
        c.setFillColor(colors.black)
        header_line_y -= 16
        c.drawCentredString(
            width / 2,
            header_line_y,
            f"Order Form expires on {exp_date} without signature",
        )
    if order.opportunity_type == "Expansion/Upsell" and order.addendum_effective_date:
        eff_date = parse_date(order.addendum_effective_date).strftime("%m/%d/%Y")
        c.setFont(font_regular, 11)
        c.setFillColor(colors.black)
        header_line_y -= 16
        c.drawCentredString(
            width / 2,
            header_line_y,
            f"This order form is an addendum to the order form with an effective date of {eff_date}.",
        )

    c.setStrokeColor(colors.HexColor("#aaaaaa"))
    divider_y = header_line_y - 18
    c.line(36, divider_y, width - 36, divider_y)

    # Customer info
    y = divider_y - 40
    c.setFillColor(colors.black)
    c.setFont(font_bold, 12)
    c.drawString(36, y, "Customer Information")

    y -= 26
    _draw_text(c, 36, y, "Customer: ", order.account_name, font_regular=font_regular, font_bold=font_bold)

    y -= 26
    _draw_text(c, 36, y, "Customer Contact: ", order.primary_contact_name, font_regular=font_regular, font_bold=font_bold)
    _draw_text(c, 312, y, "Billing Email: ", order.billing_email, font_regular=font_regular, font_bold=font_bold)

    y -= 16
    _draw_text(c, 36, y, "Email: ", order.primary_contact_email, font_regular=font_regular, font_bold=font_bold)

    y -= 24
    c.setFont(font_bold, 10)
    c.drawString(36, y, "Ship To Address:")
    c.drawString(312, y, "Bill to Address:")
    ship_lines = preserve_input_lines(order.shipping_address)
    bill_lines = preserve_input_lines(order.billing_address)
    max_addr_lines = max(len(ship_lines), len(bill_lines))
    c.setFont(font_regular, 10)
    y -= 14
    for i in range(max_addr_lines):
        c.drawString(36, y, ship_lines[i] if i < len(ship_lines) else "")
        c.drawString(312, y, bill_lines[i] if i < len(bill_lines) else "")
        y -= 12

    y -= 6
    c.setStrokeColor(colors.HexColor("#aaaaaa"))
    c.line(36, y, width - 36, y)

    # Terms
    y -= 23
    c.setFont(font_bold, 12)
    c.setFillColor(colors.black)
    c.drawString(36, y, "Terms")

    left_label_x = 36
    right_label_x = 312
    # Use a one-space-equivalent visual gap between label and value.
    gap_width = c.stringWidth(" ", font_regular, 10)

    y -= 25
    left_label = "Paid Subscription Term Start Date:"
    c.setFont(font_bold, 10)
    c.drawString(left_label_x, y, left_label)
    c.setFillColor(colors.black)
    c.setFont(font_regular, 10)
    left_value_x = left_label_x + c.stringWidth(left_label, font_bold, 10) + gap_width
    c.drawString(left_value_x, y, start_str)

    right_label = "Billing Frequency:"
    c.setFillColor(colors.black)
    c.setFont(font_bold, 10)
    c.drawString(right_label_x, y, right_label)
    c.setFillColor(colors.black)
    c.setFont(font_regular, 10)
    right_value_x = right_label_x + c.stringWidth(right_label, font_bold, 10) + gap_width
    c.drawString(right_value_x, y, order.billing_frequency)

    y -= 16
    left_label = "Paid Subscription Term End Date:"
    c.setFillColor(colors.black)
    c.setFont(font_bold, 10)
    c.drawString(left_label_x, y, left_label)
    c.setFillColor(colors.black)
    c.setFont(font_regular, 10)
    left_value_x = left_label_x + c.stringWidth(left_label, font_bold, 10) + gap_width
    c.drawString(left_value_x, y, end_str)

    right_label = "Payment Terms:"
    c.setFillColor(colors.black)
    c.setFont(font_bold, 10)
    c.drawString(right_label_x, y, right_label)
    c.setFillColor(colors.black)
    c.setFont(font_regular, 10)
    right_value_x = right_label_x + c.stringWidth(right_label, font_bold, 10) + gap_width
    c.drawString(right_value_x, y, order.payment_terms)

    y -= 16
    left_label = "Payment Method:"
    c.setFillColor(colors.black)
    c.setFont(font_bold, 10)
    c.drawString(left_label_x, y, left_label)
    c.setFont(font_regular, 10)
    left_value_x = left_label_x + c.stringWidth(left_label, font_bold, 10) + gap_width
    c.drawString(left_value_x, y, order.payment_method)

    right_label = "PO (if applicable):"
    c.setFont(font_bold, 10)
    c.drawString(right_label_x, y, right_label)
    c.setFont(font_regular, 10)
    right_value_x = right_label_x + c.stringWidth(right_label, font_bold, 10) + gap_width
    c.drawString(right_value_x, y, order.po_number)

    y -= 12
    c.setStrokeColor(colors.HexColor("#aaaaaa"))
    c.line(36, y, width - 36, y)

    # Services header
    y -= 26
    c.setFillColor(colors.black)
    c.setFont(font_bold, 12)
    c.drawString(36, y, "Services")

    # Dynamic table area
    top = y - 16
    left = 36
    right = width - 36
    table_w = right - left

    services = normalize_service_rows(services)
    column_spec = table_columns_for_warehouse(warehouse_type)
    headers = [label for _, label in column_spec]
    widths = [0.25, 0.23, 0.21, 0.12, 0.19] if len(column_spec) == 5 else [0.27, 0.33, 0.20, 0.20]
    col_w = [table_w * w for w in widths]

    rows = max(1, len(services))
    header_wrapped = [
        wrap_text_to_width(c, h, col_w[i] - 8, font_bold, 10) for i, h in enumerate(headers)
    ]
    head_h = max(24, max(len(lines) for lines in header_wrapped) * 10 + 8)

    row_wrapped: List[List[List[str]]] = []
    row_heights: List[float] = []
    for row in services:
        wrapped_cells: List[List[str]] = []
        max_lines = 1
        for i, (key, _) in enumerate(column_spec[1:], start=1):
            value = row.get(key, "")
            if key == "annual_service_fee" and value != "":
                value = fmt_money(safe_money(str(value)))
            lines = wrap_text_to_width(c, str(value), col_w[i] - 8, font_regular, 10)
            wrapped_cells.append(lines)
            max_lines = max(max_lines, len(lines))
        row_wrapped.append(wrapped_cells)
        row_heights.append(max(22, max_lines * 10 + 8))

    total_row_h = 22
    table_h = head_h + sum(row_heights) + total_row_h

    c.setStrokeColor(colors.black)
    c.setFillColor(blue)
    c.rect(left, top - head_h, table_w, head_h, stroke=1, fill=1)

    x = left
    c.setFillColor(colors.white)
    c.setFont(font_bold, 10)
    for i, lines in enumerate(header_wrapped):
        block_h = len(lines) * 10
        y0 = top - ((head_h - block_h) / 2) - 8
        for j, line in enumerate(lines):
            c.drawCentredString(x + col_w[i] / 2, y0 - j * 10, line)
        x += col_w[i]

    # Header column lines
    x = left
    for w in col_w[:-1]:
        x += w
        c.setStrokeColor(colors.black)
        c.line(x, top, x, top - head_h)

    # Body anchors
    y_total_divider = top - head_h - sum(row_heights)
    c.line(left, y_total_divider, right, y_total_divider)

    # Row values
    c.setFillColor(colors.black)
    c.setFont(font_regular, 10)
    period_text = f"{start_str} - {end_str}"
    merged_top = top - head_h
    merged_bottom = y_total_divider
    merged_mid_y = ((merged_top + merged_bottom) / 2) - 4
    c.drawCentredString(left + col_w[0] / 2, merged_mid_y, period_text)
    if period_text.strip():
        c.rect(left, merged_bottom, col_w[0], merged_top - merged_bottom, stroke=1, fill=0)

    y_cursor = top - head_h
    for idx, row in enumerate(services):
        row_top = y_cursor
        row_h = row_heights[idx]
        cell_lines = row_wrapped[idx]
        x = left + col_w[0]
        for i, lines in enumerate(cell_lines):
            key = column_spec[i + 1][0]
            raw_value = row.get(key, "")
            if key == "annual_service_fee" and raw_value != "":
                raw_value = fmt_money(safe_money(str(raw_value)))
            block_h = len(lines) * 10
            line_y = row_top - ((row_h - block_h) / 2) - 8
            for line in lines:
                if key == "annual_service_fee":
                    c.drawRightString(x + col_w[i + 1] - 4, line_y, line)
                else:
                    c.drawCentredString(x + (col_w[i + 1] / 2), line_y, line)
                line_y -= 10
            if str(raw_value).strip():
                c.rect(x, row_top - row_h, col_w[i + 1], row_h, stroke=1, fill=0)
            x += col_w[i + 1]
        y_cursor -= row_h

    # Total row
    total = sum(safe_money(str(row.get("annual_service_fee", ""))) for row in services)
    total_y = y_total_divider - 15
    last_col_left = left + sum(col_w[:-1])
    last_col_center = last_col_left + (col_w[-1] / 2)
    prev_col_center = (
        left + sum(col_w[:-2]) + (col_w[-2] / 2) if len(col_w) > 1 else last_col_center
    )
    c.setFont(font_regular, 10)
    if len(col_w) > 1:
        c.setFillColor(colors.black)
        c.drawCentredString(prev_col_center, total_y, "Total:")
        c.rect(last_col_left - col_w[-2], y_total_divider - total_row_h, col_w[-2], total_row_h, stroke=1, fill=0)
    c.setFillColor(colors.black)
    c.drawRightString(last_col_left + col_w[-1] - 4, total_y, fmt_money(total))
    c.rect(last_col_left, y_total_divider - total_row_h, col_w[-1], total_row_h, stroke=1, fill=0)

    # Footer
    y_after = top - table_h - 22
    c.setFillColor(colors.black)
    c.setFont(font_regular, 10)
    c.drawString(36, y_after, "For information on the Statsig platform, refer to https://docs.statsig.com/")
    y_after -= 14
    c.drawString(36, y_after, "For information on Statsig support, refer to https://docs.statsig.com/support-options")

    y_after -= 10
    c.setStrokeColor(colors.HexColor("#aaaaaa"))
    c.line(36, y_after, width - 36, y_after)

    computed_excess_rate = compute_excess_usage_rate(services, warehouse_type)
    auto_usage_terms = build_usage_terms_for_products(
        warehouse_type,
        products_from_services_rows(services),
        computed_excess_rate,
    )
    usage_text = auto_usage_terms if auto_usage_terms else (order.usage_terms or "")
    y_after -= 23

    c.setFont(font_bold, 12)
    c.setFillColor(colors.black)
    c.drawString(36, y_after, "Usage Terms")

    y_after -= 20
    c.setFont(font_regular, 10)
    c.setFillColor(colors.black)

    max_text_width = width - 72
    paragraph_blocks = usage_text.split("\n")
    for paragraph in paragraph_blocks:
        if not paragraph.strip():
            y_after -= 7
            continue
        lines = wrap_text_to_width(c, paragraph, max_text_width, font_regular, 10)
        for line in lines:
            if y_after <= 40:
                c.showPage()
                c.setFillColor(colors.black)
                c.setFont(font_regular, 10)
                y_after = height - 56
            c.drawString(36, y_after, line)
            y_after -= 12

    y_after -= 5
    c.setStrokeColor(colors.HexColor("#aaaaaa"))
    c.line(36, y_after, width - 36, y_after)

    # Agreement section follows directly after Usage Terms.
    if order.terms_type == "MSA":
        msa_date_display = order.msa_execution_date.strip() if order.msa_execution_date.strip() else "MM/DD/YYYY"
        agreement_section = (
            "This Order Form is subject to the Master Subscription Agreement (“MSA”) between Customer and "
            f"Statsig, Inc., executed on {msa_date_display} , governing Customer’s use of the Service described herein."
        )
    else:
        agreement_section = (
            "This Order Form is subject to the Enterprise Subscription Agreement and Data Processing Addendum "
            "(together, the“MSA”) governing Customer’s use of the Services described herein."
        )
    draw_agreement_section(
        c,
        width,
        height,
        font_regular,
        font_bold,
        agreement_section,
        y_after - 23,
    )
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def wrap_text(text: str, max_chars: int) -> List[str]:
    words = text.split()
    if not words:
        return [""]

    lines: List[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if len(candidate) <= max_chars:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def customer_fields_complete(order: OrderData) -> bool:
    upsell_ok = (
        order.opportunity_type != "Expansion/Upsell"
        or bool(order.addendum_effective_date.strip())
    )
    return all(
        [
            order.account_name.strip(),
            order.primary_contact_name.strip(),
            order.primary_contact_email.strip(),
            order.billing_email.strip(),
            order.shipping_address.strip(),
            order.billing_address.strip(),
            order.opportunity_type.strip(),
            upsell_ok,
        ]
    )


def terms_fields_complete(order: OrderData, start_valid: bool) -> bool:
    billing_id_ok = (
        order.payment_method == "Bank Transfer" or bool(order.billing_id.strip())
    )
    return all(
        [
            start_valid,
            order.subscription_term_months >= 1,
            order.billing_frequency.strip(),
            order.payment_terms.strip(),
            order.payment_method.strip(),
            billing_id_ok,
        ]
    )


def table_columns_for_warehouse(warehouse_type: str) -> List[Tuple[str, str]]:
    if warehouse_type in {"Credit", "Credit/Usage Based", "Usage Based Pricing"}:
        return [
            ("subscription_period", "Subscription Term"),
            ("service", "Services"),
            ("annual_usage_commitment", "Credits"),
            ("annual_service_fee", "Annual Fee"),
        ]
    return [
        ("subscription_period", "Subscription Period"),
        ("service", "Service"),
        ("annual_usage_commitment", "Annual Usage Commitment"),
        ("unit", "Unit"),
        ("annual_service_fee", "Annual Service Fee"),
    ]


def main() -> None:
    st.set_page_config(page_title="Statsig Order Form Generator", layout="wide")
    st.markdown(
        """
        <style>
        :root {
          --brand-blue: #194b7d;
        }
        .stButton > button,
        .stDownloadButton > button {
          background-color: var(--brand-blue) !important;
          border-color: var(--brand-blue) !important;
          color: #ffffff !important;
        }
        .stButton > button:hover,
        .stDownloadButton > button:hover {
          background-color: var(--brand-blue) !important;
          border-color: var(--brand-blue) !important;
          color: #ffffff !important;
          filter: brightness(0.95);
        }
        input[type="radio"],
        input[type="checkbox"] {
          accent-color: var(--brand-blue);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("Statsig Order Form Generator")
    st.write("Create a branded order form PDF from guided inputs or uploaded source docs.")

    if "services_rows" not in st.session_state:
        st.session_state.services_rows = DEFAULT_SERVICES
    if "order_data" not in st.session_state:
        st.session_state.order_data = asdict(OrderData())
    if "form_step" not in st.session_state:
        st.session_state.form_step = 1
    if "uploaded_once" not in st.session_state:
        st.session_state.uploaded_once = False
    if "input_source_mode" not in st.session_state:
        st.session_state.input_source_mode = "Q/A"
    if "previous_input_source_mode" not in st.session_state:
        st.session_state.previous_input_source_mode = "Q/A"
    if "warehouse_type" not in st.session_state:
        st.session_state.warehouse_type = "Cloud"
    if "selected_products" not in st.session_state:
        st.session_state.selected_products = []
    if "support_tier" not in st.session_state:
        st.session_state.support_tier = "Premium"
    if "product_signature" not in st.session_state:
        st.session_state.product_signature = ""
    if "warehouse_native_experimentation" not in st.session_state:
        st.session_state.warehouse_native_experimentation = ""
    if "show_table_errors" not in st.session_state:
        st.session_state.show_table_errors = False

    order = OrderData(**st.session_state.order_data)

    st.caption(f"Step {st.session_state.form_step} of 4")

    if st.session_state.form_step == 1:
        st.subheader("Input Source")
        st.session_state.input_source_mode = st.radio(
            "Choose input source", ["Q/A", "Upload Document"], horizontal=True
        )
        if (
            st.session_state.previous_input_source_mode == "Q/A"
            and st.session_state.input_source_mode == "Upload Document"
        ):
            order.opportunity_type = ""
        st.session_state.previous_input_source_mode = st.session_state.input_source_mode

        if st.session_state.input_source_mode == "Upload Document":
            uploaded = st.file_uploader(
                "Upload source file (.txt, .pdf, .docx)",
                type=["txt", "pdf", "docx"],
            )
            if uploaded and not st.session_state.uploaded_once:
                raw = extract_text_from_upload(uploaded)
                if raw.strip():
                    extracted = extract_order_fields(raw)
                    order.account_name = extracted.account_name or order.account_name
                    order.primary_contact_name = extracted.primary_contact_name or order.primary_contact_name
                    order.primary_contact_email = extracted.primary_contact_email or order.primary_contact_email
                    order.billing_email = extracted.billing_email or order.billing_email
                    order.shipping_address = extracted.shipping_address or order.shipping_address
                    order.billing_address = extracted.billing_address or order.billing_address
                    order.start_date = extracted.start_date or order.start_date
                    order.subscription_term_months = extracted.subscription_term_months or order.subscription_term_months
                    order.billing_frequency = extracted.billing_frequency or order.billing_frequency
                    order.payment_terms = extracted.payment_terms or order.payment_terms
                    order.payment_method = extracted.payment_method or order.payment_method
                    order.po_number = extracted.po_number or order.po_number
                    order.usage_terms = extracted.usage_terms or order.usage_terms
                    st.session_state.uploaded_once = True
                    with st.expander("Extracted text", expanded=False):
                        st.text_area("Source", raw, height=180)
                    st.info("Auto-detected values loaded. Review before continuing.")
                else:
                    st.warning("No readable text found. Use Q/A mode or another file.")
        else:
            if not order.opportunity_type:
                order.opportunity_type = "New Logo"

        st.subheader("Customer Information")
        col_left, col_right = st.columns(2)
        with col_left:
            order.account_name = st.text_input("Customer/Account Name", value=order.account_name)
            order.primary_contact_name = st.text_input("Primary Contact", value=order.primary_contact_name)
            order.primary_contact_email = st.text_input(
                "Primary Contact Email", value=order.primary_contact_email
            )
            same_email = st.checkbox(
                "Billing email same as primary contact email",
                value=(not order.billing_email or order.billing_email == order.primary_contact_email),
            )
            order.billing_email = st.text_input(
                "Billing Email",
                value=(order.primary_contact_email if same_email else order.billing_email),
                disabled=same_email,
            )
        with col_right:
            opportunity_options = ["New Logo", "Renewal", "Expansion/Upsell"]
            order.opportunity_type = st.radio(
                "Opportunity Type",
                opportunity_options,
                index=(
                    opportunity_options.index(order.opportunity_type)
                    if order.opportunity_type in opportunity_options
                    else (0 if st.session_state.input_source_mode == "Q/A" else None)
                ),
                horizontal=True,
            )
            if order.opportunity_type == "Expansion/Upsell":
                effective_default = (
                    parse_date(order.addendum_effective_date)
                    if order.addendum_effective_date
                    else None
                )
                selected_effective = st.date_input(
                    "Upsell Effective Date",
                    value=effective_default,
                    format="MM/DD/YYYY",
                )
                order.addendum_effective_date = (
                    display_date(selected_effective) if selected_effective is not None else ""
                )
            else:
                order.addendum_effective_date = ""
            order.shipping_address = st.text_area(
                "Ship To Address",
                value=order.shipping_address,
                height=100,
                help="Address field",
            )
            same_address = st.checkbox(
                "Bill to address same as shipping address",
                value=(
                    not order.billing_address
                    or order.billing_address == order.shipping_address
                    or order.billing_address == "Same as shipping address"
                ),
            )
            order.billing_address = st.text_area(
                "Bill To Address",
                value=(order.shipping_address if same_address else order.billing_address),
                height=100,
                disabled=same_address,
                help="Address field",
            )

        step1_ok = customer_fields_complete(order)
        if not step1_ok:
            st.warning("Complete all required customer information fields to continue.")
            if (
                order.opportunity_type == "Expansion/Upsell"
                and not order.addendum_effective_date.strip()
            ):
                st.warning("Upsell Effective Date is required when Opportunity Type is Expansion/Upsell.")
        if st.button("Continue to Terms", type="primary", disabled=not step1_ok):
            st.session_state.form_step = 2
            st.session_state.order_data = asdict(order)
            st.rerun()

    elif st.session_state.form_step == 2:
        st.subheader("Terms")
        left_col, right_col = st.columns(2)
        with left_col:
            start = parse_date(order.start_date) if order.start_date else default_subscription_start_date()
            selected_start = st.date_input(
                "Subscription Start Date",
                value=start,
                format="MM/DD/YYYY",
            )
            start_valid = selected_start.day == 1
            if not start_valid:
                st.error("Subscription Start Date must be the 1st day of a month.")
            order.start_date = display_date(first_of_month(selected_start))

            order.subscription_term_months = int(
                st.number_input(
                    "Subscription Term (Months)",
                    min_value=1,
                    step=1,
                    value=max(1, int(order.subscription_term_months or 12)),
                )
            )
            computed_end = compute_end_date(parse_date(order.start_date), order.subscription_term_months)
            st.text_input("Computed Subscription End Date", value=display_date(computed_end), disabled=True)

        with right_col:
            payment_method_options = ["Bank Transfer", "AWS Billing", "GCP Billing", "Azure Billing"]
            order.payment_method = st.selectbox(
                "Payment Method",
                payment_method_options,
                index=payment_method_options.index(order.payment_method)
                if order.payment_method in payment_method_options
                else 0,
            )
            if order.payment_method != "Bank Transfer":
                order.billing_id = st.text_input(
                    "Billing ID (Required for cloud billing methods)",
                    value=order.billing_id,
                )
                if not order.billing_id.strip():
                    st.error("Billing ID is required unless Payment Method is Bank Transfer.")
            else:
                order.billing_id = ""

            billing_frequency_options = ["Annual", "Semi-Annual", "Quarterly"]
            order.billing_frequency = st.selectbox(
                "Billing Frequency",
                billing_frequency_options,
                index=billing_frequency_options.index(order.billing_frequency)
                if order.billing_frequency in billing_frequency_options
                else 0,
            )

            payment_term_options = ["Net 30", "Net 45", "Net 60", "Net 90"]
            order.payment_terms = st.selectbox(
                "Payment Terms",
                payment_term_options,
                index=payment_term_options.index(order.payment_terms)
                if order.payment_terms in payment_term_options
                else 0,
            )
            order.po_number = st.text_input("PO Number (Optional)", value=order.po_number)

        nav1, nav2 = st.columns(2)
        with nav1:
            if st.button("Back to Customer Information"):
                st.session_state.form_step = 1
                st.session_state.order_data = asdict(order)
                st.rerun()
        with nav2:
            step2_ok = terms_fields_complete(order, start_valid)
            if st.button("Continue to Product Selection", type="primary", disabled=not step2_ok):
                st.session_state.form_step = 3
                st.session_state.order_data = asdict(order)
                st.rerun()

    elif st.session_state.form_step == 3:
        st.subheader("Product Selection")

        st.session_state.warehouse_type = st.selectbox(
            "1. Warehouse type",
            WAREHOUSE_TYPES,
            index=WAREHOUSE_TYPES.index(st.session_state.warehouse_type)
            if st.session_state.warehouse_type in WAREHOUSE_TYPES
            else 0,
        )

        product_options = product_options_by_warehouse(st.session_state.warehouse_type)
        if st.session_state.warehouse_type == "Usage Based Pricing":
            st.session_state.warehouse_type = "Credit/Usage Based"

        if st.session_state.warehouse_type == "Credit/Usage Based":
            st.session_state.selected_products = USAGE_BASED_REQUIRED_PRODUCTS
            st.info(
                "Credit requires: Warehouse Native and Platform Fee. Support is required via Support Tier."
            )
            st.markdown("**Products (required):**")
            for p in USAGE_BASED_REQUIRED_PRODUCTS:
                st.write(f"- {p}")
            st.session_state.warehouse_native_experimentation = ""
        elif st.session_state.warehouse_type == "Warehouse Native":
            st.markdown("**Products**")
            st.session_state.warehouse_native_experimentation = st.radio(
                "Experimentation Product (select one)",
                options=["None"] + WAREHOUSE_NATIVE_EXPERIMENTATION_OPTIONS,
                index=(
                    (["None"] + WAREHOUSE_NATIVE_EXPERIMENTATION_OPTIONS).index(
                        st.session_state.warehouse_native_experimentation
                    )
                    if st.session_state.warehouse_native_experimentation
                    in (["None"] + WAREHOUSE_NATIVE_EXPERIMENTATION_OPTIONS)
                    else 0
                ),
            )
            wn_other_options = list(WAREHOUSE_NATIVE_OTHER_PRODUCTS)
            if st.session_state.warehouse_native_experimentation == "Experimentation: Analysis Only":
                wn_other_options = [p for p in wn_other_options if p != "Feature Gates and SDKs"]
            warehouse_native_others_default = [
                p
                for p in st.session_state.selected_products
                if p in wn_other_options
            ]
            warehouse_native_others = st.multiselect(
                "Other Products (multi-select)",
                options=wn_other_options,
                default=warehouse_native_others_default,
            )
            if (
                st.session_state.warehouse_native_experimentation
                == "Experimentation: Analysis + Assignment"
                and "Feature Gates and SDKs" not in warehouse_native_others
            ):
                warehouse_native_others.append("Feature Gates and SDKs")
            st.session_state.selected_products = list(warehouse_native_others)
            if st.session_state.warehouse_native_experimentation != "None":
                st.session_state.selected_products.append(st.session_state.warehouse_native_experimentation)
        else:
            st.markdown("**Products**")
            default_selected = [
                p for p in st.session_state.selected_products if p in product_options
            ]
            st.session_state.selected_products = st.multiselect(
                "Products",
                options=product_options,
                default=default_selected,
            )
            st.session_state.warehouse_native_experimentation = ""

        st.session_state.support_tier = st.selectbox(
            "Support Tier",
            SUPPORT_TIERS,
            index=SUPPORT_TIERS.index(st.session_state.support_tier)
            if st.session_state.support_tier in SUPPORT_TIERS
            else 0,
        )
        selected_products_with_support = [
            p for p in st.session_state.selected_products if p != "Support"
        ] + ["Support"]

        product_signature = (
            f"{st.session_state.warehouse_type}|"
            f"{','.join(selected_products_with_support)}|"
            f"{st.session_state.warehouse_native_experimentation}|"
            f"{st.session_state.support_tier}"
        )
        if product_signature != st.session_state.product_signature:
            st.session_state.services_rows = build_services_rows(
                selected_products_with_support,
                st.session_state.support_tier,
                st.session_state.services_rows,
            )
            st.session_state.product_signature = product_signature

        st.markdown("**Table**")
        st.caption(
            "Please update the table with the price of each product and the annual usage commitment. "
            "The contract total and excess usage rate will be auto calculated. "
            "(rendering is slighly slow when updating the numbers, please be patient)"
        )
        column_spec = table_columns_for_warehouse(st.session_state.warehouse_type)
        input_column_spec = [
            (key, label) for key, label in column_spec if key != "subscription_period"
        ]
        rows_for_editor = normalize_service_rows(st.session_state.services_rows)
        for row in rows_for_editor:
            row["annual_service_fee"] = format_fee_display(row.get("annual_service_fee", 0.0))

        edited_services = st.data_editor(
            rows_for_editor,
            num_rows="fixed",
            use_container_width=True,
            column_order=[key for key, _ in input_column_spec],
            column_config={
                key: (
                    st.column_config.TextColumn(label, disabled=True)
                    if key == "subscription_period"
                    else st.column_config.TextColumn(
                        label, help="Use whole numbers for experimentation; other rows are N/A."
                    )
                    if key == "annual_usage_commitment"
                    else st.column_config.TextColumn(label)
                    if key == "annual_service_fee"
                    else label
                )
                for key, label in input_column_spec
            },
            key="services_editor",
        )
        services_df = rows_from_editor(edited_services)
        # Subscription period/term is derived from terms and must stay locked.
        period_text_for_rows = (
            f"{display_date(parse_date(order.start_date))} - "
            f"{display_date(compute_end_date(parse_date(order.start_date), order.subscription_term_months))}"
        )
        for row in services_df:
            row["subscription_period"] = period_text_for_rows
            service_name = str(row.get("service", "")).strip()
            if st.session_state.warehouse_type != "Credit/Usage Based" and not is_experimentation_service(
                service_name, st.session_state.warehouse_type
            ):
                row["annual_usage_commitment"] = "N/A"
            else:
                row["annual_usage_commitment"] = format_usage_commitment_value(
                    row.get("annual_usage_commitment", "")
                )
            row["annual_service_fee"] = normalize_fee_value(row.get("annual_service_fee", 0.0))
        sorted_services_df = sort_rows_by_fee_desc(services_df)
        current_order_sig = [
            (
                str(row.get("service", "")),
                str(row.get("annual_usage_commitment", "")),
                str(row.get("unit", "")),
                str(row.get("annual_service_fee", "")),
            )
            for row in services_df
        ]
        sorted_order_sig = [
            (
                str(row.get("service", "")),
                str(row.get("annual_usage_commitment", "")),
                str(row.get("unit", "")),
                str(row.get("annual_service_fee", "")),
            )
            for row in sorted_services_df
        ]
        st.session_state.services_rows = sorted_services_df
        if current_order_sig != sorted_order_sig:
            st.rerun()
        services_df = sorted_services_df
        st.session_state.order_data = asdict(order)
        contract_total = sum(safe_money(str(row.get("annual_service_fee", ""))) for row in services_df)
        excess_usage_rate = compute_excess_usage_rate(
            services_df, st.session_state.warehouse_type
        )
        st.markdown(
            f"""
            <div style='text-align: right; margin-top: 8px;'>
              <div><strong>Total: {fmt_money(contract_total)}</strong></div>
              <div>Excess usage rate: {excess_usage_rate}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        service_validation_errors = validate_services_rows(
            services_df, st.session_state.warehouse_type
        )
        if st.session_state.show_table_errors and service_validation_errors:
            for err in service_validation_errors:
                st.error(err)

        nav1, nav2 = st.columns(2)
        with nav1:
            if st.button("Back to Terms"):
                st.session_state.form_step = 2
                st.session_state.order_data = asdict(order)
                st.rerun()
        with nav2:
            can_continue = len(selected_products_with_support) > 0
            if not can_continue:
                st.warning("Select at least one product to continue.")
            clicked_continue = st.button(
                "Continue",
                type="primary",
                disabled=not can_continue,
            )
            if clicked_continue:
                if service_validation_errors:
                    st.session_state.show_table_errors = True
                    st.rerun()
                else:
                    st.session_state.show_table_errors = False
                    st.session_state.order_data = asdict(order)
                    st.session_state.form_step = 4
                    st.rerun()

    else:
        st.subheader("Agreement")
        services_df = normalize_service_rows(st.session_state.services_rows)
        selected_products_with_support = [
            p for p in st.session_state.selected_products if p != "Support"
        ] + ["Support"]
        excess_usage_rate = compute_excess_usage_rate(
            services_df, st.session_state.warehouse_type
        )
        column_spec = table_columns_for_warehouse(st.session_state.warehouse_type)
        total = sum(safe_money(str(row.get("annual_service_fee", ""))) for row in services_df)
        computed_end_date = display_date(
            compute_end_date(parse_date(order.start_date), order.subscription_term_months)
        )
        terms_options = ["Online", "MSA"]
        order.terms_type = st.radio(
            "Terms",
            terms_options,
            index=terms_options.index(order.terms_type)
            if order.terms_type in terms_options
            else 0,
            horizontal=True,
        )
        if order.terms_type == "MSA":
            msa_default = parse_date(order.msa_execution_date) if order.msa_execution_date else None
            selected_msa_date = st.date_input(
                "MSA Execution Date",
                value=msa_default,
                format="MM/DD/YYYY",
            )
            order.msa_execution_date = (
                display_date(selected_msa_date) if selected_msa_date is not None else ""
            )
        else:
            order.msa_execution_date = ""

        special_term_options = [
            "price cap",
            "no auto-renewal",
            "no logo rights",
            "one time discount",
        ]
        order.special_terms = normalize_special_terms(
            st.multiselect(
                "Legal: Special Terms",
                options=special_term_options,
                default=[t for t in order.special_terms if t in special_term_options],
                placeholder="None",
            )
        )
        usage_terms_by_products = build_usage_terms_for_products(
            st.session_state.warehouse_type,
            selected_products_with_support,
            excess_usage_rate,
        )
        if usage_terms_by_products:
            order.usage_terms = usage_terms_by_products
        else:
            order.usage_terms = ", ".join(order.special_terms) if order.special_terms else "None"

        today_date = date.today()
        order.expiration_enabled = st.checkbox(
            "Include Expiration Date",
            value=bool(order.expiration_enabled or order.expiration_date),
        )
        if order.expiration_enabled:
            expiration_default = parse_date(order.expiration_date) if order.expiration_date else None
            if expiration_default is not None and expiration_default < today_date:
                expiration_default = today_date
            selected_expiration = st.date_input(
                "Expiration Date",
                value=expiration_default,
                format="MM/DD/YYYY",
                help="Required when enabled",
            )
            if selected_expiration is not None and selected_expiration < today_date:
                st.error("Expiration Date must be today or later.")
                order.expiration_date = display_date(today_date)
            else:
                order.expiration_date = (
                    display_date(selected_expiration) if selected_expiration is not None else ""
                )
        else:
            order.expiration_date = ""

        st.markdown("**Document Preview**")
        preview_pdf = create_branded_pdf(
            order,
            services_df,
            st.session_state.warehouse_type,
        )
        preview_b64 = base64.b64encode(preview_pdf).decode("utf-8")
        components.html(
            f"""
            <iframe
              src="data:application/pdf;base64,{preview_b64}"
              width="100%"
              height="700px"
              style="border: 1px solid #ddd; border-radius: 8px;"
            ></iframe>
            """,
            height=720,
        )
        st.subheader("Preview Data")
        st.json(
            {
                **asdict(order),
                "computed_subscription_end_date": computed_end_date,
                "warehouse_type": st.session_state.warehouse_type,
                "selected_products": selected_products_with_support,
                "support_tier": st.session_state.support_tier,
                "columns": [label for _, label in column_spec],
                "services": services_df,
                "calculated_total": fmt_money(total),
                "terms_type": order.terms_type,
            }
        )

        nav1, nav2 = st.columns(2)
        with nav1:
            if st.button("Back to Product Selection"):
                st.session_state.form_step = 3
                st.session_state.order_data = asdict(order)
                st.rerun()
        with nav2:
            if st.button("Generate PDF", type="primary"):
                if order.terms_type == "MSA" and not order.msa_execution_date:
                    st.error("MSA Execution Date is required when Terms is MSA.")
                    st.stop()
                if order.expiration_enabled and not order.expiration_date:
                    st.error("Expiration Date is required when enabled.")
                    st.stop()
                st.session_state.order_data = asdict(order)
                pdf_bytes = create_branded_pdf(order, services_df, st.session_state.warehouse_type)
                output_filename = build_output_filename(order.account_name)
                st.success("Branded PDF generated.")
                st.download_button(
                    "Download PDF",
                    data=pdf_bytes,
                    file_name=output_filename,
                    mime="application/pdf",
                )


if __name__ == "__main__":
    main()
