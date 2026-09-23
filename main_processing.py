
import os
import time
import re
import json
from pathlib import Path
from typing import Optional, List, Dict, Union, Tuple
from enum import Enum

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field, field_validator
from openai import OpenAI

# Import zero-drop email extractor module
from email_extractor import fetch_extracted_email_payloads


# =====================================================================
# 0. HELPER FUNCTIONS & CONFIGURATION LOADER
# =====================================================================

def clean_excel_string(val) -> str:
    """Strips non-printable ASCII control characters that crash OpenPyXL."""
    if val is None:
        return ""
    val_str = str(val)
    return re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F-\x9F]", "", val_str)


def normalize_port_code(port_raw: str) -> str:
    """
    Cleans raw location/port string into a standard 5-character UN/LOCODE string format.
    Validates 5-character UN/LOCODE patterns (2-letter country + 3-character location).
    """
    if not port_raw:
        return "UNKNOWN"
    
    port_clean = str(port_raw).upper().strip()
    port_clean = re.sub(r"[^A-Z0-9]", "", port_clean)

    if re.match(r"^[A-Z]{2}[A-Z0-9]{3}$", port_clean):
        return port_clean

    return port_clean if port_clean else "UNKNOWN"


def consolidate_surcharge_legs(legs: List["SingleLegItem"]) -> List["SingleLegItem"]:
    """
    Consolidates multiple leg rows for surcharges/local fees sharing the same route,
    container equipment, and commodity into a SINGLE leg row with consolidated rates.
    """
    if not legs or len(legs) <= 1:
        return legs

    consolidated_dict: Dict[Tuple[str, str, str, str], SingleLegItem] = {}

    for leg in legs:
        pol = normalize_port_code(leg.pol_code)
        pod = normalize_port_code(leg.pod_code)
        cntr = str(leg.container_type or "").strip().upper()
        comm = str(leg.commodity or "").strip().lower()

        group_key = (pol, pod, cntr, comm)

        if group_key not in consolidated_dict:
            consolidated_dict[group_key] = leg.model_copy()
        else:
            existing_leg = consolidated_dict[group_key]
            
            # Consolidate rates
            existing_rate = str(existing_leg.customer_requested_rate or "—").strip()
            new_rate = str(leg.customer_requested_rate or "—").strip()

            if new_rate != "—" and new_rate not in existing_rate:
                if existing_rate == "—":
                    existing_leg.customer_requested_rate = new_rate
                else:
                    existing_leg.customer_requested_rate = f"{existing_rate} + {new_rate}"

            # Combine Value-Added Services if present
            if leg.vas:
                existing_leg.vas = list(set((existing_leg.vas or []) + leg.vas))

            # Combine Cargo Badges/Flags if present
            if leg.flags:
                existing_leg.flags = list(set((existing_leg.flags or []) + leg.flags))

    # Re-sequence leg sequence numbers
    merged_legs = list(consolidated_dict.values())
    for idx, m_leg in enumerate(merged_legs, start=1):
        m_leg.seq_num = idx

    return merged_legs


def load_config(config_path: str = "config.json") -> dict:
    config = {}
    path = Path(config_path)

    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    config = json.loads(content)
        except json.JSONDecodeError as e:
            print(f"[Config Error]: {e}")

    gateway_url = config.get("TIGER_AI_GATEWAY_URL") or os.getenv("TIGER_AI_GATEWAY_URL")
    api_key = config.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")

    if gateway_url and gateway_url.endswith("/chat/completions"):
        gateway_url = gateway_url.replace("/chat/completions", "")

    return {
        "OPENAI_API_KEY": api_key,
        "TIGER_AI_GATEWAY_URL": gateway_url,
        "MODEL_NAME": config.get("MODEL_NAME", "gpt-4o-mini"),
        "LOCAL_INBOX_DIR": config.get("LOCAL_INBOX_DIR", "./outlook_inbox"),
        "ACCOUNTS": config.get("ACCOUNTS", [{"account_name": "Outlook_Local_Desk"}])
    }


def calculate_objective_confidence_score(req: "MultiLegQuotationRequest") -> float:
    """Calculates a deterministic confidence score (0.10 to 1.00) based on field presence."""
    base_score = 1.00

    if not req.customer_name or req.customer_name.strip().lower() in ["unknown customer", "unknown", "—", ""]:
        base_score -= 0.15

    if not req.legs:
        return 0.10

    total_legs = len(req.legs)
    leg_penalties = 0.0

    for leg in req.legs:
        if not leg.pol_code or str(leg.pol_code).upper() in ["UNKNOWN", "NONE", "—", ""]:
            leg_penalties += 0.15
        if not leg.pod_code or str(leg.pod_code).upper() in ["UNKNOWN", "NONE", "—", ""]:
            leg_penalties += 0.15
        if not leg.container_type or str(leg.container_type).upper() in ["UNSPECIFIED", "UNKNOWN", "NONE", "—", ""]:
            leg_penalties += 0.15
        if not leg.cntr_count or str(leg.cntr_count).strip() in ["0", "—", "None", ""]:
            leg_penalties += 0.15
        if str(leg.weight_tons).strip() in ["—", "N/A", "UNKNOWN", "NONE", "0", ""]:
            leg_penalties += 0.15
        if not leg.commodity or str(leg.commodity).upper() in ["UNSPECIFIED", "UNKNOWN", "NONE", "—", ""]:
            leg_penalties += 0.10

    avg_leg_penalty = leg_penalties / total_legs
    final_score = base_score - avg_leg_penalty

    return round(max(0.10, min(1.00, final_score)), 2)


# =====================================================================
# 1. DATA CONTRACT & PYDANTIC SCHEMAS (SCENARIOS 1, 2, 3)
# =====================================================================

class CommunicationTopic(str, Enum):
    NEW_RATE_REQUEST = "New Rate Request"
    RATE_REVISION = "Rate Request Revision / Update"
    CUSTOMER_NEGOTIATION = "Customer Negotiation"
    DEMURRAGE_DETENTION = "Demurrage & Detention (D&D) Inquiry"
    OPERATIONAL_MATTERS = "Operational Matters (Doc/Schedule/Booking)"
    GENERAL_THREAD_REPLY = "General Thread Reply"
    EXCLUDED_NOISE = "Excluded / Non-Quotation"


class StatusEnum(str, Enum):
    NEW = "New"
    SENT_PRICER = "Sent ➔ Pricer"
    SENT_CUSTOMER = "Sent ➔ Customer"
    SENT_BOTH = "Sent ➔ Both"


class RequestType(str, Enum):
    NEW = "New Request"
    AMENDMENT = "Amendment"
    FOLLOW_UP = "Follow-Up / Negotiation"


class SingleLegItem(BaseModel):
    seq_num: int = Field(description="Leg sequence number (1, 2, 3...)")
    pol_code: Optional[str] = Field(
        default="UNKNOWN", 
        description="Port of Loading official 5-character UN/LOCODE (e.g., CNSHA, PECALL, DEHAM, SGSIN, BRSSZ, AEJEA)."
    )
    pod_code: Optional[str] = Field(
        default="UNKNOWN", 
        description="Port of Discharge official 5-character UN/LOCODE (e.g., NLRTM, GBFXT, USNYC, BEANR, KRPUS, ZADUR)."
    )
    container_type: Optional[str] = Field(
        default="—", 
        description="Standardized container type (e.g., 40'HQ, 20'GP, 40'GP, 40'RF, ISO Tank)."
    )
    
    weight_tons: Optional[Union[float, str]] = Field(
        default="—", 
        description="Gross cargo weight per container in metric tons (MT/T). Extract numbers or expressions like '24.5 MT', '18T', '22.0'. Output '—' if not specified."
    )
    
    commodity: Optional[str] = Field(default=None, description="Commodity description. Set None if missing.")
    cntr_count: int = Field(default=1, description="Container count requested")
    
    service_voyage: Optional[str] = Field(
        default="—", 
        description="Liner Service Name and Voyage string formatted strictly as 'SERVICE / V.VOYAGE' (e.g., 'AEX1 / V.142W', 'MEX2 / V.018E'). Set '—' if missing."
    )
    
    system_proposed_rate: Optional[str] = Field(
        default="—", 
        description="ALWAYS output '—'. This column is strictly reserved for calculation by downstream pricing engine."
    )
    
    customer_requested_rate: Optional[str] = Field(
        default="—", 
        description="Standardized dollar rate (e.g. '$2600', '$3400/40HQ'). Set '—' if no rate is present. Give the aggregated rate. Do not create new row for each rate."
    )
    
    flags: List[str] = Field(
        default_factory=list, 
        description="Cargo badges: 'DG', 'NOR', 'SOC', 'Tank', 'MTY', 'Reefer', 'OOG'. Leave empty [] if standard cargo."
    )
    
    vas: Optional[List[str]] = Field(
        default=[],
        description="List of specific Value-Added Services requested in the email (e.g., ['Genset', 'Fumigation', 'Customs Clearance']). Do NOT return a count."
    )
    
    free_time: Optional[str] = Field(
        default="—", 
        description=(
            "Free time requirements formatted strictly on two lines using 'Origin:' and 'Destination:'. "
            "Rules:\n"
            "- If both are available: 'Origin: <time>\nDestination: <time>'\n"
            "- If only Origin is available: 'Origin: <time>\nDestination:'\n"
            "- If only Destination is available: 'Origin:\nDestination: <time>'\n"
            "- If neither is available, output '—'. Never put '-' or '—' after Origin: or Destination: labels."
        )
    )

    @field_validator("free_time", mode="before")
    @classmethod
    def clean_free_time(cls, v: Optional[str]) -> str:
        if not v or str(v).strip().upper() in ["—", "N/A", "NONE", "UNKNOWN", "STANDARD", "UNSPECIFIED", ""]:
            return "—"
        
        v_str = str(v).strip()
        
        # Remove any literal '—' or '-' following Origin: or Destination: labels
        v_str = re.sub(r"Origin:\s*[\—\-]", "Origin:", v_str)
        v_str = re.sub(r"Destination:\s*[\—\-]", "Destination:", v_str)
        
        # If neither Origin nor Destination contains actual text, output '—'
        has_origin_val = bool(re.search(r"Origin:\s*[\w\d]+", v_str, re.IGNORECASE))
        has_dest_val = bool(re.search(r"Destination:\s*[\w\d]+", v_str, re.IGNORECASE))
        
        if not has_origin_val and not has_dest_val:
            return "—"
            
        return v_str

    @field_validator("pol_code", "pod_code", mode="before")
    @classmethod
    def enforce_unlocode(cls, v: Optional[str]) -> str:
        """Enforces 5-character UN/LOCODE standardization."""
        return normalize_port_code(v)

    @field_validator("container_type", mode="before")
    @classmethod
    def standardize_container_type(cls, v: Optional[str]) -> str:
        """Standardizes equipment codes into uniform ocean industry strings."""
        if not v or str(v).strip().upper() in ["—", "N/A", "NONE", "UNKNOWN", "UNSPECIFIED", ""]:
            return "—"
        
        v_upper = str(v).strip().upper()
        
        if re.search(r"40['\s]*(HQ|HC|HIGH\s*CUBE)", v_upper):
            return "40'HQ"
        elif re.search(r"20['\s]*(GP|DV|DRY|STD|STANDARD)", v_upper) or v_upper == "20":
            return "20'GP"
        elif re.search(r"40['\s]*(GP|DV|DRY|STD|STANDARD)", v_upper) or v_upper == "40":
            return "40'GP"
        elif re.search(r"40['\s]*(RF|REEFER|RH)", v_upper):
            return "40'RF"
        elif re.search(r"20['\s]*(RF|REEFER)", v_upper):
            return "20'RF"
        elif re.search(r"TANK|ISO\s*TANK", v_upper):
            return "ISO Tank"
        elif re.search(r"40['\s]*(OT|OPEN\s*TOP)", v_upper):
            return "40'OT"
        elif re.search(r"20['\s]*(OT|OPEN\s*TOP)", v_upper):
            return "20'OT"
        elif re.search(r"40['\s]*(FR|FLAT\s*RACK)", v_upper):
            return "40'FR"
        elif re.search(r"20['\s]*(FR|FLAT\s*RACK)", v_upper):
            return "20'FR"
        
        return v_upper

    @field_validator("customer_requested_rate", mode="before")
    @classmethod
    def standardize_dollar_rate(cls, v: Optional[str]) -> str:
        """Enforces strict $ + numeric amount output with no extra currency text."""
        if not v or str(v).strip() in ["—", "N/A", "None", "UNKNOWN", ""]:
            return "—"
        
        v_str = str(v).strip()
        
        num_match = re.search(r"(\d{1,3}(?:,\d{3})*|\d+)", v_str)
        if not num_match:
            return "—"
        
        clean_num = num_match.group(1).replace(",", "")
        
        eq_match = re.search(r"/(20GP|40GP|40HQ|40HC|20RF|40RF)", v_str.upper())
        eq_suffix = eq_match.group(0) if eq_match else ""
        
        return f"${clean_num}{eq_suffix}"

    @field_validator("weight_tons", mode="before")
    @classmethod
    def clean_weight(cls, v: Optional[Union[float, str]]) -> str:
        if v is None or str(v).strip() in ["—", "N/A", "None", "UNKNOWN", ""]:
            return "—"
        
        v_str = str(v).strip()
        v_upper = v_str.upper()

        num_match = re.search(r"(\d+(?:\.\d+)?)", v_str)
        if not num_match:
            return "—"
            
        val_float = float(num_match.group(1))

        if re.search(r"\b(MT|METRIC\s*TONS?|TONS?|T)\b", v_upper):
            pass
        elif re.search(r"\bKGS?\b", v_upper):
            val_float = val_float / 1000.0
        elif re.search(r"\bLBS?\b", v_upper):
            val_float = val_float / 2204.62
        elif val_float > 100:
            val_float = val_float / 1000.0

        return f"{val_float:.1f}" if val_float % 1 != 0 else f"{int(val_float)}"

    @field_validator("service_voyage", mode="before")
    @classmethod
    def format_service_voyage(cls, v: Optional[str]) -> str:
        if not v or str(v).strip() in ["—", "N/A", "None", "UNKNOWN", ""]:
            return "—"
        
        v = str(v).strip().upper()
        if re.match(r"^[A-Z0-9\-]+\s*/\s*V\.[A-Z0-9]+$", v):
            return v
        
        match = re.search(r"([A-Z0-9\-]+)[\s/]+(?:V(?:OY)?\.?\s*)?([0-9]+[A-Z]?)", v)
        if match:
            return f"{match.group(1)} / V.{match.group(2)}"
            
        return v

    @field_validator("system_proposed_rate", mode="before")
    @classmethod
    def enforce_blank_system_rate(cls, v: Optional[str]) -> str:
        return "—"


class MultiLegQuotationRequest(BaseModel):
    request_id: str = Field(default="", description="Shared Request ID (e.g., REQ-1001 or REQ-1001-v2)")
    customer_name: str = Field(default="Unknown Customer", description="Customer company name")
    
    validity_period: Optional[str] = Field(
        default="—", 
        description="Quotation validity date range formatted strictly with years as 'MMM DD, YYYY – MMM DD, YYYY'. Output '—' if missing."
    )

    @field_validator("validity_period", mode="before")
    @classmethod
    def clean_validity(cls, v: Optional[str]) -> str:
        if not v or str(v).strip() in ["—", "N/A", "None", "UNKNOWN", ""]:
            return "—"
        
        v_str = str(v).strip().upper()
        
        matches = re.findall(
            r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+\d{1,2}(?:,\s*|\s+)\d{4}", 
            v_str
        )
        
        if len(matches) != 2:
            return "—"

        return v_str
    
    topic: CommunicationTopic = Field(
        description="Communication topic. Use 'Operational Matters' for feasibility questions like 'Do you service X port?'."
    )
    status: StatusEnum = Field(default=StatusEnum.NEW, description="Workflow status")
    request_type: RequestType = Field(default=RequestType.NEW, description="Type of request")
    
    is_new_rate_request: bool = Field(
        description="Set TRUE ONLY if the message explicitly asks for a price/quote OR provides actionable container parameters. Set FALSE for feasibility chat like 'Do you call at Jebel Ali?'."
    )

    classification_confidence: float = Field(
        default=0.95,
        description="AI classification confidence score (0.10 to 1.00) evaluating certainty for rate request vs chatter classification."
    )
    
    ai_confidence_score: float = Field(
        default=1.00, 
        description="Confidence score calculated deterministically post-extraction based on field completeness."
    )
    assigned_owner: str = Field(default="Asia_Desk_Ops", description="Assigned desk owner")
    
    customer_remarks: Optional[str] = Field(
        default="—", 
        description=(
            "Actionable operational instructions, cargo specifics, or constraints extracted DIRECTLY from the email text. "
            "Examples:\n"
            "- 'Repeat customer, priority handling'\n"
            "- 'Class 8 corrosive, MSDS attached'\n"
            "- 'Continuous temperature log required'\n"
            "- 'Heavy weight — VGM mandatory'\n"
            "- 'Shipper-owned containers, return depot info'\n"
            "- 'Last 3 cargoes log attached'\n"
            "- 'Equipment repo — best rate required'\n"
            "- 'Origin: Bogota; bagged'\n"
            "- 'Class 9, environmentally hazardous'\n"
            "If no special operational instructions exist in the text, output strictly '—'."
        )
    )

    @field_validator("customer_remarks", mode="before")
    @classmethod
    def clean_remarks(cls, v: Optional[str]) -> str:
        if not v or str(v).strip().upper() in ["—", "N/A", "NONE", "UNKNOWN", "UNSPECIFIED", ""]:
            return "—"
        return str(v).strip()

    legs: List[SingleLegItem] = Field(description="List of 1 or more trade legs extracted")


class ThreadQuotationExtraction(BaseModel):
    requests: List[MultiLegQuotationRequest] = Field(
        description="List of all distinct rate requests identified anywhere within the email thread."
    )


# =====================================================================
# 2. EXCEL SEQUENCE RENDERER (DUAL-SHEET AUDIT & DASHBOARD PERSISTENCE)
# =====================================================================

class ExcelSequenceRenderer:
    FILE_PATH = "multileg_sequence_quotation_dashboard.xlsx"
    HEADERS = [
        "ID", "LEG SEQ #", "CUSTOMER", "POL ➔ POD", "VALIDITY", "CONTAINER", 
        "WT (T)", "COMMODITY", "# CNTR", "SERVICE / VOYAGE", 
        "SYSTEM PROPOSED RATE", "CUSTOMER REQUESTED RATE", "FLAGS", 
        "VAS", "FREE TIME", "REMARKS", "STATUS", "CONFIDENCE", "MISSING DATA?"
    ]
    AUDIT_HEADERS = [
        "RECEIVED DATE", "FILE / THREAD ID", "SENDER", "SUBJECT", 
        "COMMUNICATION TOPIC", "CLASSIFICATION VERDICT", "ACTION TAKEN"
    ]

    @classmethod
    def initialize_workbook(cls):
        """Ensures both Dashboard and Audit sheets exist with proper headers, without blank row 1."""
        if Path(cls.FILE_PATH).exists():
            wb = openpyxl.load_workbook(cls.FILE_PATH)
        else:
            wb = openpyxl.Workbook()
            ws_dash = wb.active
            ws_dash.title = "Quotation Intake Dashboard"
            ws_dash.views.sheetView[0].showGridLines = True

        header_font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        dash_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        audit_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")

        # 1. Setup Dashboard Sheet
        ws_dash = wb["Quotation Intake Dashboard"]
        if ws_dash.cell(row=1, column=1).value is None:
            for col_idx, h in enumerate(cls.HEADERS, 1):
                cell = ws_dash.cell(row=1, column=col_idx, value=h)
                cell.font = header_font
                cell.fill = dash_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_dash.row_dimensions[1].height = 28

        # 2. Setup Audit Sheet
        if "Email Ingestion Audit Log" not in wb.sheetnames:
            wb.create_sheet(title="Email Ingestion Audit Log")
        
        ws_audit = wb["Email Ingestion Audit Log"]
        if ws_audit.cell(row=1, column=1).value is None:
            for col_idx, h in enumerate(cls.AUDIT_HEADERS, 1):
                cell = ws_audit.cell(row=1, column=col_idx, value=h)
                cell.font = header_font
                cell.fill = audit_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_audit.row_dimensions[1].height = 28

        wb.save(cls.FILE_PATH)
        return wb

    @classmethod
    def log_audit_record(cls, payload: dict, req: MultiLegQuotationRequest, is_rate_request: bool, created_req_ids: List[str] = None):
        """Logs 100% of emails to the intermediate audit sheet with exact mapped REQ IDs."""
        wb = cls.initialize_workbook()
        ws = wb["Email Ingestion Audit Log"]
        
        row_idx = ws.max_row + 1
        ws.row_dimensions[row_idx].height = 22
        
        status_label = "RATE REQUEST" if is_rate_request else "CHATTER / NON-QUOTATION"
        
        if is_rate_request and created_req_ids:
            req_str = ", ".join(created_req_ids)
            action_label = f"Added to Dashboard ({req_str})"
        elif is_rate_request:
            action_label = f"Added to Dashboard ({req.request_id})"
        else:
            action_label = "Excluded from Dashboard"
        
        status_font = Font(name="Segoe UI", size=9, bold=True, color="15803D" if is_rate_request else "B91C1C")
        data_font = Font(name="Segoe UI", size=9, color="1E293B")
        
        conf_score = getattr(req, "classification_confidence", 0.95)

        audit_vals = [
            clean_excel_string(payload.get("received_date", "N/A")),
            clean_excel_string(payload.get("file_name", payload.get("parent_conversation_id", "N/A"))),
            clean_excel_string(payload.get("sender", "N/A")),
            clean_excel_string(payload.get("subject", "N/A")),
            clean_excel_string(req.topic.value),
            clean_excel_string(status_label),
            clean_excel_string(action_label)
        ]
        
        for col_idx, val in enumerate(audit_vals, 1):
            c = ws.cell(row=row_idx, column=col_idx, value=val)
            c.font = status_font if col_idx == 6 else data_font
            c.alignment = Alignment(horizontal="center" if col_idx in [1, 5, 6, 7] else "left", vertical="center")
            
        col_widths = {1: 18, 2: 25, 3: 25, 4: 35, 5: 25, 6: 22, 7: 30}
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width
            
        wb.save(cls.FILE_PATH)

    @classmethod
    def save_request(cls, req: MultiLegQuotationRequest):
        """Logs valid rate requests to the main Quotation Intake Dashboard."""
        wb = cls.initialize_workbook()
        ws = wb["Quotation Intake Dashboard"]

        start_row = ws.max_row + 1
        id_font = Font(name="Segoe UI", size=9, bold=True, color="DC2626")
        seq_font = Font(name="Segoe UI", size=9, bold=True, color="2563EB")
        data_font = Font(name="Segoe UI", size=9, color="1E293B")
        alert_font = Font(name="Segoe UI", size=9, bold=True, color="B91C1C")

        border_solid = Border(bottom=Side(style='thin', color='CBD5E1'))
        border_dashed = Border(bottom=Side(style='dashed', color='E2E8F0'))

        num_legs = len(req.legs)

        for i, leg in enumerate(req.legs):
            row_idx = start_row + i
            ws.row_dimensions[row_idx].height = 24

            # --- MISSING FIELDS VALIDATION ---
            missing_fields = []
            
            def is_empty_val(val):
                if val is None:
                    return True
                s = str(val).strip().upper()
                return s in ["", "—", "UNKNOWN", "NONE", "UNSPECIFIED", "N/A", "0"]

            if is_empty_val(req.customer_name) or req.customer_name == "Unknown Customer":
                missing_fields.append("Customer")
            if is_empty_val(leg.pol_code):
                missing_fields.append("POL")
            if is_empty_val(leg.pod_code):
                missing_fields.append("POD")
            if is_empty_val(leg.container_type):
                missing_fields.append("Container Type")
            if is_empty_val(leg.cntr_count):
                missing_fields.append("Container Qty")
            if is_empty_val(leg.weight_tons):
                missing_fields.append("Weight")
            if is_empty_val(leg.commodity):
                missing_fields.append("Commodity")

            missing_indicator = f"YES ({', '.join(missing_fields)})" if missing_fields else "NO"

            if isinstance(leg.vas, list) and leg.vas:
                clean_items = [str(v).strip() for v in leg.vas if str(v).strip() and str(v).strip().upper() not in ["NONE", "N/A", "—", "0"]]
                vas_str = ", ".join(clean_items) if clean_items else "—"
            elif isinstance(leg.vas, str) and leg.vas.strip() not in ["", "None", "0", "—", "N/A"]:
                vas_str = leg.vas.strip()
            else:
                vas_str = "—"

            flags_str = " ".join(leg.flags) if leg.flags else "—"
            
            pol_norm = normalize_port_code(leg.pol_code)
            pod_norm = normalize_port_code(leg.pod_code)
            pol_pod_str = f"{pol_norm} ➔ {pod_norm}"

            cntr_type_str = leg.container_type if leg.container_type and str(leg.container_type).strip().upper() not in ["UNSPECIFIED", "UNKNOWN", "NONE", "N/A", ""] else "—"
            free_time_str = leg.free_time if leg.free_time and str(leg.free_time).strip().upper() not in ["STANDARD", "UNSPECIFIED", "UNKNOWN", "NONE", "N/A", ""] else "—"

            row_vals = [
                clean_excel_string(req.request_id if i == 0 else ""),
                leg.seq_num,
                clean_excel_string(req.customer_name if i == 0 else ""),
                clean_excel_string(pol_pod_str),
                clean_excel_string(req.validity_period or "—"),
                clean_excel_string(cntr_type_str),
                clean_excel_string(leg.weight_tons or "—"),
                clean_excel_string(leg.commodity or "—"),
                leg.cntr_count if leg.cntr_count else "—",
                clean_excel_string(leg.service_voyage or "—"),
                "—",
                clean_excel_string(leg.customer_requested_rate or "—"),
                clean_excel_string(flags_str),
                clean_excel_string(vas_str),
                clean_excel_string(free_time_str),
                clean_excel_string(req.customer_remarks or "—"),
                clean_excel_string(req.status.value if i == 0 else ""),
                f"{int(req.ai_confidence_score * 100)}%" if i == 0 else "",
                clean_excel_string(missing_indicator if i == 0 else "")
            ]

            b_style = border_dashed if i < num_legs - 1 else border_solid

            for col_idx, val in enumerate(row_vals, 1):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.font = id_font if col_idx == 1 else (seq_font if col_idx == 2 else data_font)
                c.border = b_style

                if col_idx in [1, 2, 4, 5, 6, 9, 10, 11, 12, 13, 14, 15, 17, 18]:
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif col_idx in [7]:
                    c.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx == 19 and "YES" in str(val):
                    c.font = alert_font
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        col_widths = {
            1: 14, 2: 10, 3: 22, 4: 16, 5: 24, 6: 12, 7: 15, 8: 24,
            9: 10, 10: 16, 11: 20, 12: 22, 13: 12, 14: 8, 15: 24, 16: 28, 17: 16, 18: 12, 19: 25
        }
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        wb.save(cls.FILE_PATH)
        print(f"[Excel Persistence] Logged Request '{req.request_id}' ({num_legs} legs) to {cls.FILE_PATH}")


# =====================================================================
# 3. AI SOURCING AGENT & ROUTE-AWARE MATCHING ENGINE
# =====================================================================

class SequenceSourcingAgent:
    VALID_RATE_TOPICS = [
        CommunicationTopic.NEW_RATE_REQUEST,
        CommunicationTopic.RATE_REVISION,
        CommunicationTopic.CUSTOMER_NEGOTIATION
    ]

    def __init__(self, api_key: str, base_url: str = None, model_name: str = "gpt-4o-mini"):
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.model_name = model_name
        self.active_quotations_db: Dict[str, dict] = {}
        self.processed_email_ids: set = set()
        self.request_seq_counter = 1
        self._load_existing_excel_history()

    def _load_existing_excel_history(self):
        """Loads historical route keys, max ID sequence, AND processed file IDs to prevent duplicate runs."""
        excel_path = Path(ExcelSequenceRenderer.FILE_PATH)
        if not excel_path.exists():
            return
            
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            
            if "Email Ingestion Audit Log" in wb.sheetnames:
                ws_audit = wb["Email Ingestion Audit Log"]
                for row in ws_audit.iter_rows(min_row=2, values_only=True):
                    if row and row[1]:
                        file_id = str(row[1]).strip()
                        if file_id:
                            self.processed_email_ids.add(file_id)

            if "Quotation Intake Dashboard" in wb.sheetnames:
                ws = wb["Quotation Intake Dashboard"]
                max_seq = 0
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or not row[0]:
                        continue
                    req_id = str(row[0]).strip()
                    
                    id_match = re.search(r"REQ-(\d+)", req_id)
                    if id_match:
                        seq = int(id_match.group(1)) - 1000
                        if seq > max_seq:
                            max_seq = seq

                    cust = str(row[2]).strip().lower() if len(row) > 2 and row[2] else ""
                    pol_pod = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    cntr_type = str(row[5]).strip().upper() if len(row) > 5 and row[5] else ""
                    commodity = str(row[7]).strip().lower() if len(row) > 7 and row[7] else ""
                    cntr_qty = str(row[8]).strip() if len(row) > 8 and row[8] else ""
                    
                    if "➔" in pol_pod:
                        pol, pod = [p.strip() for p in pol_pod.split("➔", 1)]
                        pol_norm = normalize_port_code(pol)
                        pod_norm = normalize_port_code(pod)
                        
                        composite_key = f"{cust}_{pol_norm}_{pod_norm}_{cntr_type}_{commodity}_{cntr_qty}"
                        self.active_quotations_db[composite_key] = {"request_id": req_id}
                
                if max_seq > 0:
                    self.request_seq_counter = max_seq + 1
                    
            print(f"[Excel History Loaded] Loaded {len(self.processed_email_ids)} processed email IDs. Resuming counter from REQ-{1000 + self.request_seq_counter}")
        except Exception as e:
            print(f"[History Load Warning]: Could not load prior Excel keys: {e}")

    def is_exact_duplicate(self, thread_id: str, customer_name: str, leg: SingleLegItem) -> Tuple[bool, Optional[str]]:
        """
        Determines if a leg is an EXACT duplicate of an existing request in the thread/database.
        Matches ALL 5 PARAMETERS: POL + POD + Container Type + Commodity + Container Count.
        """
        pol_norm = normalize_port_code(leg.pol_code)
        pod_norm = normalize_port_code(leg.pod_code)
        cntr_type = str(leg.container_type or "").strip().upper()
        commodity = str(leg.commodity or "").strip().lower()
        cntr_qty = str(leg.cntr_count).strip()
        cust_clean = customer_name.lower().strip()

        thread_key = f"{thread_id}_{pol_norm}_{pod_norm}_{cntr_type}_{commodity}_{cntr_qty}"
        if thread_key in self.active_quotations_db:
            return True, self.active_quotations_db[thread_key]["request_id"]

        cust_key = f"{cust_clean}_{pol_norm}_{pod_norm}_{cntr_type}_{commodity}_{cntr_qty}"
        if cust_key in self.active_quotations_db:
            return True, self.active_quotations_db[cust_key]["request_id"]

        # Check for route-only match (for amendments/revisions)
        route_key = f"{thread_id}_{pol_norm}_{pod_norm}"
        if route_key in self.active_quotations_db:
            return False, self.active_quotations_db[route_key]["request_id"]

        return False, None

    def process_and_save(self, extracted_payload: dict) -> bool:
        thread_id = extracted_payload.get("parent_conversation_id", "")
        file_name = extracted_payload.get("file_name", "")
        
        file_identifier = str(file_name or thread_id).strip()
        
        if file_identifier in self.processed_email_ids:
            print(f"  └─► [Duplicate Skipped] File/Thread '{file_identifier}' was already processed. Skipping.")
            return False

        system_prompt = (
            "You are PIL's Multilingual AI Sourcing Extraction Agent. Analyze incoming email communications "
            "and enforce all scenario processing rules strictly:\n\n"
            "STRICT PORT STANDARDIZATION (UN/LOCODE DIRECTIVE):\n"
            "- ALWAYS extract POL and POD as official 5-character UN/LOCODEs using your global shipping knowledge memory.\n"
            "- EXAMPLES: Shanghai -> CNSHA, Callao -> PECALL, Hamburg -> DEHAM, Jebel Ali -> AEJEA, Santos -> BRSSZ, Durban -> ZADUR, Busan -> KRPUS.\n"
            "- If a city/port name is in Spanish or other languages, translate and output its official 5-character UN/LOCODE.\n\n"
            "PRIMARY MANDATE - 100% INQUIRY RECALL:\n"
            "- Scan the ENTIRE email thread chronologically from oldest (bottom) to newest (top).\n"
            "- Extract EVERY message or form that mentions ports, container types, cargo details, or pricing requests into the 'requests' list.\n"
            "- IF IN DOUBT, ALWAYS EXTRACT. It is far better to extract a marginal inquiry than to miss a valid customer rate request.\n"
            "- Extract historical forms at the bottom AND any new inquiries anywhere in the thread as separate request items.\n"
            "- ALWAYS translate all extracted parameters into standard English.\n\n"
            "0. RATE ROUTING & SURCHARGE CONSOLIDATION RULES (CRITICAL):\n"
            "   - ALWAYS set system_proposed_rate = '—'.\n"
            "   - ALL rates and surcharges MUST be consolidated in customer_requested_rate on a SINGLE line per leg.\n"
            "   - VALIDITY DATE RANGE: ALWAYS format validity with full year strictly as 'MMM DD, YYYY – MMM DD, YYYY' (e.g. 'OCT 01, 2026 – OCT 31, 2026'). Output '—' if missing.\n"
            "   - FREE TIME FORMATTING: Output free_time formatted strictly on two lines using 'Origin:' and 'Destination:'. "
            "If only Origin or only Destination is specified, leave the other label blank without any trailing hyphen, dash, or text (e.g. 'Origin: 7d merged\\nDestination:'). Output '—' only if neither is present.\n"
            "   - VALUE ADDED SERVICES (VAS): Extract all specifically requested VAS names as a list of strings (Customs Clearance, Inland Trucking, Cargo Insurance, Warehousing, Fumigation).\n"
            "     * Do NOT return a count of VAS. Return the actual names as a list.\n"
            "   - REMARKS / CUSTOMER REMARKS: Extract specific, high-value operational instructions directly from the text. Examples: "
            "'Repeat customer, priority handling', 'Class 8 corrosive, MSDS attached', 'Continuous temperature log required', "
            "'Heavy weight — VGM mandatory', 'Shipper-owned containers, return depot info', 'Last 3 cargoes log attached', "
            "'Equipment repo — best rate required', 'Origin: Bogota; bagged', 'Class 9, environmentally hazardous'. "
            "Do NOT dump generic email chatter. If no operational constraints exist, set customer_remarks = '—'.\n\n"
            "1. SCENARIO 1 - CHRONOLOGICAL (BOTTOM-FIRST) EXTRACTION & CHATTER FILTERING (CRITICAL):\n"
            "   - BOTTOM-TO-TOP PROCESSING ORDER:\n"
            "     * READ THE EMAIL THREAD FROM THE BOTTOM (OLDEST MESSAGE) TO THE TOP (NEWEST MESSAGE).\n"
            "     * Extract the oldest rate request at the bottom FIRST into the 'requests' list.\n"
            "     * Then scan upwards and extract any subsequent newer rate requests above it SECOND.\n"
            "   - FEASIBILITY / CHAT FILTERING:\n"
            "     * FEASIBILITY / SERVICE INQUIRIES: Classify as topic = 'Operational Matters (Doc/Schedule/Booking)' and set is_new_rate_request = FALSE.\n"
            "     * CHATTER MESSAGES: Set is_new_rate_request = FALSE.\n"
            "     * TRUE RATE REQUESTS: MUST explicitly ask for pricing/rates OR provide actionable booking details.\n\n"
            "2. SCENARIO 2 - STRICT MULTI-LEG CREATION RULES:\n"
            "   - Create separate legs (seq_num: 1, 2...) WITHIN a request ONLY IF there are different container sizes, commodities, or weights for the SAME route (SAME POL AND SAME POD).\n"
            "   - DIFFERENT ROUTES: Inquiries for distinct destinations/routes MUST be separate MultiLegQuotationRequest entries in 'requests'.\n"
            "   - INHERIT ORIGIN: If POL is omitted in upper messages, infer origin from historical thread context below.\n\n"
            "3. SCENARIO 3 - CONFIDENCE SCORING & AUDIT EVALUATION:\n"
            "   - Assign ai_confidence_score (0.00 to 1.00).\n"
        )

        user_content = (
            f"MAILBOX DESK: {extracted_payload.get('mailbox_account', 'Sample_Desk')}\n"
            f"FILE NAME: {extracted_payload.get('file_name', 'N/A')}\n"
            f"PARENT THREAD ID: {thread_id}\n"
            f"RECEIVED DATE (EMAIL TIMESTAMP): {extracted_payload.get('received_date', 'N/A')}\n"
            f"SENDER: {extracted_payload.get('sender', 'N/A')}\n"
            f"SUBJECT: {extracted_payload.get('subject', 'N/A')}\n\n"
            f"EXTRACTED EMAIL CONTENT (BODY + ATTACHMENTS):\n{extracted_payload['combined_text_payload']}"
        )

        response = self.client.beta.chat.completions.parse(
            model=self.model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            response_format=ThreadQuotationExtraction,
            temperature=0.0
        )

        extracted_container: ThreadQuotationExtraction = response.choices[0].message.parsed
        
        created_req_ids = []
        any_saved = False

        # First pass: Generate IDs and save requests to Dashboard
        for parsed_req in extracted_container.requests:
            # Consolidate multiple surcharge legs into 1 single leg row
            parsed_req.legs = consolidate_surcharge_legs(parsed_req.legs)
            parsed_req.ai_confidence_score = calculate_objective_confidence_score(parsed_req)

            is_rate_request = (parsed_req.topic in self.VALID_RATE_TOPICS) and parsed_req.is_new_rate_request

            if is_rate_request:
                # Group legs by route (POL, POD)
                route_groups: Dict[Tuple[str, str], List[SingleLegItem]] = {}
                for leg in parsed_req.legs:
                    pol = normalize_port_code(leg.pol_code)
                    pod = normalize_port_code(leg.pod_code)
                    key = (pol, pod)
                    if key not in route_groups:
                        route_groups[key] = []
                    route_groups[key].append(leg)

                for (pol, pod), leg_list in route_groups.items():
                    for idx, leg in enumerate(leg_list, start=1):
                        leg.seq_num = idx

                    sub_req = parsed_req.model_copy()
                    sub_req.legs = leg_list

                    first_leg = sub_req.legs[0]
                    is_dup, matched_req_id = self.is_exact_duplicate(thread_id, sub_req.customer_name, first_leg)

                    # 1. IDENTICAL DUPLICATE PREVENTED (ALL PARAMETERS MATCH)
                    if is_dup and sub_req.topic == CommunicationTopic.NEW_RATE_REQUEST:
                        print(f"  └─► [Duplicate Quote Prevented] Re-extracted exact identical quote ({pol} ➔ {pod}, {first_leg.container_type}, {first_leg.commodity}). Excluded from Dashboard.")
                        continue

                    # 2. HANDLE AMENDMENT/REVISION (Same route + Revision topic)
                    elif matched_req_id and (sub_req.topic in [CommunicationTopic.RATE_REVISION, CommunicationTopic.CUSTOMER_NEGOTIATION]):
                        base_id = matched_req_id.split("-v")[0]
                        sub_req.request_id = f"{base_id}-v2"
                        sub_req.request_type = RequestType.AMENDMENT
                        created_req_ids.append(sub_req.request_id)
                        ExcelSequenceRenderer.save_request(sub_req)
                        any_saved = True

                    # 3. GENUINE NEW RATE REQUEST (New route OR same route with new commodity/container)
                    else:
                        sub_req.request_id = f"REQ-{1000 + self.request_seq_counter}"
                        sub_req.request_type = RequestType.NEW
                        self.request_seq_counter += 1

                        cntr_type = str(first_leg.container_type or "").strip().upper()
                        commodity = str(first_leg.commodity or "").strip().lower()
                        cntr_qty = str(first_leg.cntr_count).strip()
                        cust_clean = sub_req.customer_name.lower().strip()

                        thread_full_key = f"{thread_id}_{pol}_{pod}_{cntr_type}_{commodity}_{cntr_qty}"
                        cust_full_key = f"{cust_clean}_{pol}_{pod}_{cntr_type}_{commodity}_{cntr_qty}"
                        route_only_key = f"{thread_id}_{pol}_{pod}"
                        
                        entry_dict = {"request_id": sub_req.request_id}

                        self.active_quotations_db[thread_full_key] = entry_dict
                        self.active_quotations_db[cust_full_key] = entry_dict
                        self.active_quotations_db[route_only_key] = entry_dict

                        created_req_ids.append(sub_req.request_id)
                        ExcelSequenceRenderer.save_request(sub_req)
                        any_saved = True

        # Second pass: Log to Audit Sheet with all generated REQ IDs mapped
        for parsed_req in extracted_container.requests:
            is_rate_request = (parsed_req.topic in self.VALID_RATE_TOPICS) and parsed_req.is_new_rate_request
            
            ExcelSequenceRenderer.log_audit_record(
                extracted_payload, 
                parsed_req, 
                is_rate_request, 
                created_req_ids=created_req_ids
            )
            
            if not is_rate_request:
                print(f"  └─► [Chatter Logged to Audit] Email '{extracted_payload['subject']}' ({parsed_req.topic.value}). Excluded from main dashboard.")

        self.processed_email_ids.add(file_identifier)
        return any_saved


# =====================================================================
# 4. MAIN ORCHESTRATOR
# =====================================================================

if __name__ == "__main__":
    start_time = time.time()  # Start the timer

    config = load_config("config.json")

    agent = SequenceSourcingAgent(
        api_key=config["OPENAI_API_KEY"],
        base_url=config["TIGER_AI_GATEWAY_URL"],
        model_name=config.get("MODEL_NAME", "gpt-4o-mini")
    )

    accounts_list = config.get("ACCOUNTS", [])

    if not accounts_list:
        print("[Error] No accounts configured in config.json.")
    else:
        print(f"\n=====================================================================")
        print(f"   STARTING BATCH PROCESSING (ZERO-DROP & AUDIT LOG ENFORCED)       ")
        print(f"=====================================================================")

        for acc in accounts_list:
            extracted_emails = fetch_extracted_email_payloads(acc, config)

            for email_payload in extracted_emails:
                file_label = email_payload.get('file_name', f"Email #{email_payload['email_sequence_index']}")
                print(f"\n[Main Processing] Ingesting Payload [{file_label}] | Subject: '{email_payload['subject']}'")

                if email_payload.get("attached_files"):
                    print(f"  └─► Attachments Parsed: {email_payload['attached_files']}")

                agent.process_and_save(email_payload)

        # Calculate elapsed execution time
        elapsed_seconds = time.time() - start_time
        minutes, seconds = divmod(elapsed_seconds, 60)

        print("\n=====================================================================")
        print("   PROCESSING COMPLETE. BOTH DASHBOARD & AUDIT LOG UPDATED:           ")
        print("   multileg_sequence_quotation_dashboard.xlsx                         ")
        print(f"   TOTAL RUNTIME: {int(minutes)}m {seconds:.2f}s ({elapsed_seconds:.2f} seconds)")
        print("=====================================================================\n")