
import os
import re
import json
from pathlib import Path
from typing import Optional, List, Dict
from enum import Enum

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field, field_validator
from openai import OpenAI

# Import extractor module and Local NER Filter
from email_extractor import LocalNERFilter, fetch_extracted_email_payloads


# =====================================================================
# 0. CONFIGURATION LOADER
# =====================================================================

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
        description="Port of Loading standard 3-letter code (e.g., SHA, NGB, HAM). IF UNABLE TO MAP, ENTER EXACT RAW LOCATION NAME MENTIONED IN EMAIL (e.g. Wuhan)."
    )
    pod_code: Optional[str] = Field(
        default="UNKNOWN", 
        description="Port of Discharge standard 3-letter code (e.g., RTM, FXT, NYC). IF UNABLE TO MAP, ENTER EXACT RAW LOCATION NAME MENTIONED IN EMAIL (e.g. Berbera Port)."
    )
    container_type: Optional[str] = Field(default=None, description="Container size/type (e.g., 40'HC, 20'GP, ISO Tank, 40'RF). Set None if missing.")
    weight_tons: Optional[float] = Field(default=None, description="Gross weight in metric tons")
    commodity: Optional[str] = Field(default=None, description="Commodity description. Set None if missing.")
    cntr_count: int = Field(default=1, description="Container count requested")
    
    service_voyage: Optional[str] = Field(
        default="—", 
        description=(
            "Liner Service Name and Voyage string formatted strictly as 'SERVICE / V.VOYAGE' "
            "(e.g., 'AEX1 / V.142W', 'MEX2 / V.018E', 'TA4 / V.331W', 'INDIA-EUR / V.061E'). "
            "Set '—' if missing."
        )
    )
    
    system_proposed_rate: Optional[str] = Field(default="—", description="Calculated or system tariff rate (e.g. $2,850 or USD 3400)")
    customer_requested_rate: Optional[str] = Field(default="—", description="Target rate requested by customer (e.g. $2,600) or '—'")
    
    flags: List[str] = Field(
        default_factory=list, 
        description=(
            "List of standard cargo badge tags detected. Use short standard codes ONLY: "
            "'DG' (Dangerous Goods), 'NOR' (Non-Operating Reefer), 'SOC' (Shipper Owned Container), "
            "'Tank' (ISO Tank), 'MTY' (Empty Container), 'Reefer', 'OOG' (Out of Gauge). "
            "Leave list empty [] if standard cargo."
        )
    )
    
    vas: Optional[str] = Field(
        default="—", 
        description=(
            "Value Added Services count or tag. If 1, 2, or 3 special service requests exist "
            "(e.g., customs clearance, inland haulage, cargo insurance, special monitoring), "
            "output '🏷️ 1', '🏷️ 2', '🏷️ 3', etc. Output '—' if no VAS requested."
        )
    )
    
    free_time: Optional[str] = Field(default="Standard", description="Origin & Destination free time string (e.g. Origin: 7d merged | Destination: Dem 7d / Det 5d)")

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
            service = match.group(1)
            voyage = match.group(2)
            return f"{service} / V.{voyage}"
            
        return v


class MultiLegQuotationRequest(BaseModel):
    request_id: str = Field(default="", description="Shared Request ID (e.g., REQ-1001 or REQ-1001-v2)")
    customer_name: str = Field(default="Unknown Customer", description="Customer company name")
    validity_period: Optional[str] = Field(default=None, description="Validity period string (e.g., May 01 – May 31)")
    topic: CommunicationTopic = Field(description="Detected email communication topic")
    status: StatusEnum = Field(default=StatusEnum.NEW, description="Workflow status")
    request_type: RequestType = Field(default=RequestType.NEW, description="Type of request")
    
    is_new_rate_request: bool = Field(
        description=(
            "Set TRUE if the email thread contains ANY rate inquiry OR rate quote offer "
            "(e.g., 'USD 3400/20GP' or 'advise export rate'), even if surrounded by carrier disclaimers or replies. "
            "Set FALSE ONLY for purely operational chatter, invoices, or vessel schedule updates."
        )
    )
    
    ai_confidence_score: float = Field(
        description="Calculated confidence score (0.00 to 1.00) based on extraction completeness, ambiguity, and text clarity."
    )
    assigned_owner: str = Field(default="Asia_Desk_Ops", description="Assigned desk owner")
    customer_remarks: Optional[str] = Field(default="—", description="Special instructions or customer remarks")
    legs: List[SingleLegItem] = Field(description="List of 1 or more trade legs extracted")


# =====================================================================
# 2. EXCEL SEQUENCE RENDERER (DASHBOARD PERSISTENCE - SCREENSHOT COLUMNS)
# =====================================================================

class ExcelSequenceRenderer:
    FILE_PATH = "multileg_sequence_quotation_dashboard.xlsx"
    HEADERS = [
        "ID", "LEG SEQ #", "CUSTOMER", "POL ➔ POD", "VALIDITY", "CONTAINER", 
        "WT (T)", "COMMODITY", "# CNTR", "SERVICE / VOYAGE", 
        "SYSTEM PROPOSED RATE", "CUSTOMER REQUESTED RATE", "FLAGS", 
        "VAS", "FREE TIME", "REMARKS", "STATUS", "CONFIDENCE", "MISSING DATA?"
    ]

    @classmethod
    def save_request(cls, req: MultiLegQuotationRequest):
        file_exists = Path(cls.FILE_PATH).exists()

        if file_exists:
            wb = openpyxl.load_workbook(cls.FILE_PATH)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Quotation Intake Dashboard"
            ws.views.sheetView[0].showGridLines = True

            header_font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            ws.append(cls.HEADERS)

            for col_idx, h in enumerate(cls.HEADERS, 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 28

        # Step 3A Requirement: Always append as a new line to preserve original audit history
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

            # Dynamic Exception Handling (Scenario 3 Validation Rules)
            missing_fields = []
            if req.customer_name == "Unknown Customer": missing_fields.append("Customer")
            if not leg.pol_code or leg.pol_code.upper() in ["UNKNOWN", "NONE"]: missing_fields.append("POL")
            if not leg.pod_code or leg.pod_code.upper() in ["UNKNOWN", "NONE"]: missing_fields.append("POD")
            if not leg.container_type or leg.container_type.upper() in ["UNKNOWN", "NONE"]: missing_fields.append("Container Type")
            if not leg.commodity or leg.commodity.upper() in ["UNKNOWN", "NONE"]: missing_fields.append("Commodity")
            if req.ai_confidence_score < 0.85: missing_fields.append("Low Confidence")

            missing_indicator = f"YES ({', '.join(missing_fields)})" if missing_fields else "NO"
            
            # Format FLAGS and VAS
            flags_str = " ".join(leg.flags) if leg.flags else "—"
            vas_str = leg.vas if leg.vas and str(leg.vas).strip() not in ["", "None", "0"] else "—"
            pol_pod_str = f"{leg.pol_code or 'UNKNOWN'} ➔ {leg.pod_code or 'UNKNOWN'}"

            row_vals = [
                req.request_id if i == 0 else "",
                leg.seq_num,
                req.customer_name if i == 0 else "",
                pol_pod_str,
                req.validity_period or "—",
                leg.container_type or "UNSPECIFIED",
                leg.weight_tons or "—",
                leg.commodity or "UNSPECIFIED",
                leg.cntr_count,
                leg.service_voyage or "—",
                leg.system_proposed_rate or "—",
                leg.customer_requested_rate or "—",
                flags_str,
                vas_str,
                leg.free_time or "Standard",
                req.customer_remarks or "—",
                req.status.value if i == 0 else "",
                f"{int(req.ai_confidence_score * 100)}%" if i == 0 else "",
                missing_indicator if i == 0 else ""
            ]

            b_style = border_dashed if i < num_legs - 1 else border_solid

            for col_idx, val in enumerate(row_vals, 1):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                
                if col_idx == 1:
                    c.font = id_font
                elif col_idx == 2:
                    c.font = seq_font
                else:
                    c.font = data_font

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
            1: 14, 2: 10, 3: 22, 4: 16, 5: 18, 6: 12, 7: 10, 8: 24, 
            9: 10, 10: 16, 11: 20, 12: 22, 13: 12, 14: 8, 15: 24, 16: 28, 17: 16, 18: 12, 19: 25
        }
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        wb.save(cls.FILE_PATH)
        print(f"[Excel Persistence] Logged Request '{req.request_id}' ({num_legs} legs) to {cls.FILE_PATH}")


# =====================================================================
# 3. AI SOURCING AGENT & MATCHING ENGINE
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

    def find_existing_quotation(self, thread_id: str, customer_name: str, legs: List[SingleLegItem]) -> Optional[str]:
        if thread_id in self.active_quotations_db:
            return self.active_quotations_db[thread_id]["request_id"]

        first_leg = legs[0] if legs else None
        if first_leg:
            route_key = f"{customer_name.lower().strip()}_{first_leg.pol_code}_{first_leg.pod_code}"
            for record in self.active_quotations_db.values():
                if record["route_key"] == route_key:
                    return record["request_id"]

        return None

    def process_and_save(self, extracted_payload: dict, request_seq_counter: int) -> bool:
        thread_id = extracted_payload.get("parent_conversation_id", "")

        system_prompt = (
            "You are PIL's Multilingual AI Sourcing Extraction Agent. Analyze incoming email communications "
            "and enforce these rules:\n\n"
            "0. MULTILINGUAL & THREAD HANDLING INSTRUCTIONS:\n"
            "   - Analyze email payload in chronological order (oldest messages at bottom to newest at top).\n"
            "   - ALWAYS translate and output all extracted parameters into standard English.\n"
            "   - IGNORE carrier operational disclaimers, bank accounts, website links, depot contacts, and penalty warnings.\n"
            "   - If a rate is quoted in the thread (e.g., 'Usd 3400/20\'GP'), extract it into system_proposed_rate.\n\n"
            "1. TOPIC CLASSIFICATION (Scenario 1 Noise Exclusion):\n"
            "   - 'New Rate Request': Brand new freight rate inquiry OR carrier rate response/quote.\n"
            "   - 'Rate Request Revision / Update': Rate amendments, container count updates, or port changes on existing quotes.\n"
            "   - 'Customer Negotiation': Counter-offers or price negotiations.\n"
            "   - 'Demurrage & Detention (D&D) Inquiry': Free time/detention questions.\n"
            "   - 'Operational Matters (Doc/Schedule/Booking)': Purely administrative emails (booking confirmation, vessel delay notice).\n"
            "   - 'General Thread Reply' / 'Excluded / Non-Quotation': General chatter.\n\n"
            "2. PORT CODE RULES (CRITICAL):\n"
            "   - Extract standard 3-letter UN/LOCODE (e.g., SHA, NGB, HAM, SIN, RTM, RGN, KHI, ZLO, VER).\n"
            "   - IF YOU CANNOT MAP A 3-LETTER CODE FOR A LOCATION, ENTER THE EXACT RAW LOCATION NAME MENTIONED IN THE EMAIL.\n\n"
            "3. SERVICE / VOYAGE & FLAGS & VAS FORMATTING RULES:\n"
            "   - SERVICE / VOYAGE: Extract service code and voyage number strictly as 'SERVICE / V.VOYAGE' (e.g., 'AEX1 / V.142W', 'MEX2 / V.018E'). Output '—' if missing.\n"
            "   - FLAGS: Extract short badges ONLY: 'DG', 'NOR', 'SOC', 'Tank', 'MTY', 'Reefer', 'OOG'. Leave empty [] if standard cargo.\n"
            "   - VAS: Count Value Added Services (customs clearance, trucking, insurance). If 1 found -> '🏷️ 1', if 2 found -> '🏷️ 2'. Output '—' if none.\n\n"
            "4. CONFIDENCE SCORING (0.00 to 1.00 - Scenario 3):\n"
            "   - Assign 0.90 – 1.00: Highly clear request with parameters present.\n"
            "   - Assign 0.70 – 0.89: Missing 1 or 2 mandatory fields or contains minor ambiguity.\n"
            "   - Assign < 0.70: High ambiguity or severe lack of parameters."
        )

        user_content = (
            f"MAILBOX DESK: {extracted_payload.get('mailbox_account', 'Sample_Desk')}\n"
            f"FILE NAME: {extracted_payload.get('file_name', 'N/A')}\n"
            f"PARENT THREAD ID: {thread_id}\n"
            f"SENDER: {extracted_payload.get('sender', 'N/A')}\n"
            f"SUBJECT: {extracted_payload.get('subject', 'N/A')}\n"
            f"LOCAL SPA CY HINTS: {json.dumps(extracted_payload.get('local_ner_hints', {}))}\n\n"
            f"EXTRACTED EMAIL CONTENT (BODY + ATTACHMENTS):\n{extracted_payload['combined_text_payload']}"
        )

        response = self.client.beta.chat.completions.parse(
            model=self.model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            response_format=MultiLegQuotationRequest,
            temperature=0.0
        )

        parsed_req: MultiLegQuotationRequest = response.choices[0].message.parsed
        matched_req_id = self.find_existing_quotation(thread_id, parsed_req.customer_name, parsed_req.legs)

        if parsed_req.topic in self.VALID_RATE_TOPICS and parsed_req.is_new_rate_request:
            if matched_req_id and (parsed_req.topic in [CommunicationTopic.RATE_REVISION, CommunicationTopic.CUSTOMER_NEGOTIATION]):
                # Step 3A Logic: Extracts the base ID ("REQ-1001" from "REQ-1001-v2") and appends a version suffix
                base_id = matched_req_id.split("-v")[0]
                parsed_req.request_id = f"{base_id}-v2"
                parsed_req.request_type = RequestType.AMENDMENT
            else:
                # Step 3B Logic: Creates a completely new request record
                parsed_req.request_id = f"REQ-{1000 + request_seq_counter}"
                parsed_req.request_type = RequestType.NEW

                first_leg = parsed_req.legs[0] if parsed_req.legs else None
                route_k = f"{parsed_req.customer_name.lower().strip()}_{first_leg.pol_code if first_leg else ''}_{first_leg.pod_code if first_leg else ''}"
                self.active_quotations_db[thread_id] = {
                    "request_id": parsed_req.request_id,
                    "route_key": route_k
                }

            ExcelSequenceRenderer.save_request(parsed_req)
            return True
        else:
            print(f"  └─► [Excluded / Thread Chatter] Email '{extracted_payload['subject']}' ({parsed_req.topic.value}). Skipped Excel dashboard.")
            return False


# =====================================================================
# 4. MAIN ORCHESTRATOR
# =====================================================================

if __name__ == "__main__":
    config = load_config("config.json")

    local_ner_filter = LocalNERFilter(model_name="xx_ent_wiki_sm")
    agent = SequenceSourcingAgent(
        api_key=config["OPENAI_API_KEY"],
        base_url=config["TIGER_AI_GATEWAY_URL"],
        model_name=config.get("MODEL_NAME", "gpt-4o-mini")
    )

    accounts_list = config.get("ACCOUNTS", [])
    request_seq_counter = 1

    if not accounts_list:
        print("[Error] No accounts configured in config.json.")
    else:
        print(f"\n=====================================================================")
        print(f"      STARTING BATCH PROCESSING (SCENARIOS 1, 2 & 3 ENFORCED)         ")
        print(f"=====================================================================")

        for acc in accounts_list:
            extracted_emails = fetch_extracted_email_payloads(acc, config, local_ner_filter)

            for email_payload in extracted_emails:
                file_label = email_payload.get('file_name', f"Email #{email_payload['email_sequence_index']}")
                print(f"\n[Main Processing] Ingesting Payload [{file_label}] | Subject: '{email_payload['subject']}'")

                if email_payload.get("attached_files"):
                    print(f"  └─► Attachments Parsed: {email_payload['attached_files']}")

                saved = agent.process_and_save(email_payload, request_seq_counter)
                if saved:
                    request_seq_counter += 1

        print("\n=====================================================================")
        print("  PROCESSING COMPLETE. QUOTATION DASHBOARD UPDATED:                 ")
        print("  multileg_sequence_quotation_dashboard.xlsx                         ")
        print("=====================================================================\n")