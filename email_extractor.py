
import os
import io
import re
import email
from pathlib import Path
from typing import List, Dict, Tuple
from email.utils import parsedate_to_datetime

import spacy
from pypdf import PdfReader
import extract_msg
from bs4 import BeautifulSoup


class LocalNERFilter:
    """Multilingual Local spaCy NER filter supporting Chinese, German, Spanish, French, Japanese, and English."""

    MULTILINGUAL_KEYWORDS = [
        # English Freight & Container Jargon
        "quote", "quotation", "rate", "freight", "price", "pricing", "rfp", "rfq", "cost", "tariff",
        "shipping", "cargo", "transport", "logistics", "shipment", "carrier", "vessel", "container",
        "pol", "pod", "por", "fnd", "origin", "destination", "reefer", "20gp", "40gp", "40hc", "20'gp", "40'hc",
        "20ft", "40ft", "iso tank", "flatrack", "open top", "teu", "cbm", "gw", "gross weight",
        
        # High-Priority Subject/Header Markers
        "exp rate", "imp rate", "export rate", "import rate", "sept rate", "aug rate", "oct rate",
        "rate inquiry", "rate request", "spot rate", "freight quote",
        
        # Spanish Freight Terms & Logistics Notation
        "cotización", "cotizacion", "cotizar", "tarifa", "flete", "contenedor", "puerto", 
        "origen", "destino", "embarque", "transporte", "enlatados", "mercaderia", "incoterm",
        "1x40", "1x20", "2x40", "2x20", "40hc", "20gp", "40'hc", "20'gp", "1x40hc", "1x20gp",
        
        # Port & Location Jargon Examples
        "yangon", "karachi", "pkkhi", "mmrgn", "shanghai", "ningbo", "rotterdam", "hamburg",
        "manzanillo", "veracruz", "guayaquil", "buenaventura", "callao", "valparaiso",
        
        # Chinese (Simplified & Traditional)
        "报价", "运费", "海运", "箱型", "港口", "目的港", "始发港", "柜型", "集装箱", "运价", "询价", "货代", "船期",
        
        # German
        "angebot", "fracht", "frachtrate", "containertyp", "hafen", "ladehafen", "löschhafen", "versand", "seefracht",
        
        # French
        "devis", "tarif", "fret", "conteneur", "port", "chargement", "déchargement", "expédition",
        
        # Japanese
        "見積", "運賃", "海上運賃", "コンテナ", "船積み", "積港", "揚港"
    ]

    def __init__(self, model_name: str = "xx_ent_wiki_sm"):
        """
        Uses 'xx_ent_wiki_sm' (spaCy's universal multilingual Wikipedia model).
        Fallback to 'en_core_web_sm' if multilingual model is not installed.
        """
        try:
            print(f"[Local NER Filter] Loading Multilingual spaCy model '{model_name}'...")
            self.nlp = spacy.load(model_name)
        except Exception:
            print(f"[Local NER Warning] '{model_name}' not found. Falling back to 'en_core_web_sm'...")
            self.nlp = spacy.load("en_core_web_sm")

    def extract_entities(self, text: str) -> Dict[str, List[str]]:
        doc = self.nlp(text)
        entities = {"locations": [], "organizations": []}

        for ent in doc.ents:
            if ent.label_ in ["LOC", "GPE"]:
                entities["locations"].append(ent.text)
            elif ent.label_ in ["ORG", "PER"]:
                entities["organizations"].append(ent.text)

        for key in entities:
            entities[key] = list(set(entities[key]))
        return entities

    def is_quotation_relevant(self, subject: str, text: str) -> Tuple[bool, Dict[str, List[str]]]:
        """
        Evaluates relevance across subject and combined body content using
        multilingual keywords, regex equipment patterns, and spaCy entity recognition.
        """
        combined_content = f"{subject}\n{text}".lower()
        entities = self.extract_entities(text)
        
        # 1. Broad Multilingual Keyword Check
        has_keywords = any(kw in combined_content for kw in self.MULTILINGUAL_KEYWORDS)
        
        # 2. Regex Pattern Check for Container Specifications (e.g., 1x40, 1x20, 20ft, 40hc, 40'hc)
        has_container_pattern = bool(re.search(r"\b\d+\s*x\s*\d+\b|\b\d+['']?\s*(hc|gp|rf|ft)\b", combined_content))
        
        # 3. Location Entity Check
        has_locations = len(entities["locations"]) > 0

        # Permissive Pass: Either keywords, container notations, or locations trigger LLM processing
        is_relevant = has_keywords or has_container_pattern or has_locations

        return is_relevant, entities


class FileAttachmentExtractor:
    """Helper methods to extract text from PDF streams."""

    @staticmethod
    def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            extracted_pages = []
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    extracted_pages.append(text)
            return "\n".join(extracted_pages)
        except Exception as e:
            print(f"  └─► [PDF Extractor Warning]: Failed to read PDF: {e}")
            return ""


class OutlookFileParser:
    """Parses text body, dates, and PDF attachments from both .eml and .msg files."""

    @staticmethod
    def clean_html_body(raw_text_or_html: str) -> str:
        """Strips HTML tags, cleans &nbsp; entities, removes null bytes, and normalizes line breaks."""
        if not raw_text_or_html:
            return ""
        
        raw_text_or_html = raw_text_or_html.replace("\x00", "")

        if "<div" in raw_text_or_html or "<html" in raw_text_or_html or "<br" in raw_text_or_html or "<p" in raw_text_or_html:
            soup = BeautifulSoup(raw_text_or_html, "html.parser")
            text = soup.get_text(separator="\n")
        else:
            text = raw_text_or_html

        text = text.replace("\xa0", " ").replace("&nbsp;", " ")
        text = re.sub(r"\n\s*\n", "\n", text)
        return text.strip()

    @staticmethod
    def split_and_reverse_thread(full_text: str) -> str:
        """
        Splits an email thread by standard Outlook reply headers (EN, ES, DE, FR)
        and reverses the blocks so the original request appears first (bottom-to-top order).
        """
        if not full_text:
            return ""

        # Clean regex pattern matching standard multilingual reply headers
        reply_pattern = r"(?i)(\n-{3,}\s*Original Message\s*-{3,}|\nFrom:\s+|\nDe:\s+|\nEnviado el:\s+|\nVon:\s+)"
        thread_parts = re.split(reply_pattern, full_text)
        
        if len(thread_parts) <= 1:
            return full_text

        reconstructed_messages = []
        reconstructed_messages.append(thread_parts[0]) 
        
        for i in range(1, len(thread_parts), 2):
            header = thread_parts[i]
            body = thread_parts[i+1] if i+1 < len(thread_parts) else ""
            reconstructed_messages.append(header + body)

        # Reverse the order: Oldest message (bottom) -> Newest reply (top)
        chronological_thread = list(reversed(reconstructed_messages))
        
        structured_payload = []
        for index, msg in enumerate(chronological_thread, start=1):
            if msg.strip():
                structured_payload.append(f"--- [THREAD STEP #{index} (CHRONOLOGICAL ORDER)] ---\n{msg.strip()}")
            
        return "\n\n".join(structured_payload)

    @classmethod
    def parse_eml_file(cls, file_path: Path) -> Tuple[str, str, str, str, List[str]]:
        """Parses standard MIME .eml files."""
        try:
            with open(file_path, "rb") as f:
                msg = email.message_from_binary_file(f)

            subject = msg.get("Subject", "No Subject")
            sender = msg.get("From", "Unknown Sender")
            raw_date = msg.get("Date", "")
            
            try:
                msg_date = str(parsedate_to_datetime(raw_date)) if raw_date else "Unknown Date"
            except Exception:
                msg_date = raw_date or "Unknown Date"

            body_text = ""
            attachment_texts = []
            filenames = []

            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                filename = part.get_filename()

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body_text += payload.decode("utf-8", errors="ignore")

                elif content_type == "text/html" and not body_text and "attachment" not in content_disposition:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body_text += payload.decode("utf-8", errors="ignore")

                elif "attachment" in content_disposition and filename:
                    if filename.lower().endswith(".pdf"):
                        pdf_bytes = part.get_payload(decode=True)
                        if pdf_bytes:
                            pdf_text = FileAttachmentExtractor.extract_text_from_pdf_bytes(pdf_bytes)
                            if pdf_text.strip():
                                attachment_texts.append(f"--- ATTACHMENT CONTENT ({filename}) ---\n{pdf_text}")
                                filenames.append(filename)

            cleaned_body = cls.clean_html_body(body_text)
            chronological_body = cls.split_and_reverse_thread(cleaned_body)

            full_text = chronological_body
            if attachment_texts:
                full_text += "\n\n" + "\n\n".join(attachment_texts)

            return subject, sender, msg_date, full_text, filenames

        except Exception as e:
            print(f"  └─► [EML Parsing Error] Failed to parse '{file_path.name}': {e}")
            return f"Error reading {file_path.name}", "Unknown Sender", "Unknown Date", "", []

    @classmethod
    def parse_msg_file(cls, file_path: Path) -> Tuple[str, str, str, str, List[str]]:
        """Parses Outlook binary .msg files using extract_msg."""
        try:
            msg = extract_msg.Message(str(file_path))
            
            subject = str(msg.subject or "No Subject")
            sender = str(msg.sender or "Unknown Sender")
            msg_date = str(msg.date or "Unknown Date")
            
            raw_body = msg.body or ""
            if not raw_body.strip() and hasattr(msg, 'htmlBody') and msg.htmlBody:
                raw_body = msg.htmlBody.decode('utf-8', errors='ignore') if isinstance(msg.htmlBody, bytes) else str(msg.htmlBody)

            body_text = cls.clean_html_body(raw_body)
            chronological_body = cls.split_and_reverse_thread(body_text)

            attachment_texts = []
            filenames = []

            if hasattr(msg, 'attachments') and msg.attachments:
                for att in msg.attachments:
                    try:
                        filename = getattr(att, 'longFilename', None) or getattr(att, 'shortFilename', None) or "attachment.pdf"
                        if str(filename).lower().endswith(".pdf"):
                            pdf_bytes = att.data
                            if pdf_bytes:
                                pdf_text = FileAttachmentExtractor.extract_text_from_pdf_bytes(pdf_bytes)
                                if pdf_text.strip():
                                    attachment_texts.append(f"--- ATTACHMENT CONTENT ({filename}) ---\n{pdf_text}")
                                    filenames.append(str(filename))
                    except Exception as att_err:
                        print(f"  └─► [MSG Attachment Warning] Skipping attachment in {file_path.name}: {att_err}")

            msg.close()

            full_text = chronological_body
            if attachment_texts:
                full_text += "\n\n" + "\n\n".join(attachment_texts)

            return subject, sender, msg_date, full_text, filenames

        except Exception as e:
            print(f"  └─► [MSG Parsing Error] Failed to parse '{file_path.name}': {e}")
            return f"Error reading {file_path.name}", "Unknown Sender", "Unknown Date", "", []


def fetch_extracted_email_payloads(account_info: dict, config: dict, ner_filter: LocalNERFilter) -> List[Dict]:
    """
    Scans directory for ALL .eml and .msg files, cleans text/attachments,
    runs spaCy pre-filtering, and returns a list of payload objects.
    """
    inbox_dir = Path(config.get("LOCAL_INBOX_DIR", "./outlook_inbox"))

    if not inbox_dir.exists():
        inbox_dir.mkdir(parents=True, exist_ok=True)
        print(f"[Outlook Extractor] Created folder at '{inbox_dir}'. Drop .eml or .msg files here.")
        return []

    sample_files = sorted(list(inbox_dir.glob("*.eml")) + list(inbox_dir.glob("*.msg")))

    if not sample_files:
        print(f"[Outlook Extractor] No .eml or .msg files found in '{inbox_dir}'.")
        return []

    print(f"\n=====================================================================")
    print(f" [Outlook Extractor] Found {len(sample_files)} sample email file(s) in '{inbox_dir}'")
    print(f"=====================================================================")

    extracted_payloads = []

    for idx, file_path in enumerate(sample_files, start=1):
        if file_path.suffix.lower() == ".eml":
            subject, sender, msg_date, full_text, filenames = OutlookFileParser.parse_eml_file(file_path)
        elif file_path.suffix.lower() == ".msg":
            subject, sender, msg_date, full_text, filenames = OutlookFileParser.parse_msg_file(file_path)
        else:
            continue

        is_relevant, ner_entities = ner_filter.is_quotation_relevant(subject, full_text)

        if not is_relevant:
            print(f"  └─► [Local NER Discard] File '{file_path.name}' contains no logistics entities. Skipped.")
            continue

        clean_subject = subject.lower().replace("re:", "").replace("fwd:", "").strip()
        parent_conv_id = f"THREAD-{abs(hash(clean_subject)) % 100000:05d}"

        extracted_payloads.append({
            "mailbox_account": account_info.get("account_name", "Outlook_Local_Desk"),
            "file_name": file_path.name,
            "email_sequence_index": idx,
            "parent_conversation_id": parent_conv_id,
            "sender": sender,
            "received_date": msg_date,
            "subject": subject,
            "combined_text_payload": full_text,
            "attached_files": filenames,
            "local_ner_hints": ner_entities
        })

    return extracted_payloads