import os
import pandas as pd
import streamlit as st
from pathlib import Path

# Import pipeline components
from email_extractor import LocalNERFilter, fetch_extracted_email_payloads
from main_processing import load_config, SequenceSourcingAgent, ExcelSequenceRenderer

# Page Configuration
st.set_page_config(
    page_title="Shipping Requests - PIL Rate Intake Engine",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS matching exact UI screenshot styling
st.markdown("""
<style>
    /* Global Page Styling */
    .block-container { padding-top: 0rem; padding-bottom: 2rem; padding-left: 2rem; padding-right: 2rem; }
    
    /* Top Header Bar */
    .top-header-bar {
        background-color: #0277BD;
        color: white;
        padding: 14px 24px;
        margin-left: -2rem;
        margin-right: -2rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 3px solid #D32F2F;
    }
    .top-header-title { font-size: 20px; font-weight: 700; display: flex; align-items: center; gap: 10px; }
    .top-header-subtitle { font-size: 13px; opacity: 0.85; margin-left: 10px; font-weight: normal; }

    /* Filter Sub-Header Bar */
    .filter-subheader-bar {
        background-color: #5196CE;
        padding: 10px 24px;
        margin-left: -2rem;
        margin-right: -2rem;
        margin-bottom: 20px;
    }

    /* Table Formatting */
    .styled-table-container {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 13px;
        color: #1E293B;
        width: 100%;
        overflow-x: auto;
    }
    
    table.custom-dashboard {
        width: 100%;
        border-collapse: collapse;
        background-color: #FFFFFF;
    }
    
    table.custom-dashboard th {
        background-color: #F1F5F9;
        color: #475569;
        font-weight: 700;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 12px 8px;
        border-bottom: 2px solid #CBD5E1;
        text-align: left;
    }

    table.custom-dashboard td {
        padding: 10px 8px;
        border-bottom: 1px dashed #E2E8F0;
        vertical-align: middle;
    }

    /* System Rate Column Blue Highlight */
    .system-rate-col {
        background-color: #E0F2FE !important;
        font-weight: 600;
        color: #0369A1;
        text-align: right;
    }

    /* Badge Tags & Pills */
    .id-red { color: #DC2626; font-weight: 700; }
    .leg-tag { font-size: 11px; color: #64748B; font-weight: 400; display: block; margin-top: 2px; }

    .badge-dg { background-color: #FEE2E2; color: #991B1B; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; border: 1px solid #FCA5A5; }
    .badge-nor { background-color: #FEF3C7; color: #92400E; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; border: 1px solid #FDE68A; }
    .badge-soc { background-color: #E0F2FE; color: #075985; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; border: 1px solid #BAE6FD; }
    .badge-tank { background-color: #F3E8FF; color: #6B21A8; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; border: 1px solid #E9D5FF; }
    .badge-mty { background-color: #F1F5F9; color: #475569; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; border: 1px solid #CBD5E1; }

    .badge-vas { background-color: #F1F5F9; color: #334155; padding: 3px 8px; border-radius: 12px; font-size: 11px; border: 1px solid #CBD5E1; }
    
    .status-new { background-color: #F1F5F9; color: #475569; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
    .status-pricer { background-color: #E0F2FE; color: #0369A1; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
    .status-customer { background-color: #E0F2FE; color: #0284C7; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
    .status-both { background-color: #F3E8FF; color: #7E22CE; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


def run_pipeline():
    """Executes the ingestion and processing pipeline."""
    config = load_config("config.json")
    
    with st.spinner("Loading Multilingual NER Filter..."):
        local_ner_filter = LocalNERFilter(model_name="xx_ent_wiki_sm")
        
    agent = SequenceSourcingAgent(
        api_key=config["OPENAI_API_KEY"],
        base_url=config["TIGER_AI_GATEWAY_URL"],
        model_name=config.get("MODEL_NAME", "gpt-4o-mini")
    )

    accounts_list = config.get("ACCOUNTS", [])
    request_seq_counter = 1
    processed_count = 0
    skipped_count = 0

    if not accounts_list:
        st.error("No accounts configured in config.json.")
        return

    progress_bar = st.progress(0)
    
    for acc in accounts_list:
        extracted_emails = fetch_extracted_email_payloads(acc, config, local_ner_filter)
        total_emails = len(extracted_emails)

        if total_emails == 0:
            st.warning("No new .eml or .msg files found in ./outlook_inbox.")
            return

        for idx, email_payload in enumerate(extracted_emails, start=1):
            saved = agent.process_and_save(email_payload, request_seq_counter)
            if saved:
                request_seq_counter += 1
                processed_count += 1
            else:
                skipped_count += 1

            progress_bar.progress(idx / total_emails)

    st.success(f"Complete! Added: {processed_count} | Skipped Chatter: {skipped_count}")


# Download & Control Functions
excel_file = ExcelSequenceRenderer.FILE_PATH
file_exists = Path(excel_file).exists()

if file_exists:
    df = pd.read_excel(excel_file, sheet_name="Quotation Intake Dashboard", dtype=str)
    df = df.fillna("—")
    total_reqs = df[df['ID'] != '—']['ID'].nunique() if 'ID' in df.columns else len(df)
    total_legs = len(df)
else:
    df = pd.DataFrame()
    total_reqs = 0
    total_legs = 0

# Top Navigation Header Bar
st.markdown(f"""
<div class="top-header-bar">
    <div class="top-header-title">
        🚢 Shipping Requests 
        <span class="top-header-subtitle">{total_reqs} of {total_reqs} requests - {total_legs} of {total_legs} legs</span>
    </div>
    <div>
        <a href="#" style="color: white; text-decoration: none; margin-right: 15px; font-size: 13px;">🏠 Home</a>
        <a href="#" style="color: white; text-decoration: none; font-size: 13px;">💬 Negotiation</a>
    </div>
</div>
""", unsafe_allow_html=True)

# Filter Sub-Header Toolbar Bar
st.markdown('<div class="filter-subheader-bar">', unsafe_allow_html=True)
fc1, fc2, fc3, fc4 = st.columns([1, 1, 1, 1])

with fc1:
    filter_container = st.selectbox("Containers", ["All containers"] + (list(df['CONTAINER'].unique()) if not df.empty else []), label_visibility="collapsed")
with fc2:
    filter_flag = st.selectbox("Flags", ["All flags", "DG", "NOR", "SOC", "Tank", "MTY"], label_visibility="collapsed")
with fc3:
    filter_status = st.selectbox("Statuses", ["All statuses", "New", "Sent ➔ Pricer", "Sent ➔ Customer", "Sent ➔ Both"], label_visibility="collapsed")
with fc4:
    if st.button("🔄 Sync Outlook Inbox", type="primary", width="stretch"):
        run_pipeline()
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Render HTML Table matching exact Screenshot Layout
if df.empty:
    st.info("No dashboard data found yet. Place `.eml` or `.msg` files in `./outlook_inbox` and click 'Sync Outlook Inbox'.")
else:
    # Filter DataFrame
    filtered_df = df.copy()
    if filter_container != "All containers":
        filtered_df = filtered_df[filtered_df['CONTAINER'] == filter_container]
    if filter_flag != "All flags":
        filtered_df = filtered_df[filtered_df['FLAGS'].astype(str).str.contains(filter_flag, case=False, na=False)]
    if filter_status != "All statuses":
        filtered_df = filtered_df[filtered_df['STATUS'] == filter_status]

    # Render HTML Rows
    html_rows = []
    
    # Calculate legs per ID for badge display
    leg_counts = df[df['ID'] != '—']['ID'].value_counts().to_dict()

    for _, r in filtered_df.iterrows():
        req_id = str(r.get('ID', '—')).strip()
        id_display = ""
        
        if req_id != "—" and req_id != "":
            count = leg_counts.get(req_id, 1)
            leg_label = f"📦 {count} legs" if count > 1 else ""
            id_display = f"<span class='id-red'>{req_id}</span><br><span class='leg-tag'>{leg_label}</span>"

        # Format Flags
        flags_val = str(r.get('FLAGS', '—')).strip()
        flags_html = "—"
        if flags_val != "—" and flags_val != "":
            badges = []
            for tag in flags_val.split():
                tag_upper = tag.upper()
                if "DG" in tag_upper: badges.append("<span class='badge-dg'>DG</span>")
                elif "NOR" in tag_upper: badges.append("<span class='badge-nor'>NOR</span>")
                elif "SOC" in tag_upper: badges.append("<span class='badge-soc'>SOC</span>")
                elif "TANK" in tag_upper: badges.append("<span class='badge-tank'>Tank</span>")
                elif "MTY" in tag_upper: badges.append("<span class='badge-mty'>MTY</span>")
                else: badges.append(f"<span class='badge-mty'>{tag}</span>")
            flags_html = " ".join(badges)

        # Format VAS
        vas_val = str(r.get('VAS', '—')).strip()
        vas_html = f"<span class='badge-vas'>🏷️ {vas_val.replace('🏷️', '').strip()}</span>" if vas_val != "—" and vas_val != "" else "—"

        # Format Status
        status_val = str(r.get('STATUS', '—')).strip()
        status_html = "—"
        if status_val == "New": status_html = "<span class='status-new'>New</span>"
        elif "Pricer" in status_val: status_html = "<span class='status-pricer'>Sent ➔ Pricer</span>"
        elif "Customer" in status_val: status_html = "<span class='status-customer'>Sent ➔ Customer</span>"
        elif "Both" in status_val: status_html = "<span class='status-both'>Sent ➔ Both</span>"

        # Single-line concatenated HTML string to prevent Streamlit indent parsing bugs
        row_str = f"<tr><td>{id_display}</td><td style='font-weight: 600;'>{r.get('CUSTOMER', '—')}</td><td style='text-align: center; font-size: 12px;'>{r.get('POL ➔ POD', '—')}</td><td style='font-size: 12px;'>{r.get('VALIDITY', '—')}</td><td style='text-align: center; font-size: 12px;'>{r.get('CONTAINER', '—')}</td><td style='text-align: right;'>{r.get('WT (T)', '—')}</td><td>{r.get('COMMODITY', '—')}</td><td style='text-align: right;'>{r.get('# CNTR', '—')}</td><td style='text-align: center; font-size: 12px; font-weight: 500;'>{r.get('SERVICE / VOYAGE', '—')}</td><td class='system-rate-col'>{r.get('SYSTEM PROPOSED RATE', '—')}</td><td style='text-align: right;'>{r.get('CUSTOMER REQUESTED RATE', '—')}</td><td style='text-align: center;'>{flags_html}</td><td style='text-align: center;'>{vas_html}</td><td style='font-size: 11px;'>{r.get('FREE TIME', '—')}</td><td style='font-size: 11px;'>{r.get('REMARKS', '—')}</td><td style='text-align: center;'>{status_html}</td></tr>"
        html_rows.append(row_str)

    table_body = "".join(html_rows)
    
    table_html = f"""<div class="styled-table-container"><table class="custom-dashboard"><thead><tr><th>ID</th><th>CUSTOMER</th><th style="text-align: center;">POL ➔ POD</th><th>VALIDITY</th><th style="text-align: center;">CONTAINER</th><th style="text-align: right;">WT (T)</th><th>COMMODITY</th><th style="text-align: right;"># CNTR</th><th style="text-align: center;">SERVICE / VOYAGE</th><th style="text-align: right; background-color: #BAE6FD; color: #0369A1;">SYSTEM PROPOSED RATE</th><th style="text-align: right;">CUSTOMER REQUESTED RATE</th><th style="text-align: center;">FLAGS</th><th style="text-align: center;">VAS</th><th>FREE TIME</th><th>REMARKS</th><th style="text-align: center;">STATUS</th></tr></thead><tbody>{table_body}</tbody></table></div>"""

    st.markdown(table_html, unsafe_allow_html=True)