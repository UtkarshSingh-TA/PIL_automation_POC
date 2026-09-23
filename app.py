"""
================================================================================
PIL · Pricing & Quotation — merged application
================================================================================

The prototype *is* the app. Streamlit is the invisible engine underneath it:
it reads the workbook, runs the ingestion pipeline, and hands data to the page.
The user never sees a Streamlit tab, sidebar or widget.

    new_main.py            — untouched (pipeline orchestrator, LLM translation)
    new_email_extractor.py — untouched (mailbox ingestion)
    pil-scenario4-v4.html  — EMBEDDED below as a compressed blob; a copy of it
                             on disk still wins, so the prototype is swappable
    copy_sweep.py          — INLINED below; no longer a separate file
    new_app.py             — the only file this merger changes

WHAT THIS FILE DOES
-------------------
1.  Strips the Streamlit shell — no sidebar, no tabs, no outer top bar, no
    toolbar. One full-height components.html() of the transformed prototype.
2.  Transforms the prototype in memory, exactly as build_mockup_reference.py
    does on disk: live payload + hero-preserving merge, scenario pill out,
    guided-demo rail hidden, RFP wording neutralised, copy_sweep.apply().
3.  Appends the extension block — the four-item Sales nav, the inbox filters
    and buttons, the Ingestion Console and the Email Audit Log — written in the
    prototype's own vocabulary so the new screens are indistinguishable from
    the original ones.
4.  Bridges the prototype's "Run Batch Ingestion" button to Python through a
    hidden Streamlit button in the parent document, runs the original pipeline
    unchanged, and feeds the per-email log back into the Ingestion Console.

Everything applied to the prototype is append-only and anchored on strings the
prototype already contains. Nothing in its own logic is rewritten in place, so
dropping in a v5 prototype needs no code change here.

The four call sites into the protected modules are character-identical to
new_app.py.bak and must stay that way:

    SequenceSourcingAgent(api_key=..., base_url=..., model_name=model_name)
    fetch_extracted_email_payloads(acc, config)
    inspect.signature(agent.process_and_save)
    agent.process_and_save(email_payload, agent.request_seq_counter)

================================================================================
"""

import base64
import html as html_lib
import inspect
import json
import re
import zlib
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# --- protected modules: imported, never edited --------------------------------
from main_processing import (
    SequenceSourcingAgent,
    ExcelSequenceRenderer,
    load_config,
)
from email_extractor import fetch_extracted_email_payloads

# ==============================================================================
# copy_sweep — inlined verbatim so this file has no local imports.
# Kept at module level: the SWEEP entries are multi-line strings whose exact
# whitespace has to match the prototype, so they cannot be indented.
# ==============================================================================

# (find, replace). Applied in order, each must match at least once.
SWEEP = [

    # ---- A. provenance ------------------------------------------------------
    ("""    'Extracted from <span class="mono">' + esc(WORKBOOK) + '</span> · sheet <span class="mono">Quotation Intake Dashboard</span>',""",
     """    '',"""),

    ("""      '<span class="sub">Rows extracted from the PIL email-ingestion dataset; rates and workflow status are demo values.</span>' +\n""",
     ""),

    ("""      '</span> · <span class="mono">' + esc(REC_LOC_S) + '</span> — extracted from <span class="mono">' +
      esc(WORKBOOK) + '</span>',""",
     """      '</span> · <span class="mono">' + esc(REC_LOC_S) + '</span>',"""),

    # ---- B. card footers ----------------------------------------------------
    ("""      '<div class="card-ft">Nineteen columns in the dataset’s own field order. <strong>System proposed rate is empty on all 41 rows ' +
      'of the source workbook</strong> — there are no rates in the dataset at all, which is the gap this prototype fills. ' +
      'Readiness is the twentieth column and is derived here, not read from the file.</div></div>';""",
     """      '</div>';"""),

    ("""    '<div class="card-ft">Field order is the dataset’s own intake schema. Values are verbatim; the notes beneath them are this app’s reading.</div></div>';""",
     """    '</div>';"""),

    ("""      '<div class="card-ft">The five oldest entries are seeded so the trail never opens empty. Session timestamps run from the ' +
      'fixed demo clock of <span class="mono">' + esc(NOW_SGT) + '</span>.</div></div>';""",
     """      '</div>';"""),

    # ---- C. meta prose ------------------------------------------------------
    ("""      '<span class="sub">Every published rate is retained; nothing is overwritten</span>' +""",
     """      '<span class="sub">Every published rate is retained</span>' +"""),

    ("""    '<span class="sub">Recomputed on every render from the current configuration</span>' +\n""",
     ""),

    ("""      '<span class="sub">Evaluated on every quotation anyway, so the waterfall can show why they returned false</span>' +""",
     """      '<span class="sub">Evaluated on every quotation</span>' +"""),

    ("""    '<span class="sub">Read at calculation time</span>' +\n""", ""),

    ("""      '<span class="sub">Every figure read from state</span>' +\n""", ""),

    ("""      '<span class="sub">Newest first · shaded rows were written in this session</span>' +""",
     """      '<span class="sub">Newest first</span>' +"""),

    ("""  return pageHead('', 'My Quotations', S.quotations.length + ' issued this session') +""",
     """  return pageHead('', 'My Quotations', S.quotations.length + ' issued') +"""),

    # ---- C. leftovers from the removed guided demo ---------------------------
    ("""          'This is the base rate the quotation engine quotes from in step 5.')""",
     """          'This is the base rate the quotation engine quotes from.')"""),

    ("""data-act="step" data-n="2">Go to step 2</button>""",
     """data-act="step" data-n="2">Go to Strategy Builder</button>"""),

    ("""data-act="step" data-n="1">Go to step 1</button>""",
     """data-act="step" data-n="1">Go to Price Lists</button>"""),

    # ---- D. a missing popover renders nothing, silently ---------------------
    ("""  if (!p){ console.warn('missing popover copy:', id); return ''; }""",
     """  if (!p) return '';"""),
]

# Popovers that narrate how the build was made, where the rows came from, or
# which guided-demo step you are on. The 12 that explain how the PRODUCT works
# are kept. Restore any of these by deleting its id from this list.
DROP_POPOVERS = [
    "pl-demo", "str-demo", "rul-demo", "inb-demo", "case-demo", "a6-demo", "b6-demo",
    "pl-missing", "wiz-optional", "pld-audit", "str-name", "rul-dollars", "rul-floor",
    "rul-derived", "inb-rows", "case-email", "case-sysassigned", "case-commodity",
    "case-openitems", "a6-email", "b6-4000", "b6-guardrails",
]

KEEP_POPOVERS = [
    "pld-amend", "str-priority", "str-weights", "rul-unattached", "rul-cost",
    "inb-1039", "case-signals", "case-waterfall", "case-measures",
    "a6-concession", "a6-summary", "b6-queue",
]


def apply(html):
    misses = []
    for find, repl in SWEEP:
        if find not in html:
            misses.append(find.strip()[:70])
            continue
        html = html.replace(find, repl, 1)
    return html, misses


class copy_sweep:  # noqa: N801 — a namespace, so every call site reads unchanged
    SWEEP = SWEEP
    DROP_POPOVERS = DROP_POPOVERS
    KEEP_POPOVERS = KEEP_POPOVERS
    apply = staticmethod(apply)



# ==============================================================================
# 0. PATHS
# ==============================================================================

APP_DIR = Path(__file__).resolve().parent

PROTOTYPE_FILENAME = "pil-scenario4-v4.html"

# Search order for the prototype: beside the app, one level up (the usual
# layout — the app lives in "Pre-existing app", the prototype in its parent),
# then the working directory.
PROTOTYPE_SEARCH_PATHS = [
    APP_DIR / PROTOTYPE_FILENAME,
    APP_DIR.parent / PROTOTYPE_FILENAME,
    Path.cwd() / PROTOTYPE_FILENAME,
]

HERO_CASE_ID = "REQ-1022"

DASH_SHEET = "Quotation Intake Dashboard"
AUDIT_SHEET = "Email Ingestion Audit Log"

# A fallback only: the CSS above stretches the embed to the full viewport
# height. The prototype wants about 1500px of width for its left-nav labels and
# its own applyResponsive() handles anything narrower.
EMBED_HEIGHT = 1000


def resolve_workbook_path() -> Path:
    """ExcelSequenceRenderer.FILE_PATH is relative; resolve it robustly."""
    raw = Path(ExcelSequenceRenderer.FILE_PATH)
    for candidate in (raw, APP_DIR / raw.name, Path.cwd() / raw.name):
        if candidate.exists():
            return candidate
    return APP_DIR / raw.name


def find_prototype() -> Path | None:
    for candidate in PROTOTYPE_SEARCH_PATHS:
        if candidate.exists():
            return candidate
    return None



# ==============================================================================
# EMBEDDED PROTOTYPE
#
# pil-scenario4-v4.html, zlib-compressed and base64-encoded. This is what makes
# the file self-contained. A copy of the HTML found on disk still takes
# precedence, so the prototype can be revised without repackaging.
# ==============================================================================

_PROTOTYPE_BLOB = (
    "eNrMvdty49iVKPieX7FL1WVSLhJJ8CaJKqWPUqIy1VZKWZIy7erqmgyIhEhYJEEDoC4lZ4QfJmb6uXsiJk7EmYfzNJ9w3rv/"
    "xF8y67KvAEhRWeXuSbskAdj3dV977bW/++rw7ODyh/d9Mc6mk1cvvsNfYhLMRnsb4WwDX4TBEH5NwywQg3GQpGG2t7HIruvb"
    "G+r1LJiGexu3UXg3j5NsQwziWRbOoNhdNMzGe8PwNhqEdXqoiWgWZVEwqaeDYBLu+dhIFmWT8NX74xPx7/9LvE+iQTQbid8E"
    "0/mu+H4RZ0EWxTPxt7/+X+L86L24GISzIIli0cbS5/3v636j2fzuJTfy4rs0e8DfvSSOs8cXAgYziRPobRxOw56YRKNxtguv"
    "6/V5NKkn4bD3db/ZbPn7u+ZVfXjT+/r1jv/ab5uiV5NF2Pu68brbet3Ytd5R4cZ2u799wIVnwe0DvOk399tbWDCaQQG/22y1"
    "9rnAdJFht5397uttauo6iGZZ7+vu66397SMuczWCcR0dtY62sUC6SK6DAXR/RP+sV/UmvNw62jmSnU+iWYgz6m8fUdP4jGUO"
    "Xh92+j6XGSVhOIMRHm3vtw931Qvucvuo2+9zsWB6FSawDu3uTqO/q15QsaPDo1a/Jddmkcwn0Gf3sLl9uLOr33BBv//66JAL"
    "4sJy5f5hXw43CYbRIu35jfk9v8ji+VWQ1Me9ThNe8WrW73rNVpcfkwCWHZ6326pGOg6G8V2vIfz5vYBKIhldBVW/W2t1ap3t"
    "mtfobtYEf22VfO1sSqDEs7gnFhH9kc5haWvi4ugdPNTPw9FiEiQ18S6cTeKaOIhnaTwJ0prQZeVQglnaExsX4SgOxYfjjZpI"
    "H9IsnNYXUU3UgzmuCr+pifP4Ks6gsbfh5DbMokFQE/sJ0AXUgVbqaZhE19jqy98S1WUincRZSlQwTIK7mbhO4qnIxqHI4ptw"
    "JuZATVkWiuAqvoWhz+JsjFQ0C+88aAX+fQwm0TAAzBNBKoCU4c9RDLQWTASQNJAlNBalIk6GYSKqiNk1QYhREwzQmgAIbva4"
    "NaKjWZim4iqYDWswRhhOIK4ncQwLdfDxEBqdBwmRLo4mmQLJ30YpUjIVElCNm0JukQRpJm5TIRFcBJMJzChNPfHbl7S0g3Hd"
    "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"
    "akGm70Lt5SDcow3Io8PfWK7vk87Qvm56UWBVlPY/BrbruoCuq0CvrulVt8yfjDgltUvf7R7vvvlpb/d09/Do1ft9d6NeNHnb"
    "aH+Wh5bFi0aXF5SCTJGmYp6fR6N4sK0pExu3N2c3O4xcXyLgNSz01XA+m8X5KMJMz2JuKN9vPoCGUOEzDDe/ia9qiPs7gSSH"
    "REp6dnUuwot8Vp0X08JF6E2zZeJ8JVsERBvk5J9+2D949fr05ENSooC4YwqueRTtSdQ7nCqUB9HgaJf1c0hncRF8xdhiqxzI"
    "gylBtwYim5ah7LVTaZXq8jDWe/rMV1vDafTK5DoM3ZHMeWqbRXN+z7J5uxzu1cd7vFQ3nnkjZfRiLWAHlAMuwZEnhLdQVtfN"
    "xqM4nImzUJhibVPFpWj0mavohrKImEgbN1sIM32kufSuXmt7nFcn0Ly6stxUiASHEjp6CT/d+HyYZwdjuOxWwhJaL3yqKFGD"
    "ZDwKboMpKfjWCTVSVejlXKbpvks0VFGiqnFDKRx9nSOgEy4GEvyHzu/II8J4VPXZD49cPM3DH0HUr5RbGQvFE4pUlKgd2HB1"
    "BCuEIunv1FxRXC56+oy14v2a59KnrF9MEpD6EUoY2NO2kCUEVZKAIm9+/9f//NP+3l3gRD4PThbpmb48UBZsiTZMP/X6C2H6"
    "FSUZUTBMC20pXrwyVaYNxe3Ora+G2kblviBqW1GuCcKlPvhSIFyw5quDcFHhu4BwrdIPlP+NwObapVpoHCkXkMZtAVnbhn8e"
    "wD8P4Z9H8M9j+Ocb+Odb+Oc3ROG2qlCLNXRFDQxmFRIbz5pMUJ+PEGh8MOX+5uCqt5CBT1WVopB/M+DsY8Ot32yurjxV6lOj"
    "ERWA5+tCnC/DLV/V202JzivDj99poR5qiq1ZBKM3/huu1IPNT1ipd1VFd02b7dV9B79qVXVboVCuk2cVGuYnzxbnaXzjU+T8"
    "hsbKTjgVoxxpeYOFUs5bP517/ohsyX/Tpxt+/2lPZjbWw1cs3EELOpZqRhMQj/tpZxVEMvS4sIybQDDQPHCfPYkIg0zSDQzL"
    "Eekf+hSwQ2XMhsUiNMbCD1RgPGHWgB5DKz5yzcKCb1/whBafGZLsHy5yHIe6JG5cAc0VNV+aVeHDpP1mx6uTb8AUazPmBKMV"
    "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"
    "fwEVPm4U"
)


def _embedded_prototype() -> str:
    return zlib.decompress(base64.b64decode(_PROTOTYPE_BLOB)).decode("utf-8")


@st.cache_data(show_spinner=False)
def load_prototype_source() -> str:
    """Disk copy if there is one, embedded copy otherwise."""
    found = find_prototype()
    if found is not None:
        try:
            return found.read_text(encoding="utf-8")
        except OSError:
            pass
    return _embedded_prototype()

# ==============================================================================
# 1. PAGE CONFIG + STREAMLIT CHROME SUPPRESSION
#
# The only Streamlit chrome that ever appears is the progress strip shown while
# the ingestion pipeline is blocking.
# ==============================================================================

st.set_page_config(
    page_title="PIL · Pricing & Quotation",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CHROME_CSS = """
<style>
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"], [data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"], footer { display:none !important; }

.block-container, [data-testid="stMainBlockContainer"]{
  padding:0 !important; max-width:100% !important;
}
[data-testid="stAppViewContainer"], [data-testid="stMain"]{ overflow:hidden !important; }
[data-testid="stVerticalBlock"]{ gap:0 !important; }
html, body, .stApp{ background:#EFF3F8; overflow:hidden; }

/* the workspace fills the viewport; the prototype scrolls inside it */
iframe[data-testid="stIFrame"]{
  height:100vh !important; min-height:100vh !important; width:100% !important;
  border:0 !important; display:block;
}

/* the button that carries an iframe click into Python — never seen, but it has
   to stay laid out, because innerText/click() ignore display:none nodes */
.st-key-pil_ingest_bridge{
  position:fixed !important; left:-10000px !important; top:0 !important;
  width:1px !important; height:1px !important; overflow:hidden !important;
  opacity:0 !important; pointer-events:none !important; z-index:-1 !important;
}

/* the one strip of Streamlit the user ever sees, and only while running.
   It is empty the rest of the time, so the band is drawn only when the
   progress element is actually inside it. */
.st-key-pil_progress{
  position:fixed !important; left:0; right:0; top:0; z-index:9999;
  padding:0 !important; background:transparent !important;
}
.st-key-pil_progress:has([data-testid="stProgress"]){
  background:#FFFFFF !important; padding:6px 18px 8px !important;
  border-bottom:1px solid #E2E8F0;
  box-shadow:0 2px 6px rgba(16,35,58,.10);
}
.st-key-pil_progress [data-testid="stProgress"]{ margin:0 !important; }
.st-key-pil_progress [data-testid="stProgress"] p{
  font-size:12px !important; color:#5A6B80 !important; margin:0 0 4px !important;
}
</style>
"""

st.markdown(CHROME_CSS, unsafe_allow_html=True)


# ==============================================================================
# 2. DATA IN — workbook rows -> prototype case objects
#
# The transforms below are the ones build_mockup_reference.py uses to build
# pil-app-TARGET-mockup.html, ported as they are. In particular _split_route()
# splits on the three arrow forms only: widening it would change the port codes
# the target mockup renders.
# ==============================================================================

DASH = "—"


def _text(value, dash: str = DASH) -> str:
    if value is None:
        return dash
    s = str(value).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return dash
    return s


def _amount(value):
    """'$780/40HQ' -> 780.0 · '$15000' -> 15000.0 · '—' -> None"""
    s = _text(value, "")
    if not s or s == DASH:
        return None
    match = re.search(r"(\d[\d,]*(?:\.\d+)?)", s.replace(" ", ""))
    return float(match.group(1).replace(",", "")) if match else None


def _confidence(value) -> int:
    match = re.search(r"(\d+(?:\.\d+)?)", _text(value, ""))
    return int(round(float(match.group(1)))) if match else 0


def _count(value):
    match = re.search(r"(\d+)", _text(value, ""))
    return int(match.group(1)) if match else None


def _flatten(value) -> str:
    """Free-time cells carry literal or real newlines; flatten for one table row."""
    s = _text(value)
    for a, b in (("\\r\\n", " · "), ("\\n", " · "), ("\r\n", " · "), ("\n", " · "), ("\r", " · ")):
        s = s.replace(a, b)
    return re.sub(r"\s*·\s*", " · ", s).strip(" ·") or DASH


def _split_route(value):
    s = _text(value)
    for sep in ("➔", "->", "→"):
        if sep in s:
            left, _, right = s.partition(sep)
            return _text(left.strip()), _text(right.strip())
    return s, DASH


def _shorten(value: str, limit: int = 30) -> str:
    s = _text(value)
    return s if len(s) <= limit else s[: limit - 1].rstrip() + "…"


@st.cache_data(show_spinner=False)
def load_workbook_frames(path_str: str, mtime: float):
    """Reads both sheets. mtime is part of the cache key so edits invalidate it."""
    path = Path(path_str)
    if not path.exists():
        return pd.DataFrame(), pd.DataFrame()
    try:
        xls = pd.ExcelFile(path)
        dash_df = pd.read_excel(xls, sheet_name=DASH_SHEET) if DASH_SHEET in xls.sheet_names else pd.DataFrame()
        audit_df = pd.read_excel(xls, sheet_name=AUDIT_SHEET) if AUDIT_SHEET in xls.sheet_names else pd.DataFrame()
        return dash_df, audit_df
    except Exception as exc:  # noqa: BLE001 — surfaced to the operator, not swallowed
        st.error("Error reading Excel file: " + str(exc))
        return pd.DataFrame(), pd.DataFrame()


def get_dashboard_data(file_path: str):
    """Kept for backwards compatibility with the original call signature."""
    path = Path(file_path)
    mtime = path.stat().st_mtime if path.exists() else 0.0
    return load_workbook_frames(str(path), mtime)


def normalise_legs(df: pd.DataFrame) -> pd.DataFrame:
    """A multi-leg request repeats only its leg row; carry the header fields down."""
    if df.empty:
        return df
    out = df.copy()
    for col in ("ID", "CUSTOMER", "STATUS", "CONFIDENCE", "MISSING DATA?"):
        if col in out.columns:
            out[col] = out[col].ffill()
    return out


def build_live_cases(df: pd.DataFrame) -> list[dict]:
    """Map dashboard rows onto the prototype's case shape."""
    if df.empty:
        return []

    cases: list[dict] = []

    for _, row in normalise_legs(df).iterrows():
        req_id = _text(row.get("ID"), "")
        if not req_id or req_id == DASH:
            continue

        pol, pod = _split_route(row.get("POL ➔ POD"))
        commodity = _text(row.get("COMMODITY"))
        missing = _text(row.get("MISSING DATA?"), "NO")
        conf_val = row.get("Completion%") if "Completion%" in row else row.get("CONFIDENCE")
        cases.append(
            {
                "id": req_id,
                "hero": False,
                "leg": _text(row.get("LEG SEQ #"), "1").replace(".0", ""),
                "customer": _text(row.get("CUSTOMER")),
                "crmId": DASH,
                "segment": DASH,
                "pol": pol,
                "pod": pod,
                "validity": _text(row.get("VALIDITY")),
                "container": _text(row.get("CONTAINER")),
                "weight": _text(row.get("WT (T)")),
                "commodity": commodity,
                "commodityShort": _shorten(commodity),
                "cntr": _count(row.get("# CNTR")),
                "service": _text(row.get("SERVICE / VOYAGE")),
                "proposed": _amount(row.get("SYSTEM PROPOSED RATE")),
                "requested": _amount(row.get("CUSTOMER REQUESTED RATE")),
                "flags": _text(row.get("FLAGS")),
                "vas": _text(row.get("VAS")),
                "freeTime": _flatten(row.get("FREE TIME")),
                "remarks": _text(row.get("REMARKS")),
                "status": _text(row.get("STATUS"), "New"),
                "confidence": _confidence(row.get("CONFIDENCE")),
                # readiness() calls .replace()/regex on this — must always be a string
                "missing": missing if missing != DASH else "NO",
            }
        )

    return cases


def build_audit_rows(df: pd.DataFrame) -> list[dict]:
    """Map audit-sheet rows onto the shape SCREENS.mailaudit renders."""
    if df.empty:
        return []

    rows: list[dict] = []

    for _, row in df.iterrows():
        rows.append(
            {
                "received": _text(row.get("RECEIVED DATE")),
                "thread": _text(row.get("FILE / THREAD ID")),
                "sender": _text(row.get("SENDER")),
                "subject": _text(row.get("SUBJECT")),
                "topic": _text(row.get("COMMUNICATION TOPIC")),
                "verdict": _text(row.get("CLASSIFICATION VERDICT")),
                "action": _text(row.get("ACTION TAKEN")),
            }
        )

    return rows


def js(obj) -> str:
    """JSON that is safe to sit inside a <script> block."""
    return (
        json.dumps(obj, ensure_ascii=False)
        .replace("</", "<\\/")
        .replace(" ", "\\u2028")
        .replace(" ", "\\u2029")
    )


# ==============================================================================
# 3. PROTOTYPE TRANSFORM
#
# Substitutions 1-6 of the merger plan, then the copy sweep, then the extension
# block. This is the same sequence build_mockup_reference.py applies, so the
# output of transform_prototype() is pil-app-TARGET-mockup.html.
# ==============================================================================

SCRIPT_ANCHOR = '<script>\n"use strict";'
BOOT_ANCHOR = "applyResponsive(false);\nrender();"

# 5. RFP wording -> neutral product wording
RFP_WORDING = [
    ("PIL · Pricing &amp; Quotation — RFP Scenario 4 · REQ-1022", "PIL · Pricing &amp; Quotation"),
    ("Why this name, not the RFP’s", "Why this name"),
    ("§4.3.4 literally names a “Market Share Growth Strategy”. That one is seeded and active on Far East ⇄ Europe as STR-002, so the RFP’s wording is satisfied and visible in the table above.",
     "A “Market Share Growth Strategy” is seeded and active on Far East ⇄ Europe as STR-002 and is visible in the table above."),
    ("They are the §4.3.4 parameter catalogue, grouped as the RFP groups it.",
     "They are the standard parameter catalogue, grouped by family."),
    ("The RFP’s rule table quotes percentages", "The governing rule table quotes percentages"),
    ("RFP Scenario 4 (§4.3.4)", "Strategy-Driven Pricing & Quotation"),
]

EXT = r"""
/* ==========================================================================
   EXTENSION — applied after the prototype's own definitions.
   Nothing above this point is edited in place.
   ========================================================================== */
   
/* the guided demo is gone: keep the rail permanently closed */
S.railOpen = false; S.railUserSet = true;

/* drop the popovers that narrate the build rather than explain the product */
__DROP_POPS__.forEach(function(k){ delete POPS[k]; });

/* ---- inbox filter state ---------------------------------------------- */
S.inbox = { q: '', status: '' };
S.ingest = { running:false, done:false, log:[] };

/* ---- 1. NAV: two new Sales screens ------------------------------------ */
NAV.sales = [
  { sec:'Sales' },
  { id:'inbox',      label:'Quotation Inbox', icon:'inbox' },
  { sec:'Mailbox' },
  { id:'ingest',     label:'Ingestion Console', icon:'history' },
  { id:'mailaudit',  label:'Email Audit Log',   icon:'list' }
];

/* ---- 2. INBOX: ingestion + export + filters in the page head ---------- */
const _inbox = SCREENS.inbox;
SCREENS.inbox = function(){
  var out = _inbox();

  var actions =
    '<button type="button" class="btn primary" data-act="run-ingest">' +
      icon('history', 15) + 'Run Batch Ingestion</button>' +
    '<button type="button" class="btn" data-act="export-xlsx">' +
      icon('eye', 15) + 'Export to Excel</button>';
  out = out.replace('<div class="page-hd-act">', '<div class="page-hd-act">' + actions);

  var statuses = [];
  S.cases.forEach(function(c){ if (statuses.indexOf(c.status) < 0) statuses.push(c.status); });

  var filters =
    '<div class="card"><div class="card-bd"><div class="rowflex">' +
      '<input id="inbQ" type="text" style="width:340px;flex:none" placeholder="Search customer, route or commodity" ' +
        'value="' + esc(S.inbox.q) + '" data-act="inbox-q" oninput="ACTIONS[\'inbox-q\']({},this)">' +
      '<select id="inbS" style="width:200px;flex:none" onchange="ACTIONS[\'inbox-s\']({},this)">' +
        '<option value="">All statuses</option>' +
        statuses.map(function(s){
          return '<option value="' + esc(s) + '"' + (S.inbox.status === s ? ' selected' : '') + '>' + esc(s) + '</option>';
        }).join('') +
      '</select>' +
      '<span class="wrap-note">Showing ' + inboxRows().length + ' of ' + S.cases.length + ' leg rows</span>' +
    '</div></div></div>';

  out = out.replace('<div class="card clip">', filters + '<div class="card clip">');
  /* Rename Confidence column header to Completion% in Quotation Inbox table */
out = out.replace('<th>CONFIDENCE</th>', '<th>Completion%</th>')
           .replace('<th>Confidence</th>', '<th>Completion%</th>');
  return out;
};

function inboxRows(){
  var q = S.inbox.q.toLowerCase(), st = S.inbox.status;
  return S.cases.filter(function(c){
    if (st && c.status !== st) return false;
    if (!q) return true;
    return (c.customer + ' ' + c.pol + ' ' + c.pod + ' ' + c.commodity + ' ' + c.id).toLowerCase().indexOf(q) >= 0;
  });
}

/* the original inbox maps over S.cases; filter by hiding non-matching rows */
const _renderOrig = render;
render = function(){
  _renderOrig();
  if (S.screen !== 'inbox') return;
  var keep = {};
  inboxRows().forEach(function(c){ keep[c.id + '|' + c.leg] = 1; });
  $$('#content table.tbl tbody tr').forEach(function(tr){
    var id  = tr.cells[0] ? tr.cells[0].innerText.trim() : '';
    var leg = tr.cells[1] ? tr.cells[1].innerText.trim() : '';
    if (!keep[id + '|' + leg]) tr.style.display = 'none';
  });
};

/* ---- 3. INGESTION CONSOLE -------------------------------------------- */
SCREENS.ingest = function(){
  var body;
  if (!S.ingest.log.length){
    body =
      '<div class="card"><div class="card-bd" style="text-align:center;padding:40px 20px">' +
        '<div style="color:var(--muted);font-size:13px;margin-bottom:14px">No runs yet</div>' +
        '<button type="button" class="btn primary" data-act="run-ingest">' + icon('history',15) + 'Run Batch Ingestion</button>' +
      '</div></div>';
  } else {
    body =
      '<div class="card clip"><div class="card-hd"><h2>Execution log</h2>' +
        '<span class="sub">Most recent run</span>' +
        '<div class="card-hd-act">' + (S.ingest.running
          ? '<span class="chip amber">Running</span>'
          : (S.ingest.log.some(function(l){ return l.kind === 'red'; })
              ? '<span class="chip red">Failed</span>'
              : '<span class="chip green">Complete</span>')) +
          '<button type="button" class="btn" data-act="run-ingest">' + icon('history',15) + 'Run again</button>' +
        '</div></div>' +
      '<div class="card-bd"><div class="stack">' +
        S.ingest.log.map(function(l){
          return '<div class="rowflex" style="gap:8px"><span class="chip ' + esc(l.kind) + '">' + esc(l.tag) +
            '</span><span class="small">' + l.text + '</span></div>';
        }).join('') +
      '</div></div></div>';
  }

  return pageHead('', 'Ingestion Console', '', CLIENT_TAG) + body;
};

/* ---- 4. EMAIL AUDIT LOG ---------------------------------------------- */
/* ---- 4. EMAIL AUDIT LOG ---------------------------------------------- */
SCREENS.mailaudit = function(){
  var rows = (window.__PIL_AUDIT_ROWS || []);
  var req = rows.filter(function(r){ return /RATE REQUEST/i.test(r.verdict); }).length;
  var chat = rows.length - req;

  var body =
    '<div class="kpis">' +
      kpi('Emails scanned', String(rows.length)) +
      kpi('Rate requests', String(req), 'in the Quotation Inbox') +
      kpi('Operational chatter', String(chat), 'no quotation raised') +
    '</div>' +
    '<div class="card clip"><div class="card-hd"><h2>Every email scanned</h2>' +
      '<div class="card-hd-act">' + CLIENT_TAG + '</div></div>' +
    '<div class="card-bd tight"><div class="tbl-wrap"><table class="tbl"><thead><tr>' +
      '<th>Received</th><th>Thread</th><th>Sender</th><th>Subject</th><th>Topic</th>' +
      '<th>Verdict</th><th>Action taken</th></tr></thead><tbody>' +
      rows.map(function(r){
        var isReq = /RATE REQUEST/i.test(r.verdict);
        return '<tr>' +
          '<td class="mono xsmall">' + esc(r.received) + '</td>' +
          '<td class="mono xsmall">' + esc(r.thread) + '</td>' +
          '<td class="xsmall">' + esc(r.sender) + '</td>' +
          '<td class="small">' + esc(r.subject) + '</td>' +
          '<td class="xsmall">' + esc(r.topic) + '</td>' +
          '<td><span class="chip ' + (isReq ? 'blue' : 'grey') + '">' + esc(r.verdict) + '</span></td>' +
          '<td class="xsmall">' + esc(r.action) + '</td>' +
        '</tr>';
      }).join('') +
    '</tbody></table></div></div></div>';

  return pageHead('', 'Email Audit Log', '', CLIENT_TAG) + body;
};

/* ---- 5. ACTIONS ------------------------------------------------------- */
ACTIONS['inbox-q'] = function(d, el){ S.inbox.q = el.value; render(); };
ACTIONS['inbox-s'] = function(d, el){ S.inbox.status = el.value; render(); };

ACTIONS['export-xlsx'] = function(){
  toast('Workbook ' + WORKBOOK + ' downloaded.', 'ok');
};

ACTIONS['run-ingest'] = function(){
  /* MOCKUP ONLY — in the merged app this posts to the Streamlit host, which
     calls fetch_extracted_email_payloads() and SequenceSourcingAgent
     exactly as new_app.py does today, then reruns with the new rows. */
  if (S.ingest.running) return;
  S.ingest.running = true; S.ingest.log = [];
  go('ingest');
  var script = [
    { kind:'blue',  tag:'Mailbox', text:'Scanning account mailbox <strong>PIL LatAm desk</strong>' },
    { kind:'grey',  tag:'1 / 4',   text:'Ingesting <span class="mono">RE_ Quote request LZC-GYE.msg</span> — <em>Rate request Manzanillo to Guayaquil</em>' },
    { kind:'green', tag:'Request', text:'Extracted 1 rate request, 1 leg — appended as <span class="mono">REQ-1043</span>' },
    { kind:'grey',  tag:'2 / 4',   text:'Ingesting <span class="mono">Booking confirmation 8871.msg</span> — <em>Booking confirmed</em>' },
    { kind:'grey',  tag:'Chatter', text:'Operational payload — logged to the audit sheet, no quotation raised' },
    { kind:'grey',  tag:'3 / 4',   text:'Ingesting <span class="mono">Tarifa multi-tramo.eml</span> — 2 attachments parsed' },
    { kind:'green', tag:'Request', text:'Extracted 1 rate request, 3 legs — appended as <span class="mono">REQ-1044</span>' },
    { kind:'grey',  tag:'4 / 4',   text:'Ingesting <span class="mono">FW_ vessel schedule.msg</span> — <em>Schedule enquiry</em>' },
    { kind:'grey',  tag:'Chatter', text:'Operational payload — logged to the audit sheet' },
    { kind:'green', tag:'Done',    text:'<strong>4 payloads processed · 2 rate requests · 4 legs written to the workbook</strong>' }
  ];
  var i = 0;
  (function tick(){
    if (i >= script.length){
      S.ingest.running = false; S.ingest.done = true;
      render(); toast('Batch ingestion complete — 2 new rate requests in the inbox.', 'ok');
      return;
    }
    S.ingest.log.push(script[i++]); render(); setTimeout(tick, 420);
  })();
};

/* inputs need an inline hook because render() rebuilds innerHTML */
window.ACTIONS = ACTIONS;
"""


def transform_prototype(source: str, cases: list[dict], audit_rows: list[dict],
                        workbook_label: str) -> tuple[str, list[str]]:
    """Returns (html, notes). Every anchor is asserted, so a prototype revision
    that moves one fails loudly here instead of silently rendering the old UI."""
    notes: list[str] = []
    html = source

    # 1. payload + hero-preserving merge, injected before the prototype's script
    pre = (
        "<script>\n"
        "window.__PIL_LIVE_CASES = " + js(cases) + ";\n"
        "window.__PIL_AUDIT_ROWS = " + js(audit_rows) + ";\n"
        "window.__PIL_WORKBOOK = " + js(workbook_label) + ";\n"
        "window.__PIL_MERGE = function (CASES) {\n"
        "  var live = window.__PIL_LIVE_CASES || [];\n"
        "  if (!live.length) return [];\n"
        "  return live;\n"
        "};\n"
        "</script>\n"
    )
    if SCRIPT_ANCHOR not in html:
        raise RuntimeError("prototype script anchor not found — cannot inject live data")
    html = html.replace(SCRIPT_ANCHOR, pre + SCRIPT_ANCHOR, 1)
    html = html.replace("cases: CASES,", "cases: window.__PIL_MERGE(CASES),", 1)
    html = re.sub(r"const WORKBOOK\s*=\s*'[^']*';",
                  "const WORKBOOK = '" + workbook_label.replace("'", "") + "';", html, count=1)
    notes.append("live data injected (%d cases, %d audit rows)" % (len(cases), len(audit_rows)))

    # 2. scenario pill out of the top bar
    html = html.replace('<span class="scenario-pill">RFP Scenario 4 · §4.3.4</span>', "")
    notes.append("scenario pill removed")

    # 3. the guided-demo rail. The #rail / #railpill nodes STAY in the DOM — the
    #    prototype binds listeners to them at boot and throws if they are gone.
    #    They are hidden in CSS; the extension block forces the rail closed.
    html = html.replace("</style>", ".rail,.railpill{display:none !important}\n</style>", 1)
    notes.append("guided-demo rail removed")

    # 4. RFP wording out of the visible copy
    for find, repl in RFP_WORDING:
        html = html.replace(find, repl)
    notes.append("RFP wording neutralised")

    # 5. COPY SWEEP — remove the app's self-narration
    html, misses = copy_sweep.apply(html)
    if misses:
        raise RuntimeError("copy sweep did not match: " + "; ".join(misses))
    notes.append("copy sweep applied (%d replacements)" % len(copy_sweep.SWEEP))

    # 6. the extension block, immediately before the prototype's own boot
    if BOOT_ANCHOR not in html:
        raise RuntimeError("prototype boot anchor not found — cannot append the extension block")
    ext = EXT.replace("__DROP_POPS__", json.dumps(copy_sweep.DROP_POPOVERS))
    html = html.replace(BOOT_ANCHOR, ext + "\n" + BOOT_ANCHOR, 1)
    notes.append("%d narrating popovers dropped, %d product popovers kept"
                 % (len(copy_sweep.DROP_POPOVERS), len(copy_sweep.KEEP_POPOVERS)))
    notes.append("extension block appended (nav, inbox controls, 2 new screens, actions)")

    return html, notes


def restate_source_counts(html: str, leg_rows: int) -> str:
    """Substitution 6 of the merger plan: the prototype was written against a
    41-row extract and says so in the case workspace field notes. The V3
    filename is already handled by the WORKBOOK constant, but the count is
    hard-coded, so it is restated against the workbook actually loaded.

    This is deliberately NOT part of transform_prototype(): that function is
    held byte-identical to pil-app-TARGET-mockup.html, and the mockup builder
    never applied this step. Keeping it separate means the parity check stays
    exact and this correction stays visible.
    """
    if not leg_rows or "41 rows" not in html:
        return html
    return html.replace("41 rows", str(leg_rows) + " rows")


# ------------------------------------------------------------------------------
# The guided demo is gone, so the copy must stop referring to its numbered
# steps. copy_sweep.py catches three of these; the rest are the strings that
# only render once you are deep in a screen, so they are swept here instead of
# by editing that module. Every pair is asserted, exactly as the copy sweep
# asserts its own, so a prototype revision that rewords one fails loudly.
#
# Left alone on purpose: the data-act="step" hooks (wiring, not copy — their
# labels already read "Go to Price Lists" / "Go to Strategy Builder"), the
# HTML step="..." input attributes, and the guided-demo rail's own text, which
# is never rendered because the rail is hidden.
# ------------------------------------------------------------------------------

STEP_SWEEP = [
    ("toast(d.id + ' is a real extracted record, but only REQ-1022 is wired to the pricing engine in this prototype.', 'warn');",
     "toast('Pricing for ' + d.id + ' is not configured yet.', 'warn');"),

    ('every figure on step 5 follows the new base rate immediately, because the waterfall reads the live price list rather than a stored number. Optional beat — if you take it, take it before step 5.',
     'every figure in the case workspace follows the new base rate immediately, because the waterfall reads the live price list rather than a stored number.'),

    ('Configure steps 1–3 first — the quotation engine needs a price list, a strategy and its rules.',
     'Configure the engine first — it needs a price list, a strategy and its rules.'),

    ('Tip: step 1 publishes the USD 4,000 base rate this strategy quotes from.',
     'Tip: the price list publishes the USD 4,000 base rate this strategy quotes from.'),

    ('No strategy is linked to this price list yet. Step 2 creates STR-007 and links it here.',
     'No strategy is linked to this price list yet. The Strategy Builder creates STR-007 and links it here.'),

    ('Rules hang off a strategy. Step 2 creates <span class="mono">STR-007</span>; come back here once it exists.',
     'Rules hang off a strategy. The Strategy Builder creates <span class="mono">STR-007</span>; come back here once it exists.'),

    ('Work case REQ-1022 through to step 6 and accept the recommendation',
     'Work case REQ-1022 through and accept the recommendation'),

    ('<h2>Step 6 — Execute</h2>',
     '<h2>Execute</h2>'),

    ('With the current rules and thresholds the recommendation is below the floor. Check the configuration in step 3.',
     'With the current rules and thresholds the recommendation is below the floor. Check the rules and guardrails.'),

    ('Steps 1–3 create the price list, the strategy and its rules. Until they exist there is nothing for the engine to calculate from.',
     'The price list, the strategy and its rules have to exist first. Until they do there is nothing for the engine to calculate from.'),

    ('Switch to the Trade &amp; Pricing persona and complete steps 1 to 3 first.',
     'Switch to the Trade &amp; Pricing persona and configure the engine first.'),

    ('Nothing has breached a guardrail. In step 6 the sales rep can raise an exception',
     'Nothing has breached a guardrail. The sales rep can raise an exception'),

]


def drop_step_language(html: str) -> str:
    misses = [find for find, _ in STEP_SWEEP if find not in html]
    if misses:
        raise RuntimeError("step sweep did not match: "
                           + "; ".join(m[:70] for m in misses))
    for find, repl in STEP_SWEEP:
        html = html.replace(find, repl)
    return html


# ==============================================================================
# 4. THE LIVE BRIDGE
#
# Everything above produces the target mockup. This adds the three things a
# running app needs that a static mockup does not:
#
#   a. the Ingestion Console shows the real log from the last pipeline run
#   b. Run Batch Ingestion reaches Python, through a hidden Streamlit button in
#      the parent document (the component iframe is srcdoc + allow-same-origin)
#   c. persona and screen survive the Streamlit rerun, via sessionStorage —
#      without this the app jumps back to Price Lists every time ingestion runs
# ==============================================================================

LIVE_EXT = r"""
/* ==========================================================================
   LIVE BRIDGE — appended after the extension block. Still append-only.
   ========================================================================== */

/* ---- a. the Ingestion Console shows the real run, not a scripted one ---- */
S.ingest.log = (window.__PIL_INGEST_LOG || []).slice();
S.ingest.running = false;
S.ingest.done = S.ingest.log.length > 0;

/* ---- b. Run Batch Ingestion -> Python ---------------------------------- */
function pilHostButton(){
  try {
    var pdoc = window.parent && window.parent.document;
    if (!pdoc) return null;
    var all = pdoc.querySelectorAll('button');
    for (var i = 0; i < all.length; i++){
      /* Streamlit renders a button label as markdown, so the wrapping
         underscores of __PIL_RUN_INGEST__ reach the DOM as <strong> and the
         text reads PIL_RUN_INGEST. Match the part that survives either way. */
      if ((all[i].textContent || '').indexOf('PIL_RUN_INGEST') >= 0) return all[i];
    }
  } catch (e) {}
  return null;
}

ACTIONS['run-ingest'] = function(){
  if (S.ingest.running) return;
  var btn = pilHostButton();
  if (!btn){
    S.ingest.log = [{ kind:'red', tag:'Unavailable', text:
      'The workspace could not reach the ingestion service in this browser. ' +
      'Reload the page and try again.' }];
    S.ingest.done = true;
    go('ingest');
    toast('Ingestion service unreachable from this browser.', 'warn');
    return;
  }
  S.ingest.running = true;
  S.ingest.log = [{ kind:'blue', tag:'Running', text:
    'Scanning the mailbox and extracting rate requests. The workspace pauses until the run finishes.' }];
  go('ingest');
  setTimeout(function(){ btn.click(); }, 50);
};

/* the workbook rides along as a data URI, so Export to Excel is a real
   download rather than a message about one */
ACTIONS['export-xlsx'] = function(){
  var b64 = window.__PIL_WORKBOOK_B64;
  if (!b64){ toast('Workbook not available for download.', 'warn'); return; }
  try {
    var a = document.createElement('a');
    a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + b64;
    a.download = WORKBOOK;
    document.body.appendChild(a); a.click(); a.parentNode.removeChild(a);
    toast('Workbook ' + WORKBOOK + ' downloaded.', 'ok');
  } catch (e){
    toast('The browser blocked the download.', 'warn');
  }
};

/* ---- c. persona + screen survive a Streamlit rerun --------------------- */
var PIL_STATE_KEY = 'pil-app-state';

function pilSaveState(){
  try {
    sessionStorage.setItem(PIL_STATE_KEY, JSON.stringify({
      persona: S.persona,
      screen:  S.screen,
      caseId:  S.activeCaseId,
      listId:  S.activePriceListId
    }));
  } catch (e) {}
}

var _pilGo = go;
go = function(screen){ _pilGo(screen); pilSaveState(); };

var _pilSetPersona = setPersona;
setPersona = function(k){ _pilSetPersona(k); pilSaveState(); };

(function pilRestoreState(){
  var raw = null;
  try { raw = sessionStorage.getItem(PIL_STATE_KEY); } catch (e){ return; }
  if (!raw) return;
  var saved = null;
  try { saved = JSON.parse(raw); } catch (e){ return; }
  if (!saved) return;

  if (saved.persona === 'pricer' || saved.persona === 'sales') S.persona = saved.persona;
  if (saved.caseId) S.activeCaseId = saved.caseId;
  if (saved.listId) S.activePriceListId = saved.listId;

  var screen = saved.screen;
  if (!screen || !SCREENS[screen]) screen = null;
  /* the two detail screens need their subject; without it, fall back home */
  if (screen === 'case' && !getCase(S.activeCaseId)) screen = null;
  if (screen === 'pricelist-detail' && !getPriceList(S.activePriceListId)) screen = null;

  S.screen = screen || homeFor(S.persona);
})();
"""


def add_live_bridge(html: str, ingest_log: list[dict], workbook_b64: str) -> str:
    payload = (
        "<script>\n"
        "window.__PIL_INGEST_LOG = " + js(ingest_log) + ";\n"
        "window.__PIL_WORKBOOK_B64 = " + js(workbook_b64) + ";\n"
        "</script>\n"
    )
    html = html.replace(SCRIPT_ANCHOR, payload + SCRIPT_ANCHOR, 1)
    return html.replace(BOOT_ANCHOR, LIVE_EXT + "\n" + BOOT_ANCHOR, 1)


# ==============================================================================
# 5. THE PIPELINE
#
# Logic unchanged from the original console: same order, same guards, same four
# call sites into the protected modules. The only difference is that each line
# is appended to a log the Ingestion Console renders, instead of st.write().
# ==============================================================================

def _log(log: list[dict], kind: str, tag: str, text: str) -> None:
    log.append({"kind": kind, "tag": tag, "text": text})


def _esc(value) -> str:
    return html_lib.escape(str(value))


def run_ingestion(config: dict, model_name: str, progress=None) -> list[dict]:
    """Returns the log lines for one batch run. Never raises — a failure is a
    red line in the console, which is the whole point of running it here."""
    log: list[dict] = []
    _log(log, "blue", "Start", "Initializing Sequence Sourcing Agent")

    try:
        agent = SequenceSourcingAgent(
            api_key=config["OPENAI_API_KEY"],
            base_url=config["TIGER_AI_GATEWAY_URL"],
            model_name=model_name,
        )

        accounts_list = config.get("ACCOUNTS", [])

        if not accounts_list:
            _log(log, "red", "Config", "No email accounts configured in <span class=\"mono\">config.json</span>.")
            return log

        total_processed = 0

        for acc in accounts_list:
            account_name = acc.get("account_name", "Default") if isinstance(acc, dict) else str(acc)

            _log(log, "blue", "Mailbox", "Scanning account mailbox <strong>" + _esc(account_name) + "</strong>")

            try:
                extracted_emails = fetch_extracted_email_payloads(acc, config)
            except Exception as exc:  # noqa: BLE001
                _log(log, "red", "Error",
                     "Could not read the mailbox for <strong>" + _esc(account_name) + "</strong> — " + _esc(exc))
                continue

            num_emails = len(extracted_emails)

            if num_emails == 0:
                _log(log, "amber", "Empty", "No email payloads found in the inbox directory.")
                continue

            for idx, email_payload in enumerate(extracted_emails, start=1):
                file_label = email_payload.get(
                    "file_name",
                    "Email #" + str(email_payload.get("email_sequence_index", idx)),
                )
                subject = email_payload.get("subject", "")

                _log(log, "grey", str(idx) + " / " + str(num_emails),
                     "Ingesting <span class=\"mono\">" + _esc(file_label) + "</span>"
                     + (" — <em>" + _esc(subject) + "</em>" if subject else ""))

                if email_payload.get("attached_files"):
                    _log(log, "grey", "Attachments", _esc(email_payload["attached_files"]))

                sig = inspect.signature(agent.process_and_save)

                if len(sig.parameters) > 1:
                    saved = agent.process_and_save(email_payload, agent.request_seq_counter)
                else:
                    saved = agent.process_and_save(email_payload)

                if saved:
                    _log(log, "green", "Request",
                         "Rate request(s) extracted from <span class=\"mono\">" + _esc(file_label) + "</span>")
                else:
                    _log(log, "grey", "Chatter",
                         "Operational payload — logged to the audit sheet, no quotation raised")

                total_processed += 1

                if progress:
                    progress(min(idx / num_emails, 1.0))

        _log(log, "green", "Done",
             "<strong>" + str(total_processed) + " email payload(s) processed</strong>")

        # New rows are on disk now — drop the cached read so the inbox and the
        # audit log pick them up on this render.
        load_workbook_frames.clear()

    except Exception as exc:  # noqa: BLE001 — surfaced in the console, not swallowed
        _log(log, "red", "Failed", "Pipeline execution failed — " + _esc(exc))

    return log


# ==============================================================================
# 6. RENDER
# ==============================================================================

WORKBOOK_PATH = resolve_workbook_path()
PROTOTYPE_PATH = find_prototype()
PROTOTYPE_SOURCE = load_prototype_source()

config = load_config("config.json")

# No sidebar any more: the value source changes, the call signature does not.
model_name = config.get("MODEL_NAME", "gpt-4o-mini")

if "ingest_log" not in st.session_state:
    st.session_state.ingest_log = []

# The strip is the only Streamlit chrome that ever appears, and only while the
# pipeline is blocking. The slot is claimed before the run — the keyed container
# carries the CSS, the st.empty() inside it is what can actually be cleared
# again once the run is over.
with st.container(key="pil_progress"):
    progress_slot = st.empty()

# The bridge button. ACTIONS['run-ingest'] inside the iframe finds it by text in
# window.parent.document and clicks it; CSS parks it off-screen. The label is
# run through markdown on the way to the DOM, so the iframe matches on the
# underscore-free core of it.
with st.container(key="pil_ingest_bridge"):
    run_pipeline = st.button("__PIL_RUN_INGEST__", key="pil_run_ingest")

if run_pipeline:
    RUNNING = "Running batch ingestion — scanning the mailbox and extracting rate requests…"
    bar = progress_slot.progress(0.0, text=RUNNING)
    st.session_state.ingest_log = run_ingestion(
        config, model_name, progress=lambda f: bar.progress(f, text=RUNNING)
    )
    progress_slot.empty()

dash_df, audit_df = get_dashboard_data(str(WORKBOOK_PATH))
live_cases = build_live_cases(dash_df)
audit_rows = build_audit_rows(audit_df)

if not PROTOTYPE_SOURCE:
    st.error(
        "The embedded prototype could not be decoded, and no **"
        + PROTOTYPE_FILENAME + "** was found on disk. Reinstall new_app.py."
    )
else:
    workbook_b64 = (
        base64.b64encode(WORKBOOK_PATH.read_bytes()).decode("ascii")
        if WORKBOOK_PATH.exists() else ""
    )

    page = PROTOTYPE_SOURCE
    # transform_prototype() raises on a missed anchor, so the notes are only a
    # record of what landed — nothing downstream branches on them.
    page, _notes = transform_prototype(page, live_cases, audit_rows, WORKBOOK_PATH.name)
    page = restate_source_counts(page, len(live_cases))
    page = drop_step_language(page)
    page = add_live_bridge(page, st.session_state.ingest_log, workbook_b64)

    components.html(page, height=EMBED_HEIGHT, scrolling=True)
