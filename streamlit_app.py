#!/usr/bin/env python
"""
Streamlit app: RMS Analysis Total AAL lookup

Enter one or more App Analysis IDs and a Perspective, and the app resolves
each App Analysis ID to its internal analysisId, pulls analysis metadata
(Portfolio Name / LOB) and /stats (Total AAL = purePremium, CV) for the
chosen perspective(s), shows the results in a table, and lets you download
them as a CSV.

This re-uses the same resolution logic as rms_analysis_total_aal.py:
  1. GET /analyses?filter=appAnalysisId={id}   -> resolves analysisId,
     Portfolio Name (exposureResourceName) and LOB (lobName)
  2. GET /analyses/{analysisId}                -> analysis metadata
  3. GET /analyses/{analysisId}/stats?perspectiveCode=...&exposureResourceId=...
     &exposureResourceType=...                  -> Total AAL (purePremium) + CV
"""

import os
import re
import time
import threading
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urlencode, quote

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter

BASE = "https://api-euw1.rms.com/platform/riskdata/v1"
API_KEY = os.environ.get(
    "RMS_API_KEY_RISKDATA",
    "cG23ZJ4i5pwGkR-RiqIXVqCHfc1dcYmnU-y5vWvzDVg",
)

PERSPECTIVES = ["GU", "GR", "RL", "RP"]  # Ground Up, Gross, Net Pre Cat, Net Post Cat
DEFAULT_PERSPECTIVE = "GR"
TIMEOUT = 120
MAX_RETRIES = 5

HEADERS = [
    "App Analysis ID", "Analysis ID", "Analysis Name", "Portfolio Name", "LOB",
    "Create Date", "Analysis Type", "Peril", "Subperil", "Perspective",
    "Exposure/Portfolio ID", "Exposure Resource Type", "Currency",
    "Total AAL", "CV", "Status",
]


# --------------------------------------------------------------------------
# Core helpers (ported from rms_analysis_total_aal.py)
# --------------------------------------------------------------------------

def parse_analysis_ids(raw):
    """Pull out every whole number from the input text - one per line, comma
    separated, space separated, or any mix. Ignores blank lines / '#' comments."""
    ids = []
    for line in str(raw).splitlines():
        line = line.split("#", 1)[0]
        for match in re.findall(r"\d+", line):
            if match not in ids:
                ids.append(match)
    return ids


def gv(obj, key):
    """Case-insensitive dict getter."""
    if not isinstance(obj, dict):
        return None
    if key in obj:
        return obj[key]
    wanted = key.lower()
    for current_key, value in obj.items():
        if str(current_key).lower() == wanted:
            return value
    return None


def text(obj, key):
    value = gv(obj, key)
    return "" if value is None or isinstance(value, (dict, list)) else str(value)


def numeric_or_blank(value):
    if value in (None, "") or isinstance(value, bool):
        return ""
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


class Api:
    """Thread-safe requests wrapper with a pooled session and SSL fallback."""

    def __init__(self, api_key, verify=True, pool_size=32):
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        self.session.headers.update({"Accept": "application/json", "Authorization": api_key})
        self.session.verify = verify
        self.lock = threading.Lock()
        self.ssl_fallback_done = not verify

    def disable_ssl_verify(self):
        with self.lock:
            if self.ssl_fallback_done:
                return
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            self.session.verify = False
            self.ssl_fallback_done = True

    def get(self, url, retries=MAX_RETRIES):
        last_exc = None
        for attempt in range(retries):
            try:
                response = self.session.get(url, timeout=TIMEOUT)
                if response.status_code == 200:
                    try:
                        return response.json(), response.status_code
                    except ValueError:
                        return None, response.status_code
                if response.status_code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 20))
                    continue
                return None, response.status_code
            except requests.exceptions.SSLError:
                if self.session.verify:
                    self.disable_ssl_verify()
                    continue
                last_exc = "SSL error persisted with verification disabled"
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 20))
                    continue
                return None, 0
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 20))
                    continue
                return None, 0
        return None, 0


def resolve_analysis_id(app_analysis_id, api):
    """Resolve App Analysis ID -> analysisId, Portfolio Name, LOB."""
    flt = quote(f"appAnalysisId={app_analysis_id}")
    payload, status = api.get(f"{BASE}/analyses?filter={flt}&limit=5")
    if status != 200 or payload is None:
        return None, status, None

    items = payload if isinstance(payload, list) else (
        gv(payload, "items") or gv(payload, "results") or gv(payload, "analyses") or []
    )
    if not isinstance(items, list) or not items:
        return None, status, None

    for item in items:
        if not isinstance(item, dict):
            continue
        found_app_id = text(item, "appAnalysisId")
        if found_app_id == str(app_analysis_id):
            resolved = text(item, "analysisId")
            if resolved:
                return resolved, status, item

    if len(items) == 1 and isinstance(items[0], dict):
        resolved = text(items[0], "analysisId")
        if resolved:
            return resolved, status, items[0]

    return None, status, None


def first_stats_item(payload):
    if isinstance(payload, list):
        return payload[0] if payload and isinstance(payload[0], dict) else None
    if isinstance(payload, dict):
        for key in ("items", "results", "statistics", "stats"):
            value = gv(payload, key)
            if isinstance(value, list) and value and isinstance(value[0], dict):
                return value[0]
        if gv(payload, "purePremium") is not None or gv(payload, "cv") is not None:
            return payload
    return None


def fetch_analysis(app_analysis_id, perspectives, api):
    """Resolve ID, pull metadata, then /stats for each requested perspective."""
    analysis_id, resolve_status, list_item = resolve_analysis_id(app_analysis_id, api)
    if not analysis_id:
        detail = ("no HTTP response (connection dropped / timed out)" if resolve_status == 0
                   else f"HTTP {resolve_status}" if resolve_status != 200
                   else "no match returned")
        return [[
            app_analysis_id, "", "", "", "",
            "", "", "", "", perspective, "", "", "",
            "", "", f"ERROR - could not resolve appAnalysisId {app_analysis_id} to an analysisId ({detail})",
        ] for perspective in perspectives]

    fallback_portfolio_name = text(list_item, "exposureResourceName") or text(list_item, "exposureName")
    fallback_lob = text(list_item, "lobName")

    metadata, status = api.get(f"{BASE}/analyses/{analysis_id}")
    if not isinstance(metadata, dict) or status != 200:
        detail = "no HTTP response (connection dropped / timed out)" if status == 0 else f"HTTP {status}"
        return [[
            app_analysis_id, analysis_id, "", fallback_portfolio_name, fallback_lob,
            "", "", "", "", perspective, "", "", "",
            "", "", f"ERROR - analysis metadata {detail}",
        ] for perspective in perspectives]

    analysis_name = text(metadata, "analysisName")
    resolved_app_id = text(metadata, "appAnalysisId") or app_analysis_id
    portfolio_name = (text(list_item, "exposureResourceName") or text(list_item, "exposureName")
                      or text(metadata, "exposureResourceName"))
    lob = text(list_item, "lobName") or text(metadata, "lobName")
    create_date = text(metadata, "createDate")
    analysis_type = text(metadata, "analysisType")
    peril = text(metadata, "peril")
    subperil = text(metadata, "subPeril") or text(metadata, "subperil")
    exposure_id = text(metadata, "exposureResourceId")
    exposure_type = text(metadata, "exposureResourceType")
    currency_node = gv(metadata, "currency")
    currency = text(currency_node, "currencyCode") if isinstance(currency_node, dict) else ""
    analysis_status = text(metadata, "analysisStatus")

    rows = []
    for perspective in perspectives:
        if not exposure_id or not exposure_type:
            rows.append([
                resolved_app_id, analysis_id, analysis_name, portfolio_name, lob, create_date,
                analysis_type, peril, subperil, perspective, exposure_id, exposure_type,
                currency, "", "", "ERROR - missing exposureResourceId or exposureResourceType",
            ])
            continue

        query = urlencode({
            "perspectiveCode": perspective,
            "exposureResourceId": exposure_id,
            "exposureResourceType": exposure_type,
        })
        payload, stats_status = api.get(f"{BASE}/analyses/{analysis_id}/stats?{query}")
        item = first_stats_item(payload)
        if item is None:
            message = f"NO STATS - HTTP {stats_status}"
            if analysis_status and analysis_status.lower() != "analyzed":
                message += f"; analysisStatus={analysis_status}"
            total_aal = cv = ""
        else:
            total_aal = numeric_or_blank(gv(item, "purePremium"))
            cv = numeric_or_blank(gv(item, "cv"))
            message = "OK"

        rows.append([
            resolved_app_id, analysis_id, analysis_name, portfolio_name, lob, create_date,
            analysis_type, peril, subperil, perspective, exposure_id, exposure_type,
            currency, total_aal, cv, message,
        ])
    return rows


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="RMS Analysis Total AAL", page_icon="📊", layout="wide")

st.title("📊 RMS Analysis Total AAL Lookup")
st.write(
    "Enter one or more **App Analysis ID**s (the ID shown in the Risk Modeler / "
    "UnderwriteIQ UI) and choose a **Perspective**, then run the lookup. "
    "Results include Total AAL (purePremium) and CV, plus Portfolio Name and LOB."
)

with st.sidebar:
    st.header("Settings")
    workers = st.slider("Parallel requests (workers)", min_value=1, max_value=16, value=8,
                         help="Lower this if you see 'HTTP 0' / dropped-connection errors.")
    verify_ssl = st.checkbox("Verify SSL", value=True,
                              help="Uncheck only if you hit corporate SSL/proxy errors.")
    st.caption("API key is read from the RMS_API_KEY_RISKDATA environment variable.")

col1, col2 = st.columns([2, 1])
with col1:
    ids_raw = st.text_area(
        "App Analysis ID(s)",
        value="",
        height=120,
        placeholder="e.g. 75857 75820 75818\nor one per line, or comma separated",
    )
with col2:
    perspectives = st.multiselect(
        "Perspective(s)",
        options=PERSPECTIVES,
        default=[DEFAULT_PERSPECTIVE],
        help="GU=Ground Up, GR=Gross, RL=Net Pre Cat, RP=Net Post Cat",
    )

run = st.button("Run lookup", type="primary")

if "results_df" not in st.session_state:
    st.session_state["results_df"] = None

if run:
    ids = parse_analysis_ids(ids_raw)
    if not ids:
        st.warning("Please enter at least one App Analysis ID.")
    elif not perspectives:
        st.warning("Please select at least one Perspective.")
    else:
        api = Api(API_KEY, verify=verify_ssl, pool_size=max(8, workers * 2))
        progress = st.progress(0.0, text=f"Fetching {len(ids)} analysis ID(s)...")
        results = []
        started = time.time()

        with ThreadPoolExecutor(max_workers=max(1, min(workers, len(ids)))) as executor:
            futures = {executor.submit(fetch_analysis, aid, perspectives, api): aid for aid in ids}
            done = 0
            for future in futures:
                results.extend(future.result())
                done += 1
                progress.progress(done / len(ids), text=f"Fetched {done}/{len(ids)} analysis ID(s)...")

        progress.empty()
        df = pd.DataFrame(results, columns=HEADERS)
        st.session_state["results_df"] = df

        elapsed = time.time() - started
        ok_count = (df["Status"] == "OK").sum()
        st.success(f"Done in {elapsed:.1f}s. {ok_count}/{len(df)} statistic row(s) returned OK.")

if st.session_state["results_df"] is not None:
    df = st.session_state["results_df"]
    st.subheader("Results")
    st.dataframe(df, use_container_width=True)

    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download as CSV",
        data=csv_bytes,
        file_name="rms_analysis_total_aal.csv",
        mime="text/csv",
    )
