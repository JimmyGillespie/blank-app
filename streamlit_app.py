#!/usr/bin/env python
"""
RMS Location Intelligence Explorer (Streamlit)

Enter an address, see it plotted on a map, and pick which RMS Location
Intelligence hazard layers you want enriched for that location.

Request shapes in this file are ported directly from a CONFIRMED-WORKING
Excel VBA macro (modFloodGeo) that successfully calls the same API:

  Host + app path : https://{host}/li/   (host = api-euw1.rms.com or api-use1.rms.com)
  Auth header     : Authorization: {api_key}   <- RAW key, no "Bearer " prefix

  1) Geocode  (POST {host}/li/geocode/{version})
     Body is the address fields FLAT at the top level (NOT wrapped in a
     "location"/"address" object):
         {"countryScheme":"ISO2A","countryCode":"US","streetAddress":"...",
          "cityName":"...","admin1Code":"...","postalCode":"..."}
     Response includes latitude, longitude, rmsGeoModelResolutionCode.

  2) Hazard lookup (POST {host}/li/{layer}/{version}, e.g. us_fema_firm,
     us_fl_depth, us_fl_susceptibility)
     Body nests the geocoded coordinates under location.address, with
     sibling "characteristics" / "coverageValues" objects, and an optional
     "layerOptions" object as a SIBLING of "location" (not nested inside it):
         {
           "location": {
             "address": {"countryScheme":"ISO2A","countryCode":"US",
                          "countryRmsCode":"US","admin1Code":"...",
                          "latitude":..,"longitude":..,
                          "rmsGeoModelResolutionCode":1},
             "characteristics": {"construction":"ATC0","occupancy":"ATC0",
                 "yearBuilt":0,"numOfStories":0,"foundationType":0,
                 "firstFloorHeightAboveGround":0,"basement":"DEFAULT","floorArea":0},
             "coverageValues": {"buildingValue":0,"contentsValue":0,
                 "businessInterruptionValue":0}
           },
           "layerOptions": {"distanceUnit":"feet","proximityDistance":100}
         }

Only the three flood layers above (FEMA FIRM, Flood Depth, Elevation/
Susceptibility) are confirmed working. Earthquake / windstorm / wildfire
layer names follow RMS's documented <region>_<peril>_hazard convention but
are NOT independently confirmed - they are included as editable best-effort
rows; verify/adjust the layer name if one returns ERR_4040 (not found).
"""

import os
import time
import json
import threading

import pandas as pd
import requests
import streamlit as st
from requests.adapters import HTTPAdapter

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------

DEFAULT_HOST = os.environ.get("RMS_LI_HOST", "api-euw1.rms.com")
API_KEY = os.environ.get("RMS_LI_API_KEY", "")
TIMEOUT = 60
MAX_RETRIES = 4

ERROR_CODES = {
    "ERR_4000": "No details found for requested layer.",
    "ERR_4001": "Insufficient geocoding information - a required field is missing.",
    "ERR_4002": "Insufficient geocoding resolution to look up requested data.",
    "ERR_4003": "Location is in a region not supported by this layer.",
    "ERR_4005": "Insufficient information to look up requested data.",
    "ERR_4041": "Invalid input.",
    "ERR_4100": "Invalid lookupType.",
    "ERR_4101": "Invalid distanceUnit.",
    "ERR_4102": "Invalid returnPeriod.",
    "ERR_4103": "Invalid defended option.",
    "ERR_4104": "Invalid proximityDistance.",
    "ERR_4105": "Invalid distanceUnit.",
    "ERR_4106": "Missing distanceUnit and/or proximityDistance.",
    "ERR_4300": "Invalid basement option.",
    "ERR_4400": "General error for the layer.",
    "ERR_4401": "Invalid address input.",
    "ERR_4010": "Forbidden.",
    "ERR_4030": "Unauthorized.",
    "ERR_4040": "Not found - check the layer name/version.",
    "ERR_4050": "Method not allowed.",
    "ERR_4060": "Not acceptable.",
    "ERR_4080": "Request timeout.",
    "ERR_4130": "Request too large.",
    "ERR_4150": "Unsupported media type.",
    "ERR_4220": "Unprocessable entity.",
    "ERR_4290": "Quota exceeded.",
    "ERR_5040": "Gateway timeout.",
}

# Confirmed-working flood layers (ported 1:1 from the VBA macro), plus
# best-effort rows for other perils using RMS's documented naming
# convention. Edit any Layer/Version/LayerOptions before running if your
# tenant differs.
DEFAULT_LAYERS = [
    {"Hazard": "US FEMA FIRM (flood zone)", "Layer": "us_fema_firm", "Version": "latest",
     "BodyStyle": "flood_basic",
     "LayerOptions": '{"distanceUnit":"feet","proximityDistance":100}', "Confirmed": "Yes", "Selected": True},
    {"Hazard": "US Flood Depth (undefended)", "Layer": "us_fl_depth", "Version": "latest",
     "BodyStyle": "flood_basic",
     "LayerOptions": '{"lookup":"floodDepth","proximityDistance":100,"distanceUnit":"feet"}',
     "Confirmed": "Yes", "Selected": True},
    {"Hazard": "Elevation / Susceptibility", "Layer": "us_fl_susceptibility", "Version": "latest",
     "BodyStyle": "full", "LayerOptions": "", "Confirmed": "Yes", "Selected": True},
    {"Hazard": "US Flood Risk Score", "Layer": "us_fl_risk_score", "Version": "latest",
     "BodyStyle": "full", "LayerOptions": "", "Confirmed": "Yes", "Selected": True},
    {"Hazard": "Earthquake (soil/liquefaction/landslide)", "Layer": "na_eq_hazard", "Version": "latest",
     "BodyStyle": "minimal", "LayerOptions": "", "Confirmed": "Yes", "Selected": True},
    {"Hazard": "Earthquake Intensity (MMI)", "Layer": "us_mmi", "Version": "latest",
     "BodyStyle": "minimal", "LayerOptions": "", "Confirmed": "Yes", "Selected": True},
    {"Hazard": "Windstorm (name unconfirmed)", "Layer": "us_hu_hazard", "Version": "latest",
     "BodyStyle": "minimal", "LayerOptions": "", "Confirmed": "No - edit before use", "Selected": False},
    {"Hazard": "Wildfire (name unconfirmed)", "Layer": "us_wf_hazard", "Version": "latest",
     "BodyStyle": "minimal", "LayerOptions": "", "Confirmed": "No - edit before use", "Selected": False},
]


# --------------------------------------------------------------------------
# API client (mirrors the confirmed-working VBA request shapes exactly)
# --------------------------------------------------------------------------

class LocationIntelligenceApi:
    """Thin, resilient wrapper around the RMS Location Intelligence API."""

    def __init__(self, host, api_key, verify=True, pool_size=16):
        self.base_url = f"https://{host}/li"
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        clean_key = (api_key or "").strip()
        if clean_key.lower().startswith("bearer "):
            clean_key = clean_key[7:].strip()
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Authorization": clean_key,  # RAW key - no "Bearer " prefix (confirmed via working VBA macro)
        })
        self.session.verify = verify
        self.lock = threading.Lock()
        self.ssl_fallback_done = not verify

    def _disable_ssl_verify(self):
        with self.lock:
            if self.ssl_fallback_done:
                return
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            self.session.verify = False
            self.ssl_fallback_done = True

    def _post(self, path, body, retries=MAX_RETRIES):
        url = f"{self.base_url}/{path}"
        last_exc = None
        for attempt in range(retries):
            try:
                response = self.session.post(url, json=body, timeout=TIMEOUT)
                if response.status_code == 200:
                    try:
                        return response.json(), response.status_code, None
                    except ValueError:
                        return None, response.status_code, "Response was not valid JSON."
                if response.status_code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                try:
                    detail = response.json()
                except ValueError:
                    detail = response.text[:500]
                return None, response.status_code, detail
            except requests.exceptions.SSLError:
                if self.session.verify:
                    self._disable_ssl_verify()
                    continue
                return None, 0, "SSL error persisted with verification disabled."
            except requests.RequestException as exc:
                last_exc = str(exc)
                if attempt < retries - 1:
                    time.sleep(min(2 ** (attempt + 1), 15))
                    continue
                return None, 0, last_exc
        return None, 0, last_exc

    def geocode(self, street, city, admin1, postal, country_code, version="latest"):
        """Flat address body - matches FG_Step1_Geocode in the VBA macro exactly."""
        country_code = (country_code or "US").upper()
        if country_code == "USA":
            country_code = "US"
        body = {"countryScheme": "ISO2A", "countryCode": country_code}
        if street:
            body["streetAddress"] = street
        if city:
            body["cityName"] = city
        if admin1:
            body["admin1Code"] = admin1
        if postal:
            body["postalCode"] = postal
        return self._post(f"geocode/{version}", body)

    def hazard_lookup(self, lat, lon, admin1_code, country_code, resolution_code,
                       layer, version, body_style="full", layer_options=None,
                       characteristics=None, coverage_values=None):
        """Builds the request body in one of three CONFIRMED-WORKING shapes,
        ported directly from two working RMS VBA macros (modFloodGeo and
        modLocationsHazard):

          "minimal"     -> {"location":{"address":{...}}}
                           Used for simple hazard layers that need only
                           coordinates: na_eq_hazard (earthquake soil/
                           liquefaction/landslide), us_mmi (earthquake
                           intensity).

          "flood_basic" -> {"location":{"address":{...},"characteristics":{},
                             "coverageValues":{}}, "layerOptions":{...}}
                           Used for us_fema_firm (flood zone) and us_fl_depth
                           (flood depth) - empty characteristics/coverage
                           objects are still required, plus layerOptions as
                           a SIBLING of "location" (not nested inside it).

          "full"        -> {"location":{"address":{...},
                             "characteristics":{full building attrs},
                             "coverageValues":{full coverage amounts}}}
                           Used for us_fl_susceptibility (elevation) and
                           us_fl_risk_score (needs real building attributes
                           to score accurately).
        """
        address = {
            "countryScheme": "ISO2A",
            "countryCode": country_code,
            "countryRmsCode": country_code,
        }
        if admin1_code:
            address["admin1Code"] = admin1_code
        address["latitude"] = lat
        address["longitude"] = lon
        address["rmsGeoModelResolutionCode"] = resolution_code

        location = {"address": address}

        if body_style == "minimal":
            pass  # location.address only - no characteristics/coverageValues at all
        elif body_style == "flood_basic":
            location["characteristics"] = {}
            location["coverageValues"] = {}
        else:  # "full"
            location["characteristics"] = characteristics or {
                "construction": "ATC0", "occupancy": "ATC0", "yearBuilt": 0,
                "numOfStories": 0, "foundationType": 0,
                "firstFloorHeightAboveGround": 0, "basement": "DEFAULT", "floorArea": 0,
            }
            location["coverageValues"] = coverage_values or {
                "buildingValue": 0, "contentsValue": 0, "businessInterruptionValue": 0,
            }

        body = {"location": location}
        if body_style == "flood_basic" and layer_options:
            body["layerOptions"] = layer_options  # sibling of "location", not nested inside it

        return self._post(f"{layer}/{version}", body)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def flatten_for_table(layer_name, results):
    if not isinstance(results, dict):
        return []
    rows = []
    for key, value in results.items():
        if isinstance(value, (dict, list)):
            value = json.dumps(value)
        rows.append({"Layer": layer_name, "Field": key, "Value": value})
    return rows


def format_latlon_caption(lat, lon):
    lat_str = "{:.6f}".format(lat)
    lon_str = "{:.6f}".format(lon)
    return "Resolved coordinates: " + lat_str + ", " + lon_str


def parse_layer_options(raw):
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return None


def resolve_admin1_code(geo_payload, fallback_admin1):
    """Hazard layers require a strict 2-letter admin1Code. Prefer the value
    RMS itself resolved during geocoding (guaranteed valid) over whatever
    free-text the user typed in the address form - this mirrors the
    confirmed-working VBA macro's 'only send if exactly 2 chars' guard,
    but is more robust because it trusts RMS's own resolution first."""
    resolved = ""
    if isinstance(geo_payload, dict):
        resolved = str(geo_payload.get("admin1Code") or "").strip().upper()
    if len(resolved) == 2:
        return resolved
    typed = str(fallback_admin1 or "").strip().upper()
    if len(typed) == 2:
        return typed
    return ""


STATUS_ICON = {"OK": "✅", "ERROR": "⚠️"}

# Confirmed mapping for rmsGeoModelResolutionCode, sourced directly from a
# working RMS VBA macro's ResDesc() function (comment: "Source:
# RMS_EDM.Lookup.GeoResolutionCode"). Lower code = higher precision.
MODEL_RESOLUTION_LABELS = {
    0: "None", 1: "Coordinate", 2: "Parcel", 3: "Street Block", 4: "Street Name",
    5: "Postcode", 6: "City District", 7: "City", 8: "District", 9: "County",
    10: "State", 11: "Accumulation Zone", 12: "Multiple", 13: "User Defined", 14: "Country",
}
# Precision tier used for the underwriting quality badge - parcel/coordinate
# level is as good as it gets for siting a risk; postcode-and-up means the
# hazard results below are increasingly an area estimate, not a site-specific read.
MODEL_RESOLUTION_TIER = {
    0: ("⚠️", "Unresolved"), 1: ("🎯", "Rooftop / coordinate-level"), 2: ("🎯", "Parcel-level"),
    3: ("✅", "Street block-level"), 4: ("✅", "Street-level"), 5: ("🟡", "Postcode-level"),
    6: ("🟡", "City district-level"), 7: ("🟡", "City-level"), 8: ("🟠", "District-level"),
    9: ("🟠", "County-level"), 10: ("🔴", "State-level"), 11: ("🔴", "Accumulation zone-level"),
    12: ("🔴", "Multiple/ambiguous match"), 13: ("🟡", "User-defined"), 14: ("🔴", "Country-level only"),
}

# Standard, publicly documented FEMA/NFIP flood zone classifications.
# https://www.fema.gov/glossary/flood-zones
FEMA_ZONE_INFO = {
    "A": ("High risk (SFHA)", "1% annual chance flood (100-yr); flood insurance federally required if mortgaged."),
    "AE": ("High risk (SFHA)", "1% annual chance flood (100-yr) with defined base flood elevations; flood insurance federally required if mortgaged."),
    "AH": ("High risk (SFHA)", "1% annual chance shallow flooding (ponding); flood insurance federally required if mortgaged."),
    "AO": ("High risk (SFHA)", "1% annual chance shallow flooding (sheet flow); flood insurance federally required if mortgaged."),
    "AR": ("High risk (SFHA)", "Temporarily elevated risk behind a flood-control system under restoration; flood insurance federally required if mortgaged."),
    "A99": ("High risk (SFHA)", "1% annual chance flood, protected by a federal flood-control system under construction; flood insurance federally required if mortgaged."),
    "V": ("High risk coastal (SFHA)", "1% annual chance coastal flood with wave hazard; flood insurance federally required if mortgaged."),
    "VE": ("High risk coastal (SFHA)", "1% annual chance coastal flood with high-velocity wave action; flood insurance federally required if mortgaged."),
    "B": ("Moderate risk", "Between 0.2%-1% annual chance flood (500-yr); outside the mandatory purchase zone but insurance recommended."),
    "X": ("Minimal to moderate risk", "Outside the 1% floodplain; exact tier depends on the annual probability value shown below."),
    "C": ("Minimal risk", "Outside the 500-yr floodplain; flood insurance not federally required."),
    "D": ("Undetermined", "Flood hazard has not been formally studied for this area."),
}


def classify_fema_zone(zone, annual_probability):
    """Combine the raw FEMA zone letter with the annualProbability field
    (which RMS already resolves precisely) to give an accurate risk tier -
    this avoids guessing 'shaded vs unshaded X' from the letter alone."""
    zone = (zone or "").strip().upper()
    info = FEMA_ZONE_INFO.get(zone)
    label, detail = info if info else ("Unknown zone", "Zone code not recognized - verify manually.")
    prob = (annual_probability or "").strip()
    if zone == "X" and prob:
        if "0.2" in prob or "less than" in prob.lower():
            label, detail = "Minimal risk", f"Outside the 1% floodplain, annual chance {prob}; flood insurance not federally required."
        else:
            label, detail = "Moderate risk", f"Zone X with elevated annual chance {prob}; consider flood insurance."
    return label, detail


# Modified Mercalli Intensity scale labels, confirmed from a working RMS
# VBA macro's mLH_MmiDict (built from standard USGS MMI descriptions).
MMI_LABELS = {
    1: "Not Felt", 2: "Weak", 3: "Weak", 4: "Light", 5: "Moderate", 6: "Strong",
    7: "Very Strong", 8: "Severe", 9: "Violent", 10: "Extreme", 11: "Extreme", 12: "Extreme",
}


def mmi_label(value):
    try:
        v = round(float(value))
    except (TypeError, ValueError):
        return "Unknown"
    v = max(1, min(12, v))
    return MMI_LABELS.get(v, "Unknown")


def soil_risk_badge(description):
    """Qualitative descriptions already come back in plain English from
    na_eq_hazard; this just adds a consistent risk-colored badge."""
    d = (description or "").lower()
    if not d or d in ("n/a", "unknown"):
        return "⚪"
    if "very low" in d or d == "low":
        return "🟢"
    if "low to moderate" in d or "moderate" in d:
        return "🟡"
    if "high" in d or "severe" in d:
        return "🔴"
    return "🟢" if "stiff" in d or "rock" in d else "🟡"


def get_cresta_zone(geo_data):
    """CRESTA (Catastrophe Risk Evaluation and Standardizing Target
    Accumulations) zone - the insurance industry's standard geographic
    accumulation-management unit for earthquake/windstorm exposure.
    RMS returns this as zone1Code in the geocode response (confirmed via
    the working VBA macro's LH_GetCresta / SubCresta logic, which reads
    the same zone1Code field)."""
    code = (geo_data or {}).get("zone1Code") or ""
    name = (geo_data or {}).get("zone1Name") or ""
    return str(code).strip(), str(name).strip()


def mmi_gauge_html(mmi_value, label):
    """Renders a 1-12 MMI scale as a horizontal colour gradient with a
    marker at the current value - green (not felt) through red (extreme)."""
    try:
        v = float(mmi_value)
    except (TypeError, ValueError):
        return ""
    pct = max(0.0, min(100.0, (v - 1) / 11 * 100))
    return f"""
    <div class="mmi-gauge-wrap">
        <div class="mmi-gauge-track">
            <div class="mmi-gauge-marker" style="left: calc({pct}% - 2px);"></div>
        </div>
        <div class="mmi-scale-labels">
            <span>I · Not felt</span><span>VI · Strong</span><span>XII · Extreme</span>
        </div>
    </div>
    """


def get_layer_results(results, layer_name):
    for entry in results or []:
        if entry.get("name") == layer_name:
            return entry.get("results") or {}, entry.get("errors") or []
    return {}, []


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="RMS Location Intelligence Explorer", page_icon="🌍", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; max-width: 1200px; }
    div[data-testid="stExpander"] { border: 1px solid rgba(135,135,135,0.25); border-radius: 10px; }

    .risk-hero {
        border-radius: 16px;
        padding: 1.6rem 1.8rem;
        margin-bottom: 1rem;
        background: linear-gradient(135deg, rgba(178,34,52,0.12) 0%, rgba(178,34,52,0.03) 100%);
        border: 1px solid rgba(178,34,52,0.35);
        border-left: 6px solid #b22234;
    }
    .risk-hero.flood {
        background: linear-gradient(135deg, rgba(31,119,180,0.10) 0%, rgba(31,119,180,0.03) 100%);
        border: 1px solid rgba(31,119,180,0.30);
        border-left: 6px solid #1f77b4;
        padding: 1.2rem 1.6rem;
    }
    .section-eyebrow {
        font-size: 0.78rem; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase;
        opacity: 0.65; margin-bottom: 0.2rem;
    }
    .section-title { font-size: 1.5rem; font-weight: 800; margin: 0 0 0.6rem 0; }
    .section-title.flood { font-size: 1.15rem; margin-bottom: 0.4rem; }

    .cresta-chip {
        display: inline-flex; align-items: center; gap: 0.4rem;
        background: rgba(178,34,52,0.15); border: 1px solid rgba(178,34,52,0.4);
        border-radius: 999px; padding: 0.35rem 0.9rem; font-weight: 700; font-size: 0.95rem;
        margin-bottom: 0.9rem;
    }
    .mini-card {
        background: rgba(135,135,135,0.07); border-radius: 12px; padding: 0.9rem 1.1rem;
        height: 100%; border: 1px solid rgba(135,135,135,0.15);
    }
    .mini-card .label { font-size: 0.75rem; opacity: 0.65; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600;}
    .mini-card .value { font-size: 1.05rem; font-weight: 700; margin-top: 0.15rem; }

    .mmi-gauge-wrap { margin: 0.8rem 0 0.3rem 0; }
    .mmi-gauge-track {
        position: relative; height: 14px; border-radius: 999px; margin-top: 0.5rem;
        background: linear-gradient(90deg, #2e7d32 0%, #9e9d24 25%, #f9a825 50%, #e65100 75%, #b71c1c 100%);
    }
    .mmi-gauge-marker {
        position: absolute; top: -7px; width: 4px; height: 28px; background: #ffffff;
        border-radius: 2px; box-shadow: 0 0 0 2px rgba(0,0,0,0.55);
    }
    .mmi-scale-labels { display: flex; justify-content: space-between; font-size: 0.68rem; opacity: 0.55; margin-top: 0.25rem; }

    .flood-strip { display: flex; flex-wrap: wrap; gap: 1.6rem; align-items: center; }
    .flood-metric .label { font-size: 0.72rem; opacity: 0.6; text-transform: uppercase; letter-spacing: 0.04em; }
    .flood-metric .value { font-size: 1rem; font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🌍 RMS Location Intelligence Explorer")
st.caption(
    "Enter an address, see it on a map, and choose which RMS hazard layers to look up for that "
    "location. Request formats are ported from a confirmed-working RMS VBA macro."
)

if "li_results" not in st.session_state:
    st.session_state["li_results"] = None
if "li_latlon" not in st.session_state:
    st.session_state["li_latlon"] = None
if "li_resolution" not in st.session_state:
    st.session_state["li_resolution"] = None

with st.sidebar:
    st.header("Connection")
    host = st.text_input("Host", value=DEFAULT_HOST, help="api-euw1.rms.com or api-use1.rms.com")
    api_key = st.text_input("API key", value=API_KEY, type="password",
                             help="Raw API key, sent as-is in the Authorization header (no 'Bearer ' prefix).")
    verify_ssl = st.checkbox("Verify SSL", value=True,
                              help="Uncheck only if you hit corporate SSL/proxy errors.")
    region_code = st.text_input("Region code (for unconfirmed layers)", value="us")
    version_default = st.text_input("Default layer version", value="latest")

    st.divider()
    test_conn = st.button("🔌 Test connection (geocode only)")
    st.caption(
        "If you still get HTTP 401 after this fix: confirm the pasted API key is EXACTLY the one "
        "working in your RMS workbook macro, with no extra spaces or truncation, and that this key "
        "is licensed for Location Intelligence specifically (not just Risk Modeler / Risk Data API)."
    )

st.subheader("1. Address")
col1, col2, col3, col4 = st.columns([3, 2, 1, 1])
with col1:
    street = st.text_input("Street address", placeholder="766 Ocean Ave")
with col2:
    city = st.text_input("City", placeholder="Sea Bright")
with col3:
    admin1 = st.text_input("State / Admin1", placeholder="NJ")
with col4:
    postal = st.text_input("Postal code", placeholder="07760")
country_code = st.text_input("Country code (ISO2)", value="US", max_chars=2)

with st.expander("🏗️ Building characteristics (optional - improves accuracy for Elevation & Flood Risk Score)"):
    bc1, bc2, bc3, bc4 = st.columns(4)
    with bc1:
        bld_construction = st.text_input("Construction code", value="ATC0")
        bld_occupancy = st.text_input("Occupancy code", value="ATC0")
    with bc2:
        bld_year_built = st.number_input("Year built", value=0, step=1)
        bld_stories = st.number_input("# of stories", value=0, step=1)
    with bc3:
        bld_value = st.number_input("Building value", value=0, step=1000)
        bld_contents_value = st.number_input("Contents value", value=0, step=1000)
    with bc4:
        bld_bi_value = st.number_input("BI value", value=0, step=1000)
    st.caption("Defaults (ATC0 / 0) match the proven macro's fallback and will still return results - "
               "fill these in only if you want building-specific risk scoring.")

if test_conn:
    if not api_key:
        st.warning("Please provide an API key (sidebar) before testing the connection.")
    else:
        test_api = LocationIntelligenceApi(host.strip(), api_key, verify=verify_ssl)
        with st.spinner("Sending a minimal test geocode request..."):
            test_payload, test_status, test_error = test_api.geocode(
                "7575 Gateway Blvd", "Newark", "CA", "94560", "US", version_default
            )
        masked = ("*" * max(len(api_key.strip()) - 4, 0)) + api_key.strip()[-4:]
        st.write(f"**Request URL:** `https://{host.strip()}/li/geocode/{version_default}`")
        st.write(f"**Authorization header sent:** `{masked}`")
        st.write(f"**HTTP status:** `{test_status}`")
        if test_status == 200:
            st.success("✅ Connection succeeded - this key and host are authorized for Location Intelligence.")
            st.json(test_payload)
        elif test_status == 401:
            st.error(
                "❌ 401 Unauthorized. Double-check this is the exact same key used successfully in "
                "your working RMS Excel macro (copy it again from the workbook's RMS_TOKEN constant "
                "or named range RMS_API_KEY, with no extra characters), and that 'host' matches "
                "exactly ('api-euw1.rms.com', no https://, no trailing slash)."
            )
            st.json(test_error)
        else:
            st.error(f"Request failed with HTTP {test_status}.")
            st.json(test_error)

st.subheader("2. Hazard layers to look up")
st.caption(
    "✅ Confirmed layers (FEMA FIRM, Flood Depth, Elevation/Susceptibility, Flood Risk Score, "
    "Earthquake, MMI) are ported directly from two working RMS workbook macros - these just work. "
    "⚠️ Windstorm/Wildfire layer names are **unconfirmed guesses** - the US product line does not "
    "follow the generic `<region>_<peril>_hazard` pattern used for other regions, so these may return "
    "'No details found for layer name...' until corrected. Find the exact name in your tenant's "
    "Postman collection / OpenAPI spec (Developer Assets on developer.rms.com), then edit the "
    "'Layer name' cell below."
)

catalog_df = pd.DataFrame(DEFAULT_LAYERS)
catalog_df["Layer"] = catalog_df["Layer"].apply(lambda s: s.format(region=region_code.strip().lower() or "us"))
edited_df = st.data_editor(
    catalog_df,
    column_config={
        "Selected": st.column_config.CheckboxColumn("Include"),
        "Hazard": st.column_config.TextColumn("Hazard", disabled=True),
        "Layer": st.column_config.TextColumn("Layer name"),
        "Version": st.column_config.TextColumn("Version"),
        "BodyStyle": st.column_config.TextColumn("Body shape", disabled=True,
                                                  help="minimal / flood_basic / full - controls the request body shape sent to RMS."),
        "LayerOptions": st.column_config.TextColumn("Layer options (JSON, optional)"),
        "Confirmed": st.column_config.TextColumn("Confirmed?", disabled=True),
    },
    hide_index=True,
    width="stretch",
    num_rows="dynamic",
    key="layer_editor",
)

run = st.button("Geocode & look up hazards", type="primary")

if run:
    if not api_key:
        st.warning("Please provide an API key (sidebar) before running a lookup.")
    elif not (street or (city and admin1) or postal):
        st.warning("Please enter at least a street address, or a city + state, or a postal code.")
    else:
        api = LocationIntelligenceApi(host.strip(), api_key, verify=verify_ssl)

        with st.spinner("Geocoding address..."):
            geo_payload, geo_status, geo_error = api.geocode(
                street, city, admin1, postal, country_code, version_default
            )

        if geo_status != 200 or geo_payload is None:
            st.error(f"Geocoding failed (HTTP {geo_status}): {geo_error}")
            st.session_state["li_results"] = None
            st.session_state["li_latlon"] = None
        else:
            lat = geo_payload.get("latitude")
            lon = geo_payload.get("longitude")
            resolution_code = geo_payload.get("rmsGeoModelResolutionCode") or 1
            resolved_admin1 = resolve_admin1_code(geo_payload, admin1)
            st.session_state["li_latlon"] = (lat, lon)
            st.session_state["li_resolution"] = resolution_code

            results = [{"name": "geocode", "version": version_default, "results": geo_payload, "errors": []}]

            if lat and lon:
                selected_layers = [
                    row for _, row in edited_df.iterrows()
                    if row.get("Selected") and str(row.get("Layer", "")).strip()
                ]
                characteristics = {
                    "construction": bld_construction or "ATC0", "occupancy": bld_occupancy or "ATC0",
                    "yearBuilt": int(bld_year_built), "numOfStories": int(bld_stories),
                    "foundationType": 0, "firstFloorHeightAboveGround": 0,
                    "basement": "DEFAULT", "floorArea": 0,
                }
                coverage_values = {
                    "buildingValue": int(bld_value), "contentsValue": int(bld_contents_value),
                    "businessInterruptionValue": int(bld_bi_value),
                }

                with st.spinner(f"Looking up {len(selected_layers)} hazard layer(s)..."):
                    for row in selected_layers:
                        layer_options = parse_layer_options(row.get("LayerOptions"))
                        body_style = row.get("BodyStyle") or "full"
                        r_payload, r_status, r_error = api.hazard_lookup(
                            lat, lon, resolved_admin1, country_code.strip().upper() or "US",
                            resolution_code, row["Layer"], row["Version"] or version_default,
                            body_style=body_style, layer_options=layer_options,
                            characteristics=characteristics, coverage_values=coverage_values,
                        )
                        if r_status == 200 and r_payload is not None:
                            results.append({"name": row["Layer"], "version": row["Version"],
                                             "results": r_payload, "errors": []})
                        else:
                            # Surface RMS's own error code/message when available, so the
                            # user sees exactly why a layer failed (e.g. "No details found
                            # for layer name X" means the layer name itself is wrong).
                            api_errors = []
                            if isinstance(r_error, dict):
                                api_errors = r_error.get("errors") or []
                            if api_errors:
                                code = api_errors[0].get("code", "ERR_HTTP")
                                message = api_errors[0].get("message") or ERROR_CODES.get(code, "Unknown error")
                                if "no details found for layer name" in message.lower():
                                    message += " — this layer name is likely incorrect for this region/product; verify it in your tenant's layer catalog."
                            else:
                                code = "ERR_4040" if r_status == 404 else "ERR_HTTP"
                                message = f"HTTP {r_status}: {r_error}"
                            results.append({"name": row["Layer"], "version": row["Version"], "results": {},
                                             "errors": [{"code": code, "message": message}]})
            else:
                st.warning("Geocoding succeeded but returned no coordinates - hazard lookups were skipped.")

            st.session_state["li_results"] = results

# --------------------------------------------------------------------------
# Results
# --------------------------------------------------------------------------

results = st.session_state["li_results"]
latlon = st.session_state["li_latlon"]

if results is not None:
    geocode_entry = next((e for e in results if e.get("name") == "geocode"), None)
    geo_data = (geocode_entry or {}).get("results") or {}
    hazard_entries = [e for e in results if e.get("name") != "geocode"]
    ok_count = sum(1 for e in hazard_entries if not e.get("errors"))
    err_count = len(hazard_entries) - ok_count

    st.divider()
    st.subheader("📍 Resolved location")

    if latlon and latlon[0] is not None and latlon[1] is not None:
        left, right = st.columns([3, 2])
        with left:
            map_df = pd.DataFrame({"lat": [latlon[0]], "lon": [latlon[1]]})
            st.map(map_df, zoom=14)
        with right:
            st.markdown(f"**{geo_data.get('streetAddress') or street or '—'}**")
            city_line = ", ".join(filter(None, [
                geo_data.get("cityName") or city,
                geo_data.get("admin1Code") or admin1,
                geo_data.get("postalCode") or postal,
            ]))
            st.markdown(city_line or "—")
            st.caption(format_latlon_caption(latlon[0], latlon[1]))

            model_code = geo_data.get("rmsGeoModelResolutionCode")
            try:
                model_code_int = int(model_code)
            except (TypeError, ValueError):
                model_code_int = None
            icon, tier_label = MODEL_RESOLUTION_TIER.get(model_code_int, ("⚪", "Unknown"))
            model_label = MODEL_RESOLUTION_LABELS.get(model_code_int, "Unknown")
            confidence = geo_data.get("geoResolutionConfidence")

            st.markdown(f"**{icon} Geocode quality: {tier_label}**")
            conf_text = f" · {confidence}% confidence" if confidence not in (None, "") else ""
            st.caption(f"Matched to {model_label.lower()} precision{conf_text}.")
            if model_code_int is not None and model_code_int >= 9:
                st.warning(
                    "This address only resolved to a broad area (county/state level or coarser) - "
                    "the hazard results below are area averages, not a site-specific read. Consider "
                    "confirming the address details before relying on this for risk selection."
                )
    else:
        st.info("No coordinates were returned for this address - check the geocode result below.")

    # ------------------------------------------------------------------
    # Underwriting Risk Summary - translates raw hazard fields into the
    # plain-English, risk-selection-oriented view an underwriter needs,
    # before the full technical detail further down.
    # ------------------------------------------------------------------
    if latlon and latlon[0] is not None and latlon[1] is not None:
        st.divider()
        st.subheader("🧭 Underwriting Risk Summary")

        fema_data, fema_errors = get_layer_results(results, "us_fema_firm")
        depth_data, depth_errors = get_layer_results(results, "us_fl_depth")
        susc_data, susc_errors = get_layer_results(results, "us_fl_susceptibility")
        score_data, score_errors = get_layer_results(results, "us_fl_risk_score")
        eq_data, eq_errors = get_layer_results(results, "na_eq_hazard")
        mmi_data, mmi_errors = get_layer_results(results, "us_mmi")

        cresta_code, cresta_name = get_cresta_zone(geo_data)

        # ================================================================
        # PRIMARY: EARTHQUAKE
        # ================================================================
        st.markdown('<div class="risk-hero">', unsafe_allow_html=True)
        st.markdown('<div class="section-eyebrow">Primary peril</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚡ Earthquake Risk Profile</div>', unsafe_allow_html=True)

        if cresta_code:
            cresta_display = f"CRESTA Zone {cresta_code}" + (f" — {cresta_name.title()}" if cresta_name else "")
            st.markdown(f'<div class="cresta-chip">📍 {cresta_display}</div>', unsafe_allow_html=True)
        else:
            st.caption("CRESTA zone not returned for this location.")

        if eq_data or mmi_data:
            soil_desc = eq_data.get("soilDescription", "") if eq_data else ""
            liq_desc = eq_data.get("liquefactionDescription", "") if eq_data else ""
            land_desc = eq_data.get("landslideDescription", "") if eq_data else ""

            c1, c2, c3 = st.columns(3)
            for col, label, desc in (
                (c1, "Soil Type", soil_desc), (c2, "Liquefaction Risk", liq_desc), (c3, "Landslide Risk", land_desc)
            ):
                with col:
                    st.markdown(
                        f'<div class="mini-card"><div class="label">{label}</div>'
                        f'<div class="value">{soil_risk_badge(desc)} {desc or "—"}</div></div>',
                        unsafe_allow_html=True,
                    )

            if mmi_data:
                mmi_475 = mmi_data.get("mmi475yr") or mmi_data.get("mmi500yr")
                if mmi_475 not in (None, ""):
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown(
                        f"**Expected shaking at the design-level event (475-yr return period):** "
                        f"{mmi_label(mmi_475)}  ·  **MMI {mmi_475}**"
                    )
                    st.markdown(mmi_gauge_html(mmi_475, mmi_label(mmi_475)), unsafe_allow_html=True)
                    st.caption("475-yr return period ≈ 10% chance of exceedance in 50 years — the standard "
                               "US building-code design-level earthquake used for site risk comparison.")
                with st.expander("📊 Shaking intensity at other return periods"):
                    mmi_rows = [{"Return period": k.replace("mmi", "").replace("yr", "-yr"),
                                 "MMI": v, "Description": mmi_label(v)}
                                for k, v in mmi_data.items() if k.startswith("mmi")]
                    if mmi_rows:
                        mmi_df = pd.DataFrame(mmi_rows)
                        try:
                            mmi_df["_sort"] = mmi_df["Return period"].str.replace("-yr", "").astype(int)
                            mmi_df = mmi_df.sort_values("_sort").drop(columns="_sort")
                        except (ValueError, KeyError):
                            pass
                        st.dataframe(mmi_df, hide_index=True, width="stretch")
        else:
            st.caption("No earthquake layers were selected/run for this lookup.")

        eq_warnings = []
        for d in (eq_data, mmi_data):
            w = d.get("warnings") if isinstance(d, dict) else None
            if w:
                eq_warnings.append(w)
        if eq_warnings:
            with st.expander("⚠️ Data quality notes (earthquake)"):
                for w in eq_warnings:
                    st.code(w if isinstance(w, str) else json.dumps(w, indent=2), language="json")

        st.markdown("</div>", unsafe_allow_html=True)

        # ================================================================
        # SECONDARY: FLOOD
        # ================================================================
        st.markdown('<div class="risk-hero flood">', unsafe_allow_html=True)
        st.markdown('<div class="section-eyebrow">Secondary peril</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title flood">🌊 Flood Risk</div>', unsafe_allow_html=True)

        if any([fema_data, depth_data, susc_data, score_data]):
            zone = fema_data.get("floodZone", "") if fema_data else ""
            prob = fema_data.get("annualProbability", "") if fema_data else ""
            label, detail = classify_fema_zone(zone, prob) if fema_data else ("—", "")
            sfha = str(fema_data.get("sfha", "")).strip().lower() if fema_data else ""
            badge = "🔴" if "high" in label.lower() else ("🟡" if "moderate" in label.lower() else "🟢")

            u100 = depth_data.get("depthInlandUndef100yr") if depth_data else None
            u500 = depth_data.get("depthInlandUndef500yr") if depth_data else None
            elevation = susc_data.get("dtmElevation") if susc_data else None
            overall_score = score_data.get("scoreOverall") if score_data else None

            if fema_data:
                st.markdown(f"**{badge} {label}** — FEMA Zone **{zone or '—'}**  ·  {detail}")

            st.markdown('<div class="flood-strip">', unsafe_allow_html=True)
            flood_metrics = [
                ("SFHA", "Yes" if sfha == "yes" else "No" if sfha == "no" else "—"),
                ("Annual chance", prob or "—"),
                ("100-yr depth", f"{u100} ft" if u100 not in (None, "") else "—"),
                ("500-yr depth", f"{u500} ft" if u500 not in (None, "") else "—"),
                ("Elevation", f"{elevation} ft" if elevation not in (None, "") else "—"),
                ("RMS Flood Score", overall_score if overall_score not in (None, "") else "—"),
            ]
            for m_label, m_value in flood_metrics:
                st.markdown(
                    f'<div class="flood-metric"><div class="label">{m_label}</div>'
                    f'<div class="value">{m_value}</div></div>',
                    unsafe_allow_html=True,
                )
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.caption("No flood layers were selected/run for this lookup.")

        flood_warnings = []
        for d in (fema_data, depth_data, susc_data, score_data):
            w = d.get("warnings") if isinstance(d, dict) else None
            if w:
                flood_warnings.append(w)
        if flood_warnings:
            with st.expander("⚠️ Data quality notes (flood)"):
                for w in flood_warnings:
                    st.code(w if isinstance(w, str) else json.dumps(w, indent=2), language="json")

        st.markdown("</div>", unsafe_allow_html=True)

    st.divider()
    header_col1, header_col2 = st.columns([3, 1])
    with header_col1:
        st.subheader("📋 Full hazard data (technical detail)")
    with header_col2:
        if hazard_entries:
            st.markdown(
                f"<div style='text-align:right; padding-top:0.6em;'>"
                f"✅ {ok_count} succeeded&nbsp;&nbsp;·&nbsp;&nbsp;⚠️ {err_count} failed</div>",
                unsafe_allow_html=True,
            )

    all_rows = []
    for entry in results:
        name = entry.get("name", "unknown")
        version = entry.get("version", "")
        errors = entry.get("errors") or []
        is_geocode = name == "geocode"
        icon = "🧭" if is_geocode else STATUS_ICON["ERROR" if errors else "OK"]
        panel_title = f"{icon}  **{name}**  ·  v{version}"
        with st.expander(panel_title, expanded=False):
            if errors:
                for err in errors:
                    code = err.get("code", "")
                    message = ERROR_CODES.get(code, err.get("message", "Unknown error"))
                    st.error(f"`{code}` — {message}")
            layer_results = entry.get("results") or {}
            if layer_results:
                rows = flatten_for_table(name, layer_results)
                all_rows.extend(rows)
                table_df = pd.DataFrame(rows).drop(columns=["Layer"]).rename(
                    columns={"Field": "Field", "Value": "Value"}
                )
                st.dataframe(table_df, width="stretch", hide_index=True)
            elif not errors:
                st.caption("No data returned for this layer.")

    if all_rows:
        results_df = pd.DataFrame(all_rows)
        st.divider()
        st.subheader("⬇️ Export")
        csv_bytes = results_df.to_csv(index=False).encode("utf-8")
        json_bytes = json.dumps(results, indent=2, default=str).encode("utf-8")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("Download as CSV", data=csv_bytes,
                                file_name="rms_location_intelligence.csv", mime="text/csv",
                                width="stretch")
        with col2:
            st.download_button("Download raw JSON", data=json_bytes,
                                file_name="rms_location_intelligence.json", mime="application/json",
                                width="stretch")
