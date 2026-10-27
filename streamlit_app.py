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
     "LayerOptions": '{"distanceUnit":"feet","proximityDistance":100}', "Confirmed": "Yes", "Selected": True},
    {"Hazard": "US Flood Depth (undefended)", "Layer": "us_fl_depth", "Version": "latest",
     "LayerOptions": '{"lookup":"floodDepth","proximityDistance":100,"distanceUnit":"feet"}',
     "Confirmed": "Yes", "Selected": True},
    {"Hazard": "Elevation / Susceptibility", "Layer": "us_fl_susceptibility", "Version": "latest",
     "LayerOptions": "", "Confirmed": "Yes", "Selected": True},
    {"Hazard": "Earthquake (name unconfirmed)", "Layer": "us_eq_sitehazard", "Version": "latest",
     "LayerOptions": "", "Confirmed": "No - edit before use", "Selected": False},
    {"Hazard": "Windstorm (name unconfirmed)", "Layer": "us_hu_hazard", "Version": "latest",
     "LayerOptions": "", "Confirmed": "No - edit before use", "Selected": False},
    {"Hazard": "Wildfire (name unconfirmed)", "Layer": "us_wf_hazard", "Version": "latest",
     "LayerOptions": "", "Confirmed": "No - edit before use", "Selected": False},
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
                       layer, version, layer_options=None):
        """Nested location/characteristics/coverageValues body - matches
        BuildBody() in the VBA macro exactly."""
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

        body = {
            "location": {
                "address": address,
                "characteristics": {
                    "construction": "ATC0", "occupancy": "ATC0", "yearBuilt": 0,
                    "numOfStories": 0, "foundationType": 0,
                    "firstFloorHeightAboveGround": 0, "basement": "DEFAULT", "floorArea": 0,
                },
                "coverageValues": {
                    "buildingValue": 0, "contentsValue": 0, "businessInterruptionValue": 0,
                },
            }
        }
        if layer_options:
            body["layerOptions"] = layer_options
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


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="RMS Location Intelligence Explorer", page_icon="🌍", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; }
    div[data-testid="stExpander"] { border: 1px solid #2b2f3a; border-radius: 8px; }
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
    "✅ The first 3 rows (FEMA FIRM, Flood Depth, Elevation/Susceptibility) use the exact request "
    "shape confirmed working in your RMS workbook macro - these just work. "
    "⚠️ The remaining rows are **unconfirmed guesses** at the layer name - the US product line does "
    "not follow the generic `<region>_<peril>_hazard` pattern used for other regions, so these will "
    "likely return 'No details found for layer name...' until you correct the name. Find the exact "
    "name in your tenant's Postman collection / OpenAPI spec (Developer Assets on developer.rms.com), "
    "or in an existing working macro/report, then edit the 'Layer name' cell directly below."
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
                with st.spinner(f"Looking up {len(selected_layers)} hazard layer(s)..."):
                    for row in selected_layers:
                        layer_options = parse_layer_options(row.get("LayerOptions"))
                        r_payload, r_status, r_error = api.hazard_lookup(
                            lat, lon, resolved_admin1, country_code.strip().upper() or "US",
                            resolution_code, row["Layer"], row["Version"] or version_default,
                            layer_options,
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
            m1, m2 = st.columns(2)
            m1.metric("Match resolution", geo_data.get("rmsGeocodingResolutionCode", "—"))
            m2.metric("Model resolution", geo_data.get("rmsGeoModelResolutionCode", "—"))
    else:
        st.info("No coordinates were returned for this address - check the geocode result below.")

    st.divider()
    header_col1, header_col2 = st.columns([3, 1])
    with header_col1:
        st.subheader("🌪️ Hazard results")
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
