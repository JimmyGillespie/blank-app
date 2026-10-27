#!/usr/bin/env python
"""
RMS Location Intelligence Explorer (Streamlit)

Enter an address, see it plotted on a map, and pick which RMS Location
Intelligence hazard layers you want enriched for that location.

Workflow (per https://developer.rms.com/location-intelligence):
  1. Geocode the address -> resolves lat/long + admin codes.
  2. Hazard lookup -> enrich the geocoded location with peril-specific hazard
     data (earthquake, windstorm, flood, wildfire, convective storm, winter
     storm, terrorism, etc.).

Both steps are combined into a single call to the Composite Lookup endpoint
(POST https://{host}/li/composite), which geocodes once and then runs every
selected hazard layer against the resolved location in one round trip:

    https://developer.rms.com/location-intelligence/reference/composite-1

Layer naming follows the documented <region>_<peril>_hazard convention
(e.g. us_eq_hazard, ac_eq_hazard, au_eq_hazard). Layer availability and
exact naming can vary by tenant/vintage, so every layer name + version is
editable in the table below before running.
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

DEFAULT_LAYERS = [
    {"Hazard": "Earthquake", "Layer": "{region}_eq_hazard", "Version": "latest", "Selected": True},
    {"Hazard": "Windstorm / Hurricane", "Layer": "{region}_ws_hazard", "Version": "latest", "Selected": True},
    {"Hazard": "Flood", "Layer": "{region}_fl_hazard", "Version": "latest", "Selected": True},
    {"Hazard": "Wildfire", "Layer": "{region}_wf_hazard", "Version": "latest", "Selected": True},
    {"Hazard": "Severe Convective Storm", "Layer": "{region}_cs_hazard", "Version": "latest", "Selected": False},
    {"Hazard": "Winter Storm", "Layer": "{region}_wt_hazard", "Version": "latest", "Selected": False},
    {"Hazard": "Terrorism", "Layer": "{region}_tr_hazard", "Version": "latest", "Selected": False},
    {"Hazard": "Earthquake Intensity (MMI)", "Layer": "{region}_mmi", "Version": "latest", "Selected": False},
    {"Hazard": "Distance to Fault", "Layer": "{region}_distance_to_fault", "Version": "latest", "Selected": False},
    {"Hazard": "Distance to Coast", "Layer": "{region}_distance_to_coast", "Version": "latest", "Selected": False},
]


# --------------------------------------------------------------------------
# API client
# --------------------------------------------------------------------------

class LocationIntelligenceApi:
    """Thin, resilient wrapper around the RMS Location Intelligence API."""

    def __init__(self, host, api_key, auth_style="Bearer token", verify=True, pool_size=16):
        self.base_url = f"https://{host}/li"
        self.session = requests.Session()
        adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        auth_value = f"Bearer {api_key}" if auth_style == "Bearer token" else api_key
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Authorization": auth_value,
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

    def geocode(self, address, version="latest"):
        body = {"location": {"address": address}}
        return self._post(f"geocode/{version}", body)

    def composite(self, address, layers):
        body = {
            "location": {"address": address},
            "layers": [{"name": l["layer"], "version": l["version"]} for l in layers],
        }
        return self._post("composite", body)

    def lookup_layer(self, address, layer, version):
        body = {"location": {"address": address}}
        return self._post(f"{layer}/{version}", body)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def build_address_payload(street, city, admin1, postal, country_code):
    address = {"countryCode": country_code, "countryScheme": "ISO2A"}
    if street:
        address["streetAddress"] = street
    if city:
        address["cityName"] = city
    if admin1:
        address["admin1Code"] = admin1
    if postal:
        address["postalCode"] = postal
    return address


def extract_latlong(geocode_results):
    if not isinstance(geocode_results, dict):
        return None, None
    lat = geocode_results.get("latitude")
    lon = geocode_results.get("longitude")
    return lat, lon


def flatten_for_table(layer_name, results):
    """Turn a layer's results dict into flat rows for display/export."""
    if not isinstance(results, dict):
        return []
    rows = []
    for key, value in results.items():
        if isinstance(value, (dict, list)):
            value = json.dumps(value)
        rows.append({"Layer": layer_name, "Field": key, "Value": value})
    return rows


def format_latlon_caption(lat, lon):
    """Build the 'Resolved coordinates' caption safely (avoids f-string nesting issues)."""
    lat_str = "{:.6f}".format(lat)
    lon_str = "{:.6f}".format(lon)
    return "Resolved coordinates: " + lat_str + ", " + lon_str


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="RMS Location Intelligence Explorer", page_icon="🌍", layout="wide")

st.title("🌍 RMS Location Intelligence Explorer")
st.write(
    "Enter an address, see it on a map, and choose which RMS hazard layers to look up for that "
    "location — earthquake, windstorm, flood, wildfire, and more."
)

if "li_results" not in st.session_state:
    st.session_state["li_results"] = None
if "li_latlon" not in st.session_state:
    st.session_state["li_latlon"] = None

with st.sidebar:
    st.header("Connection")
    host = st.text_input("Host", value=DEFAULT_HOST, help="api-euw1.rms.com or api-use1.rms.com")
    api_key = st.text_input("API key", value=API_KEY, type="password",
                             help="Defaults to the RMS_LI_API_KEY environment variable.")
    auth_style = st.selectbox("Auth header style", ["Bearer token", "API key (raw)"], index=0)
    verify_ssl = st.checkbox("Verify SSL", value=True,
                              help="Uncheck only if you hit corporate SSL/proxy errors.")
    region_code = st.text_input("Region code", value="us",
                                 help="Prefix used in <region>_<peril>_hazard layer names, "
                                      "e.g. 'us', 'ca', 'eu', 'au'.")
    version_default = st.text_input("Default layer version", value="latest")

st.subheader("1. Address")
col1, col2, col3, col4 = st.columns([3, 2, 1, 1])
with col1:
    street = st.text_input("Street address", placeholder="98 East First Street")
with col2:
    city = st.text_input("City", placeholder="Donora")
with col3:
    admin1 = st.text_input("State / Admin1", placeholder="PA")
with col4:
    postal = st.text_input("Postal code", placeholder="15033")
country_code = st.text_input("Country code (ISO2)", value="US", max_chars=2)

st.subheader("2. Hazard layers to look up")
st.caption(
    "Layer names follow RMS's documented `<region>_<peril>_hazard` convention. Edit names/versions "
    "below if your tenant's Location Intelligence layer catalog differs."
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
        selected_layers = [
            {"hazard": row["Hazard"], "layer": row["Layer"], "version": row["Version"] or version_default}
            for _, row in edited_df.iterrows()
            if row.get("Selected") and str(row.get("Layer", "")).strip()
        ]
        address = build_address_payload(street, city, admin1, postal, country_code.strip().upper() or "US")
        api = LocationIntelligenceApi(host.strip(), api_key, auth_style=auth_style, verify=verify_ssl)

        with st.spinner("Geocoding address and looking up hazard layers..."):
            payload, status, error = api.composite(address, selected_layers)

        if status != 200 or payload is None:
            st.error("Composite lookup failed (HTTP {}). Falling back to individual layer calls...".format(status))
            # Fallback: geocode first, then call each hazard layer separately.
            geo_payload, geo_status, geo_error = api.geocode(address, version_default)
            if geo_status != 200 or geo_payload is None:
                st.error("Geocoding failed (HTTP {}): {}".format(geo_status, geo_error))
            else:
                lat, lon = extract_latlong(geo_payload)
                st.session_state["li_latlon"] = (lat, lon)
                results = [{"name": "geocode", "version": version_default, "results": geo_payload, "errors": []}]
                for layer in selected_layers:
                    r_payload, r_status, r_error = api.lookup_layer(address, layer["layer"], layer["version"])
                    if r_status == 200 and r_payload is not None:
                        results.append({"name": layer["layer"], "version": layer["version"],
                                         "results": r_payload, "errors": []})
                    else:
                        err_msg = "HTTP {}: {}".format(r_status, r_error)
                        results.append({"name": layer["layer"], "version": layer["version"], "results": {},
                                         "errors": [{"code": "ERR_HTTP", "message": err_msg}]})
                st.session_state["li_results"] = results
        else:
            st.session_state["li_results"] = payload
            geocode_entry = next((item for item in payload if item.get("name") == "geocode"), None)
            if geocode_entry:
                lat, lon = extract_latlong(geocode_entry.get("results"))
                st.session_state["li_latlon"] = (lat, lon)

# --------------------------------------------------------------------------
# Results
# --------------------------------------------------------------------------

results = st.session_state["li_results"]
latlon = st.session_state["li_latlon"]

if results is not None:
    st.subheader("3. Map")
    if latlon and latlon[0] is not None and latlon[1] is not None:
        map_df = pd.DataFrame({"lat": [latlon[0]], "lon": [latlon[1]]})
        st.map(map_df, zoom=13)
        st.caption(format_latlon_caption(latlon[0], latlon[1]))
    else:
        st.info("No coordinates were returned for this address - check the geocode result below.")

    st.subheader("4. Hazard results")
    all_rows = []
    for entry in results:
        name = entry.get("name", "unknown")
        version = entry.get("version", "")
        errors = entry.get("errors") or []
        panel_title = "📍 {}  ·  v{}".format(name, version)
        with st.expander(panel_title, expanded=(name == "geocode")):
            if errors:
                for err in errors:
                    code = err.get("code", "")
                    message = ERROR_CODES.get(code, err.get("message", "Unknown error"))
                    st.warning("{}: {}".format(code, message))
            layer_results = entry.get("results") or {}
            if layer_results:
                rows = flatten_for_table(name, layer_results)
                all_rows.extend(rows)
                table_df = pd.DataFrame(rows).drop(columns=["Layer"])
                st.dataframe(table_df, width="stretch", hide_index=True)
            elif not errors:
                st.caption("No data returned for this layer.")

    if all_rows:
        results_df = pd.DataFrame(all_rows)
        st.subheader("5. Export")
        csv_bytes = results_df.to_csv(index=False).encode("utf-8")
        json_bytes = json.dumps(results, indent=2, default=str).encode("utf-8")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("⬇️ Download as CSV", data=csv_bytes,
                                file_name="rms_location_intelligence.csv", mime="text/csv")
        with col2:
            st.download_button("⬇️ Download raw JSON", data=json_bytes,
                                file_name="rms_location_intelligence.json", mime="application/json")
