import cv2
import pytesseract
from dash import Dash, dcc, html, Output, Input, State
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import base64, re, os, collections
import numpy as np
import pandas as pd
from datetime import datetime

# -----------------------------
# 1. CONFIGURATION
# -----------------------------
path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(path_to_tesseract):
    pytesseract.pytesseract.tesseract_cmd = path_to_tesseract
else:
    print("❌ Tesseract not found. Fix the path!")

cap = cv2.VideoCapture(0)
cap.set(cv2.CAP_PROP_FPS, 30)

# -----------------------------
# 2. DASH APP
# -----------------------------
app = Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])

history = {k: collections.deque(maxlen=30) for k in
           ["aqi", "pm25", "pm10", "co2", "tvoc", "hcho", "temp"]}

LOG_FILE = "air_quality_log.xlsx"

# -----------------------------
# 3. AQI COLORS
# -----------------------------
def aqi_color(aqi):
    if aqi <= 50: return "#009966"
    elif aqi <= 100: return "#FFDE33"
    elif aqi <= 150: return "#FF9933"
    elif aqi <= 200: return "#CC0033"
    elif aqi <= 300: return "#660099"
    else: return "#7E0023"

# -----------------------------
# 4. DASH LAYOUT
# -----------------------------
app.layout = html.Div(
    style={"padding": "25px", "fontFamily": "Segoe UI", "background": "#f4f7fb"},
    children=[
        html.H1("📊 Simbo Smart Air Monitor — Live Dashboard",
                style={"textAlign": "center", "color": "#222", "fontWeight": "800", "marginBottom": "25px"}),

        # TOP KPIs
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "repeat(6, 1fr)", "gap": "15px"},
            children=[
                *[html.Div(style={"background": "white", "borderRadius": "10px", "padding": "15px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)", "textAlign": "center", "height": "110px"},
                           children=[
                               html.Div(name, style={"fontSize": "13px", "color": "#555"}),
                               html.H2(id="val-" + tag, style={"color": "#1164A3", "fontWeight": "800", "marginTop": "5px"})
                           ])
                  for name, tag in [("Temperature (°C)", "temp"), ("PM 2.5 (µg/m³)", "pm25"), ("PM 10 (µg/m³)", "pm10"),
                                    ("CO₂ (ppm)", "co2"), ("TVOC (ppb)", "tvoc"), ("HCHO (mg/m³)", "hcho")]]
            ]
        ),

        html.Br(),

        # CAMERA + LARGE AQI
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "380px 1fr", "gap": "25px"},
            children=[
                html.Div(style={"background": "white", "borderRadius": "10px", "padding": "18px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)"},
                         children=[
                             html.H4("📷 Live Device Feed", style={"marginBottom": "8px", "color": "#333"}),
                             html.Img(id="camera-feed", style={"width": "100%", "borderRadius": "10px"}),
                             html.Div("OCR Extracted Text", style={"marginTop": "10px", "fontSize": "12px", "opacity": 0.6}),
                             html.Div(id="ocr-result", style={"color": "#0A84FF", "fontWeight": "700", "marginTop": "6px"})
                         ]),

                html.Div(style={"background": "white", "borderRadius": "10px", "padding": "18px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)", "textAlign": "center"},
                         children=[
                             html.H4("🌫 Air Quality Index (AQI)", style={"color": "#333"}),
                             html.H1(id="val-aqi", style={"fontSize": "65px", "fontWeight": "900", "color": "#1164A3"})
                         ])
            ]
        ),

        html.Br(),

        # CHARTS
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "repeat(3, 1fr)", "gap": "22px"},
            children=[
                html.Div(style={"background": "white", "padding": "12px", "borderRadius": "10px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)", "height": "320px"},
                         children=[html.H4("📈 Trend Line", style={"color": "#333", "fontSize": "15px"}),
                                   dcc.Graph(id="live-line-chart", config={"displayModeBar": False}, style={"height": "260px"})]),

                html.Div(style={"background": "white", "padding": "12px", "borderRadius": "10px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)", "height": "320px"},
                         children=[html.H4("📊 Latest Values", style={"color": "#333", "fontSize": "15px"}),
                                   dcc.Graph(id="live-bar-chart", config={"displayModeBar": False}, style={"height": "260px"})]),

                html.Div(style={"background": "white", "padding": "12px", "borderRadius": "10px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)", "height": "320px"},
                         children=[html.H4("🟢 Category Breakdown", style={"color": "#333", "fontSize": "15px"}),
                                   dcc.Graph(id="live-pie-chart", config={"displayModeBar": False}, style={"height": "260px"})])
            ]
        ),

        html.Br(),

        # EXPORT SECTION (UPDATED)
        html.Div(
            style={"background": "white", "borderRadius": "10px", "padding": "15px", "boxShadow": "0 2px 8px rgba(0,0,0,0.06)", "textAlign": "center", "width": "300px"},
            children=[
                html.H4("💾 Data Export", style={"color": "#333"}),
                html.Div(id="export-status", style={"marginTop": "8px", "fontSize": "14px", "marginBottom": "10px"}),
                
                # --- NEW DOWNLOAD BUTTON ---
                dbc.Button("📥 Download Excel Report", id="btn-download", color="primary", size="sm"),
                dcc.Download(id="download-dataframe-xlsx")
            ]
        ),

        dcc.Interval(id="timer", interval=2000, n_intervals=0) # Increased to 2s to reduce file locking issues
    ]
)

# -----------------------------
# 5. OCR PROCESSING
# -----------------------------
def extract_parameter(text, label):
    if label.upper() == "AQI": pattern = r"AQI\s*[:=]\s*(\d{1,4})"
    else: pattern = rf"{label}\s*[:\-]?\s*(\d+\.\d+|\d+)"
    match = re.search(pattern, text, re.IGNORECASE)
    return float(match.group(1)) if match else None

def process_ocr(frame):
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    bin_img = cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    text = pytesseract.image_to_string(bin_img)
    values = {
        "aqi": extract_parameter(text, "AQI"),
        "pm25": extract_parameter(text, "PM2.5"),
        "pm10": extract_parameter(text, "PM10"),
        "co2": extract_parameter(text, "CO2"),
        "tvoc": extract_parameter(text, "TVOC"),
        "hcho": extract_parameter(text, "HCHO"),
        "temp": extract_parameter(text, "TEMP"),
    }
    return values, text

def moving_average(data, window=3):
    arr = np.array(data)
    if len(arr) < window: return arr
    return np.convolve(arr, np.ones(window)/window, mode='valid')

# -----------------------------
# 6. CALLBACKS
# -----------------------------

# Callback 1: Handle Data Update & Logging
@app.callback(
    [*[Output("val-" + k, "children") for k in history.keys()],
     Output("ocr-result", "children"),
     Output("camera-feed", "src"),
     Output("live-line-chart", "figure"),
     Output("live-bar-chart", "figure"),
     Output("live-pie-chart", "figure"),
     Output("export-status", "children")],
    Input("timer", "n_intervals")
)
def update_metrics(n):
    ret, frame = cap.read()
    if not ret: return ["--"] * 7 + ["No Camera"] * 4 + ["No Export"]

    _, buffer = cv2.imencode(".jpg", frame)
    img_src = "data:image/jpeg;base64," + base64.b64encode(buffer).decode()

    extracted, debug = process_ocr(frame)
    latest_values = []
    
    # Update History & Handle Missing Values
    for key in history.keys():
        val = extracted[key]
        if val is None:
            # If OCR failed, use last known value or 0
            val = history[key][-1] if history[key] else 0.0
        
        val = round(float(val), 2)
        history[key].append(val)
        latest_values.append(val)

    # --- SAVE TO EXCEL ---
    # We now save even if some values are 0, so the file is created.
    row = {**dict(zip(history.keys(), latest_values)), "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    export_message = "Waiting for data..."
    
    try:
        # Create DataFrame
        df_new = pd.DataFrame([row])
        
        if os.path.exists(LOG_FILE):
            # Check if file is readable (not locked by user opening it)
            try:
                with open(LOG_FILE, "a"): pass 
                df_existing = pd.read_excel(LOG_FILE)
                df_final = pd.concat([df_existing, df_new], ignore_index=True)
                df_final.to_excel(LOG_FILE, index=False)
                export_message = f"✅ Logged 1 row at {row['Timestamp'][-8:]}"
            except PermissionError:
                export_message = "⚠️ File open in Excel! Close to save."
        else:
            df_new.to_excel(LOG_FILE, index=False)
            export_message = "✅ Created new Log File"
            
    except Exception as e:
        export_message = f"Error: {str(e)}"

    # Charts
    x_axis = list(range(len(history["aqi"])))
    line_fig = go.Figure()
    for k in history:
        y = list(history[k])
        y_smooth = moving_average(y, window=3)
        x_smooth = x_axis[len(x_axis)-len(y_smooth):]
        line_fig.add_trace(go.Scatter(x=x_smooth, y=y_smooth, mode="lines", name=k))

    bar_fig = go.Figure([go.Bar(x=list(history.keys()), y=[history[k][-1] for k in history],
                                marker_color=[aqi_color(history["aqi"][-1]) if k=="aqi" else "#1164A3" for k in history])])

    pie_fig = go.Figure(go.Pie(labels=list(history.keys()), values=[history[k][-1] for k in history], hole=0.4))

    return (*latest_values, debug, img_src, line_fig, bar_fig, pie_fig, export_message)


# Callback 2: Handle File Download
@app.callback(
    Output("download-dataframe-xlsx", "data"),
    Input("btn-download", "n_clicks"),
    prevent_initial_call=True
)
def download_excel(n_clicks):
    if os.path.exists(LOG_FILE):
        return dcc.send_file(LOG_FILE)
    return None

# -----------------------------
# RUN
# -----------------------------
if __name__ == "__main__":
    app.run(debug=True)