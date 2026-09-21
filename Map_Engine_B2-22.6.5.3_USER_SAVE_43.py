# -*- coding: utf-8 -*-
"""
Beta Farm - Map Engine B2-22

فقط موتور نقشه است.

مسئولیت:
- اجرای نقشه
- دریافت داده آماده از Dynamic Filter
- نمایش نقاط
- زوم و حرکت دستی

بدون:
- Excel
- Master.xlsx
- openpyxl
- فیلتر مستقل
- استخراج داده
- Popup
- fit خودکار
"""
import urllib.parse
from pathlib import Path
import json
import base64
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
import tkinter as tk
from tkinter import filedialog
import ctypes
import time
CAPTURE_PORT = 8765


class CaptureHandler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):

        self.send_response(204)

        self.send_header(
            "Access-Control-Allow-Origin",
            "*"
        )

        self.send_header(
            "Access-Control-Allow-Methods",
            "POST, OPTIONS"
        )

        self.send_header(
            "Access-Control-Allow-Headers",
            "Content-Type"
        )

        self.end_headers()


    def do_POST(self):

        if self.path != "/save_capture":

            self.send_response(404)
            self.end_headers()
            return


        try:

            content_length = int(
                self.headers.get(
                    "Content-Length",
                    0
                )
            )

            body = self.rfile.read(
                content_length
            )

            content_type = (
                self.headers.get(
                    "Content-Type",
                    ""
                )
            )


            if (
                "application/x-www-form-urlencoded"
                in content_type
            ):

                form_data = urllib.parse.parse_qs(
                    body.decode("utf-8")
                )

                image_data = form_data.get(
                    "image",
                    [""]
                )[0]

            else:

                data = json.loads(
                    body.decode("utf-8")
                )

                image_data = data.get(
                    "image",
                    ""
                )


            if not image_data:

                raise ValueError(
                    "Image data is empty"
                )


            if "," in image_data:

                image_data = image_data.split(
                    ",",
                    1
                )[1]


            image_bytes = base64.b64decode(
                image_data
            )


            # =================================================
            # WINDOWS SAVE AS
            # =================================================

            def bring_save_dialog_to_front():

                user32 = ctypes.windll.user32

                for _ in range(50):

                    hwnd = user32.FindWindowW(
                        None,
                        "Save Map Picture"
                    )

                    if hwnd:

                        user32.ShowWindow(
                            hwnd,
                            5
                        )

                        user32.SetForegroundWindow(
                            hwnd
                        )

                        return

                    time.sleep(0.1)


            root = tk.Tk()
            root.withdraw()

            root.attributes(
                "-topmost",
                True
            )

            root.update()


            threading.Thread(
                target=bring_save_dialog_to_front,
                daemon=True
            ).start()


            save_path = filedialog.asksaveasfilename(
                parent=root,
                title="Save Map Picture",
                defaultextension=".png",
                filetypes=[
                    ("PNG Image", "*.png")
                ],
                initialfile="Beta_Farm_Map.png"
            )


            root.destroy()


            # =================================================
            # CANCEL
            # =================================================

            if not save_path:

                self.send_response(204)

                self.send_header(
                    "Access-Control-Allow-Origin",
                    "*"
                )

                self.end_headers()

                return


            # =================================================
            # SAVE IMAGE
            # =================================================

            with open(
                save_path,
                "wb"
            ) as f:

                f.write(
                    image_bytes
                )


            self.send_response(200)

            self.send_header(
                "Access-Control-Allow-Origin",
                "*"
            )

            self.end_headers()


        except Exception as error:

            print(
                "❌ CAPTURE SAVE ERROR:",
                error
            )

            self.send_response(500)

            self.send_header(
                "Access-Control-Allow-Origin",
                "*"
            )

            self.end_headers()


def start_capture_server():

    try:

        server = HTTPServer(
            (
                "0.0.0.0",
                CAPTURE_PORT
            ),
            CaptureHandler
        )

        print(
            f"📷 Capture Server started on port "
            f"{CAPTURE_PORT}"
        )

        server.serve_forever()


    except Exception as error:

        print(
            "❌ CAPTURE SERVER ERROR:",
            error
        )


capture_server_thread = threading.Thread(
    target=start_capture_server,
    daemon=False
)

capture_server_thread.start()


# ============================================================
# مسیر پروژه
# ============================================================

CURRENT_FILE = Path(__file__).resolve()

PROJECT_CANDIDATES = [
    CURRENT_FILE.parent,
    CURRENT_FILE.parent.parent,
    CURRENT_FILE.parent.parent.parent,
]

PROJECT_DIR = None

for candidate in PROJECT_CANDIDATES:

    if (candidate / "maps" / "leaflet").exists():
        PROJECT_DIR = candidate
        break

if PROJECT_DIR is None:
    PROJECT_DIR = CURRENT_FILE.parent


MAPS_DIR = PROJECT_DIR / "maps"

LEAFLET_DIR = MAPS_DIR / "leaflet"

LEAFLET_JS = LEAFLET_DIR / "leaflet.js"

LEAFLET_CSS = LEAFLET_DIR / "leaflet.css"

OUTPUT_FILE = MAPS_DIR / "Beta_Farm_Map_B2-22.html"


# ============================================================
# بررسی Leaflet
# ============================================================

if not LEAFLET_JS.exists():

    raise FileNotFoundError(
        "leaflet.js پیدا نشد:\n"
        + str(LEAFLET_JS)
    )


if not LEAFLET_CSS.exists():

    raise FileNotFoundError(
        "leaflet.css پیدا نشد:\n"
        + str(LEAFLET_CSS)
    )


# ============================================================
# HTML
# ============================================================

html = f'''
<!DOCTYPE html>

<html lang="fa" dir="rtl">

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>
Beta Farm - Map Engine B2-22
</title>


<link
    rel="stylesheet"
    href="leaflet/leaflet.css"
>


<style>

* {{
    box-sizing: border-box;
}}


html,
body,
#map {{

    width: 100%;
    height: 100%;

    margin: 0;
    padding: 0;

}}


body {{

    overflow: hidden;

    font-family:
        Tahoma,
        Arial,
        sans-serif;

    background: #e9eef2;

}}


#map {{

    position: relative;

}}


.quarantine-list {{
    position: absolute;

    top: 12px;
    right: 12px;

    z-index: 1000;

    min-width: 260px;
    max-width: 360px;
      
    max-height: 420px;
    overflow-y: auto;
    overflow-x: hidden; 
    padding: 4px 6px;

    background:
        rgba(255,255,255,.82);

    border-radius: 5px;

    font-family:
        Tahoma,
        Arial,
        sans-serif;

    font-size: 10px;

    color: #263746;

    direction: rtl;
    text-align: right;

    pointer-events: auto;

    display: none;
}}

.map-status {{

    position: absolute;

    right: 12px;
    bottom: 10px;

    z-index: 1000;

    padding: 4px 8px;

    border-radius: 5px;

    background:
        rgba(255,255,255,.68);

    color: #53616d;

    font-size: 10px;

    pointer-events: none;

}}
.offline-map-selector {{
    position: absolute;
    top: 50%;
    right: 10px;
    transform: translateY(-50%);
    z-index: 1000;

    display: none;

    flex-direction: column;
    gap: 5px;
}}

.offline-map-button {{
    width: 42px;
    height: 34px;

    border: 1px solid #b8c0c7;
    border-radius: 6px;

    background: rgba(255,255,255,.94);

    color: #263746;

    font-family: Tahoma, Arial, sans-serif;
    font-size: 11px;
    font-weight: bold;

    cursor: pointer;
}}

.offline-map-button.active {{
    background: #17452d;
    color: white;
    border-color: #17452d;
}}

.offline-map-button:hover {{
    background: #e9eef2;
}}

.offline-map-button.active:hover {{
    background: #17452d;
}}

#captureImageMeta {{
    width: 100%;
    height: 28px;

    flex-shrink: 0;

    background: #faf7ef;
    border: 1px solid #e4dfd2;
    border-radius: 4px;

    padding: 4px 8px;

    box-sizing: border-box;

    font-family:
        Tahoma,
        Arial,
        sans-serif;

    font-size: 10px;
    color: #444;

    direction: rtl;
    text-align: right;

    white-space: nowrap;
    overflow: hidden;

      
}}

/* =========================================================
   CAPTURE PREVIEW LAYER
   ========================================================= */

#capturePanel {{
    position: absolute;

    left: 10%;
    top: 10%;

    width: 80%;
    height: 80%;

    z-index: 100000;

    background: rgba(255,255,255,.96);

    border: 1px solid #999;
    border-radius: 8px;

    box-shadow:
        0 4px 20px rgba(0,0,0,.30);

    overflow: hidden;

    pointer-events: auto;
}}


#captureImageBox {{
    position: absolute;

    left: 0;
    top: 0;

    width: 100%;
    height: 100%;

    display: flex;

    align-items: center;
    justify-content: center;

    overflow: auto;

    padding: 10px;
}}


#capturePreview {{
    display: block;

    max-width: 100%;
    max-height: 100%;

    width: auto;
    height: auto;

    object-fit: contain;
    pointer-events: none;
}}
#saveMyPictureBtn {{
    position: absolute;
    top: 6px;
    right: 36px;
    width: 25px;
    height: 25px;
    padding: 0;
    margin: 0;
    border: 1px solid #b8b8b8;
    border-radius: 4px;
    background: #f2f2f2;
    color: #333;
    font-family: Arial, sans-serif;
    font-size: 14px;
    line-height: 23px;
    text-align: center;
    cursor: pointer;
    z-index: 100002;
}}

#saveMyPictureBtn:hover {{
    background: #e2e2e2;
}}

#closeCaptureBtn {{
    position: absolute;
    top: 6px;
    right: 6px;
    width: 25px;
    height: 25px;
    padding: 0;
    margin: 0;
    border: 1px solid #b8b8b8;
    border-radius: 4px;
    background: #f2f2f2;
    color: #333;
    font-family: Arial, sans-serif;
    font-size: 17px;
    line-height: 21px;
    text-align: center;
    cursor: pointer;
    z-index: 100002;
}}

#closeCaptureBtn:hover {{
    background: #e81123;
    color: white;
    border-color: #e81123;
}}
/* =========================================================
   CAPTURE SAVE AS PANEL
   ========================================================= */

#captureSaveBox {{
    position: absolute;
    top: 55px;
    right: 15px;
    width: 170px;
    background: white;
    border: 1px solid #999;
    border-radius: 6px;
    box-shadow: 0 4px 15px rgba(0,0,0,.25);
    padding: 8px;
    z-index: 100001;
}}

#captureSaveTitle {{
    font-weight: bold;
    font-size: 14px;
    padding: 6px 8px;
    border-bottom: 1px solid #ddd;
    margin-bottom: 5px;
}}

.captureSaveItem {{
    width: 100%;
    box-sizing: border-box;
    border: 1px solid #ccc;
    background: #f7f7f7;
    padding: 7px 8px;
    margin: 3px 0;
    border-radius: 4px;
    text-align: left;
    cursor: pointer;
    font-size: 13px;
}}

.captureSaveItem:hover {{
    background: #e8e8e8;
}}

#captureReportList {{
    max-height: 250px;
    overflow-y: auto;
}}
</style>

</head>


<body>


<div id="map">


    <div
        id="quarantineList"
        class="quarantine-list"
    >

    </div>


    <div
        id="mapStatus"
        class="map-status"
    >

        در حال بارگذاری نقشه...

    </div>
    <div
    id="offlineMapSelector"
    class="offline-map-selector"
    >
    <button
        class="offline-map-button"
        data-map="001.png"
        onclick="selectOfflineMap('001.png', this)"
    >
        001
    </button>

    <button
        class="offline-map-button active"
        data-map="002.png"
        onclick="selectOfflineMap('002.png', this)"
    >
        002
    </button>

    <button
        class="offline-map-button"
        data-map="003.png"
        onclick="selectOfflineMap('003.png', this)"
    >
        003
    </button>

    <button
        class="offline-map-button"
        data-map="004.png"
        onclick="selectOfflineMap('004.png', this)"
    >
        004
    </button>
    
    </div>
    
    <div
    id="capturePanel"
    style="display:none;">

    <div id="captureImageBox">

        <button
            id="saveMyPictureBtn"
            type="button"
            title="ذخیره">
            💾
        </button>

        <button
            id="closeCaptureBtn"
            type="button"
            title="بستن">
            ×
        </button>

        <img
            id="capturePreview"
            alt="Map Capture">

    </div>

</div>


</div>


<script src="leaflet/leaflet.js"></script>

<script src="html2canvas.min.js"></script>
<script>

"use strict";
window.addEventListener(
    "load",
    function() {{

        const saveBtn =
            document.getElementById(
                "saveMyPictureBtn"
            );

        if (!saveBtn) {{

            setStatus(
                "❌ My Picture پیدا نشد"
            );

            return;
        }}

        saveBtn.onclick =
            function() {{

                setStatus(
                    "💾 My Picture کلیک شد"
                );

                const preview =
                    document.getElementById(
                        "capturePreview"
                    );

                if (!preview || !preview.src) {{

                    setStatus(
                        "❌ تصویری برای ذخیره وجود ندارد"
                    );

                    return;
                }}

                if (
                    window.parent &&
                    window.parent !== window
                ) {{

                    window.parent.postMessage(
                        {{
                            type:
                                "BETA_FARM_SAVE_CAPTURE",
                            image:
                                preview.src
                        }},
                        "*"
                    );

                    setStatus(
                        "💾 در حال باز کردن Save As..."
                    );

                }} else {{

                    setStatus(
                        "❌ Dynamic Filter پیدا نشد"
                    );
                }}
            }};
    }}
);
        const closeBtn =
            document.getElementById(
                "closeCaptureBtn"
            );

        if (closeBtn) {{
            closeBtn.onclick =
                function(event) {{

                    event.preventDefault();
                    event.stopPropagation();

                    const panel =
                        document.getElementById(
                            "capturePanel"
                        );

                    if (panel) {{
                        panel.style.display =
                            "none";
                    }}

                    setStatus(
                        "📷 Preview بسته شد"
                    );
                }};
        }}
function showCapturePreview(imageData) {{
    setStatus("🟢 وارد SHOW PREVIEW شد");
    
    const panel =
        document.getElementById(
            "capturePanel"
        );

    const preview =
        document.getElementById(
            "capturePreview"
        );

    setStatus(
        panel && preview
            ? "🟢 Panel و Preview پیدا شدند"
            : "🔴 Panel یا Preview پیدا نشد"
    );

    if (!panel || !preview) {{
          console.error(
              "❌ CAPTURE PANEL ELEMENT NOT FOUND"
        );

        return;
    }}

    preview.src = imageData;

    panel.style.display = "block";
    panel.onmousedown = function(event) {{
    event.stopPropagation();
    }};

    panel.onmouseup = function(event) {{
        event.stopPropagation();
    }};

    panel.onclick = function(event) {{
        event.stopPropagation();
    }};

    panel.ondblclick = function(event) {{
        event.stopPropagation();
    }};

    panel.onwheel = function(event) {{
       event.stopPropagation();
    }};
   
}}    
// ============================================================
// MAP CAPTURE
// ============================================================

async function captureCurrentMap() {{
    const cameraButton =
        document.querySelector(
            ".leaflet-control-camera"
        );

    try {{
        const mapElement =
            document.getElementById("map");

        if (!mapElement) {{
            console.error(
                "❌ MAP ELEMENT NOT FOUND"
            );
            return null;
        }}

        // مخفی کردن موقت دکمه Camera
        if (cameraButton) {{
            cameraButton.style.visibility =
                "hidden";
        }}

        const canvas =
            await html2canvas(
                mapElement,
                {{
                    useCORS: true,
                    allowTaint: false,
                    backgroundColor: "#e9eef2",
                    logging: false,
                    scale: 1
                }}
            );

        const imageData =
            canvas.toDataURL("image/png");

        console.log(
            "📷 MAP CAPTURE CREATED"
        );

        return imageData;

    }} catch (error) {{

        console.error(
            "❌ MAP CAPTURE ERROR:",
            error
        );

        setStatus(
            "❌ خطای Capture: " +
            (error.message || error)
        );

        return null;

    }} finally {{

        // برگرداندن Camera بعد از Capture
        if (cameraButton) {{
            cameraButton.style.visibility =
                "";
        }}
    }}
}}
// ============================================================
// وضعیت داخلی موتور
// ============================================================

let mapData = [];

let offlineImage = null;
let currentOfflineMap = "002.png";
let lastQuarantineReportCaptures = null;
const offlineBounds = [
    [35.5779816, 47.1799518],
    [37.2492219, 49.4336719]
];

function showOfflineMap002() {{

    if (offlineImage) {{

        if (map.hasLayer(offlineImage)) {{
            map.removeLayer(offlineImage);
        }}

    }}

    offlineImage =
        L.imageOverlay(
            currentOfflineMap,
            offlineBounds,
            {{
                opacity: 1,
                interactive: false
            }}
        );

    offlineImage.addTo(map);

    setStatus(
        "نقشه آفلاین: 002"
    );

    map.invalidateSize();
}}
function selectOfflineMap(fileName, button) {{

    console.log(
        "B2-22 OFFLINE MAP SELECT:",
        fileName
    );

    currentOfflineMap =
        fileName;

    if (offlineImage) {{

        if (map.hasLayer(offlineImage)) {{
            map.removeLayer(offlineImage);
        }}

    }}

    offlineImage =
        L.imageOverlay(
            currentOfflineMap,
            offlineBounds,
            {{
                opacity: 1,
                interactive: false
            }}
        );

    offlineImage.addTo(map);

    document
        .querySelectorAll(
            ".offline-map-button"
        )
        .forEach(
            function(btn) {{
                btn.classList.remove(
                    "active"
                );
            }}
        );

    if (button) {{
        button.classList.add(
            "active"
        );
    }}

    setStatus(
        "نقشه آفلاین: " +
        fileName.replace(
            ".png",
            ""
        )
    );

    setTimeout(
        function() {{
            map.invalidateSize();
        }},
        100
    );
}}
function hideOfflineMap() {{

    if (
        offlineImage &&
        map.hasLayer(offlineImage)
    ) {{
        map.removeLayer(offlineImage);
    }}
}}

function setBetaFarmViewMode(mode) {{

   if (mode === "OFFLINE") {{

    showOfflineMap002();

    const selector =
        document.getElementById(
            "offlineMapSelector"
        );

    if (selector) {{
        selector.style.display =
            "flex";
    }}

}} else {{

    hideOfflineMap();

    const selector =
        document.getElementById(
            "offlineMapSelector"
        );

    if (selector) {{
        selector.style.display =
            "none";
    }}

}}
}}

/* =========================================================
   PAGE 4 DEBUG MONITOR — DISABLED
   مانیتور و کادر شناور حذف شده‌اند.
   page4MonitorAdd فقط برای حفظ سازگاری setStatus باقی مانده.
   ========================================================= */

function page4MonitorAdd(text) {{
    return;
}}


const mapStatus =
    document.getElementById(
        "mapStatus"
    );


function setStatus(text) {{

    mapStatus.textContent =
        text;

    page4MonitorAdd(
        text
    );

}}

// ============================================================
// نقشه
// مرکز اولیه = زنجان
// ============================================================

const map =
    L.map(
        "map",
        {{

            zoomControl: false,

            preferCanvas: true

        }}
    );


map.setView(

    [
        36.6769,
        48.4850
    ],

    9

);


L.control.zoom(

    {{

        position:
            "bottomleft"

    }}

).addTo(map);
// ============================================================
// CAMERA BUTTON
// ============================================================

const cameraControl =
    L.control({{
        position:
            "topright"
    }});

cameraControl.onAdd =
    function() {{

        const button =
            L.DomUtil.create(
                "button",
                "leaflet-control-camera"
            );

        button.innerHTML = "📷";
        button.title = "تصویر از نقشه";

        button.style.width = "34px";
        button.style.height = "34px";
        button.style.padding = "0";
        button.style.margin = "0";
        button.style.border = "none";
        button.style.background = "transparent";
        button.style.boxShadow = "none";
        button.style.fontSize = "20px";
        button.style.lineHeight = "34px";
        button.style.cursor = "pointer";

        L.DomEvent.disableClickPropagation(button);

        button.onclick =
    async function() {{

        setStatus(
            "📷 دکمه دوربین کلیک شد"
        );

        const image =
            await captureCurrentMap();

        console.log(
            "📷 CAPTURE RESULT:",
            image
                ? "IMAGE CREATED"
                : "NO IMAGE"
        );

        if (!image) {{

            setStatus(
                "❌ عکس ساخته نشد"
            );

            return;
        }}

        setStatus(
            "📷 عکس ساخته شد"
        );
        showCapturePreview(image);

        console.log(
            "📷 IMAGE DATA READY"
        );

    }};
        return button;
    }};

cameraControl.addTo(map);
// ============================================================
// Tile
// ============================================================

const tiles =
    L.tileLayer(
        "https://tile.openstreetmap.de/{{z}}/{{x}}/{{y}}.png",
        {{
            maxZoom: 19,
            attribution: "© OpenStreetMap",
            crossOrigin: true
        }}
    );


/* ============================================================
   Tile Diagnostic
   ============================================================ */

let tileLoadCount = 0;
let tileErrorCount = 0;

tiles.on(
    "tileloadstart",
    function(event) {{

        console.log(
            "TILE LOAD START:",
            event.coords
        );

    }}
);


tiles.on(
    "tileload",
    function(event) {{

        tileLoadCount += 1;

        console.log(
            "TILE LOAD OK:",
            tileLoadCount,
            event.coords
        );

        console.log(
            "Tile OK: " +
            tileLoadCount
        );

    }}
);


tiles.on(
    "tileerror",
    function(event) {{

        tileErrorCount += 1;

        const coords =
            event.coords || {{}};

        console.log(
            "================================"
        );

        console.log(
            "TILE ERROR #" +
            tileErrorCount
        );

        console.log(
            "Z:",
            coords.z
        );

        console.log(
            "X:",
            coords.x
        );

        console.log(
            "Y:",
            coords.y
        );

        console.log(
            "TILE:",
            event.tile
        );

        console.log(
            "URL:",
            event.tile
                ? event.tile.src
                : "NO TILE URL"
        );

        console.log(
            "================================"
        );

        console.log(
            "Tile Error: " +
            tileErrorCount
        );

    }}
);
tiles.addTo(map);


/* ============================================================
   گزارش وضعیت اولیه
   ============================================================ */
console.log(
    "TILE LAYER ADDED"
);

console.log(
    "MAP CENTER:",
    map.getCenter()
);

console.log(
    "MAP ZOOM:",
    map.getZoom()
);

console.log(
    "TILE ERROR COUNT:",
    tileErrorCount
);

console.log(
    "TILE LOAD COUNT:",
    tileLoadCount
);

// ============================================================
// لایه نقاط
// ============================================================

const markerLayer =

    L.layerGroup()
        .addTo(map);

// ============================================================
// QUARANTINE CIRCLES
// FF001 — اولین دایره
// ============================================================

let quarantineCircles = [];

let quarantineNextId = 1;
const quarantineColors = [

    "#e74c3c",
    "#3498db",
    "#2ecc71",
    "#9b59b6",
    "#f39c12",
    "#1abc9c",
    "#e67e22",
    "#34495e",
    "#d35400",
    "#16a085",

    "#c0392b",
    "#2980b9",
    "#27ae60",
    "#8e44ad",
    "#f1c40f",
    "#138d75",
    "#ca6f1e",
    "#566573",
    "#af601a",
    "#117864",

    "#cd6155",
    "#5dade2",
    "#58d68d",
    "#bb8fce",
    "#f5b041",
    "#48c9b0",
    "#eb984e",
    "#5d6d7e",
    "#dc7633",
    "#45b39d",

    "#ec7063",
    "#85c1e9",
    "#82e0aa",
    "#d2b4de",
    "#f7dc6f",
    "#76d7c4",
    "#f0b27a",
    "#85929e",
    "#e59866",
    "#73c6b6",

    "#922b21",
    "#1f618d",
    "#1e8449",
    "#6c3483",
    "#b7950b",
    "#0e6251",
    "#a04000",
    "#2e4053",
    "#a569bd",
    "#2874a6"
];
function toggleQuarantineVisibility(index) {{

    const circle =
        quarantineCircles[index];

    if (!circle) {{
        return;
    }}

    circle.visible =
        !circle.visible;

    if (circle.visible) {{

        circle.layer.addTo(map);

    }} else {{

        map.removeLayer(
            circle.layer
        );

    }}

    updateQuarantineList();
}}
function toggleQuarantineReport(index) {{

    const circle =
        quarantineCircles[index];

    if (!circle) {{
        return;
    }}

    circle.reportActive =
        !circle.reportActive;

    updateQuarantineList();
      
}}
/* =========================================================
   B2-22 — ENGINE REPORT DATA
   ========================================================= */

let betaFarmEngineReportData = {{

    reportId: null,

    circles: [],

    activeCircle: null,

    generatedAt: null

}};


function buildEngineReportData(circle) {{

    const target =
        circle || null;

    return {{

        reportId:
            target
                ? target.id
                : null,

        activeCircle:
            target
                ? {{

                    id:
                        target.id,

                    centerLat:
                        target.centerLat,

                    centerLon:
                        target.centerLon,

                    radiusKm:
                        target.radiusKm,

                    color:
                        target.color,

                    visible:
                        target.visible,

                    reportActive:
                        target.reportActive

                }}
                : null,

        circles:
            quarantineCircles.map(
                function(item) {{

                    return {{

                        id:
                            item.id,

                        centerLat:
                            item.centerLat,

                        centerLon:
                            item.centerLon,

                        radiusKm:
                            item.radiusKm,

                        color:
                            item.color,

                        visible:
                            item.visible,

                        reportActive:
                            item.reportActive

                    }};

                }}
            ),

        generatedAt:
            new Date().toLocaleString(
                "fa-IR"
            )

    }};

}}
/* =========================================================
   B2-22 — BUILD REPORT PAGE 4
   Dynamic Filter Data + Engine Data
   ========================================================= */

function buildReportPage4(circle) {{

    const engineData =
        buildEngineReportData(
            circle
        );

    const quarantineResult =
        window.BETA_FARM_QUARANTINE_RESULT ||
        null;

    let html = "";

    html +=
        '<div id="betaFarmReportPage4" ' +
        'style="' +
            'width:1200px;' +
            'min-height:850px;' +
            'box-sizing:border-box;' +
            'padding:35px;' +
            'background:white;' +
            'color:#222;' +
            'direction:rtl;' +
            'font-family:Tahoma,Arial,sans-serif;' +
        '">' +

        '<div style="' +
            'font-size:24px;' +
            'font-weight:bold;' +
            'margin-bottom:8px;' +
            'border-bottom:2px solid #333;' +
            'padding-bottom:10px;' +
        '">' +
            'گزارش نهایی قرنطینه' +
        '</div>' +

        '<div style="' +
            'font-size:12px;' +
            'color:#666;' +
            'margin-bottom:20px;' +
        '">' +
            'Beta Farm — Map Engine B2-22' +
        '</div>';


    

    // =====================================================
    // QUARANTINE CALCULATION RESULT
    // =====================================================

    html +=
        '<div style="' +
            'font-size:18px;' +
            'font-weight:bold;' +
            'margin:20px 0 10px;' +
        '">' +
            'نتیجه محاسبات قرنطینه' +
        '</div>';


    if (quarantineResult) {{

        html +=
            '<table style="' +
                'width:100%;' +
                'border-collapse:collapse;' +
                'font-size:14px;' +
                'margin-bottom:25px;' +
            '">' +


            '<tr>' +

            '<td style="' +
                'border:1px solid #999;' +
                'padding:10px;' +
                'font-weight:bold;' +
            '">' +

            'واحدهای طیور داخل دایره' +

            '</td>' +

            '<td style="' +
                'border:1px solid #999;' +
                'padding:10px;' +
                'font-weight:bold;' +
                'font-size:18px;' +
            '">' +

            String(
                quarantineResult.poultryCount ??
                0
            ) +

            '</td>' +

            '</tr>' +


            '<tr>' +

            '<td style="' +
                'border:1px solid #999;' +
                'padding:10px;' +
                'font-weight:bold;' +
            '">' +

            'شهرهای داخل دایره' +

            '</td>' +

            '<td style="' +
                'border:1px solid #999;' +
                'padding:10px;' +
                'font-weight:bold;' +
                'font-size:18px;' +
            '">' +

            String(
                quarantineResult.cityCount ??
                0
            ) +

            '</td>' +

            '</tr>' +


            '<tr>' +

            '<td style="' +
                'border:1px solid #999;' +
                'padding:10px;' +
                'font-weight:bold;' +
            '">' +

            'روستاهای داخل دایره' +

            '</td>' +

            '<td style="' +
                'border:1px solid #999;' +
                'padding:10px;' +
                'font-weight:bold;' +
                'font-size:18px;' +
            '">' +

            String(
                quarantineResult.villageCount ??
                0
            ) +

            '</td>' +

            '</tr>' +


            '</table>';


        html +=
            '<div style="' +
                'font-size:11px;' +
                'color:#666;' +
                'margin-bottom:15px;' +
            '">' +

            'شناسه نتیجه: ' +

            String(
                quarantineResult.circleId ||
                "-"
            ) +

            '</div>';

    }} else {{

        html +=
            '<div style="' +
                'border:1px solid #aaa;' +
                'padding:15px;' +
                'text-align:center;' +
                'color:#900;' +
                'margin-bottom:20px;' +
            '">' +

            'نتیجه محاسبات قرنطینه دریافت نشد.' +

            '</div>';

    }}


    // =====================================================
    // END
    // =====================================================

    html +=
        '</div>';

    return html;

}}
async function captureReportPage4(circle) {{

    try {{

        const page4HTML =
            buildReportPage4(
                circle
            );


        if (
            !page4HTML ||
            typeof page4HTML !== "string"
        ) {{

            console.error(
                "❌ PAGE 4 HTML NOT CREATED"
            );

            return null;
        }}


        const temp =
            document.createElement(
                "div"
            );


        temp.id =
            "betaFarmReportPage4Temp";


        temp.style.position =
            "fixed";


        temp.style.left =
            "0";


        temp.style.top =
            "0";


        temp.style.width =
            "1200px";


        temp.style.background =
            "#ffffff";


        temp.style.zIndex =
            "999999";


        temp.style.visibility =
            "visible";


        temp.style.opacity =
            "0";


        temp.innerHTML =
            page4HTML;


        document.body.appendChild(
            temp
        );


        await new Promise(
            function(resolve) {{

                setTimeout(
                    resolve,
                    300
                );

            }}
        );


        const reportPage =
            temp.querySelector(
                "#betaFarmReportPage4"
            );


        if (!reportPage) {{

            console.error(
                "❌ PAGE 4 ELEMENT NOT FOUND"
            );


            temp.remove();

            return null;
        }}


        const rect =
            reportPage.getBoundingClientRect();


        const captureWidth =
            Math.ceil(
                Math.max(
                    rect.width,
                    reportPage.offsetWidth,
                    1
                )
            );


        const captureHeight =
            Math.ceil(
                Math.max(
                    rect.height,
                    reportPage.offsetHeight,
                    1
                )
            );


        console.log(
            "📐 PAGE 4 SIZE:",
            captureWidth,
            captureHeight
        );


        const canvas =
            await html2canvas(
                reportPage,
                {{

                    backgroundColor:
                        "#ffffff",

                    useCORS:
                        true,

                    allowTaint:
                        false,

                    logging:
                        false,

                    scale:
                        1,

                    width:
                        captureWidth,

                    height:
                        captureHeight,

                    windowWidth:
                        captureWidth,

                    windowHeight:
                        captureHeight

                }}
            );


        if (
            !canvas ||
            canvas.width <= 0 ||
            canvas.height <= 0
        ) {{

            console.error(
                "❌ PAGE 4 CANVAS INVALID"
            );

            temp.remove();

            return null;
        }}


        const imageData =
            canvas.toDataURL(
                "image/png"
            );


        console.log(
            "🟢 PAGE 4 REAL IMAGE CREATED"
        );


        temp.remove();


        return imageData;


    }} catch (error) {{

        console.error(
            "❌ PAGE 4 CAPTURE ERROR:",
            error
        );


        const oldTemp =
            document.getElementById(
                "betaFarmReportPage4Temp"
            );


        if (oldTemp) {{

            oldTemp.remove();

        }}


        return null;

    }}
}}
/* =========================================================
   B2-22 — FOUR PAGE QUARANTINE REPORT
   PAGE 1 = User View
   PAGE 2 = Circle + 5 km
   PAGE 3 = Circle + 10 km
   PAGE 4 = Dynamic Filter Data + Engine Data
   ========================================================= */

async function handleQuarantineReportClick(index) {{

    const circle =
        quarantineCircles[index];

    if (!circle) {{
        return;
    }}
   let quarantineReportResult = null;

await new Promise(
    function(resolve) {{

        function waitForResult(event) {{

            if (
                event.data &&
                event.data.type ===
                    "BETA_FARM_QUARANTINE_RESULT" &&
                event.data.report &&
                String(
                    event.data.report.circleId
                ) ===
                String(circle.id)
            ) {{

                window.BETA_FARM_QUARANTINE_RESULT =
                    event.data.report;

                window.removeEventListener(
                    "message",
                    waitForResult
                );

                resolve();
            }}
        }}

        window.addEventListener(
            "message",
            waitForResult
        );

        window.parent.postMessage(
            {{
                type:
                    "BETA_FARM_QUARANTINE_UPDATE_REQUEST",

                circleId:
                    circle.id
            }},
            "*"
        );
    }}
); 
    // --------------------------------------------------------
    // TOGGLE REPORT
    // --------------------------------------------------------

    circle.reportActive =
        !circle.reportActive;

    updateQuarantineList();

    if (!circle.reportActive) {{
        return;
    }}


    // --------------------------------------------------------
    // حفظ نمای اصلی کاربر
    // --------------------------------------------------------

    const originalCenter =
        map.getCenter();

    const originalZoom =
        map.getZoom();


    // ========================================================
    // PAGE 1
    // نمای دقیق کاربر در لحظه کلیک
    // ========================================================

    setStatus(
        "📸 در حال ساخت صفحه ۱..."
    );

    const page1 =
        await captureCurrentMap();

    if (!page1) {{

        alert(
            "❌ صفحه ۱ ساخته نشد"
        );

        return;
    }}


    // ========================================================
    // PAGE 2
    // دایره + ۵ کیلومتر
    // ========================================================

    setStatus(
        "📸 در حال ساخت صفحه ۲..."
    );

    const totalRadiusKm =
        Number(circle.radiusKm) + 5;

    const bounds =
        L.latLng(
            circle.centerLat,
            circle.centerLon
        ).toBounds(
            totalRadiusKm * 2000
        );

    map.fitBounds(
        bounds,
        {{
            animate: false,
            padding: [20, 20]
        }}
    );


    await new Promise(
        function(resolve) {{
            setTimeout(
                resolve,
                500
            );
        }}
    );


    const page2 =
        await captureCurrentMap();

    if (!page2) {{

        map.setView(
            originalCenter,
            originalZoom,
            {{
                animate: false
            }}
        );

        alert(
            "❌ صفحه ۲ ساخته نشد"
        );

        return;
    }}


    // ========================================================
    // PAGE 3
    // دایره + ۱۰ کیلومتر
    // ========================================================

    setStatus(
        "📸 در حال ساخت صفحه ۳..."
    );

    const totalRadiusKm3 =
        Number(circle.radiusKm) + 10;

    const bounds3 =
        L.latLng(
            circle.centerLat,
            circle.centerLon
        ).toBounds(
            totalRadiusKm3 * 2000
        );

    map.fitBounds(
        bounds3,
        {{
            animate: false,
            padding: [20, 20]
        }}
    );


    await new Promise(
        function(resolve) {{
            setTimeout(
                resolve,
                500
            );
        }}
    );


    const page3 =
        await captureCurrentMap();

    if (!page3) {{

        map.setView(
            originalCenter,
            originalZoom,
            {{
                animate: false
            }}
        );

        alert(
            "❌ صفحه ۳ ساخته نشد"
        );

        return;
    }}


    // ========================================================
    // بازگرداندن نمای اصلی کاربر
    // ========================================================

    map.setView(
        originalCenter,
        originalZoom,
        {{
            animate: false
        }}
    );


    // ========================================================
    // PAGE 4
    // Dynamic Filter + Engine
    // ========================================================

    setStatus(
        "📄 در حال ساخت صفحه ۴..."
    );


    /*
     * Snapshot اطلاعات Engine
     */
    
    const engineData =
        buildEngineReportData(
            circle
        );
        
 

    /*
     * اطلاعات Dynamic Filter
     */

    const dynamicData =
        betaFarmDynamicReportData || {{

            version: null,
            count: 0,
            units: [],
            rows: []
        }};
   

    /*
     * ذخیره آخرین گزارش
     */

      // ========================================================
    // ساخت تصویر PAGE 4
    // Dynamic Filter + Engine
    // ========================================================

    setStatus(
        "📄 در حال تبدیل صفحه ۴ به تصویر..."
    );

    const page4 =
        await captureReportPage4(
            circle
        );
   
    setStatus(
    page4
        ? "🔔 TEST — صفحه ۴ ساخته شد"
        : "🔔 TEST — صفحه ۴ ساخته نشد"
    );
    setStatus(
    "🔔 TEST — از Page 4 عبور کردیم"
    );

    if (!page4) {{

        console.error(
            "❌ PAGE 4 IMAGE NOT CREATED"
        );

        setStatus(
            "⚠️ صفحه ۴ ساخته نشد"
        );

    }} else {{

        console.log(
            "📄 PAGE 4 IMAGE READY"
        );

    }}
 

    // ========================================================
    // ذخیره آخرین گزارش چهار صفحه‌ای
    // ========================================================

    window.BETA_FARM_LAST_REPORT = {{

        circle:
            engineData,

        dynamic:
            dynamicData,

        page1:
            page1,

        page2:
            page2,

        page3:
            page3,

        page4:
            page4

    }};

    setStatus(
     "🔔 TEST — از BETA_FARM_LAST_REPORT عبور کردیم"
    ); 
    // ========================================================
    // ذخیره Captureهای گزارش
    // ========================================================

    lastQuarantineReportCaptures = {{

        userView:
            page1,

        margin5km:
            page2,

        margin10km:
            page3,

        page4:
            page4

    }};
    setStatus(
     "🔔 TEST — از lastQuarantineReportCaptures عبور کردیم"
    ); 
    // ========================================================
// B2-22 — FINAL REPORT DATA
// بعد از ساخته‌شدن هر ۴ صفحه
// ========================================================

const finalReportData = {{

    dynamic:
        dynamicData,

    engine:
        engineData,

    page1:
        page1,

    page2:
        page2,

    page3:
        page3,

    page4:
        page4,

    rowId:
        circle.id,
    mapCenter: {{
        lat:
          originalCenter.lat,

        lon:
           originalCenter.lng
    }},

mapZoom:
    originalZoom,    

    generatedAt:
        new Date().toLocaleString(
            "fa-IR"
        )

}};
setStatus(
    "🔔 TEST — finalReportData ساخته شد"
);


// ========================================================
// ALARM — ارسال واقعی گزارش
// ========================================================

setStatus(
    "🔔 ALARM — در حال ارسال Report Data به Dynamic Filter"
);


// ========================================================
// ارسال فقط یک‌بار
// ========================================================

if (
    window.parent &&
    window.parent !== window
) {{
    setStatus(
      "🟢 TEST — Page 4 در Report Data: " +
      (
          finalReportData.page4
              ? "YES — " +
                String(
                  finalReportData.page4.length
                ) +
                " chars"
              :   "NO"
      )
    ); 
    window.parent.postMessage(
        {{

            type:
                "BETA_FARM_REPORT_DATA",

            report:
                finalReportData

        }},
        "*"
    );
    setStatus(
    "🟢 TEST — Page 4 در Report Data: " +
    (
        finalReportData.page4
            ? "YES — " +
              String(
                  finalReportData.page4.length
              ) +
              " chars"
            : "NO"
    )
);

    setStatus(
        "🔔 ALARM — Report Data ارسال شد"
    );
}}


// =========================================================
// ذخیره Captureهای گزارش
// =========================================================

lastQuarantineReportCaptures = {{

    userView:
        page1,

    margin5km:
        page2,

    margin10km:
        page3,

    page4:
        page4
}};


    // ========================================================
    // نمایش صفحه ۱ فعلاً
    // ========================================================

   
    setStatus(
        "✅ گزارش ۴ صفحه‌ای آماده شد"
    );


    console.log(
        "📋 BETA FARM FOUR PAGE REPORT READY",
        window.BETA_FARM_LAST_REPORT
    );

}}
function deleteQuarantineCircle(index) {{

    const circle =
        quarantineCircles[index];

    if (!circle) {{
        return;
    }}

    if (
        circle.layer &&
        map.hasLayer(circle.layer)
    ) {{
        map.removeLayer(
            circle.layer
        );
    }}

    console.log(
        "QUARANTINE CIRCLE DELETED:",
        circle
    );

    quarantineCircles.splice(
        index,
        1
    );

    updateQuarantineList();
}}
function selectQuarantineCircle(index) {{

    const selected =
        quarantineCircles[index];

    if (!selected) {{
        return;
    }}

    quarantineCircles.forEach(
        function(circle) {{

            if (
                circle.layer &&
                map.hasLayer(circle.layer)
            ) {{

                circle.layer.setStyle({{
                    weight: 1,
                    fillOpacity: 0.18
                }});

            }}

        }}
    );

    if (
        selected.layer &&
        map.hasLayer(selected.layer)
    ) {{

        selected.layer.setStyle({{
            weight: 4,
            fillOpacity: 0.30
        }});

        map.setView(
            [
                selected.centerLat,
                selected.centerLon
            ],
            Math.max(
                map.getZoom(),
                10
            )
        );
    }}

    updateQuarantineList();
}}
function updateQuarantineList() {{

    const list =
        document.getElementById(
            "quarantineList"
        );

    if (!list) {{
        return;
    }}

    if (quarantineCircles.length === 0) {{

        list.style.display = "none";
        list.innerHTML = "";

        return;
    }}

    list.style.display = "block";

    list.innerHTML =
        quarantineCircles
            .map(
                function(circle, index) {{

                    return (
                        '<div ' +
                            'style="' +
                                'white-space:nowrap;' +
                                'line-height:20px;' +
                                'cursor:pointer;' +
                                'padding:2px 3px;' +
                            '"' +

                            'onclick="selectQuarantineCircle(' +
                                index +
                            ')"' +

                        '>' +

                        (index + 1) +
                        "-<b>" +
                        circle.id +
                        "</b> / " +

                        Number(
                            circle.centerLon
                        ).toFixed(5) +

                        "*" +

                        Number(
                            circle.centerLat
                        ).toFixed(5) +

                        " / " +

                        circle.radiusKm +
                        "km " +

                        '<span ' +
                            'style="' +
                                'cursor:pointer;' +
                                'font-size:13px;' +
                                'color:' +
                                circle.color +
                                ';' +
                            '"' +

                            'onclick="event.stopPropagation();' +
                                'toggleQuarantineVisibility(' +
                                    index +
                                ')"' +
                        '>' +

                            (
                                circle.visible
                                    ? "👁"
                                    : "◉"
                            ) +

                        '</span>' +
                        '<span ' +
    'style="' +
    'cursor:pointer;' +
    'font-size:13px;' +
    'margin-right:6px;' +
    'padding:2px 4px;' +
    'border-radius:4px;' +
    'background:' +
    (
        circle.reportActive
            ? '#ffd54f'
            : 'transparent'
    ) +
';' +
'"' +

    'onclick="event.stopPropagation();' +
        'handleQuarantineReportClick(' +
    index +
')"' +
'>' +

    '📝' +

'</span>' +

                        '<span ' +
                            'style="' +
                                'cursor:pointer;' +
                                'font-size:13px;' +
                                'margin-right:8px;' +
                                'color:#555;' +
                            '"' +

                            'onclick="event.stopPropagation();' +
                                'deleteQuarantineCircle(' +
                                    index +
                                ')"' +
                        '>' +

                            '✕' +

                        '</span>' +

                        '</div>'
                    );

                }}
            )
            .join("<br>");
}}
function createQuarantineId() {{

    const number =
        String(
            quarantineNextId
        ).padStart(3, "0");

    quarantineNextId += 1;

    return "FF" + number;
}}


function drawQuarantineCircle(
    lat,
    lon,
    radiusKm
) {{

    const circleId =
        createQuarantineId();
   const circleColor =
    quarantineColors[
        quarantineCircles.length %
        quarantineColors.length
    ]; 
    const circle =
        L.circle(
            [lat, lon],
            {{
                radius:
                    radiusKm * 1000,

                color:
                    circleColor,

                weight:
                    1,

                fillColor:
                    circleColor ,

                fillOpacity:
                    0.18
            }}
        ).addTo(map);


    const circleData = {{

        id:
            circleId,

        centerLat:
            lat,

        centerLon:
            lon,

        radiusKm:
            radiusKm,

        color:
            circleColor,

        visible:
            true,
        reportActive: false,    

        layer:
            circle

    }};


    quarantineCircles.push(
        circleData
    );


    console.log(
        "QUARANTINE CIRCLE CREATED:",
        circleData
    );


    setStatus(
        circleId +
        " — شعاع " +
        radiusKm +
        " کیلومتر"
    );
    updateQuarantineList();
    const circleMessage = {{
    type: "BETA_FARM_QUARANTINE_CIRCLE_CREATED",
    circle: {{
        id: circleData.id,
        centerLat: circleData.centerLat,
        centerLon: circleData.centerLon,
        radiusKm: circleData.radiusKm,
        color: circleData.color,
        visible: circleData.visible
    }}
}};

// ============================================================
// QUARANTINE REPORT CAPTURE
// Image 3 = current user view
// Image 1 = circle + 5 km margin
// Image 2 = circle + 10 km margin
// ============================================================

let lastQuarantineReportCaptures = {{
    userView: null,
    margin5km: null,
    margin10km: null
}};


function waitForMapMovement() {{

    return new Promise(
        function(resolve) {{

            map.once(
                "moveend",
                function() {{
                    setTimeout(
                        resolve,
                        300
                    );
                }}
            );

        }}
    );
}}


async function captureReportView(
    centerLat,
    centerLon,
    totalRadiusKm
) {{

    const bounds =
        L.latLng(
            centerLat,
            centerLon
        ).toBounds(
            totalRadiusKm * 2000
        );

    map.fitBounds(
        bounds,
        {{
            animate: false,
            padding: [20, 20]
        }}
    );

    await waitForMapMovement();

    const image =
        await captureCurrentMap();

    return image;
}}


async function createQuarantineReportCaptures(
    circle
) {{

    if (!circle) {{
        return null;
    }}

    console.log(
        "📸 REPORT CAPTURE START:",
        circle.id
    );

    const originalCenter =
        map.getCenter();

    const originalZoom =
        map.getZoom();


    // --------------------------------------------------------
    // IMAGE 3
    // دقیقاً نمای فعلی کاربر
    // --------------------------------------------------------

    setStatus(
        "📸 گزارش — تصویر ۳ از ۳"
    );

    const userViewImage =
        await captureCurrentMap();


    // --------------------------------------------------------
    // IMAGE 1
    // دایره + ۵ کیلومتر حاشیه
    // --------------------------------------------------------

    setStatus(
        "📸 گزارش — تصویر ۱ از ۳"
    );

    const margin5Image =
        await captureReportView(
            circle.centerLat,
            circle.centerLon,
            Number(circle.radiusKm) + 5
        );


    // --------------------------------------------------------
    // IMAGE 2
    // دایره + ۱۰ کیلومتر حاشیه
    // --------------------------------------------------------

    setStatus(
        "📸 گزارش — تصویر ۲ از ۳"
    );

    const margin10Image =
        await captureReportView(
            circle.centerLat,
            circle.centerLon,
            Number(circle.radiusKm) + 10
        );


    // --------------------------------------------------------
    // RESTORE USER VIEW
    // --------------------------------------------------------

    map.setView(
        originalCenter,
        originalZoom,
        {{
            animate: false
        }}
    );

    await waitForMapMovement();


    lastQuarantineReportCaptures = {{
        userView:
            userViewImage,

        margin5km:
            margin5Image,

        margin10km:
            margin10Image
    }};


    console.log(
        "📸 REPORT CAPTURES READY:",
        {{
            circleId: circle.id,
            image1_5km:
                Boolean(margin5Image),
            image2_10km:
                Boolean(margin10Image),
            image3_userView:
                Boolean(userViewImage)
        }}
    );


    setStatus(
        "✅ هر ۳ تصویر گزارش ساخته شد"
    );


    return lastQuarantineReportCaptures;
}}
function requestQuarantineReport(index) {{

    return handleQuarantineReportClick(index);

}}
console.log(
    "🔥 SENDING CIRCLE MESSAGE TO PARENT:",
    circleMessage
);

window.parent.postMessage(
    circleMessage,
    "*"
);

window.top.postMessage(
    circleMessage,
    "*"
);
}}

// ============================================================
// QUARANTINE CONTEXT MENU — STEP 2
// ظاهر شبیه منوی راست‌کلیک ویندوز
// ============================================================

map.on(
    "contextmenu",
    function(event) {{

        const lat =
            event.latlng.lat;

        const lon =
            event.latlng.lng;

        const radiusOptions = [
            3,
            5,
            10,
            15,
            20,
            30,
            50
        ];

        const menu =
            L.popup({{
                closeButton: false,
                autoClose: true,
                closeOnClick: true,
                className: "quarantine-context-menu"
                            }})
            .setLatLng(event.latlng)
            .setContent(
                '<div style="' +
                    'font-family:Tahoma,Arial,sans-serif;' +
                    'direction:rtl;' +
                    'text-align:right;' +
                    'background:#ffffff;' +
                    'border:1px solid #bdbdbd;' +
                    'border-radius:3px;' +
                    'padding:3px;' +
                    'min-width:125px;' +
                    'box-shadow:0 2px 7px rgba(0,0,0,.22);' +
                '">' +

                '<div style="' +
                    'font-size:11px;' +
                    'font-weight:bold;' +
                    'padding:5px 8px;' +
                    'border-bottom:1px solid #eeeeee;' +
                    'color:#333;' +
                '">' +
                    'شعاع قرنطینه' +
                '</div>' +

                radiusOptions.map(function(radius) {{
                    return (
                        '<button ' +
                        'type="button" ' +
                        'style="' +
                            'display:block;' +
                            'width:100%;' +
                            'border:0;' +
                            'background:transparent;' +
                            'text-align:right;' +
                            'font-family:Tahoma,Arial,sans-serif;' +
                            'font-size:11px;' +
                            'padding:4px 8px;' +
                            'margin:0;' +
                            'cursor:pointer;' +
                            'border-radius:2px;' +
                        '"' +
                        ' onmouseover="this.style.background=\\'#eeeeee\\'"' +
                        ' onmouseout="this.style.background=\\'transparent\\'"' +
                        'onclick="drawQuarantineCircle(' +
                           lat +
                           ', ' +
                           lon +
                           ', ' +
                           radius +
                        ')"' +
                        '>' +
                            radius +
                            ' کیلومتر' +
                        '</button>'
                    );
                }}).join("") +

                '</div>'
            )
            .openOn(map);
    }}
);
// ============================================================
// SIMPLE DISTANCE RULER
// ============================================================

let measureStartPoint = null;
let measureStartMarker = null;
let measureLine = null;
let measureEndPoint = null;
let measureLabel = null;
let rulerEnabled = false;

function formatDistance(distance) {{
    if (distance < 1000) {{
        return Math.round(distance) + " متر";
    }}

    return (distance / 1000).toFixed(2) + " کیلومتر";
}}
map.on("click", function(event) {{

    if (!rulerEnabled) {{
        return;
    }}

    const point = event.latlng;

    // --------------------------------------------------------
    // First click = START POINT
    // --------------------------------------------------------
    if (!measureStartPoint) {{

        // پاک کردن اندازه‌گیری قبلی
        if (measureLine) {{
            map.removeLayer(measureLine);
            measureLine = null;
        }}

        if (measureStartMarker) {{
            map.removeLayer(measureStartMarker);
            measureStartMarker = null;
        }}

        if (measureEndPoint) {{
            map.removeLayer(measureEndPoint);
            measureEndPoint = null;
        }}

        if (measureLabel) {{
            map.removeLayer(measureLabel);
            measureLabel = null;
        }}

        // ثبت مختصات نقطه شروع
        measureStartPoint = point;

        // ساخت نشانگر نقطه شروع
        measureStartMarker =
            L.circleMarker(
                point,
                {{
                    radius: 5,
                    color: "#d35400",
                    weight: 2,
                    fillColor: "#f39c12",
                    fillOpacity: 1
                }}
            ).addTo(map);

        setStatus(
            "نقطه شروع انتخاب شد؛ نقطه پایان را انتخاب کنید."
        );

        return;
    }}

    // --------------------------------------------------------
    // Second click = END POINT
    // --------------------------------------------------------

    // اگر مقصد قبلی وجود دارد، حذف شود
    if (measureEndPoint) {{
        map.removeLayer(measureEndPoint);
        measureEndPoint = null;
    }}

    // ساخت نشانگر نقطه مقصد
    measureEndPoint =
        L.circleMarker(
            point,
            {{
                radius: 5,
                color: "#8e44ad",
                weight: 2,
                fillColor: "#9b59b6",
                fillOpacity: 1
            }}
        ).addTo(map);

    // حذف خط قبلی
    if (measureLine) {{
        map.removeLayer(measureLine);
        measureLine = null;
    }}

    // رسم خط اندازه‌گیری
    measureLine =
        L.polyline(
            [measureStartPoint, point],
            {{
                color: "#e67e22",
                weight: 3,
                dashArray: "8,6"
            }}
        ).addTo(map);

    // محاسبه فاصله
    const distance =
        map.distance(
            measureStartPoint,
            point
        );

    const distanceText =
        formatDistance(distance);

    // حذف برچسب قبلی
    if (measureLabel) {{
        map.removeLayer(measureLabel);
        measureLabel = null;
    }}

    // ساخت برچسب فاصله
    measureLabel =
        L.marker(
            point,
            {{
                icon:
                    L.divIcon(
                        {{
                            className:
                                "measure-label",

                            html:
                                '<div style="background:rgba(255,255,255,.95);border:1px solid #d35400;border-radius:5px;padding:4px 7px;font-family:Tahoma,Arial,sans-serif;font-size:11px;font-weight:bold;color:#7d3c0c;white-space:nowrap;">' +
                                distanceText +
                                "</div>",

                            iconSize:
                                [100, 28],

                            iconAnchor:
                                [50, 35]
                        }}
                    ),

                interactive: false
            }}
        ).addTo(map);

    setStatus(
        "مسافت: " +
        distanceText
    );

    // آماده شدن برای اندازه‌گیری بعدی
    measureStartPoint = null;
}});
// ------------------------------------------------------------
// Mouse movement = temporary ruler line
// ------------------------------------------------------------

map.on("mousemove", function(event) {{

    if (!rulerEnabled) {{
        return;
    }}

    if (!measureStartPoint) {{
        return;
    }}

    const point = event.latlng;

    if (measureLine) {{
        map.removeLayer(measureLine);
    }}

    measureLine =
        L.polyline(
            [measureStartPoint, point],
            {{
                color: "#e67e22",
                weight: 3,
                dashArray: "8,6"
            }}
        ).addTo(map);

    const distance =
        map.distance(
            measureStartPoint,
            point
        );

    setStatus(
        "مسافت: " +
        formatDistance(distance)
    );
}});
// ============================================================
// Marker
// Dynamic Filter مختصات آماده را می‌فرستد
// ============================================================
function createMarker(item) {{

    if (!item) {{
        return null;
    }}

    let lat;
    let lon;

    if (
        item.lat !== undefined &&
        item.lon !== undefined
    ) {{

        lat = Number(item.lat);
        lon = Number(item.lon);

    }} else if (
        item.X !== undefined &&
        item.Y !== undefined
    ) {{

        lat = Number(item.y);
        lon = Number(item.x);

    }} else {{

        return null;
    }}

    if (
        !Number.isFinite(lat) ||
        !Number.isFinite(lon)
    ) {{

        return null;
    }}

   return L.circleMarker(
    [
        lat,
        lon
    ],
    {{
        radius: 7,
        color: "#ffffff",
        weight: 2,
        fillColor: "#20a44b",
        fillOpacity: 1
    }}
); 
}}

// ============================================================
// نمایش نقاط
// ============================================================

function drawMarkers(data) {{

    markerLayer.clearLayers();

    mapData =
        Array.isArray(data)
            ? data
            : [];

    let validCount = 0;

    mapData.forEach(
        function(item) {{

            const marker =
                createMarker(item);

                        if (marker) {{
                       marker.on(
    "click",
    function() {{

        if (
            window.top &&
            window.top !== window
        ) {{

            window.top.postMessage(
                {{
                    type:
                        "BETA_FARM_REPORT_DATA",

                    report:
                        finalReportData
                }},
                "*"
        );

    setStatus(
        "🔔 ALARM — Report Data ارسال شد به Dynamic"
    );
}} 

    }}
);
                                         markerLayer.addLayer(
                    marker
                );

                validCount += 1;

            }}

        }}
    );


    // ========================================================
    // تست مرکز روی اولین نقطه
    // ========================================================

   

    setStatus(
        "دریافت: " +
        mapData.length +
        " | Marker: " +
        validCount
    );


    window.setTimeout(
        function() {{

            map.invalidateSize();

        }},
        100
    );

}}
map.on(
    "dragstart",
    function() {{

        if (
            window.parent &&
            window.parent !== window
        ) {{

            window.parent.postMessage(
                {{
                    type:
                        "BETA_FARM_MAP_DRAG_START"
                }},
                "*"
            );

        }}

    }}
);
// ============================================================
// Bridge
//
// Dynamic Filter:
// BETA_FARM_MAP_UPDATE
//
// map_data = payload
// ============================================================

let betaFarmDynamicReportData = {{
    version: null,
    count: 0,
    units: [],
    rows: []
}};


window.addEventListener(

    "message",

    function(event) {{

        if (!event.data) {{

            return;

        }}
        
/* =================================================
   REPORT CHAIN TEST
   دریافت ACK از Dynamic Filter
   ================================================= */


// ====================================================
// DYNAMIC → ENGINE — نتیجه تازه دایره قرنطینه
// ====================================================

if (
    event.data.type ===
    "BETA_FARM_QUARANTINE_RESULT"
) {{

    const result =
        event.data.report;

    if (!result) {{
        return;
    }}

    betaFarmDynamicReportData =
        {{

            circleId:
                result.circleId,

            poultryCount:
                Number(
                    result.poultryCount || 0
                ),

            cityCount:
                Number(
                    result.cityCount || 0
                ),

            villageCount:
                Number(
                    result.villageCount || 0
                ),

            version:
                betaFarmDynamicReportData.version,

            count:
                betaFarmDynamicReportData.count,

            units:
                Array.isArray(
                    betaFarmDynamicReportData.units
                )
                    ? betaFarmDynamicReportData.units
                    : [],

            rows:
                Array.isArray(
                    betaFarmDynamicReportData.rows
                )
                    ? betaFarmDynamicReportData.rows
                    : []

        }};

    console.log(
        "📥 QUARANTINE RESULT RECEIVED:",
        betaFarmDynamicReportData
    );

    return;
}}

        // ====================================================
        // VIEW MODE
        // ====================================================

        if (
            event.data.type ===
            "BETA_FARM_MAP_VIEW_MODE"
        ) {{

            setBetaFarmViewMode(
                event.data.mode
            );

            return;
        }}


        // ====================================================
        // RULER
        // ====================================================

        if (
            event.data.type ===
            "BETA_FARM_MAP_RULER_TOGGLE"
        ) {{

            const enabled =
                toggleRuler();

            if (
                window.parent !==
                window
            ) {{

                window.parent.postMessage(
                    {{
                        type:
                            "BETA_FARM_MAP_RULER_STATE",

                        enabled:
                            enabled
                    }},
                    "*"
                );

            }}

            return;

        }}


        // ====================================================
        // B2-22 — DYNAMIC FILTER → MAP ENGINE
        // ====================================================

        if (
            event.data.type !==
            "BETA_FARM_MAP_UPDATE"
        ) {{

            return;

        }}


        const incoming =
            event.data.map_data;


        // ====================================================
        // PAYLOAD جدید
        // ====================================================

        if (
            incoming &&
            !Array.isArray(incoming) &&
            typeof incoming === "object"
        ) {{

            betaFarmDynamicReportData = {{

                version:
                    incoming.version || null,

                count:
                    Number(
                        incoming.count || 0
                    ),

                units:
                    Array.isArray(
                        incoming.units
                    )
                        ? incoming.units
                        : [],

                rows:
                    Array.isArray(
                        incoming.rows
                    )
                        ? incoming.rows
                        : []

            }};


            // داده نقاط برای Map
            drawMarkers(
                betaFarmDynamicReportData.units
            );
            if (
               incoming.mapView &&
               incoming.mapView.center &&
               Number.isFinite(
                   Number(
                       incoming.mapView.center.lat
                   )      
               ) &&
               Number.isFinite(
                   Number(
                       incoming.mapView.center.lon
                   )
               ) &&
               Number.isFinite(
                   Number(
                        incoming.mapView.zoom
                    )
                )
           ) {

                map.setView(
                    [
                        Number(
                            incoming.mapView.center.lat
                        ),
                        Number(
                            incoming.mapView.center.lon
                        )
                    ],
                    Number(
                        incoming.mapView.zoom
                    ),
                    {
                        animate: false
                    }
                );
             }

            console.log(
                "📥 DYNAMIC REPORT RECEIVED",
                betaFarmDynamicReportData
            );


            return;

        }}


        // ====================================================
        // سازگاری با فرمت قدیمی
        // ====================================================

        if (
            Array.isArray(incoming)
        ) {{

            betaFarmDynamicReportData = {{

                version:
                    "LEGACY",

                count:
                    incoming.length,

                units:
                    incoming,

                rows:
                    []

            }};


            drawMarkers(
                incoming
            );


            return;

        }}

    }}

);
// ============================================================
// اعلام آماده بودن Map Engine
// ============================================================

window.addEventListener(

    "load",

    function() {{

        window.setTimeout(

            function() {{

                map.invalidateSize();


                if (

                    window.parent !==
                    window

                ) {{

                    window.parent.postMessage(

                        {{

                            type:
                                "BETA_FARM_MAP_READY",

                            version:
                                "B2-22"

                        }},

                        "*"

                    );

                }}

            }},

            200

        );

    }}

);

// ============================================================
// RULER ON / OFF
// ============================================================

function clearRuler() {{

    if (measureLine) {{
        map.removeLayer(measureLine);
        measureLine = null;
    }}

    if (measureStartMarker) {{
        map.removeLayer(measureStartMarker);
        measureStartMarker = null;
    }}
    
    if (measureEndPoint) {{
        map.removeLayer(measureEndPoint);
        measureEndPoint = null;
    }}

    if (measureLabel) {{
        map.removeLayer(measureLabel);
        measureLabel = null;
    }}

    measureStartPoint = null;

    setStatus("خط‌کش خاموش است.");
}}


function setRulerEnabled(enabled) {{

    rulerEnabled = Boolean(enabled);

    if (!rulerEnabled) {{
        clearRuler();
    }}

    if (rulerEnabled) {{
        setStatus("خط‌کش فعال است؛ نقطه شروع را انتخاب کنید.");
    }}
}}


function toggleRuler() {{

    setRulerEnabled(!rulerEnabled);

    return rulerEnabled;
}}
// ============================================================
// API حداقلی
// ============================================================

window.BetaFarmMap = {{

    setData:
        function(data) {{

            drawMarkers(
                data
            );

        }},


    clear:
        function() {{

            markerLayer.clearLayers();

            mapData = [];

            setStatus(
                "تعداد نقاط: 0"
            );

        }},
        
   toggleRuler: function() {{
        return toggleRuler();
        }}, 

    getMap:
        function() {{

            return map;

        }}

}};


// ============================================================
// اندازه اولیه
// ============================================================

map.invalidateSize();

</script>


</body>

</html>
'''



# ============================================================
# ساخت پوشه خروجی
# ============================================================

MAPS_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# ذخیره HTML
# ============================================================

OUTPUT_FILE.write_text(
    html,
    encoding="utf-8"
)


# ============================================================
# گزارش
# ============================================================

print("=" * 70)

print(
    "Beta Farm - Map Engine B2-22"
)

print(
    "Created successfully"
)

print()

print(
    "Output:"
)

print(
    OUTPUT_FILE
)

print()

print(
    "Initial center: Zanjan"
)

print(
    "Bridge: BETA_FARM_MAP_UPDATE"
)

print(
    "Data: payload.units"
)

print("=" * 70)
