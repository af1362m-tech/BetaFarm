# -*- coding: utf-8 -*-
"""
Beta Farm - Dynamic Filter
Version 0.5.1-B2-13

هدف این نسخه:
- تثبیت Layout
- نوار Navigation در بالا
- Filter = 25%
- Map = 75%
- Filter و Map هم‌ارتفاع
- Map فعلاً Fake / Placeholder
- بدون اتصال به Map Builder
- منطق فیلتر Dynamic حفظ شده

منبع داده:
Farm-Data.xlsx / Data1
"""

import sys
from pathlib import Path
import json
import re
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import threading
import runpy
import base64

from openpyxl import load_workbook
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import QUrl, QMarginsF
from PySide6.QtCore import QObject, Slot
from PySide6.QtWebEngineCore import QWebEngineSettings
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtPrintSupport import QPrinter, QPrintDialog
from PySide6.QtGui import QPageSize, QPageLayout

BASE_DIR = Path(__file__).resolve().parent
# ============================================================
# Zanjan Settlements Data
# ============================================================

SETTLEMENTS_FILE = (
    BASE_DIR /
    "Zanjan_Settlements_Clean.json"
)

SETTLEMENTS = []

if SETTLEMENTS_FILE.exists():

    try:

        with open(
            SETTLEMENTS_FILE,
            "r",
            encoding="utf-8-sig"
        ) as f:

            SETTLEMENTS = json.load(f)

        print(
            "Zanjan settlements loaded:",
            len(SETTLEMENTS)
        )

    except Exception as e:

        print(
            "ERROR loading settlements:",
            e
        )

else:

    print(
        "WARNING: Zanjan_Settlements_Clean.json not found."
    )
SHEET_NAME = "Data1"
# ============================================================
# Local Map Server
# ============================================================

class MapRequestHandler(SimpleHTTPRequestHandler):

    def log_message(self, format, *args):
        return


class LocalMapServer:

    def __init__(self, maps_dir):
        self.maps_dir = Path(maps_dir).resolve()

        handler = lambda *args, **kwargs: MapRequestHandler(
            *args,
            directory=str(self.maps_dir),
            **kwargs
        )

        self.server = ThreadingHTTPServer(
            ("127.0.0.1", 0),
            handler
        )

        self.port = self.server.server_address[1]

        self.thread = threading.Thread(
            target=self.server.serve_forever,
            daemon=True
        )

        self.thread.start()

    def url(self, filename):
        return (
            "http://127.0.0.1:"
            + str(self.port)
            + "/"
            + filename
        )

    def stop(self):
        if self.server:
            self.server.shutdown()
            self.server.server_close()


def find_excel():
    candidates = [
        BASE_DIR / "Farm-Data.xlsx",
        BASE_DIR.parent / "Farm-Data.xlsx",
    ]

    for path in candidates:
        if path.exists():
            return path

    return None


def _normalize_header_name(value):
    s = "" if value is None else str(value).strip()
    s = s.replace("ي", "ی").replace("ى", "ی").replace("ك", "ک")
    s = re.sub(r"[\u200c\u200d]", " ", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s
def _is_coordinate_header(value, axis):
    n = _normalize_header_name(value)

    if axis == "x":
        return n in {
            "x",
            "x coordinate",
            "مختصات x",
            "عرض جغرافیایی",
            "latitude",
            "lat",
            "ایکس"
        }

    return n in {
        "y",
        "y coordinate",
        "طول جغرافیایی",
        "longitude",
        "lon",
        "مختصات y",
        "ایگریگ",
        "ایگرگ",
        "وای"
    }

      

def _is_id_header(value):
    n = _normalize_header_name(value)
    return n in {"fm_id", "fm id", "fmid", "id", "کد واحد", "کد واحد اپیدمیولوژیک", "شناسه", "شناسه واحد"}

def _read_sheet_payload(excel_path, sheet_name, drop_leading_index=False):
    excel_path = Path(excel_path)
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"شیت {sheet_name} در {excel_path.name} پیدا نشد.")
    ws = wb[sheet_name]
    raw_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    headers = [str(v).strip() if v is not None else f"Column_{i}" for i, v in enumerate(raw_headers, 1)]
    start_col = 2 if drop_leading_index and ws.max_column > 1 else 1
    effective_headers = headers[start_col-1:]
    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        record = {}
        for idx, col in enumerate(range(start_col, ws.max_column + 1), 1):
            record[str(idx)] = row[col-1] if col-1 < len(row) else None
        rows.append(record)
    x_col = next((i for i, h in enumerate(effective_headers, 1) if _is_coordinate_header(h, "x")), None)
    y_col = next((i for i, h in enumerate(effective_headers, 1) if _is_coordinate_header(h, "y")), None)
    id_col = next((i for i, h in enumerate(effective_headers, 1) if _is_id_header(h)), None)
    if x_col is None and len(effective_headers) >= 2: x_col = len(effective_headers) - 1
    if y_col is None and len(effective_headers) >= 2: y_col = len(effective_headers)
    if id_col is None and effective_headers: id_col = 1
    excluded = {c for c in (x_col, y_col, id_col) if c is not None}
    filter_columns = [c for c in range(1, len(effective_headers)+1) if c not in excluded]
    filters = []
    for col in filter_columns:
        values, seen = [], set()
        for record in rows:
            value = record.get(str(col))
            if value is None: continue
            text = str(value).strip()
            if not text or text in seen: continue
            seen.add(text); values.append(text)
        filters.append({"column": col, "name": effective_headers[col-1], "values": values})
    wb.close()
    return {"headers": effective_headers, "filters": filters, "rows": rows, "row_count": len(rows), "column_count": len(effective_headers), "x_column": x_col, "y_column": y_col, "id_column": id_col}

def read_data(excel_path=None):
    if excel_path is None: excel_path = find_excel()
    if excel_path is None: raise FileNotFoundError("Farm-Data.xlsx پیدا نشد.")
    return _read_sheet_payload(excel_path, SHEET_NAME, drop_leading_index=False)

def read_data_from_sheet(excel_path, sheet_name):
    return _read_sheet_payload(excel_path, sheet_name, drop_leading_index=False)


class BetaFarmBridge(QObject):

    def __init__(self, load_callback, view_callback=None):
        super().__init__()
        self.load_callback = load_callback
        self._view_callback = view_callback
    @Slot(str)
    def saveMapCapture(self, image_data):


     try:

         if "," in image_data:
             image_data = image_data.split(",", 1)[1]

         import base64
         image_bytes = base64.b64decode(image_data)

         from PySide6.QtWidgets import QFileDialog

         file_path, _ = QFileDialog.getSaveFileName(
             None,
             "Save Map Picture",
             str(BASE_DIR / "MapPicture.png"),
             "PNG Image (*.png)"
         )

         if not file_path:
             print("📷 MAP SAVE CANCELLED")
             return ""

         Path(file_path).write_bytes(image_bytes)

         print(
             "📷 MAP IMAGE SAVED:",
            file_path
         )

         return str(file_path)

     except Exception as error:

         print(
             "❌ MAP IMAGE SAVE ERROR:",
             error
         )

         return ""

    
   
    @Slot(str)
    def saveReportCapture(self, image_data):

        try:

            if "," in image_data:
                image_data = image_data.split(",", 1)[1]

            import base64

            image_bytes = base64.b64decode(
                image_data
            )

            from PySide6.QtWidgets import QFileDialog

            file_path, _ = QFileDialog.getSaveFileName(
                None,
                "Save Report Picture",
                str(
                    BASE_DIR /
                    "BetaFarm_Report.png"
                ),
                "PNG Image (*.png)"
            )

            if not file_path:

                print(
                    "📄 REPORT SAVE CANCELLED"
                )

                return ""

            Path(
                file_path
            ).write_bytes(
                image_bytes
            )

            print(
                "📄 REPORT IMAGE SAVED:",
                file_path
            )

            return str(file_path)

        except Exception as error:

            print(
                "❌ REPORT IMAGE SAVE ERROR:",
                error
            )

            return ""
    @Slot()
    def saveReportPDF(self):

        try:

            from PySide6.QtWidgets import QFileDialog

            file_path, _ = QFileDialog.getSaveFileName(
                None,
                "Save Circle Report - PDF",
                str(
                    BASE_DIR /
                    "BetaFarm_Circle_Report.pdf"
                ),
                "PDF Files (*.pdf)"
            )

            if not file_path:
                print(
                    "📄 CIRCLE PDF SAVE CANCELLED"
                )
                return ""

            if not file_path.lower().endswith(".pdf"):
                file_path += ".pdf"

            if self._view_callback:
                self._view_callback(
                    "__SAVE_CIRCLE_PDF__:" + file_path
                )

            return str(file_path)

        except Exception as error:

            print(
                "❌ CIRCLE PDF SAVE ERROR:",
                error
            )

            return ""
    @Slot(str)
    def saveReportExcel(self, report_data):

        try:

            data = json.loads(
                report_data
            )

            headers = data.get(
                "headers",
                []
            )

            rows = data.get(
                "rows",
                []
            )

            from openpyxl import Workbook
            from PySide6.QtWidgets import QFileDialog

            workbook = Workbook()

            worksheet = workbook.active

            worksheet.title = "Filter Report"

            # -----------------------------------------
            # Header
            # -----------------------------------------

            for column_index, header in enumerate(
                headers,
                1
            ):

                worksheet.cell(
                    1,
                    column_index
                ).value = header


            # -----------------------------------------
            # Data
            # -----------------------------------------

            for row_index, row in enumerate(
                rows,
                2
            ):

                for column_index in range(
                    1,
                    len(headers) + 1
                ):

                    value = ""

                    if isinstance(
                        row,
                        dict
                    ):

                        value = row.get(
                            str(column_index),
                            ""
                        )

                    elif isinstance(
                        row,
                        list
                    ):

                        if (
                            column_index - 1
                        ) < len(row):

                            value = row[
                                column_index - 1
                            ]


                    worksheet.cell(
                        row_index,
                        column_index
                    ).value = value


            # -----------------------------------------
            # Save As
            # -----------------------------------------

            file_path, _ = QFileDialog.getSaveFileName(
                None,
                "Save Filter Report - Excel",
                str(
                    BASE_DIR /
                    "BetaFarm_Filter_Report.xlsx"
                ),
                "Excel Files (*.xlsx)"
            )


            if not file_path:

                print(
                    "📊 FILTER EXCEL SAVE CANCELLED"
                )

                return ""


            if not file_path.lower().endswith(
                ".xlsx"
            ):

                file_path += ".xlsx"


            workbook.save(
                file_path
            )


            print(
                "📊 FILTER EXCEL SAVED:",
                file_path
            )

            return str(
                file_path
            )


        except Exception as error:

            print(
                "❌ FILTER EXCEL SAVE ERROR:",
                error
            )

            return ""
    @Slot()
    def openFile(self):

        from PySide6.QtWidgets import QFileDialog

        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "Open File",
            str(BASE_DIR),
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if file_path:

            print(
                "Open File:",
                file_path
            )

            self.load_callback(file_path)

    @Slot()
    def newFile(self):
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "New File - Select Excel",
            str(BASE_DIR),
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            print("New File selected:", file_path)
            self.load_callback(str(file_path))

    @Slot(int)
    def activateFile(self, file_id):
        if self._view_callback:
            self._view_callback(f"__ACTIVATE_FILE__:{file_id}")

    @Slot(int)
    def removeFile(self, file_id):
        if self._view_callback:
            self._view_callback(f"__REMOVE_FILE__:{file_id}")

    @Slot()
    def printFile(self):
        if self._view_callback:
            self._view_callback("__PRINT__")

    @Slot()
    def exitApp(self):
        # خروج کامل از برنامه
        QApplication.quit()

    @Slot(str)
    def setMapMode(self, mode):
        print("View Mode:", mode)

        if self._view_callback:
            self._view_callback(mode)


def build_ui(data, map_engine_url):

    payload = json.dumps(
        data,
        ensure_ascii=False,
        default=str
    )
    settlements_payload = json.dumps(
    SETTLEMENTS,
    ensure_ascii=False,
    default=str
    )
        # =========================================================
    # BF.Paint — Fabric.js
    # =========================================================

    fabric_path = BASE_DIR / "BF_Paint" / "fabric.min.js"

    if not fabric_path.exists():
        raise FileNotFoundError(
            f"Fabric.js پیدا نشد:\n{fabric_path}"
        )

    fabric_js = fabric_path.read_text(
        encoding="utf-8"
    )

    # جلوگیری از بسته شدن زودهنگام <script>
    fabric_js = fabric_js.replace(
        "</script>",
        "<\\/script>"
    )
    html2canvas_path = (
        BASE_DIR /
        "BF_Paint" /
        "html2canvas.min.js"
    )

    if not html2canvas_path.exists():
        raise FileNotFoundError(
            f"html2canvas پیدا نشد:\n{html2canvas_path}"
        )

    html2canvas_js =(
        html2canvas_path.read_text(
            encoding="utf-8"
        )
    )
    html2canvas_js =(
        html2canvas_js.replace(
            "</script>",
            "<\\/script>"
        )
    )   
    return r"""<!doctype html>


<html lang="fa" dir="rtl">

<head>

<meta charset="utf-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1"
>

<title>Beta Farm - Dynamic Filter 0.5.1</title>
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<script>
__FABRIC_JS__
</script>

<script>
__HTML2CANVAS_JS__
</script>


<style>

/* =========================================================
   GLOBAL
   ========================================================= */

* {
    box-sizing: border-box;
}

html,
body {
    margin: 0;
    padding: 0;
    width: 100%;
    height: 100%;
}

body {

    background: #eef1f4;

    color: #222;

    font-family:
        Tahoma,
        Arial,
        sans-serif;

    overflow: hidden;
}
/* =========================================================
   UI B1 — NAVIGATION MENU
   ========================================================= */

.ui-menu-area {
    display: flex;
    align-items: center;
    gap: 0px;
    direction: ltr;
}

.ui-menu-btn {
    height: 32px;
    padding: 0 4px;

    border: 1px solid transparent;
    border-radius: 6px;

    background: transparent;
    color: #354656;

    font-family: inherit;
    font-size: 13px;

    cursor: pointer;
}

.ui-menu-btn:hover {
    background: #edf2f6;
    border-color: #d7e0e8;
}

.ui-brand {
    margin-left: auto;
    padding-left: 12px;

    font-size: 14px;
    font-weight: bold;

    color: #34495e;
}
/* =========================================================
   UI B1 — SUB MENUS
   ========================================================= */

.ui-menu-wrapper {
    position: relative;
}


.ui-submenu {

    display: none;

    position: absolute;

    top: 100%;
    left: 0;

    min-width: 150px;

    padding: 4px 0;

    background: #ffffff;

    border: 1px solid #d5dde5;

    border-radius: 7px;

    box-shadow:
        0 4px 12px
        rgba(0,0,0,.15);

    z-index: 1000;

    direction: ltr;
}


.ui-menu-wrapper.open .ui-submenu {
    display: block;
}
.ui-menu-wrapper:hover .ui-submenu {
    display: block;
}

.ui-submenu-item {

    display: block;

    width: 100%;

    height: 32px;

    padding: 0 12px;

    border: 0;

    background: transparent;

    color: #354656;

    font-family: inherit;

    font-size: 12px;

    text-align: left;

    cursor: pointer;
}


.ui-submenu-item:hover {

    background: #edf2f6;
}

/* =========================================================
   TOP NAVIGATION BAR
   ========================================================= */

.topbar {

    height: 46px;

    width: 100%;

    background: #ffffff;

    border-bottom:
        1px solid #d5dde5;

    display: flex;

    align-items: center;

    padding:
        0 12px;

    box-shadow:
        0 1px 5px
        rgba(0,0,0,.08);

    direction: ltr;

    flex-shrink: 0;
}


.topbar-left {

    display: flex;

    align-items: center;

    gap: 4px;
}


.nav-btn {

    height: 32px;

    padding:
        0 13px;

    border:
        1px solid transparent;

    border-radius: 6px;

    background: transparent;

    color: #354656;

    font-family: inherit;

    font-size: 13px;

    cursor: pointer;
}


.nav-btn:hover {

    background: #edf2f6;

    border-color:
        #d7e0e8;
}


.nav-brand {

    margin-left: auto;

    padding-left: 12px;

    font-size: 14px;

    font-weight: bold;

    color: #34495e;
}


/* =========================================================
   MAIN PAGE
   ========================================================= */

.page {
    width: 100%;
    height: calc(100vh - 46px);

    padding: 12px;

    display: flex;
    flex-direction: column;
    gap: 10px;

    overflow: hidden;
}

/* =========================================================
   WORKSPACE
   ========================================================= */
.workspace {
    flex: 1;
    min-height: 0;

    display: grid;
    grid-template-columns: 1fr 3fr;
    gap: 10px;
    align-items: stretch;
}

/* =========================================================
   FILTER AREA
   ========================================================= */

.filter-area {

    width: 100%;

    height: 100%;

    min-height: 0;

    background:
        #f3f7fb;

    border:
        1px solid #d7e1ec;

    border-radius: 10px;

    padding: 10px;

    display: flex;

    flex-direction: column;

    overflow: hidden;
}


.filter-heading {

    font-weight: bold;

    font-size: 14px;

    margin:
        2px 4px 8px;

    color: #34495e;

    flex-shrink: 0;
}


.panel {

    background: #ffffff;

    border-radius: 9px;

    flex: 1 1 auto;

    min-height: 0;

    overflow-y: auto;

    padding:
        4px 10px;

    box-shadow:
        0 1px 5px
        rgba(0,0,0,.05);
}


.panel::-webkit-scrollbar {

    width: 8px;
}


.panel::-webkit-scrollbar-thumb {

    background: #aeb8c2;

    border-radius: 8px;
}


.filter-card {

    padding:
        9px 4px;

    border-bottom:
        1px solid #e8e8e8;
}


.filter-title {

    font-weight: bold;

    font-size: 13px;

    margin-bottom: 6px;

    color: #3d4f60;
}


select {

    width: 100%;

    height: 34px;

    border:
        1px solid #c8c8c8;

    border-radius: 6px;

    background: #ffffff;

    padding:
        0 8px;

    font-size: 13px;

    font-family: inherit;
}


.actions {

    margin-top: 8px;

    background:
        #ffffff;

    border-radius: 8px;

    padding:
        8px;

    box-shadow:
        0 1px 5px
        rgba(0,0,0,.05);

    display: flex;

    align-items: center;

    gap: 8px;

    flex: 0 0 auto;
    min-height: 50px;
}


#applyBtn {

    width: 100%;

    border: 0;

    border-radius: 7px;

    padding:
        8px 14px;

    font-family: inherit;

    font-size: 13px;

    cursor: pointer;

    background: #34495e;

    color: #ffffff;
}


#applyBtn:hover {

    opacity: .88;
}
.filter-result-box {
    margin-top: 8px;
    background: #e9f7ed;
    border: 1px solid #9bd3a8;
    border-radius: 8px;
    padding: 8px 10px;
    color: #176b2d;
    font-size: 12px;
    line-height: 1.7;
    flex: 0 0 auto;
    max-height: 100px;
    overflow-y: auto;
}

.filter-result-title {
    font-weight: bold;
    margin-bottom: 2px;
}

.filter-result-summary {
    color: #245b31;
}

/* =========================================================
   WORKSPACE AREA
   ========================================================= */

.workspace-area {

    min-width: 0;

    height: 100%;

    display: flex;
   
    flex-direction: column;

    gap: 10px;
    
}


/* =========================================================
   MAP SHELL
   ========================================================= */

.map-shell {

    position: relative;

    width: 100%;

    height: 100%;
     

    min-height: 0;

    border:
        1px solid #17452d;

    border-radius: 10px;

    overflow: hidden;

    background:
        #155c37;

    box-shadow:
        0 2px 8px
        rgba(0,0,0,.10);

    transition:
        height .25s ease;
}


/* =========================================================
   FAKE MAP
   ========================================================= */

.fake-map {

    position: absolute;

    inset: 0;

    overflow: hidden;

    background:
        #155c37;
}


/* subtle land areas */

.land {

    position: absolute;

    background:
        #1e7044;

    border:
        1px solid
        rgba(145,190,150,.22);

    opacity: .85;
}


.land-1 {

    width: 42%;

    height: 38%;

    top: 8%;

    right: 8%;

    border-radius:
        52% 35% 60% 42%;
}


.land-2 {

    width: 34%;

    height: 46%;

    bottom: 4%;

    left: 7%;

    border-radius:
        38% 58% 44% 62%;
}


.land-3 {

    width: 25%;

    height: 28%;

    top: 46%;

    right: 36%;

    border-radius:
        60% 35% 48% 55%;
}


/* fake roads */

.road {

    position: absolute;

    height: 1px;

    background:
        rgba(205,225,195,.25);

    transform-origin:
        left center;
}


.road-1 {

    width: 80%;

    top: 32%;

    left: 8%;

    transform:
        rotate(-12deg);
}


.road-2 {

    width: 75%;

    top: 60%;

    left: 13%;

    transform:
        rotate(18deg);
}


.road-3 {

    width: 55%;

    top: 20%;

    left: 35%;

    transform:
        rotate(65deg);
}


.road-4 {

    width: 58%;

    top: 70%;

    left: 25%;

    transform:
        rotate(-55deg);
}


/* map markers */

.marker {

    position: absolute;

    width: 12px;

    height: 12px;

    border-radius: 50%;

    background:
        #ffffff;

    border:
        3px solid #d64b3b;

    box-shadow:
        0 1px 5px
        rgba(0,0,0,.35);
}


.marker-1 {
    top: 28%;
    right: 27%;
}

.marker-2 {
    top: 48%;
    right: 55%;
}

.marker-3 {
    top: 64%;
    right: 18%;
}

.marker-4 {
    top: 72%;
    right: 43%;
}

.marker-5 {
    top: 38%;
    right: 74%;
}


/* map label */

.fake-map-title {

    position: absolute;

    top: 12px;

    right: 14px;

    padding:
        6px 10px;

    border-radius: 6px;

    background:
        rgba(0,0,0,.30);

    color:
        rgba(255,255,255,.92);

    font-size: 12px;

    z-index: 3;
}


/* =========================================================
   MAP CONTROLS
   ========================================================= */

.map-controls {

    position: absolute;

    top: 12px;

    left: 12px;

    display: flex;

    gap: 6px;

    z-index: 5;
}


.map-control-btn,
.restore-result-btn {

    border:
        1px solid #c4d0db;

    background:
        rgba(255,255,255,.94);

    color: #304354;

    border-radius: 7px;

    min-width: 32px;

    height: 32px;

    cursor: pointer;

    font-weight: bold;

    box-shadow:
        0 2px 7px
        rgba(0,0,0,.10);
}


/* =========================================================
   REPORT TABS — STAGE 1
   ========================================================= */

.filter-report-tabs {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    width: 100%;
    padding: 6px 0;
}

.filter-report-tab {
    display: flex;
    align-items: center;
    gap: 3px;    
   

    min-height: 28px;
    padding: 1px 5px 1px 12px;

    border: 1px solid #c4d0db;
    border-radius: 8px;

    background: #f4f7fa;
    box-sizing: border-box;
}

.filter-report-tab-title {
    border: none;
    background: transparent;

    color: #304354;
    font-family: Tahoma, sans-serif;
    font-size: 11px;
    font-weight: bold;

    white-space: nowrap;
}

.filter-report-tab-control {
    width: 20px;
    height: 20px;

    padding: 0;

    border: 1px solid #c4d0db;
    border-radius: 4px;

    background: #ffffff;
    color: #304354;

    font-size: 11px;
    font-weight: bold;

    cursor: default;
}
.betaFarmReportHeader,
#betaFarmReportHeader,
#tabGuardianLayer4 {

    /* فقط برای محدودهٔ Header */
}


#betaFarmReportHeader .filter-report-tab-control:not(:disabled),
#tabGuardianLayer4 .filter-report-tab-control:not(:disabled) {

    border: 2px solid #2A9D8F !important;

    background-color: #D8F3EE !important;

}


#betaFarmReportHeader .filter-report-tab-control:disabled,
#tabGuardianLayer4 .filter-report-tab-control:disabled {

    border: 1px solid #c4d0db !important;

    background-color: #ffffff !important;

}
/* =========================================================
   RESULT AREA
   ========================================================= */

.result-area {

    display: none;

    width: 100%;
    flex: 0 0 auto;
}
.result-area .result-box {
    display: none !important;
}

.result-area.result-visible {

    display: block;
}

/* =========================================================
   REPORT TABS
   ========================================================= */

.filter-report-tabs {

    display: flex;

    gap: 5px;

    margin-bottom: 8px;

    flex-wrap: wrap;
}


.filter-report-tab {
    display: flex;
    align-items: center
    border:
        1px solid #c4d0db;

    background:
        #ffffff;

    color:
        #304354;

    border-radius:
        7px 7px 0 0;

    padding:
        2px 6px;

    cursor:
        pointer;

    font-size:
        12px;

    font-weight:
        bold;

    box-shadow:
        0 1px 3px
        rgba(0,0,0,.08);
}


.filter-report-tab:hover {

    background:
        #f1f3f5;
}


.filter-report-tab.active {

    background:
        #d9dee3;

    color:
        #263746;

    border-color:
        #aeb8c1;

    box-shadow:
        none;
}
.filter-report-tab.maximized {
    position: fixed;
    left: 20px;
    right: 20px;
    top: 20px;
    bottom: 20px;

    width: auto;
    height: auto;

    z-index: 99999;
}

.result-box {

    background:
        #e9f7ed;

    border:
        1px solid #9bd3a8;

    border-radius: 10px;

    padding:
        10px 14px;

    color: #176b2d;

    font-size: 13px;

    line-height: 1.8;
}


.result-title {

    font-weight: bold;

    margin-bottom: 2px;
}


.result-summary {

    color: #245b31;
}


.result-table-box {

    margin-top: 8px;

    background:
        #fffaf0;

    border:
        1px solid #ead9b6;

    border-radius: 10px;

    padding:
        10px 14px;

    color: #4a3a22;
}


.result-count {

    font-size: 12px;

    margin:
        3px 0 7px;

    color: #6b5a3d;
}


.result-table {

    max-height: 220px;

    overflow: auto;

    font-size: 12px;
}


.result-table table {

    width: max-content;

    min-width: 100%;

    border-collapse: collapse;
}


.result-table th,
.result-table td {

    border-bottom:
        1px solid #eadfce;

    padding:
        6px;

    text-align: right;

    white-space: nowrap;
}


.result-table th {

    background:
        #f5ead7;

    font-weight: bold;
}


.result-table td {

    color: #4b3d2b;
}


/* =========================================================
   FOOTER
   ========================================================= */

.footer {

    margin-top: auto;

    background:
        #ffffff;

    border-radius: 8px;

    padding:
        7px;

    text-align: center;

    color: #777;

    font-size: 11px;

    box-shadow:
        0 1px 5px
        rgba(0,0,0,.04);

    flex-shrink: 0;
}


/* =========================================================
   MAP VIEW STATES
   ========================================================= */

/* حالت 1، 2 و 3:
   کادر پایین حذف می‌شود و نقشه فضای عمودی آن را می‌گیرد */
.workspace.map-view-1 .workspace-area,
.workspace.map-view-2 .workspace-area,
.workspace.map-view-3 .workspace-area {
    min-height: 0;
}

.workspace.map-view-1 .result-area,
.workspace.map-view-2 .result-area,
.workspace.map-view-3 .result-area {
    display: none !important;
}

.workspace.map-view-1 .map-shell,
.workspace.map-view-2 .map-shell,
.workspace.map-view-3 .map-shell {
    flex: 1 1 auto;
    min-height: 0;
    height: auto;
}

/* حالت 2:
   ستون فیلتر هم حذف می‌شود */
.workspace.map-view-2 {
    grid-template-columns: 1fr;
}

.workspace.map-view-2 .filter-area {
    display: none !important;
}

/* حالت 3:
   فیلتر برمی‌گردد ولی کادر پایین هنوز مخفی است */
.workspace.map-view-3 {
    grid-template-columns: 1fr 3fr;
}

.workspace.map-view-3 .filter-area {
    display: flex;
}
/* =========================================================
   RESPONSIVE
   ========================================================= */

@media (max-width: 900px) {

    body {
        overflow: auto;
    }

    .page {
        height: auto;
        overflow: visible;
    }

    .workspace {

        height: auto;

        grid-template-columns:
            1fr;
    }

    .filter-area {

        height: 420px;
    }

    .map-shell {

        height: 520px;
    }
}

</style>

</head>


<body>


<nav class="topbar"></nav>

<div id="fileCacheBar" style="display:flex;align-items:center;gap:6px;overflow-x:auto;overflow-y:hidden;white-space:nowrap;padding:6px 8px;margin:0 8px 6px 8px;min-height:38px;box-sizing:border-box;"></div>

<!-- =======================================================
     MAIN PAGE
     ======================================================= -->

<div class="page">


    <div
        class="workspace"
        id="workspace">


        <!-- =================================================
             FILTER
             ================================================= -->

        <section
            class="filter-area">

            <div
                class="filter-heading">

                فیلترهای اطلاعات

            </div>


            <div
                id="filterPanel"
                class="panel">
            </div>


            <div
                class="actions">
                

                <button
                    id="applyBtn"
                    type="button">

                    اعمال فیلترها

                </button>

            </div>
            

               
        </section>


        <!-- =================================================
             WORKSPACE / MAP
             ================================================= -->

        <section
            class="workspace-area">


            <div
                class="map-shell"
                id="mapShell">


                <!-- FAKE MAP -->

               <!-- =========================================================
     REAL MAP BUILDER 2.5
     ========================================================= -->

<div id="mapFramesContainer" style="position:absolute;inset:0;width:100%;height:100%;"></div>

                <!-- MAP CONTROLS -->

                <div
                    class="map-controls">

                    <button
                        id="expandMapBtn"
                        class="map-control-btn"
                        type="button"
                        title="بزرگ‌نمایی / بازگشت نقشه">

                        ⛶

                    </button>

                </div>


            </div>


            <!-- =================================================
                 RESULT AREA
                 ================================================= -->

            <section
                class="result-area"
                id="resultArea">
                <div
                    id="filterReportTabs"
                    class="filter-report-tabs">
                </div> 

                                 
                <div
                    id="resultTableBox"
                    class="result-table-box"
                    style="display:none;">

                    <div
                        class="result-title">

                        نتیجه فیلتر

                    </div>


                    <div
                        id="resultCount"
                        class="result-count">
                    </div>


                    <div
                        id="resultTable"
                        class="result-table">
                    </div>

                </div>


            </section>

        </section>

    </div>


    <div
        id="footer"
        class="footer">
    </div>


</div>


<script>


/* =========================================================
   DATA
   ========================================================= */

let DATA = __DATA__;
const SETTLEMENTS = __SETTLEMENTS__;
let ACTIVE_FILE_ID = 0;
/* =========================================================
   PYTHON BRIDGE
   ========================================================= */

new QWebChannel(
    qt.webChannelTransport,
    function(channel) {

        window.betaFarmBridge =
            channel.objects.betaFarmBridge;

        console.log(
            "Beta Farm Bridge: READY"
        );

    }
);
const APP_STATE = {
    mode: "INITIAL",
    selections: {},
    filteredRows: [],
    selectedUnit: null
};/* =========================================================
   BF.Paint — Fabric Engine Check
   ========================================================= */

window.BF_PAINT_ENGINE = null;

if (
    typeof fabric !== "undefined"
) {

    window.BF_PAINT_ENGINE =
        fabric;

    console.log(
        "✅ BF.Paint — Fabric.js READY"
    );

} else {

    console.error(
        "❌ BF.Paint — Fabric.js NOT LOADED"
    );

}

const FILE_STATES = {};
/* =========================================================
   MULTI FILE / MULTI MAP CACHE
   ========================================================= */
let FILE_CACHE = [];
let MAP_FRAME = null;

function ensureMapFrame(visible) {
    const container = document.getElementById("mapFramesContainer");
    if (!container) return null;
    if (!MAP_FRAME) {
        MAP_FRAME = document.createElement("iframe");
        console.log("🔥 MAP FRAME CREATED");
        console.log("🔥 MAP FRAME parent:", MAP_FRAME.parentElement);
        console.log("🔥 MAP FRAME contentWindow:", MAP_FRAME.contentWindow);

        MAP_FRAME.addEventListener("load", function () {
            console.log("🔥 MAP FRAME LOADED");
            console.log(
                "🔥 iframe contentWindow === iframe.contentWindow:",
                MAP_FRAME.contentWindow === this.contentWindow
            );
        });
        MAP_FRAME.id = "betaFarmMapFrame";
        MAP_FRAME.src = "__MAP_ENGINE_URL__";
        MAP_FRAME.style.cssText = "position:absolute;inset:0;width:100%;height:100%;border:0;border-radius:10px;";
        MAP_FRAME.style.display = visible ? "block" : "none";
        MAP_FRAME.setAttribute("loading", "eager");
        container.appendChild(MAP_FRAME);
        


       MAP_FRAME.addEventListener("load", function() {

    console.log("🔥🔥 MAP ENGINE LOADED");

    console.log(
        "🔥 iframe contentWindow:",
        MAP_FRAME.contentWindow
    );

    console.log(
        "🔥 iframe parent window:",
        window
    );

    console.log(
        "🔥 iframe src:",
        MAP_FRAME.src
    );

    sendFilteredDataToMap();

}); 
    } else {
        MAP_FRAME.style.display = visible ? "block" : "none";
    }
    return MAP_FRAME;
}

function renderFileCache(items, activeId) {
    FILE_CACHE = items || [];
    const bar = document.getElementById("fileCacheBar");
    if (!bar) return;
    bar.innerHTML = "";
    FILE_CACHE.forEach(function(item) {
        const wrap = document.createElement("div");
        wrap.style.cssText = "display:flex;align-items:center;gap:2px;flex:0 0 auto;border:1px solid #cbd5df;border-radius:6px;background:#eef1f4;padding:2px;";
        const btn = document.createElement("button");
        btn.type = "button"; btn.textContent = item.name; btn.title = item.path + "\n" + item.data_sheet;
        btn.style.cssText = String(item.id) === String(activeId) ? "background:#34495e;color:white;border:0;border-radius:5px;padding:6px 10px;cursor:pointer;" : "background:transparent;color:#354656;border:0;border-radius:5px;padding:6px 10px;cursor:pointer;";
        btn.onclick = function() { if (window.betaFarmBridge && typeof window.betaFarmBridge.activateFile === "function") window.betaFarmBridge.activateFile(item.id); };
        const close = document.createElement("button");
        close.type = "button"; close.textContent = "×"; close.title = "Close";
        close.style.cssText = "background:transparent;color:#354656;border:0;padding:4px 7px;cursor:pointer;font-size:16px;";
        close.onclick = function(event) { event.stopPropagation(); if (window.betaFarmBridge && typeof window.betaFarmBridge.removeFile === "function") window.betaFarmBridge.removeFile(item.id); };
        wrap.appendChild(btn); wrap.appendChild(close); bar.appendChild(wrap);
    });
    if (bar.scrollWidth > bar.clientWidth) bar.scrollLeft = Math.min(bar.scrollLeft, bar.scrollWidth - bar.clientWidth);
}

window.setDynamicData = function(newData, fileId) {
    DATA = newData;
    ACTIVE_FILE_ID = Number(fileId || 0);
    const savedState = FILE_STATES[ACTIVE_FILE_ID];

    if (savedState) {
        APP_STATE.mode = savedState.mode;
        APP_STATE.selections = { ...savedState.selections };
    } else {
        APP_STATE.mode = "INITIAL";
        APP_STATE.selections = {};
    }

    APP_STATE.filteredRows = [];
    APP_STATE.selectedUnit = null;
    createFilters();
    Object.keys(APP_STATE.selections).forEach(function(column) {
        const select = filterPanel.querySelector(
            'select[data-column="' + column + '"]'
        );

        if (select) {
            select.value = APP_STATE.selections[column];
        }
    });
    if (FILE_STATES[ACTIVE_FILE_ID]) {
    applyBtn.click();
    }
    updateMap();
    sendFilteredDataToMap();
    footer.textContent = "Beta Farm Dynamic Filter 0.5.1-B2-19" + "  |  ستون‌ها: " + DATA.column_count + "  |  ردیف‌ها: " + DATA.row_count + "  |  فیلترها: " + DATA.filters.length + "  |  Map Units: " + getMapRowCount();
};

window.setFileCache = function(items, activeId) { renderFileCache(items, activeId); };

const fileCacheBar = document.getElementById("fileCacheBar");
if (fileCacheBar) {
    fileCacheBar.addEventListener("wheel", function(event) {
        if (Math.abs(event.deltaY) > Math.abs(event.deltaX)) { event.preventDefault(); fileCacheBar.scrollLeft += event.deltaY; }
    }, {passive:false});
}

window.BETA_FARM_VIEW_MODE = "DEFAULT";
function getMapRows() {
    return APP_STATE.filteredRows;
}

function setBetaFarmViewMode(mode) {

    window.BETA_FARM_VIEW_MODE = mode;

    console.log("B2-22 VIEW MODE:", mode);

    if (MAP_FRAME && MAP_FRAME.contentWindow) {

        MAP_FRAME.contentWindow.postMessage(
            {
                type: "BETA_FARM_MAP_VIEW_MODE",
                mode: mode
            },
            "*"
        );

        console.log(
            "B2-22 VIEW MODE SENT TO MAP:",
            mode
        );
    }

    if (mode === "ONLINE") {

        console.log(
            "Google Maps: remote mode ONLINE"
        );

    } else if (mode === "OFFLINE") {

        console.log(
            "Offline Map: remote mode OFFLINE"
        );

    } else {

        console.log(
            "View: Default"
        );
    }
}
function getMapRowCount() {
    return getMapRows().length;
}
function normalizeValue(value) {
    if (value === undefined || value === null) {
        return "";
    }

    let s = String(value).trim();

    s = s.replace(/[يى]/g, "ی")
         .replace(/ك/g, "ک")
         .replace(/ۀ/g, "ه")
         .replace(/[\u200c\u200d]/g, " ")
         .replace(/\s+/g, " ");

    return s;
}

function normalizeRow(row) {
    const clean = {};

    Object.keys(row).forEach(function(key) {
        clean[key] = normalizeValue(row[key]);
    });

    return clean;
}

function normalizeRows(rows) {
    return rows.map(normalizeRow);
}


function getMapRows() {

    if (
        ACTIVE_REPORT_ID &&
        REPORT_CACHE[ACTIVE_REPORT_ID]
    ) {
        return REPORT_CACHE[
            ACTIVE_REPORT_ID
        ].DATA.rows || [];
    }

    return [];
}
function getMapRowCount() {
    return getMapRows().length;
}

function getMapUnits() {
    const rows = normalizeRows(getMapRows());
    const xCol = Number(DATA.x_column || 0);
    const yCol = Number(DATA.y_column || 0);
    const idCol = Number(DATA.id_column || 1);
    return rows.map(function(row, index) {
        return {
            FM_ID: row[String(idCol)] || ("Unit " + (index + 1)),
            X: row[String(xCol)],
            Y: row[String(yCol)]
        };
    }).filter(function(unit) {
        return unit.X !== undefined && unit.X !== null && unit.Y !== undefined && unit.Y !== null && String(unit.X).trim() !== "" && String(unit.Y).trim() !== "";
    });
}

/* =========================================================
   B2-21 — CALIBRATION + MAP BRIDGE
   STEP 2
   ========================================================= */

function calibrateMapUnits(units) {

    return units.map(function(unit) {

        const x = Number(unit.X);
        const y = Number(unit.Y);

        let status = "VALID";

        if (
            !Number.isFinite(x) ||
            !Number.isFinite(y)
        ) {
            status = "INVALID";
        }

        return {
            FM_ID: unit.FM_ID || null,

            X: Number.isFinite(x) ? x : null,
            Y: Number.isFinite(y) ? y : null,

            // Map Builder 2.5 convention
           lat: Number.isFinite(x) ? y : null,
           lon: Number.isFinite(y) ? x : null,

            coordinate_status: status
        };

    });
}


/* =========================================================
   ارسال داده کالیبره‌شده به Map
   ========================================================= */

function sendCalibratedMapData() {

    const units = getMapUnits();

    const calibrated =
        calibrateMapUnits(units);

    window.BETA_FARM_CALIBRATED_MAP =
        calibrated;

    window.dispatchEvent(
        new CustomEvent(
            "betaFarmCalibratedMap",
            {
                detail: {
                    version: "B2-21-Calibration-v1",
                    count: calibrated.length,
                    map_data: calibrated
                }
            }
        )
    );

    return calibrated;
}
/* B2-20 — Safe Map Adapter */
function buildMapPayload() {

    const rows = getMapRows();

    const rawUnits =
        getMapUnits();

    const calibratedUnits =
        calibrateMapUnits(rawUnits);

    return {

        version: "B2-21",

        count:
            calibratedUnits.length,

        units:
            calibratedUnits,

        rows:
            rows
    };
}

function sendFilteredDataToMap() {

    const payload =
        buildMapPayload();
    
/* =====================================================
   مرحله 15-2
   ذخیره داده Map داخل همان Report
   ===================================================== */

     if (
        ACTIVE_REPORT_ID &&
        REPORT_CACHE[ACTIVE_REPORT_ID]
    ) {

        REPORT_CACHE[
             ACTIVE_REPORT_ID
        ].MAP.units =
            Array.isArray(payload.units)
                ? [...payload.units]
                : [];

    }



    window.BETA_FARM_MAP_PAYLOAD =
        payload;


    if (
        typeof window.CustomEvent ===
        "function"
    ) {

        window.dispatchEvent(

            new CustomEvent(
                "betaFarmFilteredData",
                {
                    detail:
                        payload
                }
            )

        );
    }


    /* =====================================================
       B2-22 — ارسال Payload به Map Engine
       ===================================================== */

    const mapFrame = ensureMapFrame(true);


    if (
        mapFrame &&
        mapFrame.contentWindow
    ) {

        mapFrame.contentWindow.postMessage(

            {
                type:
                    "BETA_FARM_MAP_UPDATE",

                map_data:
                    payload,

                fit: false
            },

            "*"

        );
    }


    return payload;
}
let reportDataReceiveCount = 0;
window.addEventListener(
    "message",
    function(event) {
   if (!window.betaFarmMessageListenerId) {

    window.betaFarmMessageListenerId =
        Math.random()
            .toString(36)
            .substring(2, 8);

}

let listenerIdBox =
    document.getElementById(
        "betaFarmListenerIdTest"
    );

if (!listenerIdBox) {

    listenerIdBox =
        document.createElement("div");

    listenerIdBox.id =
        "betaFarmListenerIdTest";

    listenerIdBox.style.position =
        "fixed";

    listenerIdBox.style.top =
        "10px";

    listenerIdBox.style.left =
    
        "10px";
    listenerIdBox.style.cursor =
        "move";

    listenerIdBox.style.userSelect =
        "none";     

    listenerIdBox.style.zIndex =
        "999999";

    listenerIdBox.style.background =
        "black";

    listenerIdBox.style.color =
        "#00ff88";

    listenerIdBox.style.padding =
        "10px";

    listenerIdBox.style.fontFamily =
        "Consolas, monospace";

    document.body.appendChild(
        listenerIdBox
    );
   let isDragging =
    false;

let dragOffsetX =
    0;

let dragOffsetY =
    0;

listenerIdBox.addEventListener(
    "mousedown",
    function(event) {

        isDragging = true;

        const rect =
            listenerIdBox.getBoundingClientRect();

        dragOffsetX =
            event.clientX -
            rect.left;

        dragOffsetY =
            event.clientY -
            rect.top;

        event.preventDefault();
    }
);

document.addEventListener(
    "mousemove",
    function(event) {

        if (!isDragging) {
            return;
        }

        listenerIdBox.style.left =
            (
                event.clientX -
                dragOffsetX
            ) + "px";

        listenerIdBox.style.top =
            (
                event.clientY -
                dragOffsetY
            ) + "px";
    }
);

document.addEventListener(
    "mouseup",
    function() {

        isDragging = false;

    }
);  
}

listenerIdBox.style.display = "none";
listenerIdBox.textContent =
    "🎯 نتیجه دایره قرنطینه\n\n" +
    "منتظر ایجاد دایره...";

 
    
        let listenerTestCount =
    Number(
        window.listenerTestCount || 0
    );

listenerTestCount++;

window.listenerTestCount =
    listenerTestCount;

window.messageReceiveCount =
    Number(
        window.messageReceiveCount || 0
    ) + 1;

const messageNow =
    new Date().toLocaleTimeString();

const messageType =
    event.data?.type || "NO TYPE";

if (!Array.isArray(window.messageMessageLog)) {
    window.messageMessageLog = [];
}

window.messageMessageLog.push(
    messageNow +
    " — #" +
    String(
        window.messageReceiveCount
    ) +
    " — " +
    String(
        messageType
    )
);

listenerIdBox.textContent =
    window.messageMessageLog.join("\n");

/* =====================================================
   TEST — دریافت درخواست UPDATE دایره
   ===================================================== */

if (
    event.data.type ===
    "BETA_FARM_QUARANTINE_UPDATE_REQUEST"
) {

    const updateCircleId =
        event.data.circleId;

    const updateCircle =
        Array.isArray(window.quarantineCircles)
            ? window.quarantineCircles.find(
                function(circle) {
                    return String(circle.id) ===
                           String(updateCircleId);
                }
            )
            : null;

    if (!updateCircle) {

        alert(
            "❌ دایره پیدا نشد\n\n" +
            "کد درخواست:\n" +
            String(updateCircleId)
        );

        return;
    }

    /*
     * همان دایره موجود را وارد
     * مسیر محاسبه فعلی می‌کنیم.
     */
    event.data.circle =
        updateCircle;

}
/* =====================================================
   SAVE CAPTURE
   Engine → Dynamic → Capture Server
   ===================================================== */

if (
    event.data &&
    event.data.type ===
        "BETA_FARM_SAVE_CAPTURE"
) {
    
    if (typeof setStatus === "function") {
        setStatus("🔥 پیام Save به Dynamic رسید");
    }

    const image =
        event.data.image || null;
    
    if (!image) {

        if (typeof setStatus === "function") {
            setStatus(
                "❌ تصویر برای ذخیره دریافت نشد"
            );
        }

        return;
    }

    try {

        fetch(
            "http://127.0.0.1:8765/save_capture",
            {
                method: "POST",

                headers: {
                    "Content-Type":
                        "application/json"
                },

                body:
                    JSON.stringify({
                        image:
                            image
                    })
            }
        )
        .then(
            function(response) {

                if (
                    response.ok
                ) {

                    if (
                        typeof setStatus ===
                        "function"
                    ) {

                        setStatus(
                            "💾 SERVER RESPONSE:"+ response.status
                        );
                    }

                } else {

                    if (
                        typeof setStatus ===
                        "function"
                    ) {

                        setStatus(
                            "❌ خطا در Save As"
                        );
                    }
                }
            }
        )
        .catch(
            function(error) {

                console.error(
                    "❌ SAVE CAPTURE ERROR:",
                    error
                );

                if (
                    typeof setStatus ===
                    "function"
                ) {

                    setStatus(
                        "❌ ارتباط با Save Server برقرار نشد"
                    );
                }
            }
        );

    } catch (error) {

        console.error(
            "❌ SAVE CAPTURE EXCEPTION:",
            error
        );

        if (
            typeof setStatus ===
            "function"
        ) {

            setStatus(
                "❌ خطای ذخیره تصویر"
            );
        }
    }
}
/* =========================================================
   B2-22 — RECEIVE FINAL REPORT DATA
   مسیر واحد:
   Engine → Dynamic → Report Cache → Report Tab
   ========================================================= */

if (
    event.data &&
    event.data.type ===
        "BETA_FARM_REPORT_DATA"
) {

    const reportData =
        event.data.report || null;

    /* -----------------------------------------------------
       ACK روی مانیتور Dynamic
       ----------------------------------------------------- */

    listenerIdBox.textContent =
        "📥 REPORT DATA RECEIVED\n\n" +
        "Type: BETA_FARM_REPORT_DATA\n" +
        "Page 1: " +
        (
            reportData &&
            reportData.page1
                ? "YES"
                : "NO"
        ) +
        "\n" +
        "Page 2: " +
        (
            reportData &&
            reportData.page2
                ? "YES"
                : "NO"
        ) +
        "\n" +
        "Page 3: " +
        (
            reportData &&
            reportData.page3
                ? "YES"
                : "NO"
        ) +
        "\n" +
        "Page 4: " +
        (
            reportData &&
            reportData.page4
                ? "YES"
                : "NO"
        );

    if (!reportData) {

        listenerIdBox.textContent +=
            "\n\n❌ REPORT PAYLOAD EMPTY";

        return;
    }


    /* -----------------------------------------------------
       گزارش فیلتر فعال در لحظه تولید گزارش
       ----------------------------------------------------- */

    const sourceReport =
        (
            ACTIVE_REPORT_ID &&
            REPORT_CACHE[ACTIVE_REPORT_ID]
        )
            ? REPORT_CACHE[ACTIVE_REPORT_ID]
            : null;


    /* -----------------------------------------------------
       Dynamic rows که Engine از آن استفاده کرده
       ----------------------------------------------------- */

    const incomingDynamic =
        reportData.dynamic || {};

    const reportRows =
        Array.isArray(
            incomingDynamic.rows
        )
            ? [...incomingDynamic.rows]
            : (
                sourceReport &&
                sourceReport.DATA &&
                Array.isArray(
                    sourceReport.DATA.rows
                )
                    ? [
                        ...sourceReport.DATA.rows
                    ]
                    : []
            );


    /* -----------------------------------------------------
       Header همان داده‌ای که Dynamic در اختیار دارد
       ----------------------------------------------------- */

    const reportHeaders =
        sourceReport &&
        sourceReport.DATA &&
        Array.isArray(
            sourceReport.DATA.headers
        )
            ? [
                ...sourceReport.DATA.headers
            ]
            : (
                sourceReport &&
                Array.isArray(
                    sourceReport.headers
                )
                    ? [
                        ...sourceReport.headers
                    ]
                    : []
            );


    /* -----------------------------------------------------
       Selection همان فیلتر مبنا
       ----------------------------------------------------- */

    const reportSelections =
        sourceReport &&
        sourceReport.DATA &&
        sourceReport.DATA.selections
            ? {
                ...sourceReport.DATA.selections
            }
            : (
                sourceReport &&
                sourceReport.selections
                    ? {
                        ...sourceReport.selections
                    }
                    : {}
            );


    /* -----------------------------------------------------
       عنوان گزارش
       ----------------------------------------------------- */

    const circleId =
        reportData.rowId ||
        (
            reportData.engine &&
            reportData.engine.reportId
        ) ||
        (
            incomingDynamic &&
            incomingDynamic.circleId
        ) ||
        "";


    const reportTitle =
        circleId
            ? "گزارش قرنطینه " + circleId
            : "گزارش قرنطینه";


    /* -----------------------------------------------------
       CREATE REPORT CACHE
       ----------------------------------------------------- */

    const reportId =
        createReportCache({
            title:
                reportTitle,

            selections:
                reportSelections,

            rows:
                reportRows,

            headers:
                reportHeaders,

            summary:
                sourceReport &&
                sourceReport.DATA
                    ? (
                        sourceReport.DATA.summary ||
                        ""
                    )
                    : ""
        });


    /* -----------------------------------------------------
       ذخیره اطلاعات Engine داخل همان Cache
       ----------------------------------------------------- */
    const sourceReportId =
        ACTIVE_REPORT_ID;
        
    ACTIVE_REPORT_ID =
        reportId;

    REPORT_CACHE[reportId].ENGINE =
        reportData.engine ||
        null;


    /* -----------------------------------------------------
       ذخیره چهار Capture
       ----------------------------------------------------- */

    REPORT_CACHE[reportId].CAPTURES = {

        page1:
            reportData.page1 ||
            null,

        page2:
            reportData.page2 ||
            null,

        page3:
            reportData.page3 ||
            null,

        page4:
            reportData.page4 ||
            null
    };
    REPORT_CACHE[reportId].SOURCE_REPORT_ID =
    sourceReportId;
    /* =========================================================
   CREATE REPORT TAB — STAGE 1
   PHYSICAL UI ONLY
   ========================================================= */

const tab =
    document.createElement("div");

tab.className =
    "filter-report-tab";

tab.dataset.reportId =
    reportId;


/* عنوان همان عنوان موجود گزارش */

const tabTitle =
    document.createElement("span");

tabTitle.className =
    "filter-report-tab-title";

tabTitle.textContent =
    REPORT_CACHE[reportId].title;


/* Exit — فقط ظاهر */

const exitBtn =
    document.createElement("button");

exitBtn.type =
    "button";

exitBtn.className =
    "filter-report-tab-control";

exitBtn.textContent =
    "×";

exitBtn.title =
    "Exit";


    /* =========================================================
       MAXIMIZE BUTTON
       ========================================================= */

    let maximizeBtn = null;

    const existingButtons =
        tab.querySelectorAll(
            ".filter-report-tab-control"
        );

    existingButtons.forEach(
        function(button) {

            if (
                button.textContent.trim() === "□" &&
                !maximizeBtn
            ) {

                maximizeBtn = button;
            }

        }
    );


    /* اگر تب Maximize نداشت، بساز */

    if (!maximizeBtn) {

        maximizeBtn =
            document.createElement("button");

        maximizeBtn.type =
            "button";

        maximizeBtn.className =
            "filter-report-tab-control";

        maximizeBtn.textContent =
            "□";

        maximizeBtn.title =
            "Maximize";

        tab.appendChild(
            maximizeBtn
        );
    }


    /* مشخص کردن دکمه Maximize */

    maximizeBtn.dataset.action =
        "maximize";
     
   
/* ساخت فیزیکی تب */

tab.appendChild(
    tabTitle
);

tab.appendChild(
    maximizeBtn
);

tab.appendChild(
    exitBtn
);

filterReportTabs.appendChild(
    tab
);
TabGuardian(tab);

/* =========================================================
   CIRCLE REPORT TAB — CLICK / ACTIVE
   ========================================================= */

tab.addEventListener(
    "click",
    function() {

        ACTIVE_REPORT_ID =
            tab.dataset.reportId;

        const allReportTabs =
            filterReportTabs.querySelectorAll(
                ".filter-report-tab"
            );

        allReportTabs.forEach(
            function(reportTab) {

                reportTab.classList.remove(
                    "active"
                );

            }
        );

        tab.classList.add(
            "active"
        );

        renderReportFromCache(
            ACTIVE_REPORT_ID
        );
    }
);  
   
    /* -----------------------------------------------------
       نمایش گزارش تازه
       ----------------------------------------------------- */

    renderReportFromCache(
        reportId
    );


    /* -----------------------------------------------------
       ACK نهایی
       ----------------------------------------------------- */

    listenerIdBox.textContent =
        "✅ REPORT CACHE CREATED\n\n" +
        "Report ID: " +
        reportId +
        "\n" +
        "Circle: " +
        (
            circleId ||
            "—"
        ) +
        "\n" +
        "Rows: " +
        reportRows.length +
        "\n" +
        "Page 1: " +
        (
            reportData.page1
                ? "YES"
                : "NO"
        ) +
        "\n" +
        "Page 2: " +
        (
            reportData.page2
                ? "YES"
                : "NO"
        ) +
        "\n" +
        "Page 3: " +
        (
            reportData.page3
                ? "YES"
                : "NO"
        ) +
        "\n" +
        "Page 4: " +
        (
            reportData.page4
                ? "YES"
                : "NO"
        );


    console.log(
        "📥 NEW QUARANTINE REPORT RECEIVED:",
        reportData
    );

    return;
}
   
                /* =====================================================
           STAGE 13-2 — دریافت دایره قرنطینه
           ===================================================== */
        if (
    event.data.type ===
        "BETA_FARM_QUARANTINE_CIRCLE_CREATED" ||
    event.data.type ===
        "BETA_FARM_QUARANTINE_UPDATE_REQUEST"
) {

    let circle =
        event.data.circle;

    if (
        event.data.type ===
        "BETA_FARM_QUARANTINE_CIRCLE_CREATED"
    ) {

        if (!circle) {
            return;
        }

        if (
            !Array.isArray(
                window.quarantineCircles
            )
        ) {
            window.quarantineCircles = [];
        }

        window.quarantineCircles.push(
            circle
        );

        updateQuarantineList();

    } else {

        const updateCircleId =
            event.data.circleId;
        listenerIdBox.textContent =
              "🔄 UPDATE دریافت شد\n\n" +
              "کد دایره: " +
              String(updateCircleId);     

        if (
            !Array.isArray(
                window.quarantineCircles
            )
        ) {
            window.quarantineCircles = [];
        }

        circle =
            window.quarantineCircles.find(
                function(item) {
                    return String(item.id) ===
                           String(updateCircleId);
                }
            );

        if (!circle) {

            alert(
                "❌ دایره پیدا نشد\n\n" +
                "کد درخواست:\n" +
                String(updateCircleId)
            );

            return;
        }
    }        
            if (
             event.data.type ===
             "BETA_FARM_QUARANTINE_CIRCLE_CREATED"
            ) {

                if (
                    !Array.isArray(
                        window.quarantineCircles
                    )
                ) {
                    window.quarantineCircles = [];
                }

                window.quarantineCircles.push(
                     circle
                );

                updateQuarantineList();
            } 

            const centerLat =
                Number(circle.centerLat);

            const centerLon =
                Number(circle.centerLon);

            const radiusKm =
                Number(circle.radiusKm);

            if (
                !Number.isFinite(centerLat) ||
                !Number.isFinite(centerLon) ||
                !Number.isFinite(radiusKm)
            ) {
                return;
            }

            const units =
                buildMapPayload().units;

            let poultryCount = 0;

            units.forEach(
                function(unit) {

                    const lat =
                        Number(unit.lat);

                    const lon =
                        Number(unit.lon);

                    if (
                        !Number.isFinite(lat) ||
                        !Number.isFinite(lon)
                    ) {
                        return;
                    }

                    const R =
                        6371;

                    const dLat =
                        (lat - centerLat) *
                        Math.PI / 180;

                    const dLon =
                        (lon - centerLon) *
                        Math.PI / 180;

                    const a =
                        Math.sin(dLat / 2) *
                        Math.sin(dLat / 2) +
                        Math.cos(
                            centerLat *
                            Math.PI / 180
                        ) *
                        Math.cos(
                            lat *
                            Math.PI / 180
                        ) *
                        Math.sin(dLon / 2) *
                        Math.sin(dLon / 2);

                    const distanceKm =
                        2 *
                        R *
                        Math.atan2(
                            Math.sqrt(a),
                            Math.sqrt(1 - a)
                        );

                    if (
                        distanceKm <= radiusKm
                    ) {
                        poultryCount += 1;
                    }
                 
                }
            );
            /* =====================================================
   STAGE 13-2 — شمارش شهرها و روستاها
   ===================================================== */

let cityCount = 0;
let villageCount = 0;

SETTLEMENTS.forEach(
    function(place) {

        const lat =
            Number(place.lat);

        const lon =
            Number(place.lon);

        if (
            !Number.isFinite(lat) ||
            !Number.isFinite(lon)
        ) {
            return;
        }

        const R =
            6371;

        const dLat =
            (lat - centerLat) *
            Math.PI / 180;

        const dLon =
            (lon - centerLon) *
            Math.PI / 180;

        const a =
            Math.sin(dLat / 2) *
            Math.sin(dLat / 2) +
            Math.cos(
                centerLat *
                Math.PI / 180
            ) *
            Math.cos(
                lat *
                Math.PI / 180
            ) *
            Math.sin(dLon / 2) *
            Math.sin(dLon / 2);

        const distanceKm =
            2 *
            R *
            Math.atan2(
                Math.sqrt(a),
                Math.sqrt(1 - a)
            );

        if (
            distanceKm <= radiusKm
        ) {

            if (
                place.type === "city" ||
                place.type === "town"
            ) {

                cityCount += 1;

            }

            else if (
                place.type === "village"
            ) {

                villageCount += 1;

            }

        }

    }
);
const mapFrame =
    ensureMapFrame(true);

if (
    mapFrame &&
    mapFrame.contentWindow
) {

    mapFrame.contentWindow.postMessage(
        {
            type:
                "BETA_FARM_QUARANTINE_RESULT",

            report: {
                circleId:
                    circle.id,

                poultryCount:
                    poultryCount,

                cityCount:
                    cityCount,

                villageCount:
                    villageCount
            }
        },
        "*"
    );
}


            listenerIdBox.textContent =
                
                "🎯 نتیجه دایره قرنطینه\n\n" +
                "کد دایره: " +
                circle.id +
                "\n\n" +
                "واحدهای طیور داخل دایره: " +
                poultryCount +
                "\n" +
                "شهرهای داخل دایره: " +
                cityCount +
                "\n" +
                "روستاهای داخل دایره: " +
                villageCount;


            console.log(
                "CIRCLE RESULT:",
                circle.id,
                poultryCount
            );

            return;
        }
       
        
        
                // =====================================================
        // B2-22 — DIAGNOSTIC: دریافت پیام خام از Map Engine
        // =====================================================

        console.log(
            "=== BETA FARM MAP MESSAGE RECEIVED ==="
        );

        console.log(
            "event.data:",
            event.data
        );

        console.log(
            "message type:",
            event.data.type
        );

        try {
            console.log(
                "message JSON:",
                JSON.stringify(
                    event.data,
                    null,
                    2
                )
            );
        } catch (e) {
            console.log(
                "message JSON ERROR:",
                e
            );
        }


 
// =====================================================
// B2-22 — ثبت دایره قرنطینه در Dynamic Filter
// =====================================================

if (
    event.data.type ===
    "BETA_FARM_QUARANTINE_CIRCLE_CREATED"
) {

    const circle =
        event.data.circle;

    if (!circle) {
        return;
    }

    // ایجاد آرایه در صورت نبودن
    if (
        !Array.isArray(
            window.quarantineCircles
        )
    ) {
        window.quarantineCircles = [];
    }

    // ثبت دایره
    window.quarantineCircles.push(
        circle
    );

    updateQuarantineList(); 
    console.log(
        "🔥 CIRCLE SAVED IN DYNAMIC FILTER:",
        circle
    );

    console.log(
        "🔥 TOTAL CIRCLES:",
        window.quarantineCircles.length
    );

    return;
}
        if (
            event.data.type ===
           "BETA_FARM_MAP_DRAG_START"
        ) {

        // =====================================================
        // B2-22 — نمایش فقط پیام‌های غیر اولیه
        // =====================================================

        if (
            event.data.type !== "BETA_FARM_MAP_VERSION"
        ) {

            console.log(
                "NON-VERSION MESSAGE:",
                event.data
            );

        }

        

          const infoBox =
              document.getElementById(
                  "testMapOverlay"
              );

          if (infoBox) {
              infoBox.remove();
          }

          return;
        } 

        if (
            event.data.type !==
            "BETA_FARM_MAP_MARKER_CLICK"
        ) {
            return;
        }

        const fmId =
            event.data.FM_ID;

        if (!fmId) {
            return;
        }

        const mapShell =
            document.getElementById(
                "mapShell"
            );

        if (!mapShell) {
            return;
        }

        const oldBox =
            document.getElementById(
                "testMapOverlay"
            );

        if (oldBox) {
            oldBox.remove();
        }

        /*
         * پیدا کردن ردیف واقعی واحد
         * ابتدا از نتایج فیلترشده استفاده می‌کنیم.
         * اگر پیدا نشد، کل DATA.rows را بررسی می‌کنیم.
         */
        const rows =
            APP_STATE.filteredRows || [];

        let selectedRow = null;

        for (
            let i = 0;
            i < rows.length;
            i++
        ) {
            const row = rows[i];

            const idCol = Number(DATA.id_column || 1);

            if (
                String(row[String(idCol)] ?? "").trim() ===
                String(fmId).trim()
            ) {
                selectedRow = row;
                break;
            }
        }

        if (!selectedRow) {
            for (
                let i = 0;
                i < DATA.rows.length;
                i++
            ) {
                const row =
                    DATA.rows[i];

                if (
                    String(row["1"] ?? "").trim() ===
                    String(fmId).trim()
                ) {
                    selectedRow = row;
                    break;
                }
            }
        }

        const infoBox =
            document.createElement(
                "div"
            );

        infoBox.id =
            "testMapOverlay";

        infoBox.style.position =
            "absolute";

        infoBox.style.top =
            "20px";

        infoBox.style.right =
            "20px";

        infoBox.style.zIndex =
            "1000";

        infoBox.style.width =
            "240px";

        infoBox.style.maxHeight =
            "55%";

        infoBox.style.overflowY =
            "auto";

        infoBox.style.padding =
            "9px";

        infoBox.style.background =
            "rgba(255,255,255,.97)";

        infoBox.style.border =
            "2px solid #17452d";

        infoBox.style.borderRadius =
            "10px";

        infoBox.style.boxShadow =
            "0 5px 18px rgba(0,0,0,.30)";

        infoBox.style.color =
            "#263746";

        infoBox.style.fontSize =
            "11px";

        infoBox.style.lineHeight =
            "1.55";

        infoBox.style.direction =
            "rtl";

        infoBox.style.textAlign =
            "right";

        let html =
            "<div style='" +
            "font-size:14px;" +
            "font-weight:bold;" +
            "margin-bottom:8px;" +
            "'>" +
            "مشخصات واحد" +
            "</div>";

        html +=
            "<div style='" +
            "padding:5px 0;" +
            "margin-bottom:5px;" +
            "border-bottom:2px solid #17452d;" +
            "'>" +
            "<strong>کد واحد:</strong> " +
            escapeHtml(
                String(fmId)
            ) +
            "</div>";

        if (selectedRow) {

            DATA.headers.forEach(
                function(name, index) {

                    const value =
                        selectedRow[
                            String(index + 1)
                        ] ?? "";

                    if (
                        String(value).trim() === ""
                    ) {
                        return;
                    }

                    html +=
                        "<div style='" +
                        "border-bottom:1px solid #e5e5e5;" +
                        "padding:3px 0;" +
                        "'>" +
                        "<strong>" +
                        escapeHtml(
                            String(name)
                        ) +
                        ":</strong> " +
                        escapeHtml(
                            String(value)
                        ) +
                        "</div>";
                }
            );

        } else {

            html +=
                "<div style='color:#a33;'>" +
                "اطلاعات این واحد در داده‌های فعلی پیدا نشد." +
                "</div>";
        }

        infoBox.innerHTML =
            html;

        mapShell.appendChild(
            infoBox
        );

    }
);

function updateQuarantineList() {

    if (
        !Array.isArray(
            window.quarantineCircles
        )
    ) {
        window.quarantineCircles = [];
    }

    const list =
        document.getElementById(
            "quarantineCircleList"
        );

    if (!list) {
        return;
    }

    list.innerHTML = "";

    window.quarantineCircles.forEach(
        function(circle) {

            const item =
                document.createElement("div");

            item.style.padding = "6px";
            item.style.marginBottom = "5px";
            item.style.borderBottom =
                "1px solid #ddd";
            item.style.direction = "rtl";

            item.textContent =
                circle.id +
                " — شعاع " +
                circle.radiusKm +
                " کیلومتر";

            list.appendChild(item);
        }
    );
}

function updateMap() {
    const units = getMapUnits();

    console.log("MAP UPDATE");
    console.log("Units:", units);
    console.log("Count:", units.length);

    const container =
        document.getElementById("dynamicMapMarkers");

    if (!container) {
        return units;
    }

    container.innerHTML = "";

    if (units.length === 0) {
        return units;
    }

    const xs = units.map(function(u) { return Number(u.X); })
                    .filter(function(v) { return Number.isFinite(v); });
    const ys = units.map(function(u) { return Number(u.Y); })
                    .filter(function(v) { return Number.isFinite(v); });

    if (xs.length === 0 || ys.length === 0) {
        return units;
    }

    const minX = Math.min.apply(null, xs);
    const maxX = Math.max.apply(null, xs);
    const minY = Math.min.apply(null, ys);
    const maxY = Math.max.apply(null, ys);

    units.forEach(function(unit, index) {

        const x = Number(unit.X);
        const y = Number(unit.Y);

        if (!Number.isFinite(x) || !Number.isFinite(y)) {
            return;
        }

        const marker = document.createElement("div");

        const left = maxX === minX
            ? 50
            : ((x - minX) / (maxX - minX)) * 86 + 7;

        const top = maxY === minY
            ? 50
            : (1 - ((y - minY) / (maxY - minY))) * 82 + 9;

        marker.textContent =
            unit.FM_ID || ("Unit " + (index + 1));

        marker.style.position = "absolute";
        marker.style.left = left + "%";
        marker.style.top = top + "%";
        marker.style.transform = "translate(-50%, -50%)";
        marker.style.zIndex = "10";
        marker.style.width = "18px";
        marker.style.height = "18px";
        marker.style.padding = "0";
        marker.style.borderRadius = "50%";
        marker.style.background = "#ffffff";
        marker.style.border = "4px solid #d64b3b";
        marker.style.color = "#ffffff";
        marker.style.fontSize = "0";
        marker.style.whiteSpace = "nowrap";
        marker.style.boxShadow = "0 2px 6px rgba(0,0,0,.35)";
        marker.style.cursor = "pointer";
        marker.title = unit.FM_ID || ("Unit " + (index + 1));

        const row = APP_STATE.filteredRows[index] || {};
        const infoBox = document.createElement("div");

        infoBox.style.position = "absolute";
        infoBox.style.display = "none";
        infoBox.style.minWidth = "180px";
        infoBox.style.maxWidth = "260px";
        infoBox.style.padding = "8px 10px";
        infoBox.style.borderRadius = "8px";
        infoBox.style.background = "rgba(255,255,255,.97)";
        infoBox.style.border = "1px solid #c8d3dc";
        infoBox.style.boxShadow = "0 3px 12px rgba(0,0,0,.25)";
        infoBox.style.color = "#263746";
        infoBox.style.fontSize = "11px";
        infoBox.style.lineHeight = "1.7";
        infoBox.style.textAlign = "right";
        infoBox.style.direction = "rtl";
        infoBox.style.zIndex = "100";
        infoBox.style.pointerEvents = "none";

        const infoTitle = document.createElement("div");
        infoTitle.style.fontWeight = "bold";
        infoTitle.style.marginBottom = "3px";
        infoTitle.textContent = unit.FM_ID || ("Unit " + (index + 1));
        infoBox.appendChild(infoTitle);

        DATA.headers.forEach(function(name, headerIndex) {
            if (headerIndex === 21 || headerIndex === 22) {
                return;
            }

            const value = row[String(headerIndex + 1)];

            if (value === undefined || value === null || String(value).trim() === "") {
                return;
            }

            const line = document.createElement("div");
            line.textContent = String(name) + ": " + String(value);
            infoBox.appendChild(line);
        });

        container.appendChild(infoBox);

        function showInfo() {
            marker.style.transform = "translate(-50%, -50%) scale(1.25)";
            marker.style.zIndex = "150";

            infoBox.style.display = "block";
            infoBox.style.left = left + "%";
            infoBox.style.top = top + "%";
            infoBox.style.transform = "translate(-50%, calc(-100% - 14px))";
        }

        function hideInfo() {
            marker.style.transform = "translate(-50%, -50%)";
            marker.style.zIndex = "10";
            infoBox.style.display = "none";
        }

        marker.addEventListener("mouseenter", showInfo);
        marker.addEventListener("mouseleave", hideInfo);
        marker.addEventListener("click", function(event) {
            event.stopPropagation();
            showInfo();
        });

        infoBox.addEventListener("mouseleave", hideInfo);

        container.appendChild(marker);
    });

    return units;
}
function testMapData() {
    const units = getMapUnits();

    console.log("MAP TEST");
    console.log("Map units:", units);
    console.log("Map unit count:", units.length);
    return units;
}

/* =========================================================
   ELEMENTS
   ========================================================= */

const filterPanel =
    document.getElementById(
        "filterPanel"
    );

const applyBtn =
    document.getElementById(
        "applyBtn"
    );


const resultTableBox =
    document.getElementById(
        "resultTableBox"
    );

const resultCount =
    document.getElementById(
        "resultCount"
    );

const resultTable =
    document.getElementById(
        "resultTable"
    );

const workspace =
    document.getElementById(
        "workspace"
    );

const resultArea =
    document.getElementById(
        "resultArea"
    );

const expandMapBtn =
    document.getElementById(
        "expandMapBtn"
    );

const footer =
    document.getElementById(
        "footer"
    );
/* =========================================================
   REPORT CACHE
   مرحله 15 — پایه ذخیره گزارش‌های فیلتر
   ========================================================= */

const REPORT_CACHE = {};

let REPORT_SEQUENCE = 0;
let ACTIVE_REPORT_ID = null;
/* =========================================================
   REPORT CACHE
   ساخت یک گزارش جدید در Cache
   ========================================================= */
function createReportCache(reportData) {

    REPORT_SEQUENCE++;

    const reportId =
        "REPORT-" +
        REPORT_SEQUENCE;

    REPORT_CACHE[reportId] = {

        id:
            reportId,

        title:
            reportData.title ||
            ("گزارش " + REPORT_SEQUENCE),

        DATA: {

            selections:
                {
                    ...(reportData.selections || {})
                },

            rows:
                Array.isArray(reportData.rows)
                    ? [...reportData.rows]
                    : [],

            headers:
                Array.isArray(reportData.headers)
                    ? [...reportData.headers]
                    : [],

            summary:
                reportData.summary || ""
        },

        MAP: {

            units: []
        },

        QUARANTINE: {

            circles: []
        },

        createdAt:
            Date.now()
    };

    return reportId;
}


/* =========================================================
   REPORT CACHE
   نمایش گزارش ذخیره‌شده از Cache
   Dynamic Filter + Engine
   ========================================================= */

function renderReportFromCache(reportId) {

    const report =
        REPORT_CACHE[reportId];

    if (!report) {
        return;
    }

    const rows =
        report.DATA.rows || [];

    const headers =
        report.DATA.headers || [];

    const engine =
        report.ENGINE || null;

    const captures =
        report.CAPTURES || {};

    /*
     * تشخیص نوع گزارش
     *
     * Filter Report:
     * فقط DATA دارد.
     *
     * Quarantine Report:
     * علاوه بر DATA،
     * ENGINE یا CAPTURES دارد.
     */

    const isQuarantineReport =
        !!(
            engine ||
            captures.page1 ||
            captures.page2 ||
            captures.page3 ||
            captures.page4
        );


    /* =====================================================
       HTML اصلی
       ===================================================== */

    let html = "";


    /* =====================================================
       FILTER REPORT
       رفتار قبلی بدون تغییر
       ===================================================== */

    if (!isQuarantineReport) {

        resultCount.textContent =
            "تعداد ردیف‌های نتیجه: " +
            rows.length;


        /* -----------------------------------------------
           جدول اصلی Dynamic Filter
           ----------------------------------------------- */

        let filterHtml =
            "<table>" +
            "<thead>" +
            "<tr>";


        headers.forEach(
            function(name) {

                filterHtml +=
                    "<th>" +
                    escapeHtml(
                        String(name)
                    ) +
                    "</th>";

            }
        );


        filterHtml +=
            "</tr>" +
            "</thead>" +
            "<tbody>";


        rows.forEach(
            function(row) {

                filterHtml +=
                    "<tr>";


                headers.forEach(
                    function(name, index) {

                        const value =
                            row[
                                String(
                                    index + 1
                                )
                            ] ?? "";


                        filterHtml +=
                            "<td>" +
                            escapeHtml(
                                String(value)
                            ) +
                            "</td>";

                    }
                );


                filterHtml +=
                    "</tr>";

            }
        );


        filterHtml +=
            "</tbody>" +
            "</table>";


        if (
            rows.length === 0
        ) {

            filterHtml =
                '<div style="padding:12px;text-align:center;">' +
                'هیچ ردیفی با فیلترهای انتخاب‌شده پیدا نشد.' +
                '</div>';

        }


        html +=
            filterHtml;

    }


    /* =====================================================
       QUARANTINE REPORT
       هدر اختصاصی گزارش قرنطینه
       ===================================================== */

    if (isQuarantineReport) {

        let reportTitle =
            "گزارش قرنطینه";

        let reportIdText = "";


        if (
            engine &&
            engine.activeCircle
        ) {

            const active =
                engine.activeCircle;


            if (
                active.id
            ) {

                reportIdText =
                    String(
                        active.id
                    );

            }


            if (
                active.id
            ) {

                reportTitle =
                    "گزارش قرنطینه " +
                    String(
                        active.id
                    );

            }

        }


        html +=
            '<div style="' +
            'margin-bottom:20px;' +
            'padding:15px;' +
            'border:1px solid #ccc;' +
            'border-radius:8px;' +
            'background:#f7f7f7;' +
            '">' +


            '<div style="' +
            'font-size:20px;' +
            'font-weight:bold;' +
            'text-align:center;' +
            'margin-bottom:10px;' +
            '">' +

            escapeHtml(
                reportTitle
            ) +

            '</div>';


        if (
            engine &&
            engine.activeCircle
        ) {

            const active =
                engine.activeCircle;


            html +=
                '<div style="' +
                'text-align:center;' +
                'font-size:15px;' +
                'margin-top:5px;' +
                '">' +

                'شعاع قرنطینه: ' +

                escapeHtml(
                    String(
                        active.radiusKm ?? ""
                    )
                ) +

                ' کیلومتر' +

                '</div>';

        }


        html +=
            '</div>';

    }

    /* =====================================================
       B2-22
       اطلاعات Map Engine
       فقط برای گزارش قرنطینه
       ===================================================== */

    if (
        isQuarantineReport &&
        engine
    ) {

        html +=
            '<div style="' +
            'margin-top:20px;' +
            'padding:15px;' +
            'border:1px solid #ccc;' +
            'border-radius:8px;' +
            'background:#fafafa;' +
            '">' +

            '<div style="' +
            'font-weight:bold;' +
            'font-size:17px;' +
            'padding-bottom:10px;' +
            'border-bottom:2px solid #333;' +
            'margin-bottom:12px;' +
            '">' +

            'اطلاعات قرنطینه' +

            '</div>';


        /* -----------------------------------------------
           دایره فعال گزارش
           ----------------------------------------------- */

        if (
            engine.activeCircle
        ) {

            const active =
                engine.activeCircle;


            html +=
                '<table style="' +
                'width:100%;' +
                '">' +

                '<tbody>' +


                '<tr>' +
                '<th style="text-align:right;">شناسه دایره</th>' +
                '<td>' +
                escapeHtml(
                    String(
                        active.id ?? ""
                    )
                ) +
                '</td>' +
                '</tr>' +


                '<tr>' +
                '<th style="text-align:right;">عرض جغرافیایی</th>' +
                '<td>' +
                escapeHtml(
                    String(
                        active.centerLat ?? ""
                    )
                ) +
                '</td>' +
                '</tr>' +


                '<tr>' +
                '<th style="text-align:right;">طول جغرافیایی</th>' +
                '<td>' +
                escapeHtml(
                    String(
                        active.centerLon ?? ""
                    )
                ) +
                '</td>' +
                '</tr>' +


                '<tr>' +
                '<th style="text-align:right;">شعاع قرنطینه</th>' +
                '<td>' +
                escapeHtml(
                    String(
                        active.radiusKm ?? ""
                    )
                ) +
                ' کیلومتر' +
                '</td>' +
                '</tr>' +


                '<tr>' +
                '<th style="text-align:right;">رنگ دایره</th>' +
                '<td>' +
                escapeHtml(
                    String(
                        active.color ?? ""
                    )
                ) +
                '</td>' +
                '</tr>' +


                '<tr>' +
                '<th style="text-align:right;">وضعیت نمایش</th>' +
                '<td>' +
                escapeHtml(
                    String(
                        active.visible ?? ""
                    )
                ) +
                '</td>' +
                '</tr>' +


                '</tbody>' +
                '</table>';

        }


        /* -----------------------------------------------
           زمان تولید گزارش
           ----------------------------------------------- */

        if (
            engine.generatedAt
        ) {

            html +=
                '<div style="' +
                'padding-top:12px;' +
                'margin-top:12px;' +
                'border-top:1px solid #ddd;' +
                'font-size:12px;' +
                'color:#555;' +
                '">' +

                'زمان تولید گزارش: ' +

                escapeHtml(
                    String(
                        engine.generatedAt
                    )
                ) +

                '</div>';

        }


        html +=
            '</div>';

    }
   
    /* =====================================================
       REPORT CAPTURES — PAGE 1..4
       ===================================================== */

    if (
        isQuarantineReport &&
        captures &&
        (
            captures.page1 ||
            captures.page2 ||
            captures.page3 ||
            captures.page4
        )
    ) {

        html += `
            <div style="
                margin-top:25px;
                padding-top:15px;
                border-top:2px solid #ccc;
            ">

                <h3 style="
                    margin:0 0 15px 0;
                    text-align:center;
                ">
                    تصاویر گزارش
                </h3>
        `;


        const capturePages = [
            {
                key: "page1",
                title: "صفحه ۱ — نمای کاربر"
            },
            {
                key: "page2",
                title: "صفحه ۲ — شعاع ۵ کیلومتر"
            },
            {
                key: "page3",
                title: "صفحه ۳ — شعاع ۱۰ کیلومتر"
            },
            {
                key: "page4",
                title: "صفحه ۴ — گزارش نهایی"
            }
        ];


        capturePages.forEach(
            function(page) {

                const imageData =
                    captures[
                        page.key
                    ];


                if (!imageData) {
                    return;
                }


                html += `
                    <div style="
                        margin:20px 0 30px 0;
                        padding:12px;
                        border:1px solid #ccc;
                        border-radius:8px;
                        background:#fff;
                    ">

                        <div style="
                            font-weight:bold;
                            margin-bottom:10px;
                            text-align:center;
                        ">
                            ${page.title}
                        </div>

                        <img
                            src="${imageData}"
                            style="
                                display:block;
                                width:100%;
                                height:auto;
                                border:1px solid #ddd;
                                margin:auto;
                            "
                        >

                    </div>
                `;

            }
        );


        html += `
            </div>
        `;

    }


    /* =====================================================
       نمایش نهایی Report Box
       ===================================================== */

    /*
     * در Report Tab دیگر نتیجه فیلتر نمایش داده نمی‌شود.
     * برای حفظ رفتار Filter Tab، فقط در گزارش قرنطینه
     * resultCount را خالی می‌کنیم.
     */

    if (
        isQuarantineReport
    ) {

        resultCount.textContent =
            "";

    }


    resultTable.innerHTML =
        html;


    resultTableBox.style.display =
        "block";


    resultArea.classList.add(
        "result-visible"
    );


    resultArea.style.display =
        "block";
}

/* =========================================================
   UI B1
   ساخت نوار Navigation و زیرمنوها
   بدون اتصال به Filter و Map
   ========================================================= */

function buildUI() {

    const topbar =
        document.querySelector(".topbar");

    if (!topbar) {
        return;
    }

    topbar.innerHTML = "";

    const menuArea =
        document.createElement("div");

    menuArea.className =
        "ui-menu-area";


    /* =====================================================
       تعریف منوها
       ===================================================== */

    const menuDefinitions = {

        "Home": [],

        "File": [
            "Open",
            "New File",
            "Print",
            "Exit"
        ],

        "View": [
            "Default",
            "Online (Google Maps)",
            "Offline"
        ],

        "Share": [],

        "Edit": ["خط‌کش",
                 "شعاع قرنطینه"
        ]
            
    };


    /* =====================================================
       ساخت منوها
       ===================================================== */

    Object.keys(menuDefinitions).forEach(
        function(menuName) {

            const wrapper =
                document.createElement("div");

            wrapper.className =
                "ui-menu-wrapper";


            const button =
                document.createElement("div");

            button.type = "button";

            button.className =
                "ui-menu-btn";

            button.textContent =
                menuName;


            wrapper.appendChild(button);


            const submenuItems =
                menuDefinitions[menuName];

            if (submenuItems.length > 0) {
                button.addEventListener("click", function(event) {
                    event.stopPropagation();
                    document.querySelectorAll(".ui-menu-wrapper.open").forEach(function(w) {
                        if (w !== wrapper) w.classList.remove("open");
                    });
                    wrapper.classList.toggle("open");
                });
            }


            /* =================================================
               ساخت زیرمنو فقط برای منوهایی که گزینه دارند
               ================================================= */

            if (submenuItems.length > 0) {

                const submenu =
                    document.createElement("div");

                submenu.className =
                    "ui-submenu";


                submenuItems.forEach(
                    function(itemName) {

                        const item =
                            document.createElement(
                                "button"
                            );

                        item.type = "button";

                        item.className =
                            "ui-submenu-item";
                                                item.addEventListener("click", function() {
                            wrapper.classList.remove("open");
                        });

                        item.textContent =
                            itemName;
                        item.setAttribute(
                            "data-beta-menu-item",
                            menuName + "::" + itemName
                        );     

                                                /* =====================================================
                           B2-22 — TEST EXIT STYLE
                           ===================================================== */

                        if (
                            menuName === "File" &&
                            itemName === "Exit"
                        ) {

                            item.textContent =
                                "⏻  EXIT";

                            item.style.setProperty(
                                "background-color",
                                "#000000",
                                "important"
                            );

                            item.style.setProperty(
                                "color",
                                "#ffff00",
                                "important"
                            );

                            item.style.setProperty(
                                "font-weight",
                                "bold",
                                "important"
                            );

                            item.style.setProperty(
                                "border",
                                "2px solid #ff0000",
                                "important"
                            );
                        }
                        // B1 — فعال‌سازی گزینه‌های Navigation
                        if (menuName === "File" && itemName === "New File") {
                                                
                        
                            item.addEventListener("click", function() {
                                if (window.betaFarmBridge &&
                                    typeof window.betaFarmBridge.newFile === "function") {
                                    window.betaFarmBridge.newFile();
                                }
                            });
                        }

                        if (menuName === "File" && itemName === "Print") {
                            item.addEventListener("click", function() {
                                if (window.betaFarmBridge &&
                                    typeof window.betaFarmBridge.printFile === "function") {
                                    window.betaFarmBridge.printFile();
                                }
                            });
                        }

                        if (menuName === "File" && itemName === "Exit") {
                            item.addEventListener("click", function() {
                                if (window.betaFarmBridge &&
                                    typeof window.betaFarmBridge.exitApp === "function") {
                                    window.betaFarmBridge.exitApp();
                                }
                            });
                        }

                        if (menuName === "View") {
                            item.addEventListener("click", function() {
                                if (window.betaFarmBridge &&
                                    typeof window.betaFarmBridge.setMapMode === "function") {
                                    window.betaFarmBridge.setMapMode(itemName);
                                }
                            });
                        }
                         

if (
    menuName === "Edit" &&
    itemName === "خط‌کش"
) {
    
    item.addEventListener(
        "click",
        function() {

            if (MAP_FRAME && MAP_FRAME.contentWindow) {

                MAP_FRAME.contentWindow.postMessage(
                    {
                        type:
                            "BETA_FARM_MAP_RULER_TOGGLE"
                    },
                    "*"
                );

            }

        }
    );

}



                         
                        submenu.appendChild(item);
                        if (
    menuName === "File" &&
    itemName === "Open"
) {

    item.addEventListener(
        "click",
        function() {

            if (
                window.betaFarmBridge &&
                typeof window.betaFarmBridge.openFile ===
                "function"
            ) {

                window.betaFarmBridge.openFile();

            }

        }
    );

}

                    }
                );


                wrapper.appendChild(submenu);
               /* =============================================
                مخفی شدن زیرمنو هنگام خروج موس
                 ============================================= */

                wrapper.addEventListener(
                    "mouseleave",
                    function() {

                        wrapper.classList.remove("open");

                        submenu.style.display =
                            "none";

    }
                );
 

                /* =============================================
                   باز و بسته کردن زیرمنو
                   ============================================= */

                button.addEventListener(
                    "click",
                    function(event) {

                        event.stopPropagation();


                        document
                            .querySelectorAll(
                                ".ui-submenu"
                            )
                            .forEach(
                                function(menu) {

                                    if (
                                        menu !== submenu
                                    ) {

                                        menu.style.display =
                                            "none";

                                    }

                                }
                            );


                        if (
                            submenu.style.display ===
                            "block"
                        ) {

                            submenu.style.display =
                                "none";

                        } else {

                            submenu.style.display =
                                "block";

                        }

                    }
                );

            }


            menuArea.appendChild(wrapper);

        }
    );


    /* =====================================================
       نام برنامه
       ===================================================== */

    const brand =
        document.createElement("div");

    brand.className =
        "ui-brand";

    brand.textContent =
        "Beta Farm";


    topbar.appendChild(menuArea);

    topbar.appendChild(brand);

}

document.addEventListener("click", function() {
    document.querySelectorAll(".ui-menu-wrapper.open").forEach(function(w) {
        w.classList.remove("open");
    });
});

/* =========================================================
   CREATE FILTERS
   ========================================================= */

function createFilters() {

    filterPanel.innerHTML = "";


    DATA.filters.forEach(
        function(filter, index) {


            const card =
                document.createElement(
                    "div"
                );

            card.className =
                "filter-card";


            const title =
                document.createElement(
                    "div"
                );

            title.className =
                "filter-title";


            const duplicateCount =
                DATA.filters.filter(
                    function(f) {
                        return f.name ===
                            filter.name;
                    }
                ).length;


            let displayName =
                filter.name;


            if (duplicateCount > 1) {

                displayName =
                    filter.name +
                    " (" +
                    (index + 1) +
                    ")";
            }


            title.textContent =
                displayName;


            const select =
                document.createElement(
                    "select"
                );


            select.dataset.column =
                String(filter.column);


            const allOption =
                document.createElement(
                    "option"
                );

            allOption.value = "";

            allOption.textContent =
                "همه";

            select.appendChild(
                allOption
            );


            filter.values.forEach(
                function(value) {

                    const option =
                        document.createElement(
                            "option"
                        );

                    option.value =
                        value;

                    option.textContent =
                        value;

                    select.appendChild(
                        option
                    );
                }
            );


            card.appendChild(title);

            card.appendChild(select);

            filterPanel.appendChild(card);

        }
    );
}


/* =========================================================
   GET SELECTIONS
   ========================================================= */

function getSelections() {

    const selections = {};


    filterPanel
        .querySelectorAll("select")
        .forEach(
            function(select) {

                if (select.value !== "") {

                    selections[
                        select.dataset.column
                    ] =
                        select.value;
                }

            }
        );


    return selections;
}


/* =========================================================
   APPLY FILTER
   ========================================================= */

applyBtn.addEventListener(
    "click",
    function() {


        const selections =
            getSelections();
        APP_STATE.selections = selections;
        APP_STATE.mode = "FILTERED";

        FILE_STATES[ACTIVE_FILE_ID] = {
            selections: { ...selections },
            mode: "FILTERED"
        }; 
        const filteredRows =
            DATA.rows.filter(
                function(row) {

                    return Object.keys(
                        selections
                    ).every(
                        function(column) {

                            return String(
                                row[column] ?? ""
                            ).trim()
                            ===
                            String(
                                selections[column]
                            ).trim();

                        }
                    );

                }
            );
       APP_STATE.filteredRows = filteredRows;     
       
       footer.textContent =
           "Beta Farm Dynamic Filter 0.5.1-B2-19" +
           "  |  ستون‌ها: " +
           DATA.column_count +
           "  |  ردیف‌ها: " +
           DATA.row_count +
           "  |  فیلترها: " +
           DATA.filters.length +
           "  |  Map Units: " +
           getMapRowCount();
      updateMap();
      sendFilteredDataToMap();
               const selectedNames = [];


        filterPanel
            .querySelectorAll("select")
            .forEach(
                function(select) {

                    if (select.value !== "") {

                        const title =
                            select
                                .parentElement
                                .querySelector(
                                    ".filter-title"
                                )
                                .textContent;


                        selectedNames.push(
                            title +
                            ": " +
                            select.value
                        );

                    }

                }
            );


        let summary =
            "تعداد واحدهای منطبق: " +
            filteredRows.length;


        if (selectedNames.length > 0) {

            summary +=
                "<br>" +
                selectedNames.join(
                    " &nbsp; | &nbsp; "
                );

        } else {

            summary +=
                "<br>" +
                "همه واحدها؛ هیچ فیلتری محدود نشده است.";
        }
        /* =========================================================
   CREATE REPORT CACHE
   ========================================================= */

const reportId =
    createReportCache({

        title:
            "گزارش " +
            (REPORT_SEQUENCE + 1),

        selections:
            selections,

        rows:
            filteredRows,

        headers:
            DATA.headers,

        summary:
            summary

    });
    
ACTIVE_REPORT_ID = reportId;

/* =========================================================
   CREATE REPORT TAB
   ========================================================= */

const tab =
    document.createElement("button");

tab.type = "button";

tab.className =
    "filter-report-tab";

tab.textContent =
    REPORT_CACHE[reportId].title;

tab.dataset.reportId =
    reportId;


/* =========================================================
   TAB → CACHE
   ========================================================= */
tab.addEventListener(
    "click",
    function() {

        ACTIVE_REPORT_ID =
            tab.dataset.reportId;

        const allReportTabs =
            filterReportTabs.querySelectorAll(
                ".filter-report-tab"
            );

        allReportTabs.forEach(
            function(reportTab) {

                reportTab.classList.remove(
                    "active"
                );

            }
        );

        tab.classList.add(
            "active"
        );

        renderReportFromCache(
            ACTIVE_REPORT_ID
        );

        sendFilteredDataToMap();
    }
);


filterReportTabs.appendChild(
    tab
);
TabGuardian(tab);


/* نمایش همان گزارش تازه ساخته‌شده */

renderReportFromCache(
    reportId
); 
        
        resultCount.textContent =
            "تعداد ردیف‌های نتیجه: " +
            filteredRows.length;


        let html =
            "<table>" +
            "<thead>" +
            "<tr>";


        DATA.headers.forEach(
            function(name) {

                html +=
                    "<th>" +
                    escapeHtml(
                        String(name)
                    ) +
                    "</th>";

            }
        );


        html +=
            "</tr>" +
            "</thead>" +
            "<tbody>";


        filteredRows.forEach(
            function(row) {

                html += "<tr>";


                DATA.headers.forEach(
                    function(name, index) {

                        const value =
                            row[
                                String(
                                    index + 1
                                )
                            ] ?? "";


                        html +=
                            "<td>" +
                            escapeHtml(
                                String(value)
                            ) +
                            "</td>";

                    }
                );


                html +=
                    "</tr>";

            }
        );


        html +=
            "</tbody>" +
            "</table>";


        if (
            filteredRows.length === 0
        ) {

            html =
                '<div style="padding:12px;text-align:center;">' +
                'هیچ ردیفی با فیلترهای انتخاب‌شده پیدا نشد.' +
                '</div>';
        }


        resultTable.innerHTML =
            html;


        resultTableBox.style.display =
            "block";


        resultArea.classList.add(
            "result-visible"
        );
        resultArea.style.display = "block";

    }
);


/* =========================================================
   ESCAPE HTML
   ========================================================= */

function escapeHtml(value) {

    return value
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}
/* =========================================================
   TAB GUARDIAN — REPORT BOX FRAME
   فقط وقتی Report Box فعال است
   ========================================================= */

function TabGuardianReportFrame() {

    if (
        !resultTableBox ||
        !resultTable
    ) {
        return;
    }


    /* اگر قاب قبلاً ساخته شده، دوباره نساز */

    if (
        document.getElementById(
            "betaFarmReportHeader"
        )
    ) {
        return;
    }


    /* =====================================================
       Report Box
       ===================================================== */

    resultTableBox.style.display =
        "flex";

    resultTableBox.style.flexDirection =
        "column";

    resultTableBox.style.overflow =
        "hidden";

    resultTableBox.style.boxSizing =
        "border-box";


    /* =====================================================
       نوار آبی بالای گزارش
       ===================================================== */

    const header =
        document.createElement(
            "div"
        );

    header.id =
        "betaFarmReportHeader";

    header.style.flex =
        "0 0 32px";

    header.style.height =
        "32px";

    header.style.width =
        "100%";

    header.style.display =
        "flex";

    header.style.alignItems =
        "center";

    header.style.justifyContent =
        "flex-start";

    header.style.gap =
        "5px";

    header.style.padding =
        "3px 6px";

    header.style.background =
        "#3f6f9f";

    header.style.borderBottom =
        "1px solid #315b83";

    header.style.boxSizing =
        "border-box";

    header.style.direction =
        "ltr";


    /* =====================================================
       ساخت آیکون
       ===================================================== */

    function createControl(
        icon,
        title
    ) {

        const button =
            document.createElement(
                "button"
            );

        button.type =
            "button";

        button.textContent =
            icon;

        button.title =
            title;

        button.className =
            "filter-report-tab-control";

        button.style.width =
            "24px";

        button.style.height =
            "24px";

        button.style.padding =
            "0";

        button.style.margin =
            "0";

        button.style.display =
            "flex";

        button.style.alignItems =
            "center";

        button.style.justifyContent =
            "center";

        button.style.boxSizing =
            "border-box";

        button.style.background =
            "#ffffff";

        button.style.border =
            "1px solid #c4d0db";

        button.style.borderRadius =
            "4px";

        button.style.transition =
            "background-color 0.15s ease, border-color 0.15s ease";

        return button;
    }


    /* =====================================================
       پنج آیکون
       ===================================================== */

    const minimizeBtn =
        createControl(
            "□",
            "Maximize"
        );


    const exitBtn =
        createControl(
            "×",
            "Exit"
        );

    exitBtn.style.color =
        "#d62828";


    const toolsBtn =
        createControl(
            "🛠",
            "Tools"
        );


    const printBtn =
        createControl(
            "🖨",
            "Print"
        );


    const shareBtn =
        createControl(
            "🔗",
            "Share"
        );
       

    header.appendChild(
        minimizeBtn
    );

    header.appendChild(
        exitBtn
    );

    header.appendChild(
        toolsBtn
    );

    header.appendChild(
        printBtn
    );

    header.appendChild(
        shareBtn
    );


    /* =====================================================
       ناحیه محتوای گزارش
       ===================================================== */

    const body =
        document.createElement(
            "div"
        );

    body.id =
        "betaFarmReportBody";

    body.style.flex =
        "1 1 auto";

    body.style.minHeight =
        "0";

    body.style.minWidth =
        "0";

    body.style.width =
        "100%";

    body.style.overflow =
        "auto";

    body.style.boxSizing =
        "border-box";

    body.style.background =
        "#ffffff";


    /* =====================================================
       همان resultTable موجود
       ===================================================== */

    resultTable.style.display =
        "block";

    resultTable.style.width =
        "100%";

    resultTable.style.maxWidth =
        "none";

    resultTable.style.boxSizing =
        "border-box";


    body.appendChild(
        resultTable
    );


    /* =====================================================
       قرار دادن Header + Body داخل Report Box
       ===================================================== */

       resultTableBox.appendChild(
        body
    );

        /* =====================================================
       MAXIMIZE / RESTORE
       ===================================================== */

    minimizeBtn.addEventListener(
        "click",
        function(event) {

            event.stopPropagation();


            /* =================================================
               اگر Maximize باز است → Restore
               ================================================= */

            const existingLayer4 =
                document.getElementById(
                    "tabGuardianLayer4"
                );


            if (
                existingLayer4
            ) {

                existingLayer4.remove();


                minimizeBtn.textContent =
                    "□";

                minimizeBtn.title =
                    "Maximize";

                return;
            }


            /* =================================================
               ساخت پنجره Maximize
               ================================================= */

            const layer4 =
                document.createElement(
                    "div"
                );

            layer4.id =
                "tabGuardianLayer4";


            layer4.style.position =
                "fixed";

            layer4.style.left =
                "20px";

            layer4.style.right =
                "20px";

            layer4.style.top =
                "20px";

            layer4.style.bottom =
                "20px";

            layer4.style.zIndex =
                "99999";

            layer4.style.background =
                "#ffffff";

            layer4.style.border =
                "1px solid #aeb8c1";

            layer4.style.borderRadius =
                "8px";

            layer4.style.boxShadow =
                "0 4px 20px rgba(0,0,0,0.25)";

            layer4.style.display =
                "flex";

            layer4.style.flexDirection =
                "column";

            layer4.style.boxSizing =
                "border-box";


            /* =================================================
               نوار آبی Maximize
               ================================================= */

            const layer4Header =
                document.createElement(
                    "div"
                );

            layer4Header.style.flex =
                "0 0 32px";

            layer4Header.style.height =
                "32px";

            layer4Header.style.width =
                "100%";

            layer4Header.style.display =
                "flex";

            layer4Header.style.alignItems =
                "center";

            layer4Header.style.justifyContent =
                "flex-start";

            layer4Header.style.gap =
                "5px";

            layer4Header.style.padding =
                "3px 6px";

            layer4Header.style.background =
                "#3f6f9f";

            layer4Header.style.borderBottom =
                "1px solid #315b83";

            layer4Header.style.boxSizing =
                "border-box";

            layer4Header.style.direction =
                "ltr";


            /* =================================================
               پنج آیکون
               ================================================= */

            const restoreBtn =
                createControl(
                    "❐",
                    "Restore"
                );


            const maxExitBtn =
                createControl(
                    "×",
                    "Exit"
                );

            maxExitBtn.style.color =
                "#d62828";


            const maxToolsBtn =
                createControl(
                    "🛠",
                    "Tools"
                );


            const maxPrintBtn =
                createControl(
                    "🖨",
                    "Print"
                );


            const maxShareBtn =
                createControl(
                    "🔗",
                    "Share"
                );
            
                        /* =================================================
               Maximize:
               هر پنج آیکون فعال
               ================================================= */

            
            layer4Header.appendChild(
                restoreBtn
            );

            layer4Header.appendChild(
                maxExitBtn
            );

            layer4Header.appendChild(
                maxToolsBtn
            );

            layer4Header.appendChild(
                maxPrintBtn
            );

            layer4Header.appendChild(
                maxShareBtn
            );

                        /* =================================================
               بعد از Maximize هر پنج آیکون فعال هستند
               ================================================= */

            maxExitBtn.disabled =
                false;

            maxToolsBtn.disabled =
                false;

            maxPrintBtn.disabled =
                false;

            maxShareBtn.disabled =
                false;


            maxExitBtn.style.cursor =
                "pointer";

            maxToolsBtn.style.cursor =
                "pointer";

            maxPrintBtn.style.cursor =
                "pointer";

            maxShareBtn.style.cursor =
                "pointer";
            
                        /* =================================================
               ناحیهٔ محتوای گزارش
               ================================================= */

            const layer4Content =
                document.createElement(
                    "div"
                );
            window.BF_LAYER4_CONTENT = layer4Content; 
            
            layer4Content.style.flex =
                "1 1 auto";

            layer4Content.style.minHeight =
                "0";

            layer4Content.style.minWidth =
                "0";

            layer4Content.style.width =
                "100%";

            layer4Content.style.overflow =
                "auto";

            layer4Content.style.padding =
                "10px";

            layer4Content.style.boxSizing =
                "border-box";


            /* =================================================
               سطح گزارش
               ================================================= */
            const reportSurface =
                document.createElement(
                    "div"
                );

            reportSurface.style.position = "relative";
            reportSurface.style.width = "100%";
            reportSurface.style.height = "auto";
            reportSurface.style.minHeight = "100%";
            reportSurface.style.boxSizing = "border-box";

            /* =================================================
               کپی محتوای فعلی گزارش
               ================================================= */

            const reportCopy =
                resultTable.cloneNode(
                    true
                );

            reportCopy.style.display =
                "block";

            reportCopy.style.width =
                "100%";

            reportCopy.style.maxWidth =
                "none";

            reportCopy.style.minWidth =
                "100%";

            reportCopy.style.height =
                "auto";

            reportCopy.style.maxHeight =
                "none";

            reportCopy.style.overflow =
                "visible";

            reportCopy.style.boxSizing =
                "border-box";


            const copiedTables =
                reportCopy.querySelectorAll(
                    "table"
                );


            copiedTables.forEach(
                function(table) {

                    table.style.width =
                        "100%";

                    table.style.maxWidth =
                        "none";

                    table.style.boxSizing =
                        "border-box";

                }
            );


            reportSurface.appendChild(
                reportCopy
            );

            
            /* =================================================
               قرار دادن سطح گزارش در ناحیه اسکرول
               ================================================= */

            layer4Content.appendChild(
                reportSurface
            );
            /* =================================================
               اتصال Header + Body
               ================================================= */

            layer4.appendChild(
                layer4Header
            );

            layer4.appendChild(
                layer4Content
            );


            document.body.appendChild(
                layer4
            );
                        /* =================================================
               TOOLS
               فقط بعد از کلیک روی 🛠
               ================================================= */

            /* =================================================
   TOOLS
   Fabric.js → Layer 5
   ================================================= */

maxToolsBtn.addEventListener(
    "click",
    function(event) {

        event.stopPropagation();

        /* =================================================
           LAYER 5
           فقط با کلیک روی Tools ساخته می‌شود
           ================================================= */

        let layer5 =
            document.getElementById(
                "betaFarmLayer5"
            );

        if (!layer5) {

            layer5 =
                document.createElement(
                    "canvas"
                );

            layer5.id =
                "betaFarmLayer5";

            layer5.style.position =
                "absolute";

            layer5.style.left =
                "0";

            layer5.style.top =
                "0";

            layer5.style.width =
                "100%";

            layer5.style.height =
                "100%";

            layer5.style.zIndex =
                "5";

            layer5.style.background =
                "transparent";

            layer5.style.pointerEvents =
                "auto";

            reportSurface.appendChild(
                layer5
            );
        }

          /* =================================================
           FABRIC CANVAS
           فقط یک بار ساخته شود
           ================================================= */

        requestAnimationFrame(
            function() {

                const width =
                    Math.max(
                        1,
                        reportSurface.scrollWidth
                    );

                const height =
                    Math.max(
                        1,
                        reportSurface.scrollHeight,
                        layer4Content.clientHeight
                    );


                /* =================================================
                   اندازه Layer 5
                   ================================================= */

                layer5.width =
                    width;

                layer5.height =
                    height;

                layer5.style.width =
                    width + "px";

                layer5.style.height =
                    height + "px";


                /* =================================================
                   FABRIC CANVAS
                   ================================================= */

                if (
                    window.BF_PAINT_CANVAS &&
                    window.BF_PAINT_CANVAS.wrapperEl &&
                    !document.body.contains(
                        window.BF_PAINT_CANVAS.wrapperEl
                    )
                ) {

                    try {

                        window.BF_PAINT_CANVAS.dispose();

                    } catch (e) {

                        console.log(
                            "BF.Paint dispose warning:",
                            e
                        );

                    }

                    window.BF_PAINT_CANVAS =
                        null;
                }


                /* =================================================
                   ساخت Fabric فقط یک بار
                   ================================================= */

                if (
                    !window.BF_PAINT_CANVAS &&
                    typeof window.fabric !==
                        "undefined"
                ) {

                    window.BF_PAINT_CANVAS =
                        new window.fabric.Canvas(
                            layer5,
                            {
                                selection:
                                    false,

                                preserveObjectStacking:
                                    true
                            }
                        );


                    window.BF_PAINT_CANVAS.setDimensions(
                        {
                            width:
                                width,

                            height:
                                height
                        }
                    );


                    /* =================================================
                       سطح تعاملی Fabric روی گزارش
                       ================================================= */

                    if (
                        window.BF_PAINT_CANVAS.wrapperEl
                    ) {

                        window.BF_PAINT_CANVAS.wrapperEl.style.position =
                            "absolute";

                        window.BF_PAINT_CANVAS.wrapperEl.style.left =
                            "0";

                        window.BF_PAINT_CANVAS.wrapperEl.style.top =
                            "0";

                        window.BF_PAINT_CANVAS.wrapperEl.style.width =
                            width + "px";

                        window.BF_PAINT_CANVAS.wrapperEl.style.height =
                            height + "px";

                        window.BF_PAINT_CANVAS.wrapperEl.style.zIndex =
                            "1000";

                        window.BF_PAINT_CANVAS.wrapperEl.style.pointerEvents =
                            "auto";
                    }


                    if (
                        window.BF_PAINT_CANVAS.upperCanvasEl
                    ) {

                        window.BF_PAINT_CANVAS.upperCanvasEl.style.position =
                            "absolute";

                        window.BF_PAINT_CANVAS.upperCanvasEl.style.left =
                            "0";

                        window.BF_PAINT_CANVAS.upperCanvasEl.style.top =
                            "0";

                        window.BF_PAINT_CANVAS.upperCanvasEl.style.zIndex =
                            "1001";

                        window.BF_PAINT_CANVAS.upperCanvasEl.style.pointerEvents =
                            "auto";
                    }


                    /* =================================================
                       PEN BRUSH
                       ================================================= */

                    window.BF_PAINT_CANVAS.freeDrawingBrush =
                        new window.fabric.PencilBrush(
                            window.BF_PAINT_CANVAS
                        );

                    window.BF_PAINT_CANVAS.freeDrawingBrush.width =
                        3;


                    window.BF_PAINT_CANVAS.isDrawingMode =
                        false;


                    window.BF_PAINT_TOOL =
                        "NONE";

                                    /* =================================================
                   BF.Paint — HISTORY
                   Undo / Redo
                   ================================================= */

                window.BF_PAINT_HISTORY = [];

                window.BF_PAINT_HISTORY_INDEX =
                    -1;

                window.BF_PAINT_HISTORY_TIMER =
                    null;

                window.BF_PAINT_HISTORY_LOADING =
                    false;


                /* -----------------------------------------
                   ثبت وضعیت Canvas
                   ----------------------------------------- */

                window.BF_PAINT_SAVE_HISTORY =
                    function() {

                        const canvas =
                            window.BF_PAINT_CANVAS;

                        if (
                            !canvas ||
                            window.BF_PAINT_HISTORY_LOADING
                        ) {

                            return;

                        }


                        if (
                            window.BF_PAINT_HISTORY_TIMER
                        ) {

                            clearTimeout(
                                window.BF_PAINT_HISTORY_TIMER
                            );

                        }


                        window.BF_PAINT_HISTORY_TIMER =
                            setTimeout(
                                function() {

                                    if (
                                        window.BF_PAINT_HISTORY_LOADING
                                    ) {

                                        return;

                                    }


                                    const state =
                                        canvas.toJSON();

                                    const stateText =
                                        JSON.stringify(
                                            state
                                        );


                                    const lastState =
                                        window.BF_PAINT_HISTORY.length >
                                        0 &&
                                        window.BF_PAINT_HISTORY[
                                            window.BF_PAINT_HISTORY.length - 1
                                        ];


                                    if (
                                        lastState &&
                                        lastState.text ===
                                            stateText
                                    ) {

                                        return;

                                    }


                                    /* ---------------------------------
                                       اگر Undo شده و تغییر جدیدی انجام شد
                                       Redoهای بعدی حذف شوند
                                       --------------------------------- */

                                    if (
                                        window.BF_PAINT_HISTORY_INDEX <
                                        window.BF_PAINT_HISTORY.length - 1
                                    ) {

                                        window.BF_PAINT_HISTORY =
                                            window.BF_PAINT_HISTORY.slice(
                                                0,
                                                window.BF_PAINT_HISTORY_INDEX + 1
                                            );

                                    }


                                    window.BF_PAINT_HISTORY.push(
                                        {
                                            text:
                                                stateText,

                                            data:
                                                state
                                        }
                                    );


                                    /* حداکثر 50 وضعیت */

                                    if (
                                        window.BF_PAINT_HISTORY.length >
                                        50
                                    ) {

                                        window.BF_PAINT_HISTORY.shift();

                                    }


                                    window.BF_PAINT_HISTORY_INDEX =
                                        window.BF_PAINT_HISTORY.length - 1;


                                },
                                100
                            );

                    };


                /* -----------------------------------------
                   وضعیت اولیه Canvas
                   ----------------------------------------- */

                window.BF_PAINT_HISTORY.push(
                    {
                        text:
                            JSON.stringify(
                                window.BF_PAINT_CANVAS.toJSON()
                            ),

                        data:
                            window.BF_PAINT_CANVAS.toJSON()
                    }
                );


                window.BF_PAINT_HISTORY_INDEX =
                    0;


                /* -----------------------------------------
                   تغییرات آبجکت
                   ----------------------------------------- */

                window.BF_PAINT_CANVAS.on(
                    "object:added",
                    function() {

                        window.BF_PAINT_SAVE_HISTORY();

                    }
                );


                window.BF_PAINT_CANVAS.on(
                    "object:modified",
                    function() {

                        window.BF_PAINT_SAVE_HISTORY();

                    }
                );


                window.BF_PAINT_CANVAS.on(
                    "object:removed",
                    function() {

                        window.BF_PAINT_SAVE_HISTORY();

                    }
                );


                /* -----------------------------------------
                   پایان رسم ابزارهای Drawing
                   ----------------------------------------- */

                window.BF_PAINT_CANVAS.on(
                    "mouse:up",
                    function() {

                        window.BF_PAINT_SAVE_HISTORY();

                    }
                );


                window.BF_PAINT_CANVAS.on(
                    "path:created",
                    function() {

                        window.BF_PAINT_SAVE_HISTORY();

                    }
                );


                /* -----------------------------------------
                   UNDO
                   ----------------------------------------- */

                window.BF_PAINT_UNDO =
                    function() {

                        const canvas =
                            window.BF_PAINT_CANVAS;

                        if (
                            !canvas
                        ) {

                            return;

                        }


                        if (
                            window.BF_PAINT_HISTORY_INDEX <=
                            0
                        ) {

                            return;

                        }


                        window.BF_PAINT_HISTORY_INDEX--;


                        const snapshot =
                            window.BF_PAINT_HISTORY[
                                window.BF_PAINT_HISTORY_INDEX
                            ];


                        window.BF_PAINT_HISTORY_LOADING =
                            true;


                        canvas.discardActiveObject();


                        canvas.loadFromJSON(
                            snapshot.data
                        ).then(
                            function() {

                                canvas.requestRenderAll();

                                window.BF_PAINT_HISTORY_LOADING =
                                    false;

                            }
                        );

                    };


                /* -----------------------------------------
                   REDO
                   ----------------------------------------- */

                window.BF_PAINT_REDO =
                    function() {

                        const canvas =
                            window.BF_PAINT_CANVAS;

                        if (
                            !canvas
                        ) {

                            return;

                        }


                        if (
                            window.BF_PAINT_HISTORY_INDEX >=
                            window.BF_PAINT_HISTORY.length - 1
                        ) {

                            return;

                        }


                        window.BF_PAINT_HISTORY_INDEX++;


                        const snapshot =
                            window.BF_PAINT_HISTORY[
                                window.BF_PAINT_HISTORY_INDEX
                            ];


                        window.BF_PAINT_HISTORY_LOADING =
                            true;


                        canvas.discardActiveObject();


                        canvas.loadFromJSON(
                            snapshot.data
                        ).then(
                            function() {

                                canvas.requestRenderAll();

                                window.BF_PAINT_HISTORY_LOADING =
                                    false;

                            }
                        );

                    };  
                    console.log(
                        "BF.Paint — Fabric Canvas READY"
                    );

                }

            }
        );


         

        /* =================================================
           اگر Toolbar قبلاً ساخته شده،
           دوباره نساز
           ================================================= */

        if (
            document.getElementById(
                "betaFarmPaintToolbar"
            )
        ) {

            return;

        }


        /* =================================================
           ساخت Toolbar
           ================================================= */

        const paintToolbar =
            document.createElement(
                "div"
            );

        paintToolbar.id =
            "betaFarmPaintToolbar";

        paintToolbar.style.flex =
            "0 0 auto";

        paintToolbar.style.width =
            "100%";

        paintToolbar.style.display =
            "flex";

        paintToolbar.style.alignItems =
            "center";

        paintToolbar.style.gap =
            "4px";

        paintToolbar.style.padding =
            "4px 6px";

        paintToolbar.style.background =
            "#f4f6f8";

        paintToolbar.style.borderBottom =
            "1px solid #c6ced5";

        paintToolbar.style.boxSizing =
            "border-box";

        paintToolbar.style.direction =
            "ltr";


        /* =================================================
           ابزارهای Paint
           ================================================= */

        const paintTools = [
            ["✎", "Pen"],
            ["⌫", "Eraser"],
            ["╱", "Line"],
            ["○", "Circle"],
            ["□", "Rectangle"],
            ["➜", "Arrow"],
            ["T", "Text"],
            ["↶", "Undo"],
            ["↷", "Redo"],
            ["🗑", "Clear"],
            ["📷", "Camera"],
            ["💾", "Save"]
        ];


        paintTools.forEach(
            function(toolData) {

                const toolBtn =
                    document.createElement(
                        "button"
                    );

                toolBtn.type =
                    "button";

                toolBtn.className =
                    "filter-report-tab-control";

                toolBtn.textContent =
                    toolData[0];

                toolBtn.title =
                    toolData[1];

                toolBtn.dataset.betaPaintTool =
                    toolData[1];

                toolBtn.style.width =
                    "28px";

                toolBtn.style.height =
                    "28px";

                toolBtn.style.padding =
                    "0";

                toolBtn.style.display =
                    "flex";

                toolBtn.style.alignItems =
                    "center";

                toolBtn.style.justifyContent =
                    "center";

                toolBtn.style.background =
                    "#ffffff";

                toolBtn.style.border =
                    "1px solid #b9c5ce";

                toolBtn.style.borderRadius =
                    "4px";

                toolBtn.style.cursor =
                    "pointer";


                /* -----------------------------------------
                   PEN
                   فعلاً فقط همین ابزار فعال است
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Pen"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            const canvas =
                                window.BF_PAINT_CANVAS;

                            if (
                                !canvas
                            ) {

                                return;

                            }


                            canvas.isDrawingMode =
                                true;


                            window.BF_PAINT_TOOL =
                                "PEN";


                            canvas.selection =
                                false;


                            canvas.freeDrawingBrush =
                                canvas.freeDrawingBrush ||
                                new window.fabric.PencilBrush(
                                    canvas
                                );


                            canvas.freeDrawingBrush.width =
                                3;


                            console.log(
                                "BF.Paint TOOL: PEN"
                            );

                        }
                    );

                }

                /* -----------------------------------------
                   ERASER
                   پاک کردن آبجکت‌های Paint
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Eraser"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            const canvas =
                                window.BF_PAINT_CANVAS;

                            if (
                                !canvas
                            ) {

                                return;

                            }


                            canvas.isDrawingMode =
                                false;

                            canvas.selection =
                                false;


                            window.BF_PAINT_TOOL =
                                "ERASER";


                            /* ---------------------------------
                               رویدادهای Eraser فقط یک بار وصل شوند
                               --------------------------------- */

                            if (
                                canvas.__BF_ERASER_BOUND
                            ) {

                                console.log(
                                    "BF.Paint TOOL: ERASER"
                                );

                                return;

                            }


                            canvas.__BF_ERASER_BOUND =
                                true;


                            canvas.__BF_ERASER_DOWN =
                                false;


                            canvas.on(
                                "mouse:down",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "ERASER"
                                    ) {
                                        return;
                                    }


                                    canvas.__BF_ERASER_DOWN =
                                        true;


                                    const target =
                                        canvas.findTarget(
                                            opt.e
                                        );


                                    if (
                                        target
                                    ) {

                                        canvas.remove(
                                            target
                                        );

                                        canvas.requestRenderAll();
                                    }

                                }
                            );


                            canvas.on(
                                "mouse:move",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "ERASER"
                                    ) {
                                        return;
                                    }


                                    if (
                                        !canvas.__BF_ERASER_DOWN
                                    ) {
                                        return;
                                    }


                                    if (
                                        !opt.e ||
                                        !opt.e.buttons
                                    ) {
                                        return;
                                    }


                                    const target =
                                        canvas.findTarget(
                                            opt.e
                                        );


                                    if (
                                        target
                                    ) {

                                        canvas.remove(
                                            target
                                        );

                                        canvas.requestRenderAll();
                                    }

                                }
                            );


                            canvas.on(
                                "mouse:up",
                                function() {

                                    canvas.__BF_ERASER_DOWN =
                                        false;

                                }
                            );


                            console.log(
                                "BF.Paint TOOL: ERASER"
                            );

                        }
                    );

                } 
                /* -----------------------------------------
   LINE
   رسم خط آزاد با Fabric
   ----------------------------------------- */

if (
    toolData[1] ===
    "Line"
) {

    toolBtn.addEventListener(
        "click",
        function() {

            const canvas =
                window.BF_PAINT_CANVAS;

            if (
                !canvas
            ) {

                return;

            }


            /* ---------------------------------
               فعال کردن ابزار Line
               --------------------------------- */

            canvas.isDrawingMode =
                false;

            canvas.selection =
                false;

            canvas.discardActiveObject();


            /* ---------------------------------
               آبجکت‌های قبلی دوباره قابل انتخاب
               باشند
               --------------------------------- */

            canvas.forEachObject(
                function(object) {

                    object.selectable =
                        true;

                    object.evented =
                        true;

                }
            );


            window.BF_PAINT_TOOL =
                "LINE";
            /* ---------------------------------
   اتصال رویدادهای Line فقط یک بار
   --------------------------------- */

if (
    !canvas.__BF_LINE_BOUND
) {

    canvas.__BF_LINE_BOUND =
        true;

    canvas.__BF_LINE_START =
        null;

    canvas.__BF_LINE_OBJECT =
        null;


    canvas.on(
        "mouse:down",
        function(opt) {

            if (
                window.BF_PAINT_TOOL !==
                "LINE"
            ) {

                return;

            }


            const pointer =
                opt.scenePoint;
                


            canvas.__BF_LINE_START = {
                x: pointer.x,
                y: pointer.y
            };


            canvas.__BF_LINE_OBJECT =
                new window.fabric.Line(
                    [
                        pointer.x,
                        pointer.y,
                        pointer.x,
                        pointer.y
                    ],
                    {
                        stroke: "#222222",
                        strokeWidth: 3,
                        selectable: true,
                        evented: true
                    }
                );


            canvas.add(
                canvas.__BF_LINE_OBJECT
            );

            canvas.requestRenderAll();

        }
    );


    canvas.on(
        "mouse:move",
        function(opt) {

            if (
                window.BF_PAINT_TOOL !==
                "LINE"
            ) {

                return;

            }


            if (
                !canvas.__BF_LINE_OBJECT
            ) {

                return;

            }


            const pointer =
                opt.scenePoint;
                


            canvas.__BF_LINE_OBJECT.set({
                x2:
                    pointer.x,

                y2:
                    pointer.y
            });


            canvas.requestRenderAll();

        }
    );


    canvas.on(
        "mouse:up",
        function() {

            if (
                window.BF_PAINT_TOOL !==
                "LINE"
            ) {

                return;

            }


            canvas.__BF_LINE_START =
                null;

            canvas.__BF_LINE_OBJECT =
                null;

            canvas.requestRenderAll();

        }
    );

} 

            console.log(
                "BF.Paint TOOL: LINE"
            );

        }
    );

} 

/* -----------------------------------------
   CAMERA
   ذخیره نمای قابل‌دید فعلی گزارش
   ----------------------------------------- */

if (
    toolData[1] ===
    "Camera"
) {

    toolBtn.addEventListener(
        "click",
        function(event) {

            event.stopPropagation();

            if (
                typeof window.html2canvas !==
                "function"
            ) {

                console.error(
                    "BF.Paint: html2canvas NOT READY"
                );

                return;
            }

            if (
                !reportSurface ||
                !layer4Content
            ) {

                return;
            }

            const contentStyle =
                window.getComputedStyle(
                    layer4Content
                );

            const paddingLeft =
                parseFloat(
                    contentStyle.paddingLeft
                ) || 0;

            const paddingRight =
                parseFloat(
                    contentStyle.paddingRight
                ) || 0;

            const paddingTop =
                parseFloat(
                    contentStyle.paddingTop
                ) || 0;

            const paddingBottom =
                parseFloat(
                    contentStyle.paddingBottom
                ) || 0;

            const visibleWidth =
                Math.max(
                    1,
                    layer4Content.clientWidth -
                    paddingLeft -
                    paddingRight
                );

            const visibleHeight =
                Math.max(
                    1,
                    layer4Content.clientHeight -
                    paddingTop -
                    paddingBottom
                );

            const surfaceRect =
                reportSurface.getBoundingClientRect();

            const viewportRect =
                layer4Content.getBoundingClientRect();

            const scale =
                2;

            window.html2canvas(
                reportSurface,
                {
                    backgroundColor:
                        "#ffffff",

                    useCORS:
                        true,

                    allowTaint:
                        false,

                    scale:
                        scale,

                    logging:
                        false,

                    imageTimeout:
                        0
                }
            ).then(
                function(fullCanvas) {
                   
                    console.log(
                         "BF CIRCLE fullCanvas:",
                         "width =", fullCanvas.width,
                         "height =", fullCanvas.height,
                         "canvas =", fullCanvas
                    );

 
                    const sourceX =
                        Math.max(
                            0,
                            viewportRect.left -
                            surfaceRect.left
                        ) *
                        scale;

                    const sourceY =
                        Math.max(
                            0,
                            viewportRect.top -
                            surfaceRect.top
                        ) *
                        scale;

                    const sourceWidth =
                        visibleWidth *
                        scale;

                    const sourceHeight =
                        visibleHeight *
                        scale;

                    const finalCanvas =
                        document.createElement(
                            "canvas"
                        );

                    finalCanvas.width =
                        sourceWidth;

                    finalCanvas.height =
                        sourceHeight;

                    const finalContext =
                        finalCanvas.getContext(
                            "2d"
                        );

                    if (
                        !finalContext
                    ) {

                        return;
                    }

                    finalContext.fillStyle =
                        "#ffffff";

                    finalContext.fillRect(
                        0,
                        0,
                        sourceWidth,
                        sourceHeight
                    );

                    finalContext.drawImage(
                        fullCanvas,

                        sourceX,
                        sourceY,

                        sourceWidth,
                        sourceHeight,

                        0,
                        0,

                        sourceWidth,
                        sourceHeight
                    );

                    const imageData =
                        finalCanvas.toDataURL(
                            "image/png"
                        );

                    if (
                        window.betaFarmBridge &&
                        typeof window.betaFarmBridge.saveReportCapture ===
                            "function"
                    ) {

                        window.betaFarmBridge.saveReportCapture(
                            imageData
                        );

                    }

                }
            ).catch(
                function(error) {

                    console.error(
                        "BF.Paint CAMERA ERROR:",
                        error
                    );

                }
            );

        }
    );

}

       
                /* -----------------------------------------
                   SAVE
                   Filter  → Excel / Image
                   Circle  → Image only
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Save"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function(event) {

                            event.stopPropagation();


                            /* =================================================
                               تشخیص نوع گزارش
                               ================================================= */

                            const isQuarantineReport =
                                !resultCount.textContent.trim();


                            /* =================================================
                               تابع ذخیره تصویر نمای قابل‌دید
                               ================================================= */

                            function saveVisibleReportImage() {

                                if (
                                    typeof window.html2canvas !==
                                    "function"
                                ) {

                                    console.error(
                                        "BF.Paint: html2canvas NOT READY"
                                    );

                                    return;

                                }


                                if (
                                    !reportSurface ||
                                    !layer4Content
                                ) {

                                    return;

                                }


                                const contentStyle =
                                    window.getComputedStyle(
                                        layer4Content
                                    );


                                const paddingLeft =
                                    parseFloat(
                                        contentStyle.paddingLeft
                                    ) || 0;

                                const paddingRight =
                                    parseFloat(
                                        contentStyle.paddingRight
                                    ) || 0;

                                const paddingTop =
                                    parseFloat(
                                        contentStyle.paddingTop
                                    ) || 0;

                                const paddingBottom =
                                    parseFloat(
                                        contentStyle.paddingBottom
                                    ) || 0;


                                const visibleWidth =
                                    Math.max(
                                        1,
                                        layer4Content.clientWidth -
                                        paddingLeft -
                                        paddingRight
                                    );


                                const visibleHeight =
                                    Math.max(
                                        1,
                                        layer4Content.clientHeight -
                                        paddingTop -
                                        paddingBottom
                                    );
                                const surfaceRect =
                                    reportSurface.getBoundingClientRect();

                                const viewportRect =
                                    layer4Content.getBoundingClientRect(); 

                                

                                /* -----------------------------------------
                                   Render کامل Report
                                   ----------------------------------------- */
                                 
                                window.html2canvas(
                                    reportSurface,
                                    {
                                        backgroundColor:
                                            "#ffffff",

                                        useCORS:
                                            true,

                                        allowTaint:
                                            false,

                                        scale:
                                            2,

                                        logging:
                                            false,

                                        imageTimeout:
                                            0
                                    }
                                ).then(
                                    function(fullCanvas) {
                                            /* =========================================
           TEST 1
           ذخیره مستقیم fullCanvas
           فقط برای Circle
           ========================================= */

        if (
            window.__BF_CIRCLE_FULL_TEST
        ) {

            window.__BF_CIRCLE_FULL_TEST =
                false;


            const imageData =
                fullCanvas.toDataURL(
                    "image/png"
                );


            if (
                window.betaFarmBridge &&
                typeof window.betaFarmBridge.saveReportCapture ===
                    "function"
            ) {

                window.betaFarmBridge.saveReportCapture(
                    imageData
                );

            }


            return;

        }

                                        const scale =
                                            2;
                                        const sourceX =
                                            Math.max(
                                                0,
                                                viewportRect.left -
                                                surfaceRect.left
                                            ) *
                                            scale;

                                        const sourceY =
                                            Math.max(
                                                0,
                                                viewportRect.top -
                                                surfaceRect.top
                                            ) *
                                            scale;

                                        const sourceWidth =
                                            visibleWidth *
                                            scale;

                                        const sourceHeight =
                                            visibleHeight *
                                            scale;

                                                                   

                                        /* ---------------------------------
                                           Canvas نهایی Viewport
                                           --------------------------------- */

                                        const finalCanvas =
                                            document.createElement(
                                                "canvas"
                                            );


                                        finalCanvas.width =
                                            sourceWidth;

                                        finalCanvas.height =
                                            sourceHeight;


                                        const finalContext =
                                            finalCanvas.getContext(
                                                "2d"
                                            );


                                        if (
                                            !finalContext
                                        ) {

                                            return;

                                        }


                                        finalContext.fillStyle =
                                            "#ffffff";

                                        finalContext.fillRect(
                                            0,
                                            0,
                                            sourceWidth,
                                            sourceHeight
                                        );


                                        finalContext.drawImage(
                                            fullCanvas,

                                            sourceX,
                                            sourceY,

                                            sourceWidth,
                                            sourceHeight,

                                            0,
                                            0,

                                            sourceWidth,
                                            sourceHeight
                                        );


                                        const imageData =
                                            finalCanvas.toDataURL(
                                                "image/png"
                                            );


                                        if (
                                            window.betaFarmBridge &&
                                            typeof window.betaFarmBridge.saveReportCapture ===
                                                "function"
                                        ) {

                                            window.betaFarmBridge.saveReportCapture(
                                                imageData
                                            );

                                        }

                                    }
                                ).catch(
                                    function(error) {

                                        console.error(
                                            "BF.Paint SAVE IMAGE ERROR:",
                                            error
                                        );

                                    }
                                );

                            }


                            /* =================================================
                               CIRCLE REPORT
                               فقط تصویر
                               ================================================= */

                            if (
    isQuarantineReport
) {

    const oldCircleMenu =
        document.getElementById(
            "betaFarmCircleSaveMenu"
        );

    if (
        oldCircleMenu
    ) {

        oldCircleMenu.remove();

    }


    paintToolbar.style.position =
        "relative";


    const circleSaveMenu =
        document.createElement(
            "div"
        );

    circleSaveMenu.id =
        "betaFarmCircleSaveMenu";

    circleSaveMenu.style.position =
        "absolute";

    circleSaveMenu.style.left =
        toolBtn.offsetLeft +
        "px";

    circleSaveMenu.style.top =
        "34px";

    circleSaveMenu.style.width =
        "82px";

    circleSaveMenu.style.padding =
        "3px";

    circleSaveMenu.style.background =
        "#ffffff";

    circleSaveMenu.style.border =
        "1px solid #b9c5ce";

    circleSaveMenu.style.borderRadius =
        "6px";

    circleSaveMenu.style.boxShadow =
        "0 3px 10px rgba(0,0,0,0.18)";

    circleSaveMenu.style.zIndex =
        "2000";

    circleSaveMenu.style.boxSizing =
        "border-box";


    /* -----------------------------------------
       عکس
       ----------------------------------------- */

    const circleImageOption =
        document.createElement(
            "button"
        );

    circleImageOption.type =
        "button";

    circleImageOption.textContent =
        "عکس";

    circleImageOption.style.width =
        "100%";

    circleImageOption.style.height =
        "26px";

    circleImageOption.style.padding =
        "0";

    circleImageOption.style.margin =
        "0";

    circleImageOption.style.background =
        "#ffffff";

    circleImageOption.style.border =
        "0";

    circleImageOption.style.borderRadius =
        "4px";

    circleImageOption.style.cursor =
        "pointer";

    circleImageOption.style.fontFamily =
        "Tahoma";

    circleImageOption.style.fontSize =
        "12px";


    /* -----------------------------------------
       PDF
       ----------------------------------------- */

    const circlePdfOption =
        document.createElement(
            "button"
        );

    circlePdfOption.type =
        "button";

    circlePdfOption.textContent =
        "PDF";

    circlePdfOption.style.width =
        "100%";

    circlePdfOption.style.height =
        "26px";

    circlePdfOption.style.padding =
        "0";

    circlePdfOption.style.margin =
        "2px 0 0 0";

    circlePdfOption.style.background =
        "#ffffff";

    circlePdfOption.style.border =
        "0";

    circlePdfOption.style.borderRadius =
        "4px";

    circlePdfOption.style.cursor =
        "pointer";

    circlePdfOption.style.fontFamily =
        "Tahoma";

    circlePdfOption.style.fontSize =
        "12px";


    circleSaveMenu.appendChild(
        circleImageOption
    );

    circleSaveMenu.appendChild(
        circlePdfOption
    );


    paintToolbar.appendChild(
        circleSaveMenu
    );


    /* -----------------------------------------
       عکس
       ----------------------------------------- */

    circleImageOption.addEventListener(
        "click",
        function(event) {

            event.stopPropagation();

            circleSaveMenu.remove();

            window.__BF_CIRCLE_FULL_TEST =
                true;

            saveVisibleReportImage();

        }
    );


    /* -----------------------------------------
       PDF
       ----------------------------------------- */

    circlePdfOption.addEventListener(
    "click",
    function(event) {

        event.stopPropagation();

        circleSaveMenu.remove();

        if (
            window.betaFarmBridge &&
            typeof window.betaFarmBridge.saveReportPDF ===
                "function"
        ) {

            window.betaFarmBridge.saveReportPDF();

        } else {

            console.error(
                "BF.Paint: saveReportPDF NOT READY"
            );

        }

    }
);

    return;

}


                            /* =================================================
                               FILTER REPORT
                               Excel / Image
                               ================================================= */

                            const oldMenu =
                                document.getElementById(
                                    "betaFarmSaveMenu"
                                );


                            if (
                                oldMenu
                            ) {

                                oldMenu.remove();

                            }


                            paintToolbar.style.position =
                                "relative";


                            const saveMenu =
                                document.createElement(
                                    "div"
                                );

                            saveMenu.id =
                                "betaFarmSaveMenu";

                            saveMenu.style.position =
                                "absolute";

                            saveMenu.style.left =
                                toolBtn.offsetLeft +
                                "px";

                            saveMenu.style.top =
                                "34px";

                            saveMenu.style.width =
                                "82px";

                            saveMenu.style.padding =
                                "3px";

                            saveMenu.style.background =
                                "#ffffff";

                            saveMenu.style.border =
                                "1px solid #b9c5ce";

                            saveMenu.style.borderRadius =
                                "6px";

                            saveMenu.style.boxShadow =
                                "0 3px 10px rgba(0,0,0,0.18)";

                            saveMenu.style.zIndex =
                                "2000";

                            saveMenu.style.boxSizing =
                                "border-box";


                            /* -----------------------------------------
                               Excel
                               ----------------------------------------- */

                            const excelOption =
                                document.createElement(
                                    "button"
                                );

                            excelOption.type =
                                "button";

                            excelOption.textContent =
                                "اکسل";

                            excelOption.style.width =
                                "100%";

                            excelOption.style.height =
                                "26px";

                            excelOption.style.padding =
                                "0";

                            excelOption.style.margin =
                                "0";

                            excelOption.style.background =
                                "#ffffff";

                            excelOption.style.border =
                                "0";

                            excelOption.style.borderRadius =
                                "4px";

                            excelOption.style.cursor =
                                "pointer";

                            excelOption.style.fontFamily =
                                "Tahoma";

                            excelOption.style.fontSize =
                                "12px";


                            /* -----------------------------------------
                               عکس
                               ----------------------------------------- */

                            const imageOption =
                                document.createElement(
                                    "button"
                                );

                            imageOption.type =
                                "button";

                            imageOption.textContent =
                                "عکس";

                            imageOption.style.width =
                                "100%";

                            imageOption.style.height =
                                "26px";

                            imageOption.style.padding =
                                "0";

                            imageOption.style.margin =
                                "2px 0 0 0";

                            imageOption.style.background =
                                "#ffffff";

                            imageOption.style.border =
                                "0";

                            imageOption.style.borderRadius =
                                "4px";

                            imageOption.style.cursor =
                                "pointer";

                            imageOption.style.fontFamily =
                                "Tahoma";

                            imageOption.style.fontSize =
                                "12px";


                            saveMenu.appendChild(
                                excelOption
                            );

                            saveMenu.appendChild(
                                imageOption
                            );


                            paintToolbar.appendChild(
                                saveMenu
                            );


                            /* =================================================
                               Excel
                               ================================================= */

                            excelOption.addEventListener(
                                "click",
                                function(event) {

                                    event.stopPropagation();


                                    saveMenu.remove();


                                    const activeReport =
                                        REPORT_CACHE[
                                            ACTIVE_REPORT_ID
                                        ];


                                    if (
                                        !activeReport
                                    ) {

                                        return;

                                    }


                                    const reportData =
                                        activeReport.DATA ||
                                        activeReport;


                                    const excelData =
                                        JSON.stringify(
                                            {
                                                headers:
                                                    reportData.headers || [],

                                                rows:
                                                    reportData.rows || []
                                            }
                                        );


                                    if (
                                        window.betaFarmBridge &&
                                        typeof window.betaFarmBridge.saveReportExcel ===
                                            "function"
                                    ) {

                                        window.betaFarmBridge.saveReportExcel(
                                            excelData
                                        );

                                    }

                                }
                            );


                            /* =================================================
                               Image
                               ================================================= */

                            imageOption.addEventListener(
                                "click",
                                function(event) {

                                    event.stopPropagation();

                                    saveMenu.remove();

                                    saveVisibleReportImage();

                                }
                            );

                        }
                    );

                }


                paintToolbar.appendChild(
                    toolBtn
                );

         
                /* -----------------------------------------
                   CIRCLE
                   Right Drag = Create
                   Left Click = Select / Move / Resize
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Circle"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            const canvas =
                                window.BF_PAINT_CANVAS;

                            if (
                                !canvas
                            ) {

                                return;

                            }


                            /* ---------------------------------
                               فعال کردن Circle
                               --------------------------------- */

                            canvas.isDrawingMode =
                                false;

                            canvas.selection =
                                true;


                            canvas.discardActiveObject();


                            /* ---------------------------------
                               آبجکت‌های قبلی قابل انتخاب باشند
                               --------------------------------- */

                            canvas.forEachObject(
                                function(object) {

                                    object.selectable =
                                        true;

                                    object.evented =
                                        true;

                                }
                            );


                            window.BF_PAINT_TOOL =
                                "CIRCLE";


                            /* ---------------------------------
                               وضعیت رسم
                               --------------------------------- */

                            canvas.__BF_CIRCLE_DRAWING =
                                false;

                            canvas.__BF_CIRCLE_OBJECT =
                                null;

                            canvas.__BF_CIRCLE_CENTER =
                                null;


                            /* ---------------------------------
                               جلوگیری از Context Menu
                               فقط هنگام فعال بودن Circle
                               --------------------------------- */

                            if (
                                !canvas.__BF_CIRCLE_CONTEXT_BOUND
                            ) {

                                canvas.__BF_CIRCLE_CONTEXT_BOUND =
                                    true;


                                canvas.upperCanvasEl.addEventListener(
                                    "contextmenu",
                                    function(event) {

                                        if (
                                            window.BF_PAINT_TOOL ===
                                            "CIRCLE"
                                        ) {

                                            event.preventDefault();

                                        }

                                    }
                                );

                            }


                            /* ---------------------------------
                               رویدادهای Circle فقط یک بار
                               --------------------------------- */

                            if (
                                canvas.__BF_CIRCLE_BOUND
                            ) {

                                console.log(
                                    "BF.Paint TOOL: CIRCLE"
                                );

                                return;

                            }


                            canvas.__BF_CIRCLE_BOUND =
                                true;


                            /* ---------------------------------
                               شروع ساخت Circle
                               فقط با کلیک راست
                               --------------------------------- */

                            canvas.on(
                                "mouse:down",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "CIRCLE"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !opt.e ||
                                        opt.e.button !==
                                            2
                                    ) {

                                        return;

                                    }


                                    opt.e.preventDefault();


                                    const pointer =
                                        opt.scenePoint;


                                    canvas.__BF_CIRCLE_DRAWING =
                                        true;


                                    canvas.__BF_CIRCLE_CENTER = {
                                        x:
                                            pointer.x,

                                        y:
                                            pointer.y
                                    };


                                    const circle =
                                        new window.fabric.Circle(
                                            {
                                                left:
                                                    pointer.x,

                                                top:
                                                    pointer.y,

                                                radius:
                                                    1,

                                                originX:
                                                    "center",

                                                originY:
                                                    "center",

                                                fill:
                                                    "rgba(34,34,34,0.04)",

                                                stroke:
                                                    "#222222",

                                                strokeWidth:
                                                    3,

                                                selectable:
                                                    false,

                                                evented:
                                                    false
                                            }
                                        );


                                    canvas.__BF_CIRCLE_OBJECT =
                                        circle;


                                    canvas.add(
                                        circle
                                    );


                                    canvas.requestRenderAll();

                                }
                            );


                            /* ---------------------------------
                               تغییر شعاع فقط هنگام
                               Drag راست
                               --------------------------------- */

                            canvas.on(
                                "mouse:move",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "CIRCLE"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !canvas.__BF_CIRCLE_DRAWING
                                    ) {

                                        return;

                                    }


                                    if (
                                        !opt.e ||
                                        !opt.e.buttons ||
                                        (
                                            (opt.e.buttons & 2)
                                            !== 2
                                        )
                                    ) {

                                        return;

                                    }


                                    const circle =
                                        canvas.__BF_CIRCLE_OBJECT;

                                    const center =
                                        canvas.__BF_CIRCLE_CENTER;


                                    if (
                                        !circle ||
                                        !center
                                    ) {

                                        return;

                                    }


                                    const pointer =
                                        opt.scenePoint;


                                    const dx =
                                        pointer.x -
                                        center.x;

                                    const dy =
                                        pointer.y -
                                        center.y;


                                    const radius =
                                        Math.sqrt(
                                            (dx * dx) +
                                            (dy * dy)
                                        );


                                    circle.set({
                                        radius:
                                            Math.max(
                                                1,
                                                radius
                                            ),

                                        left:
                                            center.x,

                                        top:
                                            center.y
                                    });


                                    circle.setCoords();

                                    canvas.requestRenderAll();

                                }
                            );


                            /* ---------------------------------
                               پایان ساخت Circle
                               --------------------------------- */

                            canvas.on(
                                "mouse:up",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "CIRCLE"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !canvas.__BF_CIRCLE_DRAWING
                                    ) {

                                        return;

                                    }


                                    if (
                                        opt.e &&
                                        opt.e.button !==
                                            2
                                    ) {

                                        return;

                                    }


                                    const circle =
                                        canvas.__BF_CIRCLE_OBJECT;


                                    if (
                                        circle
                                    ) {

                                        circle.set({
                                            selectable:
                                                true,

                                            evented:
                                                true
                                        });


                                        circle.setCoords();

                                    }


                                    canvas.__BF_CIRCLE_DRAWING =
                                        false;

                                    canvas.__BF_CIRCLE_OBJECT =
                                        null;

                                    canvas.__BF_CIRCLE_CENTER =
                                        null;


                                    canvas.requestRenderAll();

                                }
                            );


                            console.log(
                                "BF.Paint TOOL: CIRCLE — RIGHT DRAG CREATE"
                            );

                        }
                    );

                }
                /* -----------------------------------------
   RECTANGLE
   Right Drag = Create
   Left Click = Select / Move / Resize
   ----------------------------------------- */

if (
    toolData[1] ===
    "Rectangle"
) {

    toolBtn.addEventListener(
        "click",
        function() {

            const canvas =
                window.BF_PAINT_CANVAS;

            if (
                !canvas
            ) {

                return;

            }


            /* ---------------------------------
               فعال کردن Rectangle
               --------------------------------- */

            canvas.isDrawingMode =
                false;

            canvas.selection =
                true;

            canvas.discardActiveObject();


            /* ---------------------------------
               آبجکت‌های قبلی قابل انتخاب باشند
               --------------------------------- */

            canvas.forEachObject(
                function(object) {

                    object.selectable =
                        true;

                    object.evented =
                        true;

                }
            );


            window.BF_PAINT_TOOL =
                "RECTANGLE";


            /* ---------------------------------
               وضعیت رسم
               --------------------------------- */

            canvas.__BF_RECT_DRAWING =
                false;

            canvas.__BF_RECT_OBJECT =
                null;

            canvas.__BF_RECT_START =
                null;


            /* ---------------------------------
               جلوگیری از Context Menu
               --------------------------------- */

            if (
                !canvas.__BF_RECT_CONTEXT_BOUND
            ) {

                canvas.__BF_RECT_CONTEXT_BOUND =
                    true;

                canvas.upperCanvasEl.addEventListener(
                    "contextmenu",
                    function(event) {

                        if (
                            window.BF_PAINT_TOOL ===
                            "RECTANGLE"
                        ) {

                            event.preventDefault();

                        }

                    }
                );

            }


            /* ---------------------------------
               رویدادهای Rectangle فقط یک بار
               --------------------------------- */

            if (
                canvas.__BF_RECT_BOUND
            ) {

                console.log(
                    "BF.Paint TOOL: RECTANGLE"
                );

                return;

            }


            canvas.__BF_RECT_BOUND =
                true;


            /* ---------------------------------
               شروع ساخت Rectangle
               فقط با کلیک راست
               --------------------------------- */

            canvas.on(
                "mouse:down",
                function(opt) {

                    if (
                        window.BF_PAINT_TOOL !==
                        "RECTANGLE"
                    ) {

                        return;

                    }


                    if (
                        !opt.e ||
                        opt.e.button !==
                            2
                    ) {

                        return;

                    }


                    opt.e.preventDefault();


                    const pointer =
                        opt.scenePoint;


                    canvas.__BF_RECT_DRAWING =
                        true;


                    canvas.__BF_RECT_START = {
                        x:
                            pointer.x,

                        y:
                            pointer.y
                    };


                    const rectangle =
                        new window.fabric.Rect(
                            {
                                left:
                                    pointer.x,

                                top:
                                    pointer.y,

                                width:
                                    1,

                                height:
                                    1,

                                originX:
                                    "left",

                                originY:
                                    "top",

                                fill:
                                    "rgba(34,34,34,0.04)",

                                stroke:
                                    "#222222",

                                strokeWidth:
                                    3,

                                selectable:
                                    false,

                                evented:
                                    false
                            }
                        );


                    canvas.__BF_RECT_OBJECT =
                        rectangle;


                    canvas.add(
                        rectangle
                    );


                    canvas.requestRenderAll();

                }
            );


            /* ---------------------------------
               تغییر اندازه هنگام Drag راست
               --------------------------------- */

            canvas.on(
                "mouse:move",
                function(opt) {

                    if (
                        window.BF_PAINT_TOOL !==
                        "RECTANGLE"
                    ) {

                        return;

                    }


                    if (
                        !canvas.__BF_RECT_DRAWING
                    ) {

                        return;

                    }


                    if (
                        !opt.e ||
                        !opt.e.buttons ||
                        (
                            (opt.e.buttons & 2)
                            !== 2
                        )
                    ) {

                        return;

                    }


                    const rectangle =
                        canvas.__BF_RECT_OBJECT;

                    const start =
                        canvas.__BF_RECT_START;


                    if (
                        !rectangle ||
                        !start
                    ) {

                        return;

                    }


                    const pointer =
                        opt.scenePoint;


                    const left =
                        Math.min(
                            start.x,
                            pointer.x
                        );

                    const top =
                        Math.min(
                            start.y,
                            pointer.y
                        );

                    const width =
                        Math.abs(
                            pointer.x -
                            start.x
                        );

                    const height =
                        Math.abs(
                            pointer.y -
                            start.y
                        );


                    rectangle.set({
                        left:
                            left,

                        top:
                            top,

                        width:
                            Math.max(
                                1,
                                width
                            ),

                        height:
                            Math.max(
                                1,
                                height
                            )
                    });


                    rectangle.setCoords();

                    canvas.requestRenderAll();

                }
            );


            /* ---------------------------------
               پایان ساخت Rectangle
               --------------------------------- */

            canvas.on(
                "mouse:up",
                function(opt) {

                    if (
                        window.BF_PAINT_TOOL !==
                        "RECTANGLE"
                    ) {

                        return;

                    }


                    if (
                        !canvas.__BF_RECT_DRAWING
                    ) {

                        return;

                    }


                    if (
                        opt.e &&
                        opt.e.button !==
                            2
                    ) {

                        return;

                    }


                    const rectangle =
                        canvas.__BF_RECT_OBJECT;


                    if (
                        rectangle
                    ) {

                        rectangle.set({
                            selectable:
                                true,

                            evented:
                                true
                        });

                        rectangle.setCoords();

                    }


                    canvas.__BF_RECT_DRAWING =
                        false;

                    canvas.__BF_RECT_OBJECT =
                        null;

                    canvas.__BF_RECT_START =
                        null;


                    canvas.requestRenderAll();

                }
            );


            console.log(
                "BF.Paint TOOL: RECTANGLE — RIGHT DRAG CREATE"
            );

        }
    );

} 
                              /* -----------------------------------------
                   ARROW
                   Right Drag = Create
                   Left Click = Select / Move / Resize
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Arrow"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            const canvas =
                                window.BF_PAINT_CANVAS;

                            if (
                                !canvas
                            ) {

                                return;

                            }


                            /* ---------------------------------
                               فعال کردن Arrow
                               --------------------------------- */

                            canvas.isDrawingMode =
                                false;

                            canvas.selection =
                                true;

                            canvas.discardActiveObject();


                            /* ---------------------------------
                               آبجکت‌های قبلی قابل انتخاب باشند
                               --------------------------------- */

                            canvas.forEachObject(
                                function(object) {

                                    object.selectable =
                                        true;

                                    object.evented =
                                        true;

                                }
                            );


                            window.BF_PAINT_TOOL =
                                "ARROW";


                            /* ---------------------------------
                               وضعیت رسم
                               --------------------------------- */

                            canvas.__BF_ARROW_DRAWING =
                                false;

                            canvas.__BF_ARROW_LINE =
                                null;

                            canvas.__BF_ARROW_HEAD =
                                null;


                            /* ---------------------------------
                               جلوگیری از Context Menu
                               --------------------------------- */

                            if (
                                !canvas.__BF_ARROW_CONTEXT_BOUND
                            ) {

                                canvas.__BF_ARROW_CONTEXT_BOUND =
                                    true;

                                canvas.upperCanvasEl.addEventListener(
                                    "contextmenu",
                                    function(event) {

                                        if (
                                            window.BF_PAINT_TOOL ===
                                            "ARROW"
                                        ) {

                                            event.preventDefault();

                                        }

                                    }
                                );

                            }


                            /* ---------------------------------
                               رویدادهای Arrow فقط یک بار
                               --------------------------------- */

                            if (
                                canvas.__BF_ARROW_BOUND
                            ) {

                                console.log(
                                    "BF.Paint TOOL: ARROW"
                                );

                                return;

                            }


                            canvas.__BF_ARROW_BOUND =
                                true;


                            /* ---------------------------------
                               شروع ساخت Arrow
                               فقط با راست‌کلیک
                               --------------------------------- */

                            canvas.on(
                                "mouse:down",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "ARROW"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !opt.e ||
                                        opt.e.button !==
                                            2
                                    ) {

                                        return;

                                    }


                                    opt.e.preventDefault();


                                    const pointer =
                                        opt.scenePoint;


                                    canvas.__BF_ARROW_DRAWING =
                                        true;


                                    /* ---------------------------------
                                       خط فلش
                                       --------------------------------- */

                                    const line =
                                        new window.fabric.Line(
                                            [
                                                pointer.x,
                                                pointer.y,
                                                pointer.x,
                                                pointer.y
                                            ],
                                            {
                                                stroke:
                                                    "#222222",

                                                strokeWidth:
                                                    3,

                                                selectable:
                                                    false,

                                                evented:
                                                    false,

                                                originX:
                                                    "center",

                                                originY:
                                                    "center"
                                            }
                                        );


                                    /* ---------------------------------
                                       نوک فلش
                                       --------------------------------- */

                                    const head =
                                        new window.fabric.Triangle(
                                            {
                                                left:
                                                    pointer.x,

                                                top:
                                                    pointer.y,

                                                width:
                                                    12,

                                                height:
                                                    12,

                                                originX:
                                                    "center",

                                                originY:
                                                    "center",

                                                fill:
                                                    "#222222",

                                                angle:
                                                    90,

                                                selectable:
                                                    false,

                                                evented:
                                                    false
                                            }
                                        );


                                    canvas.__BF_ARROW_LINE =
                                        line;

                                    canvas.__BF_ARROW_HEAD =
                                        head;


                                    canvas.add(
                                        line
                                    );

                                    canvas.add(
                                        head
                                    );


                                    canvas.requestRenderAll();

                                }
                            );


                            /* ---------------------------------
                               تغییر Arrow هنگام Drag راست
                               --------------------------------- */

                            canvas.on(
                                "mouse:move",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "ARROW"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !canvas.__BF_ARROW_DRAWING
                                    ) {

                                        return;

                                    }


                                    if (
                                        !opt.e ||
                                        !opt.e.buttons ||
                                        (
                                            (opt.e.buttons & 2)
                                            !== 2
                                        )
                                    ) {

                                        return;

                                    }


                                    const line =
                                        canvas.__BF_ARROW_LINE;

                                    const head =
                                        canvas.__BF_ARROW_HEAD;


                                    if (
                                        !line ||
                                        !head
                                    ) {

                                        return;

                                    }


                                    const pointer =
                                        opt.scenePoint;


                                    line.set({
                                        x2:
                                            pointer.x,

                                        y2:
                                            pointer.y
                                    });


                                    const dx =
                                        pointer.x -
                                        line.x1;

                                    const dy =
                                        pointer.y -
                                        line.y1;


                                    const angle =
                                        Math.atan2(
                                            dy,
                                            dx
                                        ) *
                                        180 /
                                        Math.PI;


                                    head.set({
                                        left:
                                            pointer.x,

                                        top:
                                            pointer.y,

                                        angle:
                                            angle +
                                            90
                                    });


                                    line.setCoords();

                                    head.setCoords();

                                    canvas.requestRenderAll();

                                }
                            );


                            /* ---------------------------------
                               پایان ساخت Arrow
                               تبدیل به یک Group
                               --------------------------------- */

                            canvas.on(
                                "mouse:up",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "ARROW"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !canvas.__BF_ARROW_DRAWING
                                    ) {

                                        return;

                                    }


                                    if (
                                        opt.e &&
                                        opt.e.button !==
                                            2
                                    ) {

                                        return;

                                    }


                                    const line =
                                        canvas.__BF_ARROW_LINE;

                                    const head =
                                        canvas.__BF_ARROW_HEAD;


                                    if (
                                        line &&
                                        head
                                    ) {

                                        const arrowGroup =
                                            new window.fabric.Group(
                                                [
                                                    line,
                                                    head
                                                ],
                                                {
                                                    selectable:
                                                        true,

                                                    evented:
                                                        true,

                                                    originX:
                                                        "center",

                                                    originY:
                                                        "center"
                                                }
                                            );


                                        canvas.remove(
                                            line
                                        );

                                        canvas.remove(
                                            head
                                        );

                                        canvas.add(
                                            arrowGroup
                                        );


                                        arrowGroup.setCoords();


                                        canvas.setActiveObject(
                                            arrowGroup
                                        );

                                    }


                                    canvas.__BF_ARROW_DRAWING =
                                        false;

                                    canvas.__BF_ARROW_LINE =
                                        null;

                                    canvas.__BF_ARROW_HEAD =
                                        null;


                                    canvas.requestRenderAll();

                                }
                            );


                            console.log(
                                "BF.Paint TOOL: ARROW — RIGHT DRAG CREATE"
                            );

                        }
                    );

                
                }
               
                /* -----------------------------------------
                   TEXT
                   Right Click = Create / Edit
                   Left Click = Select / Move / Resize
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Text"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            const canvas =
                                window.BF_PAINT_CANVAS;

                            if (
                                !canvas
                            ) {

                                return;

                            }


                            /* ---------------------------------
                               فعال کردن Text
                               --------------------------------- */

                            canvas.isDrawingMode =
                                false;

                            canvas.selection =
                                true;

                            canvas.discardActiveObject();


                            /* ---------------------------------
                               آبجکت‌های قبلی قابل انتخاب باشند
                               --------------------------------- */

                            canvas.forEachObject(
                                function(object) {

                                    object.selectable =
                                        true;

                                    object.evented =
                                        true;

                                }
                            );


                            window.BF_PAINT_TOOL =
                                "TEXT";


                            /* ---------------------------------
                               جلوگیری از Context Menu
                               --------------------------------- */

                            if (
                                !canvas.__BF_TEXT_CONTEXT_BOUND
                            ) {

                                canvas.__BF_TEXT_CONTEXT_BOUND =
                                    true;

                                canvas.upperCanvasEl.addEventListener(
                                    "contextmenu",
                                    function(event) {

                                        if (
                                            window.BF_PAINT_TOOL ===
                                            "TEXT"
                                        ) {

                                            event.preventDefault();

                                        }

                                    }
                                );

                            }


                            /* ---------------------------------
                               رویداد Text فقط یک بار
                               --------------------------------- */

                            if (
                                canvas.__BF_TEXT_BOUND
                            ) {

                                console.log(
                                    "BF.Paint TOOL: TEXT"
                                );

                                return;

                            }


                            canvas.__BF_TEXT_BOUND =
                                true;


                            /* ---------------------------------
                               ساخت Text
                               فقط با کلیک راست
                               --------------------------------- */

                            canvas.on(
                                "mouse:down",
                                function(opt) {

                                    if (
                                        window.BF_PAINT_TOOL !==
                                        "TEXT"
                                    ) {

                                        return;

                                    }


                                    if (
                                        !opt.e ||
                                        opt.e.button !==
                                            2
                                    ) {

                                        return;

                                    }


                                    opt.e.preventDefault();


                                    const pointer =
                                        opt.scenePoint;


                                    const text =
                                        new window.fabric.IText(
                                            "متن",
                                            {
                                                left:
                                                    pointer.x,

                                                top:
                                                    pointer.y,

                                                fontSize:
                                                    24,

                                                fill:
                                                    "#222222",

                                                fontFamily:
                                                    "Tahoma",

                                                selectable:
                                                    true,

                                                evented:
                                                    true,

                                                originX:
                                                    "left",

                                                originY:
                                                    "top"
                                            }
                                        );


                                    canvas.add(
                                        text
                                    );


                                    canvas.setActiveObject(
                                        text
                                    );


                                    text.enterEditing();

                                    text.selectAll();


                                    canvas.requestRenderAll();

                                }
                            );


                            console.log(
                                "BF.Paint TOOL: TEXT — RIGHT CLICK CREATE"
                            );

                        }
                    );

                }

                              /* -----------------------------------------
                   UNDO
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Undo"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            if (
                                typeof window.BF_PAINT_UNDO ===
                                "function"
                            ) {

                                window.BF_PAINT_UNDO();

                            }

                        }
                    );

                }


                /* -----------------------------------------
                   REDO
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Redo"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            if (
                                typeof window.BF_PAINT_REDO ===
                                "function"
                            ) {

                                window.BF_PAINT_REDO();

                            }

                        }
                    );

                } 
                              /* -----------------------------------------
                   CLEAR
                   پاک کردن تمام نقاشی‌های Layer 5
                   ----------------------------------------- */

                if (
                    toolData[1] ===
                    "Clear"
                ) {

                    toolBtn.addEventListener(
                        "click",
                        function() {

                            const canvas =
                                window.BF_PAINT_CANVAS;

                            if (
                                !canvas
                            ) {

                                return;

                            }


                            /* ---------------------------------
                               پاک کردن انتخاب فعلی
                               --------------------------------- */

                            canvas.discardActiveObject();


                            /* ---------------------------------
                               پاک کردن فقط Layer 5
                               --------------------------------- */

                            canvas.clear();


                            /* ---------------------------------
                               حفظ اندازه Canvas
                               --------------------------------- */

                            const reportSurface =
                                canvas.lowerCanvasEl &&
                                canvas.lowerCanvasEl.parentElement;

                            if (
                                reportSurface
                            ) {

                                const width =
                                    Math.max(
                                        1,
                                        reportSurface.scrollWidth
                                    );

                                const height =
                                    Math.max(
                                        1,
                                        reportSurface.scrollHeight
                                    );

                                canvas.setDimensions(
                                    {
                                        width:
                                            width,

                                        height:
                                            height
                                    }
                                );

                            }


                            canvas.requestRenderAll();


                            /* ---------------------------------
                               ثبت Clear در History
                               --------------------------------- */

                            if (
                                typeof window.BF_PAINT_SAVE_HISTORY ===
                                "function"
                            ) {

                                window.BF_PAINT_SAVE_HISTORY();

                            }


                            console.log(
                                "BF.Paint TOOL: CLEAR"
                            );

                        }
                    );

                }  
            }
        );


        /* =================================================
           Toolbar را بین Header و Report Body قرار بده
           ================================================= */

        layer4.insertBefore(
            paintToolbar,
            layer4Content
        );


        /* =================================================
           علامت فعال شدن Tools
           ================================================= */

        maxToolsBtn.style.background =
            "#dfe7ed";

        maxToolsBtn.title =
            "Tools فعال";

    }
);

            /* =================================================
               RESTORE
               ================================================= */

            restoreBtn.addEventListener(
                "click",
                function(event) {

                    event.stopPropagation();

                    layer4.remove();


                    minimizeBtn.textContent =
                        "□";

                    minimizeBtn.title =
                        "Maximize";

                }
            );


            /* وضعیت دکمه اصلی */

            minimizeBtn.textContent =
                "❐";

            minimizeBtn.title =
                "Restore";

        }
    );

        /* =====================================================
       وضعیت آیکون‌ها در حالت Minimized
       ===================================================== */

    minimizeBtn.disabled =
        false;

    exitBtn.disabled =
        false;

    toolsBtn.disabled =
        true;

    printBtn.disabled =
        true;

    shareBtn.disabled =
        true;
}    
/* =========================================================
   TAB GUARDIAN
   ========================================================= */
/* =========================================================
   TAB GUARDIAN
   فقط مدیریت Tab و باز کردن Report Box
   ========================================================= */

function TabGuardian(tab) {

    if (!tab) {
        return;
    }


    /* =====================================================
       حذف کنترل‌های موجود از خود Tab
       ===================================================== */

    const tabControls =
        tab.querySelectorAll(
            ".filter-report-tab-control"
        );

    tabControls.forEach(
        function(button) {

            button.remove();

        }
    );


    /* =====================================================
       فعال / غیرفعال
       ===================================================== */

    tab.addEventListener(
        "click",
        function() {

            const container =
                tab.parentElement;

            if (!container) {
                return;
            }


            const tabs =
                container.querySelectorAll(
                    ".filter-report-tab"
                );


            tabs.forEach(
                function(item) {

                    item.classList.remove(
                        "active"
                    );

                }
            );


            tab.classList.add(
                "active"
            );


            /* =================================================
               باز کردن Report Box مربوط به همین Tab
               ================================================= */

            setTimeout(
                function() {

                    TabGuardianReportFrame();

                },
                0
            );

        }
    );

}

/* =========================================================
   MAP VIEW STATES
   ========================================================= */

let mapViewState = 0;

function applyMapViewState() {

    workspace.classList.remove(
        "map-view-1",
        "map-view-2",
        "map-view-3"
    );

    if (mapViewState === 1) {

        workspace.classList.add(
            "map-view-1"
        );

    } else if (mapViewState === 2) {

        workspace.classList.add(
            "map-view-2"
        );

    } else if (mapViewState === 3) {

        workspace.classList.add(
            "map-view-3"
        );
    }
}

expandMapBtn.addEventListener(
    "click",
    function() {

        mapViewState =
            (mapViewState + 1) % 4;

        applyMapViewState();
    }
);
/* =========================================================
   INITIALIZE
   ========================================================= */
createFilters();
testMapData();
buildUI();
ensureMapFrame(true);

footer.textContent =
    "Beta Farm Dynamic Filter 0.5.1-B2-19" +
    "  |  ستون‌ها: " +
    DATA.column_count +
    "  |  ردیف‌ها: " +
    DATA.row_count +
    "  |  فیلترها: " +
    DATA.filters.length +
    "  |  Map Units: " +
    getMapRowCount();

</script>


</body>

</html>

 """.replace(
    "__DATA__",
    payload
).replace(
    "__SETTLEMENTS__",
    settlements_payload
).replace(
    "__MAP_ENGINE_URL__",
    map_engine_url
).replace(
    "__FABRIC_JS__",
    fabric_js
).replace(
    "__HTML2CANVAS_JS__",
    html2canvas_js
)


class DynamicFilterWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle(
            "Beta Farm - Dynamic Filter 0.5.1-B2-19"
        )

        self.resize(1050, 700)

        # Multi-file cache: each imported Excel gets its own DataN.
        self.files = []
        self.active_file_id = 0

        try:
            self._prepare_farm_data()
            self._sheet1_snapshot = self._snapshot_sheet1()

            # =========================================================
            # Map Engine B2-22.6.5 — اتصال واقعی و فقط یک بار Load
            # =========================================================
            self._ensure_map_engine_6_5()

            self.map_server = LocalMapServer(
                BASE_DIR / "maps"
            )

            map_engine_url = self.map_server.url(
                "Beta_Farm_Map_B2-22.html"
            )

            data = read_data()

            page = build_ui(
                data,
                map_engine_url
            )

            self.web_view = QWebEngineView()
            self.web_view.loadFinished.connect(
                self._test_bf_paint
            )

            self.beta_farm_bridge = BetaFarmBridge(
                self.load_excel_file,
                self.set_map_mode
            )

            self.channel = QWebChannel(
                self.web_view.page()
            )

            self.channel.registerObject(
                "betaFarmBridge",
                self.beta_farm_bridge
            )

            self.web_view.page().setWebChannel(
                self.channel
            )

            settings = self.web_view.settings()

            settings.setAttribute(
                QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls,
                True
            )

            settings.setAttribute(
                QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls,
                True
            )

            base_url = QUrl.fromLocalFile(
                str(BASE_DIR) + "/"
            )

            self.web_view.setHtml(
                page,
                base_url
            )
            self.web_view.page().runJavaScript(
                "alert('INLINE BF.Paint TEST')"
            )
            # =========================================================
            # BF.Paint — Fabric.js DIAGNOSTIC TEST
            # =========================================================
            
            self.setCentralWidget(
                self.web_view
            )

            print("=" * 60)
            print("Beta Farm - Dynamic Filter 0.5.1-B2-19")
            print("=" * 60)
            print("Excel:", find_excel())
            print("Sheet:", SHEET_NAME)
            print("Columns:", data["column_count"])
            print("Rows:", data["row_count"])
            print("Filters:", len(data["filters"]))

        except Exception as exc:

            QMessageBox.critical(
                self,
                "خطا",
                str(exc)
            )

            raise
    def _test_bf_paint(self, ok):

        print(
            "BF.Paint PAGE LOAD:",
            ok
        )

        if not ok:
            return

        self.web_view.page().runJavaScript(
            """
            JSON.stringify({
            fabric: typeof window.fabric,
            htmlCanvas: typeof HTMLCanvasElement,
            version:
                typeof window.fabric !== 'undefined'
                    ? window.fabric.version
                    : 'NO_FABRIC'
            })
            """,
            lambda result: print(
                "BF.Paint TEST:", 
                result
            )
        )
    
                 
    def _ensure_map_engine_6_5(self):
        """
        Map Engine B2-22.6.5 را به‌عنوان موتور واقعی پروژه آماده می‌کند.
        خود Engine تغییر داده نمی‌شود؛ فقط در صورت نبود HTML، نسخه 6.5
        اجرا می‌شود تا Beta_Farm_Map_B2-22.html ساخته شود.
        """
        maps_dir = BASE_DIR / "maps"
        output_html = maps_dir / "Beta_Farm_Map_B2-22.html"

        if output_html.exists():
            print("MAP ENGINE HTML FOUND:", output_html)
            return

        candidates = [
            BASE_DIR / "Map_Engine_B2-22.6.5  .py",
            BASE_DIR / "Map_Engine_B2-22.6.5.py",
            BASE_DIR / "Map_Engine_B2-22.6.py",
        ]

        engine_path = next((x for x in candidates if x.exists()), None)

        if engine_path is None:
            raise FileNotFoundError(
                "Map Engine B2-22.6.5 پیدا نشد.\n\n"
                "فایل Map_Engine_B2-22.6.5.py را کنار Dynamic Filter قرار دهید."
            )

        print("MAP ENGINE PREPARE:", engine_path)
        runpy.run_path(str(engine_path), run_name="__main__")

        if not output_html.exists():
            raise FileNotFoundError(
                "Map Engine اجرا شد اما Beta_Farm_Map_B2-22.html ساخته نشد."
            )

        print("MAP ENGINE READY: B2-22.6.5")

    def _snapshot_sheet1(self):
        farm_data_path = BASE_DIR / "Farm-Data.xlsx"
        wb = load_workbook(farm_data_path, read_only=False, data_only=False)
        if "Sheet1" not in wb.sheetnames:
            wb.close()
            return []
        ws = wb["Sheet1"]
        snapshot = [list(row) for row in ws.iter_rows(values_only=True)]
        wb.close()
        return snapshot

    def _prepare_farm_data(self):
        farm_data_path = BASE_DIR / "Farm-Data.xlsx"
        wb = load_workbook(farm_data_path)
        # پاکسازی Data2...DataN باقی‌مانده از اجرای قبلی
        for name in list(wb.sheetnames):
            m = re.fullmatch(r"Data(\d+)", name, re.IGNORECASE)
            if m and int(m.group(1)) >= 2:
                wb.remove(wb[name])
        if "Data1" not in wb.sheetnames:
            wb.create_sheet("Data1")
        if "Sheet1" not in wb.sheetnames:
            wb.create_sheet("Sheet1", 0)
        wb.save(farm_data_path)
        wb.close()

    def _restore_farm_data(self):
        farm_data_path = BASE_DIR / "Farm-Data.xlsx"
        wb = load_workbook(farm_data_path)
        for name in list(wb.sheetnames):
            m = re.fullmatch(r"Data(\d+)", name, re.IGNORECASE)
            if m and int(m.group(1)) >= 2:
                wb.remove(wb[name])
        if "Sheet1" not in wb.sheetnames:
            wb.create_sheet("Sheet1", 0)
        ws = wb["Sheet1"]
        for row in ws.iter_rows():
            for cell in row:
                cell.value = None
        for r_idx, values in enumerate(self._sheet1_snapshot, 1):
            for c_idx, value in enumerate(values, 1):
                ws.cell(r_idx, c_idx).value = value
        wb.save(farm_data_path)
        wb.close()
        print("FARM-DATA CLEANUP: Data2...DataN REMOVED; Sheet1 RESTORED")

    def _read_data_sheet(self, sheet_name):
        return read_data_from_sheet(BASE_DIR / "Farm-Data.xlsx", sheet_name)

    def _next_data_sheet_number(self):
        farm_data_path = BASE_DIR / "Farm-Data.xlsx"
        wb = load_workbook(farm_data_path, read_only=True, data_only=True)
        numbers = []
        for name in wb.sheetnames:
            match = re.fullmatch(r"Data(\d+)", name, re.IGNORECASE)
            if match: numbers.append(int(match.group(1)))
        wb.close()
        return max(numbers) + 1 if numbers else 2

    def _import_excel_to_data(self, source_path, data_sheet):
        farm_data_path = BASE_DIR / "Farm-Data.xlsx"
        source_wb = load_workbook(source_path, read_only=True, data_only=True)
        source_ws = source_wb[source_wb.sheetnames[0]]
        farm_wb = load_workbook(farm_data_path)
        if data_sheet in farm_wb.sheetnames:
            source_wb.close(); farm_wb.close()
            raise ValueError(f"شیت {data_sheet} از قبل وجود دارد و بازنویسی نمی‌شود.")
        if "Sheet1" in farm_wb.sheetnames:
            farm_wb.remove(farm_wb["Sheet1"])
        target_ws = farm_wb.create_sheet("Sheet1", 0)
        for row in source_ws.iter_rows(values_only=True):
            target_ws.append(list(row))
        # فقط یک ستون اندیسی ابتداییِ واقعی را حذف می‌کنیم؛ ساختارهای دیگر دست‌نخورده می‌مانند.
        first = target_ws.cell(1, 1).value
        drop_first = first is None or str(first).strip().lower() in {"", "unnamed: 0", "index", "ردیف", "row", "no", "#"}
        data_ws = farm_wb.create_sheet(data_sheet)
        for row in target_ws.iter_rows(values_only=True):
            values = list(row)
            if drop_first and values:
                values = values[1:]
            data_ws.append(values)
        farm_wb.save(farm_data_path)
        source_wb.close(); farm_wb.close()
        print("SHEET1 ->", data_sheet, "CREATED")

    def _send_file_cache(self):
        items = [{"id":i["id"],"name":i["name"],"path":i["path"],"data_sheet":i["data_sheet"]} for i in self.files]
        payload=json.dumps(items, ensure_ascii=False, default=str)
        self.web_view.page().runJavaScript(f"window.setFileCache({payload}, {self.active_file_id or 0});")

    def add_file(self, file_path):
        path=Path(file_path).resolve()
        farm_path=(BASE_DIR / "Farm-Data.xlsx").resolve()
        if not path.exists(): raise FileNotFoundError(f"فایل انتخاب‌شده پیدا نشد:\n{path}")
        if path == farm_path:
            raise ValueError("Farm-Data.xlsx فایل مرجع برنامه است و نمی‌تواند به‌عنوان New File وارد شود.")
        for item in self.files:
            if Path(item["path"]).resolve()==path:
                self.activate_file(item["id"]); print("File already exists:", path); return
        file_id=max([item["id"] for item in self.files], default=0)+1
        data_sheet=f"Data{self._next_data_sheet_number()}"
        self._import_excel_to_data(path,data_sheet)
        data=self._read_data_sheet(data_sheet)
        item={"id":file_id,"name":path.name,"path":str(path),"data_sheet":data_sheet,"data":data}
        self.files.append(item)
        print("="*60); print("NEW FILE ADDED"); print("ID:",file_id); print("Name:",path.name); print("Data Sheet:",data_sheet); print("Files:",len(self.files)); print("="*60)
        self.activate_file(file_id)

    def load_excel_file(self, file_path):
        try: self.add_file(file_path)
        except Exception as exc: QMessageBox.critical(self,"خطا در بارگذاری فایل",str(exc))

    def activate_file(self, file_id):
        print("ACTIVATE REQUEST:", file_id, type(file_id))
        target=next((item for item in self.files if str(item["id"])==str(file_id)),None)
        if target is None: return
        self.active_file_id=file_id
        payload=json.dumps(target["data"], ensure_ascii=False, default=str)
        self.web_view.page().runJavaScript(f"window.setDynamicData({payload}, {file_id});")
        self._send_file_cache()
        print("ACTIVE FILE:",target["name"]); print("PATH:",target["path"]); print("DATA SHEET:",target["data_sheet"])

    def remove_file(self, file_id):
        target=next((item for item in self.files if item["id"]==file_id),None)
        if target is None: return
        self.files.remove(target)
        if self.active_file_id==file_id:
            if self.files: self.activate_file(self.files[-1]["id"])
            else:
                self.active_file_id=0
                self._send_file_cache()
        else: self._send_file_cache()
        print("FILE REMOVED:",target["name"])

    def set_map_mode(self, mode):
        if mode.startswith("__ACTIVATE_FILE__:"):
            file_id = int(mode.split(":", 1)[1])
            self.activate_file(file_id)
            return 
        if mode == "__PRINT__":

            printer = QPrinter(
                QPrinter.HighResolution
            )

            dialog = QPrintDialog(
                printer,
                self
            )

            if dialog.exec() == QPrintDialog.Accepted:

                self.web_view.print(
                    printer
                )

            return
        
        if mode.startswith(
            "__SAVE_CIRCLE_PDF__:"
        ):

            file_path = mode.split(
                ":",
                1
            )[1]
            
            page_layout = QPageLayout(
                QPageSize(
                    QPageSize.A4
                ),
                QPageLayout.Portrait,
                QMarginsF(
                    0,
                    0,
                    0,
                    0
                )
            )

            def pdf_finished(
                finished_path,
                success
            ):

                print(
                    "📄 PDF PRINT FINISHED:",
                    finished_path,
                    success
                )

                if success:

                    print(
                        "📄 CIRCLE PDF SAVED:",
                        finished_path
                    )

                else:

                    print(
                        "❌ CIRCLE PDF SAVE FAILED"
                    )

            self.web_view.pdfPrintingFinished.connect(
                pdf_finished
            )

            self.web_view.page().runJavaScript(
                """
                (function() {

                const content =
                    window.BF_LAYER4_CONTENT;

                if (!content) {
                    return false;
                }

                content.dataset.bfPdfOverflow =
                      content.style.overflow;

                content.dataset.bfPdfHeight =
                      content.style.height;

                content.dataset.bfPdfFlex =
                      content.style.flex;

                content.style.overflow =
                     "visible";

                content.style.height =
                     "auto";

                content.style.flex =
                     "none";

                 return true;

             })();
             """,
    lambda result:
        self.web_view.printToPdf(
            file_path,
            page_layout
        )
)

            return

           
        
        print(
            "Selected View:",
            mode
        )

        if mode == "Online (Google Maps)":

            self.web_view.page().runJavaScript(
                "setBetaFarmViewMode('ONLINE');"
            )

        elif mode == "Offline":

            self.web_view.page().runJavaScript(
                "setBetaFarmViewMode('OFFLINE');"
            )

        else:

            self.web_view.page().runJavaScript(
                "setBetaFarmViewMode('DEFAULT');"
            )
    def check_online_map(self):

        self.online_attempts += 1

        print(
            "Online Map Attempt:",
            self.online_attempts,
            "/",
            self.online_max_attempts
        )

        # فعلاً فقط شمارش می‌کنیم.
        # فعال‌سازی Offline در مرحله بعد انجام می‌شود.

        if self.online_attempts >= self.online_max_attempts:

            self.online_watchdog.stop()

            print(
                "ONLINE MAP: 40 ATTEMPTS REACHED"
            )

            print(
                "OFFLINE MODE: ACTIVATING"
            )

            self.web_view.page().runJavaScript(
                "setBetaFarmViewMode('OFFLINE');"
            )  
    def closeEvent(self, event):

        try:
            self._restore_farm_data()
        except Exception as exc:
            print("FARM-DATA CLEANUP ERROR:", exc)

        try:
            if hasattr(self, "map_server"):
                self.map_server.stop()
        except Exception:
            pass

        event.accept()


def main():

    app = QApplication(
        sys.argv
    )

    window = DynamicFilterWindow()

    window.show()

    sys.exit(
        app.exec()
    )


if __name__ == "__main__":

    main()
