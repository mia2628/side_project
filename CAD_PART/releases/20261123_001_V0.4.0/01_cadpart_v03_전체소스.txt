# -*- coding: utf-8 -*-
"""CAD-PART: AutoCAD 2019 viewport capture + Excel contact sheet.

Run: python cadpart_v03.py
Requirements: Windows, AutoCAD 2019, Excel, pywin32, Pillow.
Only run with approved DWGs / software on an authorized workstation.
"""
from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import sys
import time
import traceback
from pathlib import Path

VERSION = "0.4.0"
ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "CAD_INPUT"
OUTPUT = ROOT / "CAD_RESULT"
IMAGES = OUTPUT / "captures"
CONFIG = ROOT / "cadpart_config.json"
DEFAULT = {
    "captures_per_dwg": 3,
    "capture_mode": "model_plus_layouts",
    "capture_delay_seconds": 1.6,
    "capture_bbox_normalized": [0.045, 0.075, 0.97, 0.865],
    "excel_image_width_px": 320,
    "excel_image_height_px": 205,
    "recursive_search": True,
    "close_dwg_after_capture": True,
    "save_interval_drawings": 25,
    "show_autocad": True,
}


def config():
    if CONFIG.is_file():
        try:
            data = json.loads(CONFIG.read_text(encoding="utf-8-sig"))
            return {**DEFAULT, **data}
        except Exception as e:
            print("WARN CONFIG_INVALID", str(e))
    return DEFAULT.copy()


def safe_name(value):
    return re.sub(r'[^\w.\-]+', '_', value, flags=re.UNICODE)[:95].strip("._") or "drawing"


def get_screen_bbox(hwnd, normalized):
    import win32gui
    l, t, r, b = win32gui.GetClientRect(hwnd)
    origin = win32gui.ClientToScreen(hwnd, (0, 0))
    x0, y0 = origin
    w, h = r - l, b - t
    a, c, d, e = [float(z) for z in normalized]
    bbox = (round(x0 + w * a), round(y0 + h * c), round(x0 + w * d), round(y0 + h * e))
    if bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
        raise ValueError("INVALID_CAPTURE_BBOX")
    return bbox


def take_screenshot(hwnd, path, cfg):
    from PIL import ImageGrab
    import win32gui
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, 9)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(float(cfg["capture_delay_seconds"]))
    bbox = get_screen_bbox(hwnd, cfg["capture_bbox_normalized"])
    image = ImageGrab.grab(bbox=bbox, all_screens=True)
    if image.size[0] < 100 or image.size[1] < 100:
        raise RuntimeError("CAPTURE_TOO_SMALL")
    image.convert("RGB").save(str(path), "PNG", optimize=True)


def get_autocad():
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        app = win32com.client.Dispatch("AutoCAD.Application")
    app.Visible = True
    return app


def available_views(doc, cfg):
    views = [("MODEL", None)]
    if cfg.get("capture_mode") == "model_only":
        return views
    try:
        layouts = []
        for lay in doc.Layouts:
            name = str(lay.Name)
            if name.lower() != "model":
                layouts.append(name)
        for name in sorted(layouts, key=str.casefold):
            views.append(("LAYOUT", name))
    except Exception:
        pass
    return views[:int(cfg["captures_per_dwg"])]


def switch_view(app, doc, kind, name):
    try:
        doc.Activate()
        if kind == "MODEL":
            doc.ActiveSpace = 1  # acModelSpace
            doc.SetVariable("TILEMODE", 1)
        else:
            doc.SetVariable("TILEMODE", 0)
            doc.ActiveLayout = doc.Layouts.Item(name)
            doc.ActiveSpace = 0  # acPaperSpace
        app.ZoomExtents()
        time.sleep(0.35)
    except Exception as e:
        raise RuntimeError("VIEW_SETUP_FAILED: " + str(e))


def get_excel():
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    return app


def initialize_sheet(book, cfg):
    ws = book.Worksheets(1)
    ws.Name = "DRAWING_GALLERY"
    ws.Range("A1:D1").Merge()
    ws.Range("A1").Value = "CAD-PART | DWG IMAGE GALLERY"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 17
    ws.Range("A1:D1").Interior.Color = 0x492A16  # Excel BGR dark blue
    ws.Range("A1:D1").Font.Color = 0xFFFFFF
    ws.Rows(1).RowHeight = 36
    ws.Range("A2:D2").Value = [["DWG 파일명", "첫 번째 캡처", "두 번째 캡처", "세 번째 캡처"]]
    ws.Range("A2:D2").Font.Bold = True
    ws.Range("A2:D2").Interior.Color = 0xE9DCC5
    ws.Rows(2).RowHeight = 28
    ws.Columns("A").ColumnWidth = 33
    for letter in "BCD":
        ws.Columns(letter).ColumnWidth = 48
    ws.Range("A:D").VerticalAlignment = -4108  # xlCenter
    ws.Range("A2:D2").HorizontalAlignment = -4108
    ws.Application.ActiveWindow.DisplayGridlines = False
    return ws


def place_image(ws, cell, image_path, cfg):
    pic = ws.Shapes.AddPicture(str(image_path), False, True, 0, 0, -1, -1)
    pic.LockAspectRatio = -1
    maxw = float(cell.Width) - 12
    maxh = float(cell.Height) - 12
    factor = min(maxw / max(float(pic.Width), 1), maxh / max(float(pic.Height), 1), 1.0)
    pic.Width = float(pic.Width) * factor
    pic.Height = float(pic.Height) * factor
    pic.Left = float(cell.Left) + (float(cell.Width) - float(pic.Width)) / 2
    pic.Top = float(cell.Top) + (float(cell.Height) - float(pic.Height)) / 2
    pic.Placement = 1  # xlMoveAndSize


def process():
    cfg = config()
    if not INPUT.exists():
        INPUT.mkdir(parents=True, exist_ok=True)
        print("E01 INPUT_EMPTY | CAD_INPUT 폴더에 DWG 파일을 넣으세요")
        return 2
    dwgs = sorted(INPUT.rglob("*.dwg") if cfg["recursive_search"] else INPUT.glob("*.dwg"), key=lambda p: str(p).casefold())
    if not dwgs:
        print("E01 INPUT_EMPTY | DWG 0개")
        return 2
    OUTPUT.mkdir(exist_ok=True)
    IMAGES.mkdir(exist_ok=True)
    log = []
    try:
        acad = get_autocad()
        hwnd = int(acad.HWND)
    except Exception as e:
        print("E02 AUTOCAD_CONNECT_FAILED |", str(e))
        return 3
    try:
        excel = get_excel()
        book = excel.Workbooks.Add()
        ws = initialize_sheet(book, cfg)
    except Exception as e:
        print("E03 EXCEL_CONNECT_FAILED |", str(e))
        return 4

    ok = 0
    errors = 0
    images = 0
    workbook_path = OUTPUT / "CAD_DRAWING_GALLERY.xlsx"
    try:
        for i, dwg in enumerate(dwgs, 1):
            row = i + 2
            ws.Rows(row).RowHeight = min(390, max(150, float(cfg["excel_image_height_px"]) * 0.85))
            ws.Cells(row, 1).Value = dwg.name
            ws.Cells(row, 1).Font.Bold = True
            doc = None
            captured = 0
            err = ""
            try:
                # ReadOnly is supported by AutoCAD Documents.Open API.
                doc = acad.Documents.Open(str(dwg.resolve()), True)
                views = available_views(doc, cfg)
                for slot, (kind, name) in enumerate(views, 1):
                    try:
                        switch_view(acad, doc, kind, name)
                        code = hashlib.sha1(str(dwg).encode("utf-8")).hexdigest()[:8]
                        filename = safe_name(dwg.stem) + "_" + code + "_%02d.png" % slot
                        path = IMAGES / filename
                        take_screenshot(hwnd, path, cfg)
                        cell = ws.Cells(row, slot + 1)
                        place_image(ws, cell, path, cfg)
                        captured += 1
                        images += 1
                    except Exception as e:
                        msg = "E05 CAPTURE_FAILED " + str(e)
                        ws.Cells(row, slot+1).Value = msg[:120]
                        err += ("; " if err else "") + msg
                for missing in range(len(views)+1, int(cfg["captures_per_dwg"])+1):
                    ws.Cells(row, missing+1).Value = "캡처 없음 (배치 없음)"
                if captured:
                    ok += 1
                else:
                    errors += 1
                    if not err:
                        err = "E05 NO_CAPTURES"
            except Exception as e:
                errors += 1
                err = "E04 DWG_OPEN_FAILED " + str(e)
                ws.Cells(row, 2).Value = err[:170]
            finally:
                if doc is not None and cfg["close_dwg_after_capture"]:
                    try:
                        doc.Close(False)
                    except Exception:
                        err += "; E06 CLOSE_FAILED"
                log.append({"drawing":dwg.name, "captures":captured, "status":"OK" if captured else "FAIL", "code":err})
            print(f"{i}/{len(dwgs)} | {dwg.name} | {captured} capture | {err[:110]}")
            if i % max(1, int(cfg["save_interval_drawings"])) == 0:
                try:
                    book.SaveAs(str(workbook_path.resolve()), FileFormat=51)
                except Exception as e:
                    print("WARN E07 PARTIAL_SAVE_FAILED", str(e))
        ws.Range("A2:D2").AutoFilter()
        book.SaveAs(str(workbook_path.resolve()), FileFormat=51)
    finally:
        try:
            book.Close(False)
            excel.Quit()
        except Exception:
            pass
    with (OUTPUT / "CAPTURE_LOG.csv").open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=["drawing","captures","status","code"])
        writer.writeheader()
        writer.writerows(log)
    summary = {"version":VERSION,"dwg":len(dwgs),"success":ok,"failed":errors,"images":images,"xlsx":str(workbook_path)}
    (OUTPUT / "SUMMARY.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"V0.4 | DWG {len(dwgs)} | 성공 {ok} | 실패 {errors} | 이미지 {images} | {('E00' if errors==0 else 'E08')}")
    return 0 if errors == 0 else 1


if __name__ == "__main__":
    try:
        sys.exit(process())
    except Exception:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        (OUTPUT / "FATAL_ERROR.txt").write_text(traceback.format_exc(), encoding="utf-8")
        print("E99 FATAL_ERROR | CAD_RESULT/FATAL_ERROR.txt 확인")
        sys.exit(99)