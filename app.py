import base64
import ctypes
import ctypes.wintypes
import io
import json
import logging
import os
import sys

import win32con
import win32gui
import win32print
from PIL import Image, UnidentifiedImageError
from PyQt6.QtCore import QObject, QRunnable, QThreadPool, QUrl, pyqtSignal, pyqtSlot
from PyQt6.QtWebChannel import QWebChannel
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWidgets import QApplication, QMainWindow


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def get_resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def configure_devmode(
    devmode, width_mm, height_mm, quality, paper_type, paper_size_id=None
):
    portrait_width = min(width_mm, height_mm)
    portrait_height = max(width_mm, height_mm)
    paper_fields = win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION
    if paper_size_id is None:
        paper_fields |= win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH
    else:
        devmode.Fields &= ~(win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)
    devmode.Fields |= paper_fields
    if paper_size_id is None:
        paper_size_id = {
            (148, 210): win32con.DMPAPER_A5,
            (210, 297): win32con.DMPAPER_A4,
            (297, 420): win32con.DMPAPER_A3,
        }.get((round(portrait_width), round(portrait_height)), win32con.DMPAPER_USER)
        devmode.PaperWidth = round(portrait_width * 10)
        devmode.PaperLength = round(portrait_height * 10)
    devmode.PaperSize = paper_size_id
    devmode.Orientation = (
        win32con.DMORIENT_LANDSCAPE
        if width_mm > height_mm
        else win32con.DMORIENT_PORTRAIT
    )

    devmode.Fields |= win32con.DM_PRINTQUALITY
    devmode.PrintQuality = {
        "high": win32con.DMRES_HIGH,
        "standard": win32con.DMRES_MEDIUM,
        "draft": win32con.DMRES_DRAFT,
    }.get(quality, win32con.DMRES_HIGH)

    if hasattr(devmode, "MediaType"):
        devmode.Fields |= win32con.DM_MEDIATYPE
        devmode.MediaType = {
            "glossy": win32con.DMMEDIA_GLOSSY,
            "matte": win32con.DMMEDIA_STANDARD,
            "plain": win32con.DMMEDIA_STANDARD,
        }.get(paper_type, win32con.DMMEDIA_GLOSSY)


def find_supported_paper(printer_name, port_name, width_mm, height_mm):
    paper_ids = win32print.DeviceCapabilities(
        printer_name, port_name, win32con.DC_PAPERS, None
    )
    paper_sizes = win32print.DeviceCapabilities(
        printer_name, port_name, win32con.DC_PAPERSIZE, None
    )
    paper_names = win32print.DeviceCapabilities(
        printer_name, port_name, win32con.DC_PAPERNAMES, None
    )
    target_width, target_height = sorted((width_mm, height_mm))
    candidates = []
    for paper_id, size, name in zip(paper_ids, paper_sizes, paper_names):
        paper_width = min(size["x"], size["y"]) / 10
        paper_height = max(size["x"], size["y"]) / 10
        if paper_width <= 0 or paper_height <= 0:
            continue
        width_error = abs(paper_width - target_width) / target_width
        height_error = abs(paper_height - target_height) / target_height
        candidates.append(
            (
                max(width_error, height_error),
                width_error + height_error,
                paper_id,
                paper_width,
                paper_height,
                name.strip(),
            )
        )
    if not candidates:
        return None
    nearest = min(candidates)
    if nearest[0] > 0.03:
        return None
    return nearest[2:]


class PrintSignals(QObject):
    finished = pyqtSignal(str)


def print_image(payload):
    printer_name = payload.get("printer")
    if not isinstance(printer_name, str) or not printer_name.strip():
        raise ValueError("Выберите установленный принтер Windows.")

    width_mm = float(payload.get("width_mm", 0))
    height_mm = float(payload.get("height_mm", 0))
    if not (0 < width_mm <= 2000 and 0 < height_mm <= 2000):
        raise ValueError("Размеры бумаги должны быть больше 0 и не превышать 2000 мм.")

    encoded_image = payload.get("image_base64")
    if not isinstance(encoded_image, str):
        raise ValueError("В задании отсутствует изображение для печати.")
    image_data = encoded_image.split(",", 1)[-1]
    try:
        image_bytes = base64.b64decode(image_data, validate=True)
        with Image.open(io.BytesIO(image_bytes)) as source:
            image = source.convert("RGB")
    except (ValueError, UnidentifiedImageError) as error:
        raise ValueError(f"Не удалось прочитать изображение для печати: {error}") from error

    printer_handle = None
    device_context = None
    document_started = False
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(printer_handle, 2)
        devmode = printer_info.get("pDevMode")
        if devmode is None:
            raise RuntimeError("Драйвер принтера не предоставил параметры DEVMODE.")

        quality = payload.get("quality", "high")
        paper_type = payload.get("paper_type", "glossy")
        configure_devmode(devmode, width_mm, height_mm, quality, paper_type)

        device_context = win32gui.CreateDC("WINSPOOL", printer_name, None)
        reset_context = win32gui.ResetDC(device_context, devmode)
        if not reset_context:
            raise RuntimeError("Драйвер принтера не применил параметры DEVMODE.")
        device_context = reset_context

        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
        get_device_caps = gdi32.GetDeviceCaps
        get_device_caps.argtypes = [ctypes.wintypes.HDC, ctypes.c_int]
        get_device_caps.restype = ctypes.c_int
        dpi_x = get_device_caps(device_context, win32con.LOGPIXELSX)
        dpi_y = get_device_caps(device_context, win32con.LOGPIXELSY)
        if dpi_x <= 0 or dpi_y <= 0:
            raise RuntimeError("Драйвер принтера сообщил недопустимое разрешение печати.")
        page_width_px = get_device_caps(device_context, win32con.PHYSICALWIDTH)
        page_height_px = get_device_caps(device_context, win32con.PHYSICALHEIGHT)
        page_width_mm = page_width_px * 25.4 / dpi_x
        page_height_mm = page_height_px * 25.4 / dpi_y
        requested_page_size = sorted((width_mm, height_mm))
        actual_page_size = sorted((page_width_mm, page_height_mm))
        paper_form = None
        page_size_error = max(
            abs(actual_page_size[0] - requested_page_size[0]) / requested_page_size[0],
            abs(actual_page_size[1] - requested_page_size[1]) / requested_page_size[1],
        )
        if page_size_error > 0.03:
            paper_form = find_supported_paper(
                printer_name,
                printer_info.get("pPortName"),
                width_mm,
                height_mm,
            )
            if paper_form is None:
                raise RuntimeError(
                    f"Принтер «{printer_name}» не поддерживает формат "
                    f"{width_mm:g} × {height_mm:g} мм и не предлагает совместимый "
                    "фотоформат. Задание отменено, чтобы не печатать на A4."
                )
            paper_size_id, form_width_mm, form_height_mm, form_name = paper_form
            fallback_devmode = win32print.GetPrinter(printer_handle, 2).get("pDevMode")
            if fallback_devmode is None:
                raise RuntimeError("Не удалось получить DEVMODE для выбранного формата бумаги.")
            configure_devmode(
                fallback_devmode,
                width_mm,
                height_mm,
                quality,
                paper_type,
                paper_size_id=paper_size_id,
            )
            reset_context = win32gui.ResetDC(device_context, fallback_devmode)
            if not reset_context:
                raise RuntimeError(
                    f"Не удалось выбрать поддерживаемый формат бумаги «{form_name}»."
                )
            device_context = reset_context
            page_width_px = get_device_caps(device_context, win32con.PHYSICALWIDTH)
            page_height_px = get_device_caps(device_context, win32con.PHYSICALHEIGHT)
            page_width_mm = page_width_px * 25.4 / dpi_x
            page_height_mm = page_height_px * 25.4 / dpi_y
            actual_page_size = sorted((page_width_mm, page_height_mm))
            expected_form_size = sorted((form_width_mm, form_height_mm))
            if max(
                abs(actual_page_size[0] - expected_form_size[0]) / expected_form_size[0],
                abs(actual_page_size[1] - expected_form_size[1]) / expected_form_size[1],
            ) > 0.01:
                raise RuntimeError(
                    f"Драйвер принтера не применил поддерживаемый формат «{form_name}». "
                    "Задание отменено, чтобы избежать печати на A4."
                )
        destination_width = page_width_px
        destination_height = page_height_px
        offset_x = get_device_caps(device_context, win32con.PHYSICALOFFSETX)
        offset_y = get_device_caps(device_context, win32con.PHYSICALOFFSETY)

        class DocInfo(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.wintypes.LONG),
                ("lpszDocName", ctypes.wintypes.LPCWSTR),
                ("lpszOutput", ctypes.wintypes.LPCWSTR),
                ("lpszDatatype", ctypes.wintypes.LPCWSTR),
                ("fwType", ctypes.wintypes.DWORD),
            ]

        class BitmapInfoHeader(ctypes.Structure):
            _fields_ = [
                ("biSize", ctypes.wintypes.DWORD),
                ("biWidth", ctypes.wintypes.LONG),
                ("biHeight", ctypes.wintypes.LONG),
                ("biPlanes", ctypes.wintypes.WORD),
                ("biBitCount", ctypes.wintypes.WORD),
                ("biCompression", ctypes.wintypes.DWORD),
                ("biSizeImage", ctypes.wintypes.DWORD),
                ("biXPelsPerMeter", ctypes.wintypes.LONG),
                ("biYPelsPerMeter", ctypes.wintypes.LONG),
                ("biClrUsed", ctypes.wintypes.DWORD),
                ("biClrImportant", ctypes.wintypes.DWORD),
            ]

        class BitmapInfo(ctypes.Structure):
            _fields_ = [
                ("bmiHeader", BitmapInfoHeader),
                ("bmiColors", ctypes.wintypes.DWORD * 3),
            ]

        start_doc = gdi32.StartDocW
        start_doc.argtypes = [ctypes.wintypes.HDC, ctypes.POINTER(DocInfo)]
        start_doc.restype = ctypes.wintypes.INT
        start_page = gdi32.StartPage
        start_page.argtypes = [ctypes.wintypes.HDC]
        start_page.restype = ctypes.wintypes.INT
        end_page = gdi32.EndPage
        end_page.argtypes = [ctypes.wintypes.HDC]
        end_page.restype = ctypes.wintypes.INT
        end_doc = gdi32.EndDoc
        end_doc.argtypes = [ctypes.wintypes.HDC]
        end_doc.restype = ctypes.wintypes.INT
        abort_doc = gdi32.AbortDoc
        abort_doc.argtypes = [ctypes.wintypes.HDC]
        abort_doc.restype = ctypes.wintypes.INT
        stretch_dibits = gdi32.StretchDIBits
        stretch_dibits.argtypes = [
            ctypes.wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.wintypes.UINT,
            ctypes.wintypes.DWORD,
        ]
        stretch_dibits.restype = ctypes.c_int

        document_info = DocInfo(
            ctypes.sizeof(DocInfo), "STASH Print Studio", None, None, 0
        )
        if start_doc(device_context, ctypes.byref(document_info)) <= 0:
            raise ctypes.WinError(ctypes.get_last_error(), "Не удалось начать задание печати")
        document_started = True
        if start_page(device_context) <= 0:
            raise ctypes.WinError(ctypes.get_last_error(), "Не удалось начать страницу печати")

        image_width, image_height = image.size
        bitmap_info = BitmapInfo()
        row_stride = (image_width * 3 + 3) & ~3
        bitmap_info.bmiHeader = BitmapInfoHeader(
            ctypes.sizeof(BitmapInfoHeader),
            image_width,
            -image_height,
            1,
            24,
            win32con.BI_RGB,
            row_stride * image_height,
            0,
            0,
            0,
            0,
        )
        unpadded_stride = image_width * 3
        pixel_rows = image.tobytes("raw", "BGR")
        if row_stride == unpadded_stride:
            pixel_data = bytearray(pixel_rows)
        else:
            pixel_data = bytearray(row_stride * image_height)
            source_view = memoryview(pixel_rows)
            target_view = memoryview(pixel_data)
            for row in range(image_height):
                source_start = row * unpadded_stride
                target_start = row * row_stride
                target_view[target_start:target_start + unpadded_stride] = source_view[
                    source_start:source_start + unpadded_stride
                ]
        pixel_buffer = (ctypes.c_ubyte * len(pixel_data)).from_buffer(pixel_data)
        result = stretch_dibits(
            device_context,
            -offset_x,
            -offset_y,
            destination_width,
            destination_height,
            0,
            0,
            image_width,
            image_height,
            pixel_buffer,
            ctypes.byref(bitmap_info),
            win32con.DIB_RGB_COLORS,
            win32con.SRCCOPY,
        )
        if result <= 0:
            raise ctypes.WinError(ctypes.get_last_error(), "Не удалось передать изображение драйверу принтера")
        if end_page(device_context) <= 0:
            raise ctypes.WinError(ctypes.get_last_error(), "Не удалось завершить страницу печати")
        if end_doc(device_context) <= 0:
            raise ctypes.WinError(ctypes.get_last_error(), "Не удалось завершить задание печати")
        document_started = False
        return {
            "paper_width_mm": page_width_mm,
            "paper_height_mm": page_height_mm,
            "paper_fallback": paper_form is not None,
        }
    finally:
        if document_started and device_context:
            gdi32.AbortDoc(device_context)
        if device_context:
            win32gui.DeleteDC(device_context)
        if printer_handle:
            win32print.ClosePrinter(printer_handle)


class PrintTask(QRunnable):
    def __init__(self, payload):
        super().__init__()
        self.payload = payload
        self.signals = PrintSignals()

    def run(self):
        request_id = self.payload.get("request_id")
        try:
            print_details = print_image(self.payload)
            result = {
                "status": "success",
                "request_id": request_id,
                **print_details,
            }
        except Exception as error:
            logging.exception("Не удалось отправить задание на печать")
            result = {
                "status": "error",
                "request_id": request_id,
                "message": str(error),
            }
        self.signals.finished.emit(json.dumps(result, ensure_ascii=False))


class PrintBridge(QObject):
    print_finished = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.thread_pool = QThreadPool(self)
        self.thread_pool.setMaxThreadCount(1)
        self._tasks = {}

    @pyqtSlot(result=str)
    def get_printers(self):
        try:
            printers = [
                printer[2]
                for printer in win32print.EnumPrinters(
                    win32print.PRINTER_ENUM_LOCAL
                    | win32print.PRINTER_ENUM_CONNECTIONS
                )
            ]
            default_printer = win32print.GetDefaultPrinter()
            return json.dumps(
                {"printers": printers, "default": default_printer},
                ensure_ascii=False,
            )
        except Exception as error:
            logging.exception("Не удалось получить список принтеров")
            return json.dumps(
                {"printers": [], "default": "", "error": str(error)},
                ensure_ascii=False,
            )

    @pyqtSlot(str, result=str)
    def send_to_printer(self, payload_json):
        try:
            payload = json.loads(payload_json)
            if not isinstance(payload, dict):
                raise ValueError("Некорректный формат задания печати.")
            if not payload.get("request_id"):
                raise ValueError("В задании отсутствует идентификатор запроса.")
            task = PrintTask(payload)
            request_id = payload["request_id"]
            self._tasks[request_id] = task
            task.signals.finished.connect(self._on_print_finished)
            self.thread_pool.start(task)
            return json.dumps(
                {"status": "queued", "request_id": payload["request_id"]},
                ensure_ascii=False,
            )
        except Exception as error:
            logging.exception("Не удалось поставить задание в очередь печати")
            return json.dumps(
                {"status": "error", "message": str(error)},
                ensure_ascii=False,
            )

    @pyqtSlot(str)
    def _on_print_finished(self, result_json):
        request_id = json.loads(result_json).get("request_id")
        self._tasks.pop(request_id, None)
        self.print_finished.emit(result_json)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("STASH Print Studio — Панель оператора")
        self.resize(1440, 900)
        self.setMinimumSize(1024, 680)

        self.browser = QWebEngineView()
        self.setCentralWidget(self.browser)

        self.channel = QWebChannel()
        self.bridge = PrintBridge()
        self.channel.registerObject("pyBridge", self.bridge)
        self.browser.page().setWebChannel(self.channel)

        html_file = get_resource_path("index.html")
        if not os.path.isfile(html_file):
            raise FileNotFoundError(f"Не найден файл интерфейса: {html_file}")
        self.browser.load(QUrl.fromLocalFile(html_file))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
