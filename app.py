import base64
import ctypes
import io
import json
import logging
import os
import struct
import sys

try:
    import ctypes.wintypes as ctypes_wintypes
except ImportError:  # pragma: no cover - non-Windows runtime
    ctypes_wintypes = None

try:
    import win32con
    import win32gui
    import win32print
except ImportError:  # pragma: no cover - non-Windows runtime
    win32con = None
    win32gui = None
    win32print = None

from epson_print_ticket import (
    create_devmode as create_epson_devmode,
    get_media_type_options as get_epson_media_type_options,
    get_quality_options as get_epson_quality_options,
)

try:
    from PIL import Image, UnidentifiedImageError
except ImportError:  # pragma: no cover - optional dependency for tests
    Image = None
    UnidentifiedImageError = ValueError

try:
    from PyQt6.QtCore import QObject, QRunnable, QThreadPool, QUrl, pyqtSignal, pyqtSlot
    from PyQt6.QtWebChannel import QWebChannel
    from PyQt6.QtWebEngineWidgets import QWebEngineView
    from PyQt6.QtWidgets import QApplication, QMainWindow
except ImportError:  # pragma: no cover - optional UI dependency for tests
    class QObject:
        def __init__(self, *args, **kwargs):
            pass

    class QRunnable(QObject):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)

    class QThreadPool(QObject):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)

        def setMaxThreadCount(self, *args, **kwargs):
            return None

        def start(self, *args, **kwargs):
            return None

    class QUrl:
        @staticmethod
        def fromLocalFile(path):
            return path

    class _NoSignal:
        def __call__(self, *args, **kwargs):
            return None

    def pyqtSignal(*args, **kwargs):
        return _NoSignal()

    def pyqtSlot(*args, **kwargs):
        def decorator(func):
            return func

        return decorator

    class QWebChannel(QObject):
        def registerObject(self, *args, **kwargs):
            return None

    class QWebEngineView(QObject):
        def page(self):
            return self

        def setWebChannel(self, *args, **kwargs):
            return None

        def load(self, *args, **kwargs):
            return None

    class QApplication(QObject):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)

        def exec(self):
            return 0

    class QMainWindow(QObject):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)

        def setWindowTitle(self, *args, **kwargs):
            return None

        def resize(self, *args, **kwargs):
            return None

        def setMinimumSize(self, *args, **kwargs):
            return None

        def setCentralWidget(self, *args, **kwargs):
            return None

        def show(self):
            return None

        def setCentralWidget(self, *args, **kwargs):
            return None

    class _NoQApp:
        pass

    QApplication = QApplication
    QMainWindow = QMainWindow
    QWebChannel = QWebChannel
    QWebEngineView = QWebEngineView
    QObject = QObject
    QRunnable = QRunnable
    QThreadPool = QThreadPool
    QUrl = QUrl


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def _require_windows_printer_runtime():
    missing = []
    if ctypes_wintypes is None:
        missing.append("ctypes.wintypes")
    if win32con is None:
        missing.append("win32con")
    if win32gui is None:
        missing.append("win32gui")
    if win32print is None:
        missing.append("win32print")
    if Image is None:
        missing.append("Pillow")
    if missing:
        raise RuntimeError(
            "Эта функция доступна только в Windows-среде с установленными зависимостями: "
            + ", ".join(missing)
        )


def get_resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def is_epson_printer(printer_name):
    return isinstance(printer_name, str) and printer_name.casefold().startswith(
        "epson "
    )


def configure_paper_size(devmode, width_mm, height_mm, paper_size_id=None):
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


def configure_devmode(
    devmode,
    width_mm,
    height_mm,
    quality,
    paper_type,
    paper_size_id=None,
    printer_name=None,
    port_name=None,
):
    configure_paper_size(devmode, width_mm, height_mm, paper_size_id)
    if quality is not None:
        if quality not in ("high", "standard", "draft"):
            raise ValueError("Выбрано неизвестное качество печати.")
        devmode.Fields |= win32con.DM_PRINTQUALITY
        devmode.PrintQuality = {
            "high": win32con.DMRES_HIGH,
            "standard": win32con.DMRES_MEDIUM,
            "draft": win32con.DMRES_DRAFT,
        }[quality]

    if hasattr(devmode, "MediaType"):
        devmode.Fields |= win32con.DM_MEDIATYPE
        if printer_name and port_name:
            media_type = resolve_driver_media_type(
                printer_name, port_name, paper_type
            )
            if media_type is None:
                if isinstance(paper_type, str) and paper_type.startswith("driver:"):
                    raise ValueError(
                        "Драйвер не сообщил поддерживаемые варианты типа бумаги."
                    )
                devmode.MediaType = {
                    "glossy": win32con.DMMEDIA_GLOSSY,
                    "matte": win32con.DMMEDIA_STANDARD,
                    "plain": win32con.DMMEDIA_STANDARD,
                }.get(paper_type, win32con.DMMEDIA_GLOSSY)
            else:
                devmode.MediaType = media_type
        elif paper_type.startswith("driver:"):
            raise ValueError(
                "Нельзя применить тип бумаги без сведений о принтере."
            )
        else:
            devmode.MediaType = {
                "glossy": win32con.DMMEDIA_GLOSSY,
                "matte": win32con.DMMEDIA_STANDARD,
                "plain": win32con.DMMEDIA_STANDARD,
            }.get(paper_type, win32con.DMMEDIA_GLOSSY)


def find_supported_paper(
    printer_name, port_name, width_mm, height_mm, preferred_name_parts=()
):
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
    preferred = [
        candidate
        for candidate in candidates
        if any(
            name_part in "".join(candidate[5].casefold().split())
            for name_part in preferred_name_parts
        )
    ]
    nearest = min(preferred or candidates)
    if nearest[0] > 0.03:
        return None
    return nearest[2:]


def get_driver_media_types(printer_name, port_name):
    media_types = win32print.DeviceCapabilities(
        printer_name, port_name, win32con.DC_MEDIATYPES, None
    )
    media_names = win32print.DeviceCapabilities(
        printer_name, port_name, win32con.DC_MEDIATYPENAMES, None
    )
    return [
        {"id": media_type, "name": name.strip()}
        for media_type, name in zip(media_types, media_names)
        if name.strip()
    ]


def add_print_ticket_media_options(printer_name, media_types):
    ticket_options = get_epson_media_type_options(printer_name)
    options_by_name = {}
    for option in ticket_options:
        key = " ".join(option["name"].casefold().split())
        options_by_name.setdefault(key, []).append(option["id"])

    for media in media_types:
        key = " ".join(media["name"].casefold().split())
        matches = options_by_name.get(key, [])
        if len(matches) != 1:
            raise RuntimeError(
                f"Не удалось однозначно сопоставить тип бумаги "
                f"«{media['name']}» с настройками Epson."
            )
        media["ticket_option"] = matches[0]
    return media_types


def is_photo_media_name(media_name):
    normalized_name = " ".join(media_name.casefold().split())
    if any(term in normalized_name for term in ("cd/dvd", "envelope", "конверт")):
        return False
    return any(
        term in normalized_name
        for term in ("photo", "glossy", "matte", "semigloss", "фото", "глянц", "матов")
    )


def resolve_driver_media_type(printer_name, port_name, paper_type):
    media_types = get_driver_media_types(printer_name, port_name)
    if isinstance(paper_type, str) and paper_type.startswith("driver:"):
        try:
            requested_type = int(paper_type[len("driver:"):])
        except ValueError as error:
            raise ValueError("Выбран неизвестный тип бумаги.") from error
        if any(media["id"] == requested_type for media in media_types):
            return requested_type
        raise ValueError("Выбранный тип бумаги не поддерживается этим драйвером.")
    preferred_names = {
        "glossy": (
            "photo paper glossy",
            "epson glossy",
            "epson premium glossy",
            "epson ultra glossy",
        ),
        "matte": ("epson matte", "epson archival matte"),
        "plain": ("plain paper", "простая бумага"),
    }.get(paper_type, ())
    for preferred_name in preferred_names:
        for media in media_types:
            if " ".join(media["name"].casefold().split()) == preferred_name:
                return media["id"]
    return None


def find_driver_paper_form(printer_name, port_name, width_mm, height_mm):
    format_names = {
        (100, 150): ("10x15", "4x6"),
        (130, 180): ("13x18", "5x7"),
        (150, 200): ("15x20",),
    }
    target_format = tuple(sorted((round(width_mm), round(height_mm))))
    return find_supported_paper(
        printer_name,
        port_name,
        width_mm,
        height_mm,
        preferred_name_parts=format_names.get(target_format, ()),
    )


def configure_raw_devmode(
    devmode_pointer,
    devmode_size,
    width_mm,
    height_mm,
    paper_size_id,
    media_type,
):
    if devmode_size < 220:
        raise RuntimeError("Драйвер Epson вернул неполный DEVMODE.")
    if not (0 < width_mm <= 2000 and 0 < height_mm <= 2000):
        raise ValueError("Размеры бумаги должны быть больше 0 и не превышать 2000 мм.")
    if media_type is None:
        raise ValueError("Драйвер не сообщил выбранный тип бумаги Epson.")

    raw_devmode = (ctypes.c_ubyte * devmode_size).from_address(
        devmode_pointer.value
    )
    portrait_width = min(width_mm, height_mm)
    portrait_height = max(width_mm, height_mm)
    fields = struct.unpack_from("<I", raw_devmode, 72)[0]
    fields |= win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION

    if paper_size_id is None:
        fields |= win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH
        paper_size_id = {
            (148, 210): win32con.DMPAPER_A5,
            (210, 297): win32con.DMPAPER_A4,
            (297, 420): win32con.DMPAPER_A3,
        }.get((round(portrait_width), round(portrait_height)), win32con.DMPAPER_USER)
        struct.pack_into("<h", raw_devmode, 80, round(portrait_height * 10))
        struct.pack_into("<h", raw_devmode, 82, round(portrait_width * 10))
    else:
        fields &= ~(win32con.DM_PAPERWIDTH | win32con.DM_PAPERLENGTH)

    fields |= win32con.DM_MEDIATYPE
    struct.pack_into("<I", raw_devmode, 72, fields)
    struct.pack_into("<h", raw_devmode, 76, (
        win32con.DMORIENT_LANDSCAPE
        if width_mm > height_mm
        else win32con.DMORIENT_PORTRAIT
    ))
    struct.pack_into("<h", raw_devmode, 78, paper_size_id)
    struct.pack_into("<I", raw_devmode, 196, media_type)


def reset_device_context_with_devmode(device_context, devmode_buffer):
    gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
    reset_dc = gdi32.ResetDCW
    reset_dc.argtypes = [ctypes_wintypes.HDC, ctypes.c_void_p]
    reset_dc.restype = ctypes_wintypes.HDC
    reset_context = reset_dc(device_context, devmode_buffer.pointer)
    if not reset_context:
        raise RuntimeError("Драйвер Epson не применил выбранные параметры PrintTicket.")
    return reset_context


def prepare_driver_devmode(printer_name, port_name, devmode, options):
    width_mm = float(options.get("width_mm", 0))
    height_mm = float(options.get("height_mm", 0))
    if not (0 < width_mm <= 2000 and 0 < height_mm <= 2000):
        raise ValueError("Размеры бумаги должны быть больше 0 и не превышать 2000 мм.")

    quality = options.get("quality", "high")
    if quality is not None and quality not in ("high", "standard", "draft"):
        raise ValueError("Выбрано неизвестное качество печати.")
    paper_type = options.get("paper_type", "glossy")
    if not isinstance(paper_type, str) or not (
        paper_type in ("glossy", "matte", "plain")
        or paper_type.startswith("driver:")
    ):
        raise ValueError("Выбран неизвестный тип бумаги.")

    paper_form = find_driver_paper_form(
        printer_name,
        port_name,
        width_mm,
        height_mm,
    )
    if paper_form is None:
        configure_paper_size(devmode, width_mm, height_mm)
    else:
        configure_paper_size(
            devmode, width_mm, height_mm, paper_size_id=paper_form[0]
        )

    if quality is not None:
        devmode.Fields |= win32con.DM_PRINTQUALITY
        devmode.PrintQuality = {
            "high": win32con.DMRES_HIGH,
            "standard": win32con.DMRES_MEDIUM,
            "draft": win32con.DMRES_DRAFT,
        }[quality]
    if hasattr(devmode, "MediaType"):
        media_type = resolve_driver_media_type(
            printer_name, port_name, paper_type
        )
        if options.get("borderless") is True:
            selected_media = next(
                (
                    item
                    for item in get_driver_media_types(printer_name, port_name)
                    if item["id"] == media_type
                ),
                None,
            )
            if selected_media is None or not is_photo_media_name(
                selected_media["name"]
            ):
                raise ValueError(
                    "Печать без полей доступна только для фотобумаги. "
                    "Выберите фотобумагу или включите поля."
                )
        devmode.Fields |= win32con.DM_MEDIATYPE
        if media_type is None:
            if isinstance(paper_type, str) and paper_type.startswith("driver:"):
                raise ValueError(
                    "Драйвер не сообщил поддерживаемые варианты типа бумаги."
                )
            media_type = {
                "glossy": win32con.DMMEDIA_GLOSSY,
                "matte": win32con.DMMEDIA_STANDARD,
                "plain": win32con.DMMEDIA_STANDARD,
            }.get(paper_type, win32con.DMMEDIA_GLOSSY)
            devmode.MediaType = media_type
        else:
            devmode.MediaType = media_type


def open_printer_settings(printer_name, owner_window, options):
    _require_windows_printer_runtime()
    if not isinstance(printer_name, str) or not printer_name.strip():
        raise ValueError("Выберите установленный принтер Windows.")

    printer_handle = win32print.OpenPrinter(printer_name)
    try:
        printer_info = win32print.GetPrinter(printer_handle, 2)
        input_devmode = printer_info.get("pDevMode")
        output_devmode = win32print.GetPrinter(printer_handle, 2).get("pDevMode")
        if input_devmode is None or output_devmode is None:
            raise RuntimeError("Драйвер принтера не предоставил параметры DEVMODE.")
        prepare_driver_devmode(
            printer_name,
            printer_info.get("pPortName"),
            input_devmode,
            options,
        )

        result = win32print.DocumentProperties(
            owner_window,
            printer_handle,
            printer_name,
            output_devmode,
            input_devmode,
            win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER | win32con.DM_IN_PROMPT,
        )
        if result == win32con.IDCANCEL:
            return None
        if result != win32con.IDOK:
            raise RuntimeError(
                f"Драйвер принтера не подтвердил настройки (код {result})."
            )
        return output_devmode
    finally:
        win32print.ClosePrinter(printer_handle)


class PrintSignals(QObject):
    finished = pyqtSignal(str)


def print_image(payload, selected_devmode=None):
    _require_windows_printer_runtime()
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
    epson_devmode = None
    document_started = False
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(printer_handle, 2)
        paper_type = payload.get("paper_type", "glossy")
        is_epson = is_epson_printer(printer_name) or is_epson_printer(
            printer_info.get("pDriverName")
        )
        if is_epson:
            media_type_option = payload.get("paper_type_option")
            borderless = payload.get("borderless") is True
            media_type = resolve_driver_media_type(
                printer_name,
                printer_info.get("pPortName"),
                paper_type,
            )
            if borderless:
                selected_media = next(
                    (
                        item
                        for item in get_driver_media_types(
                            printer_name, printer_info.get("pPortName")
                        )
                        if item["id"] == media_type
                    ),
                    None,
                )
                if selected_media is None or not is_photo_media_name(
                    selected_media["name"]
                ):
                    raise ValueError(
                        "Печать без полей доступна только для фотобумаги. "
                        "Выберите фотобумагу или включите поля."
                    )
            epson_devmode = create_epson_devmode(
                printer_name,
                media_type_option,
                payload.get("quality"),
                borderless,
            )
            paper_form = find_driver_paper_form(
                printer_name,
                printer_info.get("pPortName"),
                width_mm,
                height_mm,
            )
            configure_raw_devmode(
                epson_devmode.pointer,
                epson_devmode.size,
                width_mm,
                height_mm,
                paper_form[0] if paper_form else None,
                media_type,
            )
            devmode = None
        else:
            devmode = selected_devmode or printer_info.get("pDevMode")
            if devmode is None:
                raise RuntimeError(
                    "Драйвер принтера не предоставил параметры DEVMODE."
                )
            configure_devmode(
                devmode,
                width_mm,
                height_mm,
                payload.get("quality", "high"),
                paper_type,
                printer_name=printer_name,
                port_name=printer_info.get("pPortName"),
            )

        device_context = win32gui.CreateDC("WINSPOOL", printer_name, None)
        reset_context = (
            reset_device_context_with_devmode(device_context, epson_devmode)
            if is_epson
            else win32gui.ResetDC(device_context, devmode)
        )
        if not reset_context:
            raise RuntimeError("Драйвер принтера не применил параметры DEVMODE.")
        device_context = reset_context

        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
        get_device_caps = gdi32.GetDeviceCaps
        get_device_caps.argtypes = [ctypes_wintypes.HDC, ctypes.c_int]
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
        orientation_mismatch = (width_mm > height_mm) != (
            page_width_mm > page_height_mm
        )
        if page_size_error > 0.03 or orientation_mismatch:
            if selected_devmode is not None:
                raise RuntimeError(
                    "Формат бумаги в настройках драйвера не совпадает с форматом "
                    f"приложения ({page_width_mm:.1f} × {page_height_mm:.1f} мм против "
                    f"{width_mm:g} × {height_mm:g} мм). Проверьте одинаковые формат "
                    "и ориентацию в приложении и в настройках драйвера."
                )
            paper_form = find_driver_paper_form(
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
            if is_epson:
                configure_raw_devmode(
                    epson_devmode.pointer,
                    epson_devmode.size,
                    width_mm,
                    height_mm,
                    paper_size_id,
                    media_type,
                )
                reset_context = reset_device_context_with_devmode(
                    device_context,
                    epson_devmode,
                )
            else:
                fallback_devmode = win32print.GetPrinter(
                    printer_handle, 2
                ).get("pDevMode")
                if fallback_devmode is None:
                    raise RuntimeError(
                        "Не удалось получить DEVMODE для выбранного формата бумаги."
                    )
                configure_devmode(
                    fallback_devmode,
                    width_mm,
                    height_mm,
                    payload.get("quality", "high"),
                    paper_type,
                    paper_size_id=paper_size_id,
                    printer_name=printer_name,
                    port_name=printer_info.get("pPortName"),
                )
                reset_context = win32gui.ResetDC(
                    device_context,
                    fallback_devmode,
                )
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
                ("cbSize", ctypes_wintypes.LONG),
                ("lpszDocName", ctypes_wintypes.LPCWSTR),
                ("lpszOutput", ctypes_wintypes.LPCWSTR),
                ("lpszDatatype", ctypes_wintypes.LPCWSTR),
                ("fwType", ctypes_wintypes.DWORD),
            ]

        class BitmapInfoHeader(ctypes.Structure):
            _fields_ = [
                ("biSize", ctypes_wintypes.DWORD),
                ("biWidth", ctypes_wintypes.LONG),
                ("biHeight", ctypes_wintypes.LONG),
                ("biPlanes", ctypes_wintypes.WORD),
                ("biBitCount", ctypes_wintypes.WORD),
                ("biCompression", ctypes_wintypes.DWORD),
                ("biSizeImage", ctypes_wintypes.DWORD),
                ("biXPelsPerMeter", ctypes_wintypes.LONG),
                ("biYPelsPerMeter", ctypes_wintypes.LONG),
                ("biClrUsed", ctypes_wintypes.DWORD),
                ("biClrImportant", ctypes_wintypes.DWORD),
            ]

        class BitmapInfo(ctypes.Structure):
            _fields_ = [
                ("bmiHeader", BitmapInfoHeader),
                ("bmiColors", ctypes_wintypes.DWORD * 3),
            ]

        start_doc = gdi32.StartDocW
        start_doc.argtypes = [ctypes_wintypes.HDC, ctypes.POINTER(DocInfo)]
        start_doc.restype = ctypes_wintypes.INT
        start_page = gdi32.StartPage
        start_page.argtypes = [ctypes_wintypes.HDC]
        start_page.restype = ctypes_wintypes.INT
        end_page = gdi32.EndPage
        end_page.argtypes = [ctypes_wintypes.HDC]
        end_page.restype = ctypes_wintypes.INT
        end_doc = gdi32.EndDoc
        end_doc.argtypes = [ctypes_wintypes.HDC]
        end_doc.restype = ctypes_wintypes.INT
        abort_doc = gdi32.AbortDoc
        abort_doc.argtypes = [ctypes_wintypes.HDC]
        abort_doc.restype = ctypes_wintypes.INT
        stretch_dibits = gdi32.StretchDIBits
        stretch_dibits.argtypes = [
            ctypes_wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes_wintypes.UINT,
            ctypes_wintypes.DWORD,
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
        if epson_devmode:
            epson_devmode.close()
        if printer_handle:
            win32print.ClosePrinter(printer_handle)


class PrintTask(QRunnable):
    def __init__(self, payload, selected_devmode=None):
        super().__init__()
        self.payload = payload
        self.selected_devmode = selected_devmode
        self.signals = PrintSignals()

    def run(self):
        request_id = self.payload.get("request_id")
        try:
            print_details = print_image(self.payload, self.selected_devmode)
            result = {
                "status": "success",
                "request_id": request_id,
                "driver_settings_applied": (
                    self.selected_devmode is not None
                    or is_epson_printer(self.payload.get("printer"))
                ),
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
        self._printer_devmodes = {}

    @pyqtSlot(str, str, result=str)
    def configure_printer(self, printer_name, options_json):
        try:
            options = json.loads(options_json)
            if not isinstance(options, dict):
                raise ValueError("Некорректный формат настроек принтера.")
            owner_window = win32gui.GetForegroundWindow()
            devmode = open_printer_settings(
                printer_name, owner_window, options
            )
            if devmode is None:
                return json.dumps({"status": "cancelled"}, ensure_ascii=False)
            self._printer_devmodes[printer_name] = devmode
            return json.dumps(
                {
                    "status": "configured",
                    "printer": printer_name,
                    "media_type": getattr(devmode, "MediaType", None),
                },
                ensure_ascii=False,
            )
        except Exception as error:
            logging.exception("Не удалось открыть настройки драйвера принтера")
            return json.dumps(
                {"status": "error", "message": str(error)},
                ensure_ascii=False,
            )

    @pyqtSlot(str, result=str)
    def get_printer_media_types(self, printer_name):
        printer_handle = None
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            media_types = get_driver_media_types(
                printer_name, printer_info.get("pPortName")
            )
            if not media_types:
                raise RuntimeError(
                    "Драйвер принтера не сообщил поддерживаемые типы бумаги."
                )
            if is_epson_printer(
                printer_info.get("pDriverName")
            ) or is_epson_printer(printer_name):
                media_types = add_print_ticket_media_options(
                    printer_name,
                    media_types,
                )
            return json.dumps(
                {"media_types": media_types}, ensure_ascii=False
            )
        except Exception as error:
            logging.exception("Не удалось получить типы бумаги принтера")
            return json.dumps(
                {"error": str(error), "media_types": []},
                ensure_ascii=False,
            )
        finally:
            if printer_handle:
                win32print.ClosePrinter(printer_handle)

    @pyqtSlot(str, str, bool, result=str)
    def get_epson_quality_modes(
        self,
        printer_name,
        media_type_option,
        borderless,
    ):
        try:
            if not is_epson_printer(printer_name):
                raise ValueError("Выбранный принтер не является Epson.")
            quality_modes = get_epson_quality_options(
                printer_name,
                media_type_option,
                borderless,
            )
            if not quality_modes:
                raise RuntimeError(
                    "Драйвер Epson не сообщил доступные режимы качества."
                )
            return json.dumps(
                {"quality_modes": quality_modes},
                ensure_ascii=False,
            )
        except Exception as error:
            logging.exception("Не удалось получить режимы качества Epson")
            return json.dumps(
                {"error": str(error), "quality_modes": []},
                ensure_ascii=False,
            )

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
            printer_name = payload.get("printer")
            selected_devmode = (
                None
                if is_epson_printer(printer_name)
                else self._printer_devmodes.get(printer_name)
            )
            task = PrintTask(
                payload,
                selected_devmode,
            )
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
    if win32con is None or win32gui is None or win32print is None:
        raise RuntimeError("STASH Print Studio требует Windows и зависимости pywin32.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
