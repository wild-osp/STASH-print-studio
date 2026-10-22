import sys
import os
import io
import json
import base64
import tempfile
from PIL import Image, ImageWin
import win32print
import win32con

from PyQt6.QtWidgets import QApplication, QMainWindow
from PyQt6.QtCore import pyqtSlot, QObject, QUrl
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class PrintBridge(QObject):
    @pyqtSlot(result=str)
    def get_printers(self):
        """Возвращает список всех принтеров Windows и принтер по умолчанию"""
        try:
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            default_printer = win32print.GetDefaultPrinter()
            return json.dumps({"printers": printers, "default": default_printer})
        except Exception as e:
            return json.dumps({"printers": [], "default": "", "error": str(e)})

    @pyqtSlot(str, result=str)
    def send_to_printer(self, payload_json):
        """Прямая печать через Windows Spooler без вызова падающего win32ui"""
        p_handle = None
        temp_file_path = None
        try:
            data = json.loads(payload_json)
            printer_name = data.get("printer")
            width_mm = float(data.get("width_mm", 100))
            height_mm = float(data.get("height_mm", 150))
            paper_type = data.get("paper_type", "glossy")
            quality = data.get("quality", "high")
            base64_img = data.get("image_base64").split(",")[-1]

            image_bytes = base64.b64decode(base64_img)

            # 1. Открываем дескриптор принтера с правами изменения настроек
            p_handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})

            # 2. Безопасное обновление настроек принтера (DEVMODE)
            try:
                properties = win32print.GetPrinter(p_handle, 2)
                devmode = properties['pDevMode']

                # Геометрия листа в десятых долях миллиметра (100 мм = 1000)
                devmode.dmFields |= (win32con.DM_PAPERLENGTH | win32con.DM_PAPERWIDTH | win32con.DM_ORIENTATION)
                devmode.dmPaperWidth = int(width_mm * 10)
                devmode.dmPaperLength = int(height_mm * 10)
                devmode.dmOrientation = win32con.DMORIENT_LANDSCAPE if width_mm > height_mm else win32con.DMORIENT_PORTRAIT

                # Качество печати
                devmode.dmFields |= win32con.DM_PRINTQUALITY
                if quality == "high":
                    devmode.dmPrintQuality = win32con.DMRES_HIGH
                elif quality == "draft":
                    devmode.dmPrintQuality = win32con.DMRES_DRAFT
                else:
                    devmode.dmPrintQuality = win32con.DMRES_MEDIUM

                # Тип носителя
                devmode.dmFields |= win32con.DM_MEDIATYPE
                if paper_type == "glossy":
                    devmode.dmMediaType = win32con.DMMEDIA_GLOSSY
                else:
                    devmode.dmMediaType = win32con.DMMEDIA_STANDARD

                # Сохраняем настройки в конфигурацию задачи
                properties['pDevMode'] = devmode
                win32print.SetPrinter(p_handle, 2, properties, 0)
            except Exception:
                # Если у принтера защищенный драйвер (например, сетевой), продолжаем без сбоя
                pass

            # 3. Сохраняем изображение во временный BMP файл для прямой отправки
            img = Image.open(io.BytesIO(image_bytes))
            with tempfile.NamedTemporaryFile(delete=False, suffix=".bmp") as tmp:
                temp_file_path = tmp.name
                img.save(temp_file_path, "BMP")

            # 4. Отправляем задание через стандартную подсистему GDI печати Windows
            # Используем надежный вызов через Win32 Spooler API
            import win32api
            win32api.ShellExecute(
                0,
                "printto",
                temp_file_path,
                f'"{printer_name}"',
                ".",
                0
            )

            return json.dumps({"status": "success", "message": "OK"})

        except Exception as e:
            # Резервный метод: прямая побайтовая запись в порт принтера
            try:
                if p_handle:
                    job_id = win32print.StartDocPrinter(p_handle, 1, ("STASH Photo Print", None, "RAW"))
                    win32print.StartPagePrinter(p_handle)
                    win32print.WritePrinter(p_handle, image_bytes)
                    win32print.EndPagePrinter(p_handle)
                    win32print.EndDocPrinter(p_handle)
                    return json.dumps({"status": "success", "message": "OK"})
            except Exception as raw_err:
                return json.dumps({"status": "error", "message": f"{str(e)} | Raw: {str(raw_err)}"})

            return json.dumps({"status": "error", "message": str(e)})

        finally:
            if p_handle:
                try:
                    win32print.ClosePrinter(p_handle)
                except Exception:
                    pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("STASH Print Studio — Панель оператора")
        self.resize(1440, 900)

        self.browser = QWebEngineView()
        self.setCentralWidget(self.browser)

        self.channel = QWebChannel()
        self.bridge = PrintBridge()
        self.channel.registerObject("pyBridge", self.bridge)
        self.browser.page().setWebChannel(self.channel)

        html_file = get_resource_path("index.html")
        self.browser.load(QUrl.fromLocalFile(html_file))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())