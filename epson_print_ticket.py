import ctypes
import io
import re
import xml.etree.ElementTree as ET
from ctypes import wintypes


PRINT_SCHEMA_FRAMEWORK = (
    "http://schemas.microsoft.com/windows/2003/08/printing/"
    "printschemaframework"
)
PRINT_SCHEMA_KEYWORDS = (
    "http://schemas.microsoft.com/windows/2003/08/printing/"
    "printschemakeywords"
)
PRINT_TICKET_SCOPE_JOB = 2
DEFAULT_DEVMODE_USER = 0
_QNAME_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*$")


class PrintTicketError(RuntimeError):
    pass


class DevModeBuffer:
    def __init__(self, pointer, size, release_memory):
        self.pointer = pointer
        self.size = size
        self._release_memory = release_memory

    def close(self):
        if self.pointer:
            pointer = self.pointer
            self.pointer = None
            _check_hresult(self._release_memory(pointer), "освободить DEVMODE")


class _PrintTicketProvider:
    def __init__(self, printer_name):
        import pythoncom

        self._pythoncom = pythoncom
        self._pythoncom.CoInitialize()
        self._provider = ctypes.c_void_p()
        self._library = None
        try:
            self._library = ctypes.WinDLL("prntvpt.dll", use_last_error=True)
            self._open_provider = self._bind(
                "PTOpenProvider",
                [
                    wintypes.LPCWSTR,
                    wintypes.DWORD,
                    ctypes.POINTER(ctypes.c_void_p),
                ],
            )
            self._close_provider = self._bind(
                "PTCloseProvider", [ctypes.c_void_p]
            )
            self._release_memory = self._bind(
                "PTReleaseMemory", [ctypes.c_void_p]
            )
            self._get_capabilities = self._bind(
                "PTGetPrintCapabilities",
                [
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                ],
            )
            self._convert_to_devmode = self._bind(
                "PTConvertPrintTicketToDevMode",
                [
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                    ctypes.c_int,
                    ctypes.c_int,
                    ctypes.POINTER(wintypes.ULONG),
                    ctypes.POINTER(ctypes.c_void_p),
                    ctypes.c_void_p,
                ],
            )
            self._convert_to_ticket = self._bind(
                "PTConvertDevModeToPrintTicket",
                [
                    ctypes.c_void_p,
                    wintypes.ULONG,
                    ctypes.c_void_p,
                    ctypes.c_int,
                    ctypes.c_void_p,
                ],
            )
            _check_hresult(
                self._open_provider(printer_name, 1, ctypes.byref(self._provider)),
                "открыть PrintTicket-провайдер Epson",
            )
        except Exception:
            self.close()
            raise

    def _bind(self, name, argument_types):
        function = getattr(self._library, name)
        function.argtypes = argument_types
        function.restype = ctypes.c_long
        return function

    def close(self):
        try:
            if self._provider:
                provider = self._provider
                self._provider = ctypes.c_void_p()
                _check_hresult(
                    self._close_provider(provider),
                    "закрыть PrintTicket-провайдер Epson",
                )
        finally:
            if self._pythoncom is not None:
                pythoncom = self._pythoncom
                self._pythoncom = None
                pythoncom.CoUninitialize()

    def __enter__(self):
        return self

    def __exit__(self, error_type, error, traceback):
        self.close()
        return False

    @property
    def release_memory(self):
        return self._release_memory

    def get_capabilities(self, ticket=None):
        ticket_stream = _create_stream(ticket) if ticket is not None else None
        output_stream = None
        try:
            output_stream = _create_stream()
            _check_hresult(
                self._get_capabilities(
                    self._provider,
                    ticket_stream,
                    output_stream,
                    None,
                ),
                "получить возможности PrintTicket Epson",
            )
            return _parse_capabilities(_read_stream(output_stream))
        finally:
            if ticket_stream:
                _release_stream(ticket_stream)
            if output_stream:
                _release_stream(output_stream)

    def convert_to_devmode(self, ticket):
        ticket_stream = _create_stream(ticket)
        size = wintypes.ULONG()
        devmode = ctypes.c_void_p()
        try:
            _check_hresult(
                self._convert_to_devmode(
                    self._provider,
                    ticket_stream,
                    DEFAULT_DEVMODE_USER,
                    PRINT_TICKET_SCOPE_JOB,
                    ctypes.byref(size),
                    ctypes.byref(devmode),
                    None,
                ),
                "преобразовать настройки Epson в DEVMODE",
            )
            if not devmode.value or size.value < 220:
                raise PrintTicketError(
                    "Драйвер Epson вернул неполный DEVMODE."
                )
            result = DevModeBuffer(
                devmode,
                size.value,
                self._release_memory,
            )
            devmode = ctypes.c_void_p()
            return result
        finally:
            if devmode.value:
                _check_hresult(
                    self._release_memory(devmode),
                    "освободить некорректный DEVMODE",
                )
            _release_stream(ticket_stream)

    def convert_to_ticket(self, devmode):
        output_stream = _create_stream()
        try:
            _check_hresult(
                self._convert_to_ticket(
                    self._provider,
                    devmode.size,
                    devmode.pointer,
                    PRINT_TICKET_SCOPE_JOB,
                    output_stream,
                ),
                "проверить настройки Epson",
            )
            return _read_stream(output_stream)
        finally:
            _release_stream(output_stream)


def _check_hresult(result, action):
    if result < 0:
        raise PrintTicketError(
            f"Не удалось {action} (HRESULT 0x{result & 0xFFFFFFFF:08X})."
        )


def _create_stream(data=None):
    shlwapi = ctypes.WinDLL("shlwapi.dll", use_last_error=True)
    create_stream = shlwapi.SHCreateMemStream
    create_stream.argtypes = [ctypes.c_void_p, wintypes.UINT]
    create_stream.restype = ctypes.c_void_p
    buffer = ctypes.create_string_buffer(data) if data else None
    stream = create_stream(buffer, len(data) if data else 0)
    if not stream:
        raise ctypes.WinError(
            ctypes.get_last_error(),
            "Не удалось создать поток PrintTicket.",
        )
    return stream


def _stream_method(stream, index, result_type, argument_types):
    vtable = ctypes.cast(
        stream,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)),
    ).contents
    return ctypes.WINFUNCTYPE(result_type, *argument_types)(vtable[index])


def _release_stream(stream):
    release = _stream_method(
        stream,
        2,
        wintypes.ULONG,
        [ctypes.c_void_p],
    )
    release(stream)


def _read_stream(stream):
    seek = _stream_method(
        stream,
        5,
        ctypes.c_long,
        [
            ctypes.c_void_p,
            ctypes.c_longlong,
            wintypes.DWORD,
            ctypes.POINTER(ctypes.c_ulonglong),
        ],
    )
    read = _stream_method(
        stream,
        3,
        ctypes.c_long,
        [
            ctypes.c_void_p,
            ctypes.c_void_p,
            wintypes.ULONG,
            ctypes.POINTER(wintypes.ULONG),
        ],
    )
    size = ctypes.c_ulonglong()
    _check_hresult(
        seek(stream, 0, 2, ctypes.byref(size)),
        "определить размер PrintTicket",
    )
    if size.value > 4 * 1024 * 1024:
        raise PrintTicketError("Драйвер вернул слишком большой PrintTicket.")
    position = ctypes.c_ulonglong()
    _check_hresult(
        seek(stream, 0, 0, ctypes.byref(position)),
        "прочитать PrintTicket",
    )
    buffer = ctypes.create_string_buffer(size.value)
    read_size = wintypes.ULONG()
    _check_hresult(
        read(stream, buffer, size.value, ctypes.byref(read_size)),
        "прочитать PrintTicket",
    )
    return buffer.raw[:read_size.value]


def _parse_capabilities(xml_data):
    try:
        root = ET.fromstring(xml_data)
        namespaces = dict(
            (prefix or "", uri)
            for _, (prefix, uri) in ET.iterparse(
                io.BytesIO(xml_data),
                events=("start-ns",),
            )
        )
    except ET.ParseError as error:
        raise PrintTicketError(
            f"Драйвер вернул некорректные возможности PrintTicket: {error}"
        ) from error
    return root, namespaces


def _feature_options(root, feature_name):
    for feature in root.iter():
        if feature.tag.rsplit("}", 1)[-1] != "Feature":
            continue
        if feature.get("name", "").rsplit(":", 1)[-1] == feature_name:
            options = []
            for option in feature:
                if option.tag.rsplit("}", 1)[-1] != "Option":
                    continue
                name = option.get("name", "")
                if not _QNAME_PATTERN.fullmatch(name):
                    continue
                display_name = ""
                for prop in option:
                    if (
                        prop.tag.rsplit("}", 1)[-1] == "Property"
                        and prop.get("name", "").rsplit(":", 1)[-1]
                        == "DisplayName"
                    ):
                        value = next(
                            (
                                child.text
                                for child in prop
                                if child.tag.rsplit("}", 1)[-1] == "Value"
                            ),
                            "",
                        )
                        display_name = (value or "").strip()
                        break
                options.append(
                    {
                        "id": name,
                        "name": display_name or name,
                        "constrained": option.get("constrained", "psk:None"),
                    }
                )
            return options
    return []


def _available_options(options):
    return [
        option
        for option in options
        if option["constrained"].rsplit(":", 1)[-1] == "None"
    ]


def _build_ticket(namespaces, features):
    required_namespaces = {
        "psf": PRINT_SCHEMA_FRAMEWORK,
        "psk": PRINT_SCHEMA_KEYWORDS,
    }
    for option in features.values():
        prefix, separator, _ = option.partition(":")
        if not separator or prefix not in namespaces:
            raise PrintTicketError(
                "Драйвер Epson сообщил неизвестный формат параметра."
            )
        required_namespaces[prefix] = namespaces[prefix]

    namespace_attributes = " ".join(
        f'xmlns:{prefix}="{uri}"'
        for prefix, uri in required_namespaces.items()
    )
    feature_elements = "".join(
        f'<psf:Feature name="psk:{feature_name}">'
        f'<psf:Option name="{option}"/></psf:Feature>'
        for feature_name, option in features.items()
    )
    return (
        f"<psf:PrintTicket {namespace_attributes} version=\"1\">"
        f"{feature_elements}</psf:PrintTicket>"
    ).encode("utf-8")


def _print_ticket_option(ticket_data, feature_name):
    root, namespaces = _parse_capabilities(ticket_data)
    options = _feature_options(root, feature_name)
    if not options:
        return None
    return _option_identity(namespaces, options[0]["id"])


def _option_identity(namespaces, option):
    prefix, local_name = option.split(":", 1)
    namespace = namespaces.get(prefix)
    if namespace is None:
        raise PrintTicketError(
            "Драйвер Epson вернул параметр с неизвестным namespace."
        )
    return namespace, local_name


def _roundtrip_options_match(provider, devmode, expected, expected_namespaces):
    roundtrip = provider.convert_to_ticket(devmode)
    root, namespaces = _parse_capabilities(roundtrip)
    for feature_name, option in expected.items():
        actual_options = _feature_options(root, feature_name)
        if not actual_options or (
            _option_identity(namespaces, actual_options[0]["id"])
            != _option_identity(expected_namespaces, option)
        ):
            return False
    return True


def get_media_type_options(printer_name):
    with _PrintTicketProvider(printer_name) as provider:
        root, _ = provider.get_capabilities()
        return _available_options(_feature_options(root, "PageMediaType"))


def get_quality_options(printer_name, media_type_option, borderless):
    if not isinstance(media_type_option, str) or not media_type_option:
        raise ValueError("Выберите тип бумаги Epson.")

    with _PrintTicketProvider(printer_name) as provider:
        root, namespaces = provider.get_capabilities()
        media_options = _available_options(
            _feature_options(root, "PageMediaType")
        )
        if media_type_option not in {option["id"] for option in media_options}:
            raise ValueError(
                "Выбранный тип бумаги больше не поддерживается драйвером."
            )
        context = {
            "PageMediaType": media_type_option,
            "PageBorderless": "psk:Borderless" if borderless else "psk:None",
        }
        context_ticket = _build_ticket(namespaces, context)
        context_root, _ = provider.get_capabilities(context_ticket)
        qualities = _available_options(
            _feature_options(context_root, "PageOutputQuality")
        )
        qualities = [
            option
            for option in qualities
            if option["id"].rsplit(":", 1)[-1].casefold()
            != "advancedsetting"
        ]
        available = []
        for quality in qualities:
            ticket = _build_ticket(
                namespaces,
                {
                    "PageMediaType": media_type_option,
                    "PageOutputQuality": quality["id"],
                    "PageBorderless": (
                        "psk:Borderless" if borderless else "psk:None"
                    ),
                },
            )
            devmode = provider.convert_to_devmode(ticket)
            try:
                if _roundtrip_options_match(
                    provider,
                    devmode,
                    {
                        "PageMediaType": media_type_option,
                        "PageOutputQuality": quality["id"],
                        "PageBorderless": (
                            "psk:Borderless" if borderless else "psk:None"
                        ),
                    },
                    namespaces,
                ):
                    available.append(quality)
            finally:
                devmode.close()
        return available


def create_devmode(printer_name, media_type_option, quality_option, borderless):
    if not isinstance(media_type_option, str) or not media_type_option:
        raise ValueError("Выберите тип бумаги Epson.")
    if not isinstance(quality_option, str) or not quality_option:
        raise ValueError("Выберите качество печати Epson.")

    with _PrintTicketProvider(printer_name) as provider:
        root, namespaces = provider.get_capabilities()
        media_options = _available_options(
            _feature_options(root, "PageMediaType")
        )
        if media_type_option not in {option["id"] for option in media_options}:
            raise ValueError(
                "Выбранный тип бумаги больше не поддерживается драйвером."
            )

        borderless_option = "psk:Borderless" if borderless else "psk:None"
        context_ticket = _build_ticket(
            namespaces,
            {
                "PageMediaType": media_type_option,
                "PageBorderless": borderless_option,
            },
        )
        context_root, _ = provider.get_capabilities(context_ticket)
        quality_options = _available_options(
            _feature_options(context_root, "PageOutputQuality")
        )
        quality_options = [
            option
            for option in quality_options
            if option["id"].rsplit(":", 1)[-1].casefold()
            != "advancedsetting"
        ]
        if quality_option not in {option["id"] for option in quality_options}:
            raise ValueError(
                "Выбранное качество недоступно для этой бумаги и режима полей."
            )

        borderless_options = _available_options(
            _feature_options(context_root, "PageBorderless")
        )
        if borderless_option not in {
            option["id"] for option in borderless_options
        }:
            raise ValueError(
                "Драйвер Epson не поддерживает выбранный режим печати без полей."
            )

        ticket = _build_ticket(
            namespaces,
            {
                "PageMediaType": media_type_option,
                "PageOutputQuality": quality_option,
                "PageBorderless": borderless_option,
            },
        )
        devmode = provider.convert_to_devmode(ticket)
        try:
            expected = {
                "PageMediaType": media_type_option,
                "PageOutputQuality": quality_option,
                "PageBorderless": borderless_option,
            }
            if not _roundtrip_options_match(
                provider,
                devmode,
                expected,
                namespaces,
            ):
                raise ValueError(
                    "Драйвер Epson заменил выбранное качество или режим полей; "
                    "задание не отправлено."
                )
            return devmode
        except Exception:
            devmode.close()
            raise
