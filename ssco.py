"""Lectura y cruce de registros de compras con el padrón de SUNAT."""
import csv
import hashlib
import math
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO, StringIO
from itertools import chain, islice
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SUNAT_PAGE = "https://www.sunat.gob.pe/padronesnotificaciones/sujeSinCapacidadOperativa.html"
SUNAT_XLSX = "https://www.sunat.gob.pe/padronesnotificaciones/ssco/sujesincapacidadOperativa.xlsx"
MAX_FILE_MB = 15
MAX_FILE = MAX_FILE_MB * 1024**2
MAX_UNPACKED = 90 * 1024**2
MAX_ROWS = 40_000
MAX_COLUMNS = 120
PURCHASE_EXTENSIONS = (".xlsx", ".zip", ".txt", ".csv")


class DataError(ValueError):
    """Error que se puede mostrar al usuario."""


@dataclass(frozen=True)
class Purchase:
    source: str
    section: str
    row: int
    layout: str
    issued: date | None
    kind: str
    series: str
    number: str
    provider: str
    identity_type: str
    identity: str
    total: Decimal | None
    vat: Decimal | None
    currency: str
    period: str = ""

    @property
    def ruc(self) -> str:
        if self.identity_type not in {"", "6", "06", "RUC"}:
            return ""
        return normalize_ruc(self.identity)


@dataclass(frozen=True)
class Subject:
    ruc: str
    name: str
    resolution: str
    published: date | None
    final: date | None


@dataclass(frozen=True)
class Registry:
    subjects: dict[str, tuple[Subject, ...]]
    latest_publication: date | None
    source: str
    fingerprint: str


@dataclass(frozen=True)
class Review:
    purchase: Purchase
    subjects: tuple[Subject, ...]

    @property
    def status(self) -> str:
        if self.subjects:
            return "Coincidencia SSCO"
        if self.purchase.ruc:
            return "Sin coincidencia"
        return "Sin RUC consultable"

    @property
    def subject(self) -> Subject | None:
        return self.subjects[0] if self.subjects else None


def as_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_ruc(value) -> str:
    value = re.sub(r"[\s-]+", "", as_text(value))
    if re.fullmatch(r"[0-9]{11}(?:\.0+)?", value):
        return value[:11]
    return ""


def label(value) -> str:
    plain = unicodedata.normalize("NFKD", as_text(value))
    plain = "".join(char for char in plain if not unicodedata.combining(char))
    return " ".join(re.findall(r"[a-z0-9]+", plain.lower()))


def parse_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(as_text(value), pattern).date()
        except ValueError:
            continue
    return None


def amount(value) -> Decimal | None:
    text = as_text(value)
    if not text:
        return None
    try:
        number = Decimal(text.replace(",", ""))
    except InvalidOperation:
        return None
    return number if number.is_finite() and number.adjusted() < 15 else None


def tax(row, columns):
    amounts = [amount(row[index]) for index in columns]
    return sum((value for value in amounts if value is not None), Decimal(0)) if any(
        value is not None for value in amounts
    ) else None


def file_bytes(data):
    if len(data) > MAX_FILE:
        raise DataError(f"El archivo supera el máximo de {MAX_FILE_MB} MB.")
    if not data:
        raise DataError("El archivo está vacío.")


def archive_members(archive):
    members = archive.infolist()
    if len(members) > 2000 or sum(item.file_size for item in members) > MAX_UNPACKED:
        raise DataError("El archivo comprimido supera el tamaño permitido al abrirlo.")
    if any(item.flag_bits & 1 for item in members):
        raise DataError("El archivo comprimido contiene archivos protegidos con contraseña.")
    return members


def open_excel(data):
    file_bytes(data)
    try:
        with ZipFile(BytesIO(data)) as archive:
            archive_members(archive)
        return load_workbook(BytesIO(data), read_only=True, data_only=True, keep_links=False)
    except DataError:
        raise
    except Exception as exc:
        raise DataError("No se pudo abrir el Excel. Verifica que sea un archivo .xlsx válido.") from exc


def sheet_rows(sheet):
    if (sheet.max_column or 0) > MAX_COLUMNS or (sheet.max_row or 0) > MAX_ROWS:
        raise DataError(f"La hoja «{sheet.title}» supera el máximo de {MAX_ROWS:,} filas o {MAX_COLUMNS} columnas.")
    for index, values in enumerate(sheet.iter_rows(values_only=True, max_col=MAX_COLUMNS), 1):
        if index > MAX_ROWS:
            raise DataError(f"La hoja «{sheet.title}» supera el máximo de {MAX_ROWS:,} filas.")
        yield index, values


def official_columns(headers):
    if "ruc" not in headers or "razon social" not in headers:
        return {}
    fields = {
        "ruc": "ruc",
        "name": "razon social",
        "resolution": "resolucion de atribucion como ssco",
        "published": "fecha de publicacion 2",
        "final": "fecha en la que la resolucion de atribucion quedo firme",
    }
    if fields["resolution"] not in headers:
        return {}
    if "fecha de publicacion" in headers:
        fields["published"] = "fecha de publicacion"
    return {key: headers.index(value) for key, value in fields.items() if value in headers}


def read_registry(data: bytes, source: str) -> Registry:
    groups = defaultdict(list)
    book = open_excel(data)
    try:
        for sheet in book.worksheets:
            rows = sheet_rows(sheet)
            for _, row in islice(rows, 20):
                cols = official_columns([label(value) for value in row])
                if cols:
                    break
            else:
                continue
            for row_index, row in rows:
                identity = as_text(row[cols["ruc"]])
                if not identity:
                    continue
                ruc = normalize_ruc(identity)
                if not ruc:
                    raise DataError(f"El padrón contiene un RUC inválido en la fila {row_index}.")
                subject = Subject(
                    ruc=ruc,
                    name=as_text(row[cols["name"]]),
                    resolution=as_text(row[cols["resolution"]]),
                    published=parse_date(row[cols["published"]]) if "published" in cols else None,
                    final=parse_date(row[cols["final"]]) if "final" in cols else None,
                )
                if subject not in groups[ruc]:
                    groups[ruc].append(subject)
    except DataError:
        raise
    except Exception as exc:
        raise DataError("No se pudo leer el contenido del padrón SSCO.") from exc
    finally:
        book.close()
    if not groups:
        raise DataError("No se encontró el listado SSCO. Carga el Excel de la relación oficial de SUNAT.")
    subjects = {
        ruc: tuple(sorted(entries, key=lambda entry: entry.published or date.min, reverse=True))
        for ruc, entries in groups.items()
    }
    dates = (entry.published for entries in subjects.values() for entry in entries)
    return Registry(subjects, max((item for item in dates if item), default=None), source,
                    hashlib.sha256(data).hexdigest())


def fetch_registry() -> Registry:
    try:
        with requests.get(SUNAT_XLSX, stream=True, timeout=(5, 10)) as response:
            response.raise_for_status()
            if not response.url.startswith("https://www.sunat.gob.pe/"):
                raise DataError("La descarga del padrón redirigió fuera del sitio de SUNAT.")
            buffer = bytearray()
            for part in response.iter_content(chunk_size=65536):
                buffer.extend(part)
                if len(buffer) > MAX_FILE:
                    raise DataError("La descarga del padrón supera el tamaño permitido.")
    except requests.RequestException as exc:
        raise DataError("SUNAT no respondió a la descarga del padrón.") from exc
    return read_registry(bytes(buffer), "Descarga oficial de SUNAT")


# Las posiciones se cuentan desde cero. Las firmas comprueban el diseño antes de usarlas.
ACCOUNTING = {
    "issued": 4, "kind": 6, "series": 7, "number": 9, "identity_type": 10,
    "identity": 11, "provider": 12, "total": 22, "vat": (14, 16, 18),
}
RESULT = {
    "issued": 0, "kind": 2, "series": 3, "number": 5, "identity_type": 6,
    "identity": 7, "provider": 8, "total": 18, "vat": (10, 12, 14), "currency": 27,
}


def generic_columns(headers):
    aliases = {
        "identity": ("ruc proveedor", "ruc del proveedor", "ruc emisor", "ruc del emisor",
                     "numero documento proveedor", "num doc prov", "ruc proveedor emisor"),
        "identity_type": ("tipo doc prov", "tipo doc proveedor", "tipo doc identidad", "tipo documento proveedor"),
        "provider": ("razon social proveedor", "proveedor", "razon social", "nombre proveedor"),
        "issued": ("fecha emision", "fecha de emision", "fecha comprobante"),
        "kind": ("tipo doc", "tipo cp doc", "tipo comprobante", "tipo cdp"),
        "series": ("serie", "serie cdp", "serie del cdp"),
        "number": ("numero", "nro comprobante", "numero comprobante", "nro de comprobante", "numero doc"),
        "total": ("total", "importe total", "monto total", "total cp"),
        "currency": ("moneda", "cod moneda"),
        "period": ("periodo", "periodo tributario"),
    }
    columns = {key: headers.index(name) for key, names in aliases.items() for name in names if name in headers}
    if "identity" not in columns:
        return {}
    columns["vat"] = tuple(index for index, value in enumerate(headers) if value in {"igv", "igv mixta", "igv no grav"})
    return columns


def field(row, columns, key):
    index = columns.get(key)
    return row[index] if index is not None and index < len(row) else None


def make_purchase(row, columns, source, section, index, layout):
    get = lambda key: field(row, columns, key)
    currency = as_text(get("currency")).upper()
    currency = {"S": "PEN", "D": "USD"}.get(currency, currency)
    if layout == "NewContaSis" and currency == "USD":
        currency = "Ver origen"
    return Purchase(
        source=source, section=section, row=index, layout=layout,
        issued=parse_date(get("issued")), kind=as_text(get("kind")),
        series=as_text(get("series")), number=as_text(get("number")),
        provider=as_text(get("provider")), identity_type=as_text(get("identity_type")).upper(),
        identity=as_text(get("identity")), total=amount(get("total")),
        vat=tax(row, columns.get("vat", ())), currency=currency, period=as_text(get("period")),
    )


def purchase_row(row, columns):
    return bool(
        as_text(field(row, columns, "number"))
        and (as_text(field(row, columns, "kind")) or parse_date(field(row, columns, "issued")))
    )


def excel_layout(sheet, preview):
    for index, row in preview:
        headers = [label(value) for value in row]
        if headers[10] == "informacion del proveedor" and headers[22] == "importe total":
            return "Registro contable", ACCOUNTING, index + 3, False
        generic = generic_columns(headers)
        if generic:
            if headers[:9] == ["fecha emision", "fecha venc pago", "tipo doc", "serie", "ano dua",
                               "numero", "tipo doc prov", "num doc prov", "razon social"]:
                return "NewContaSis", RESULT, index, False
            return "Excel con encabezados", generic, index, True
    if label(sheet.title) == "comprobantes":
        for index, row in preview:
            if (index >= 4 and parse_date(row[0]) and as_text(row[2]) and as_text(row[5])
                    and as_text(row[6]).isdigit() and (as_text(row[7]) or as_text(row[8]))):
                return "NewContaSis", RESULT, index - 1, False
    return None


def read_purchases_excel(data, source):
    results = []
    formats = []
    book = open_excel(data)
    try:
        for sheet in book.worksheets:
            rows = sheet_rows(sheet)
            preview = list(islice(rows, 30))
            layout = excel_layout(sheet, preview)
            if layout is None:
                continue
            name, columns, last_header, generic = layout
            formats.append(name)
            for index, row in chain(preview, rows):
                if index <= last_header:
                    continue
                has_identity = as_text(field(row, columns, "identity")) or as_text(field(row, columns, "provider"))
                if purchase_row(row, columns) or (generic and has_identity):
                    results.append(make_purchase(row, columns, source, sheet.title, index, name))
                if len(results) > MAX_ROWS:
                    raise DataError(f"El archivo contiene más de {MAX_ROWS:,} comprobantes.")
    except DataError:
        raise
    except Exception as exc:
        raise DataError("No se pudo leer el contenido del Excel de compras.") from exc
    finally:
        book.close()
    if not formats:
        raise DataError("Formato de Excel no reconocido. Usa el registro contable, NewContaSis o una columna «RUC proveedor».")
    return results


def sire_columns(headers):
    if headers[:4] != ["ruc", "apellidos y nombres o razon social", "periodo", "car sunat"]:
        return {}
    fields = {
        "issued": "fecha de emision", "kind": "tipo cp doc", "series": "serie del cdp",
        "number": "nro cp o doc nro inicial rango", "identity_type": "tipo doc identidad",
        "identity": "nro doc identidad", "provider": "apellidos nombres razon social",
        "total": "total cp", "currency": "moneda", "period": "periodo",
    }
    if not all(value in headers for value in fields.values()):
        return {}
    columns = {key: headers.index(value) for key, value in fields.items()}
    columns["vat"] = tuple(index for index, value in enumerate(headers) if value in {
        "igv ipm dg", "igv ipm dgng", "igv ipm dng"
    })
    return columns


def read_sire(data, source, section):
    file_bytes(data)
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            decoded = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    first_line = next((line for line in decoded.splitlines() if line.strip()), "")
    delimiter = max(("|", ",", ";", "\t"), key=first_line.count)
    quoting = csv.QUOTE_NONE if delimiter == "|" else csv.QUOTE_MINIMAL
    reader = csv.reader(StringIO(decoded, newline=""), delimiter=delimiter, quoting=quoting, strict=True)
    try:
        headers = next((row for row in reader if any(as_text(value) for value in row)), [])
        # Algunas exportaciones delimitadas por barra incluyen un separador al final de cada línea.
        trailing_pipe = delimiter == "|" and headers and headers[-1] == ""
        if trailing_pipe:
            headers.pop()
        columns = sire_columns([label(value) for value in headers])
        if not columns:
            raise DataError("El TXT o CSV no tiene los encabezados de una propuesta de compras SIRE.")
        results = []
        for row in reader:
            index = reader.line_num
            if not any(as_text(value) for value in row):
                continue
            if trailing_pipe and len(row) == len(headers) + 1 and row[-1] == "":
                row.pop()
            if len(row) != len(headers):
                raise DataError(f"La propuesta tiene un número inesperado de columnas en la línea {index}.")
            results.append(make_purchase(row, columns, source, section, index, "Propuesta SIRE"))
            if len(results) > MAX_ROWS:
                raise DataError(f"La propuesta contiene más de {MAX_ROWS:,} comprobantes.")
        return results
    except csv.Error as exc:
        raise DataError(f"No se pudo leer la propuesta cerca de la línea {reader.line_num}.") from exc


def read_purchases(data: bytes, name: str, source: str | None = None) -> list[Purchase]:
    file_bytes(data)
    source = source or name
    extension = PurePosixPath(name).suffix.lower()
    if extension == ".xlsx":
        return read_purchases_excel(data, source)
    if extension in (".csv", ".txt"):
        return read_sire(data, source, "")
    if extension != ".zip":
        raise DataError("Carga un XLSX o una propuesta SIRE en ZIP, TXT o CSV.")
    try:
        with ZipFile(BytesIO(data)) as archive:
            members = [
                item for item in archive_members(archive)
                if not item.is_dir() and PurePosixPath(item.filename).suffix.lower() in {".txt", ".csv"}
            ]
            if not members:
                raise DataError("El ZIP debe contener una propuesta SIRE en formato TXT o CSV.")
            result = []
            errors = []
            seen = set()
            for item in members:
                if item.file_size > MAX_FILE:
                    raise DataError(f"Un archivo del ZIP supera el máximo de {MAX_FILE_MB} MB.")
                content = archive.read(item)
                fingerprint = hashlib.sha256(content).digest()
                if fingerprint in seen:
                    continue
                try:
                    result.extend(read_sire(content, source, PurePosixPath(item.filename).name))
                    seen.add(fingerprint)
                except DataError as exc:
                    errors.append(f"«{PurePosixPath(item.filename).name}»: {exc}")
                if len(result) > MAX_ROWS:
                    raise DataError(f"El ZIP contiene más de {MAX_ROWS:,} comprobantes.")
            if errors:
                raise DataError(" ".join(errors))
            return result
    except (BadZipFile, RuntimeError, NotImplementedError) as exc:
        raise DataError("No se pudo abrir el ZIP de la propuesta.") from exc


def analyze(purchases, registry):
    return [Review(purchase, registry.subjects.get(purchase.ruc, ())) for purchase in purchases]


REPORT_HEADERS = (
    "Estado", "RUC proveedor", "Proveedor", "Emisión", "Tipo CP", "Serie", "Número",
    "Periodo SIRE", "Total informado", "IGV informado", "Moneda indicada", "Razón social SUNAT",
    "Resolución SSCO", "Publicación SSCO", "Resolución firme", "Resoluciones en el padrón",
    "Tipo doc. proveedor", "Documento proveedor", "Formato", "Archivo", "Hoja o archivo interno", "Fila",
)


def report_values(review):
    purchase, subject = review.purchase, review.subject
    return (
        review.status, purchase.ruc, purchase.provider, purchase.issued, purchase.kind,
        purchase.series, purchase.number, purchase.period, purchase.total, purchase.vat, purchase.currency,
        subject.name if subject else "", subject.resolution if subject else "",
        subject.published if subject else None, subject.final if subject else None,
        len(review.subjects), purchase.identity_type, purchase.identity, purchase.layout,
        purchase.source, purchase.section, purchase.row,
    )


def export_report(reviews, registry, consulted, inputs):
    book = Workbook(write_only=True)
    matches = [review for review in reviews if review.subject]
    pending = [review for review in reviews if not review.purchase.ruc]
    sections = (
        ("Consulta", ("Campo", "Valor"), (
            ("Consulta en hora de Perú", consulted),
            ("Origen del padrón", registry.source),
            ("Última publicación en el padrón", registry.latest_publication),
            ("RUC distintos en el padrón", len(registry.subjects)),
            ("Huella SHA-256 del padrón", registry.fingerprint),
            ("Filas leídas", len(reviews)),
            ("Coincidencias SSCO", len(matches)),
            ("Filas sin RUC consultable", len(pending)),
            ("Criterio", "Coincidencia exacta con el padrón consultado. La coincidencia marca el comprobante para revisión."),
        )),
        ("Archivos", ("Archivo", "Estado", "Filas"), inputs),
        ("Coincidencias SSCO", REPORT_HEADERS, map(report_values, matches)),
        ("Cruce completo", REPORT_HEADERS, map(report_values, reviews)),
        ("Sin RUC consultable", REPORT_HEADERS, map(report_values, pending)),
    )
    heading_font = Font(bold=True, color="FFFFFF")
    heading_fill = PatternFill("solid", fgColor="153C36")
    for name, headers, values in sections:
        sheet = book.create_sheet(name)
        sheet.freeze_panes = "A2"
        for index, title in enumerate(headers, 1):
            sheet.column_dimensions[get_column_letter(index)].width = max(14, len(title) + 3)
        if name == "Consulta":
            sheet.column_dimensions["A"].width = 38
            sheet.column_dimensions["B"].width = 84
        elif name == "Archivos":
            sheet.column_dimensions["A"].width = 48
            sheet.column_dimensions["B"].width = 75
        heading_cells = []
        for title in headers:
            cell = WriteOnlyCell(sheet, value=title)
            cell.font = heading_font
            cell.fill = heading_fill
            heading_cells.append(cell)
        sheet.append(heading_cells)
        row_number = 1
        for row_number, row in enumerate(values, 2):
            cells = []
            for value in row:
                cell = WriteOnlyCell(sheet, value=value)
                if isinstance(value, str):
                    cell.data_type = "s"
                elif isinstance(value, date):
                    cell.number_format = "dd/mm/yyyy"
                elif isinstance(value, Decimal):
                    cell.number_format = '#,##0.00'
                cells.append(cell)
            sheet.append(cells)
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{row_number}"
    result = BytesIO()
    book.save(result)
    return result.getvalue()
