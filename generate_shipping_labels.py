from __future__ import annotations

import argparse
import io
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Iterable

from openpyxl import load_workbook

REPORTLAB_IMPORT_ERROR: ModuleNotFoundError | None = None

try:
    from reportlab.graphics import renderPDF
    from reportlab.graphics.barcode import code128
    from reportlab.graphics.barcode import createBarcodeDrawing
    from reportlab.lib.colors import HexColor
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas
except ModuleNotFoundError as exc:
    REPORTLAB_IMPORT_ERROR = exc
    renderPDF = None
    code128 = None
    createBarcodeDrawing = None
    HexColor = None
    canvas = None

    def stringWidth(*args: object, **kwargs: object) -> float:
        raise ModuleNotFoundError(
            "Falta la dependencia 'reportlab'. Instalala con: pip install reportlab"
        ) from REPORTLAB_IMPORT_ERROR


POINTS_PER_INCH = 72.0
INCH = POINTS_PER_INCH
PAGE_WIDTH = 6 * POINTS_PER_INCH
PAGE_HEIGHT = 4 * POINTS_PER_INCH
PAGE_SIZE = (PAGE_WIDTH, PAGE_HEIGHT)
MARGIN = 0.14 * POINTS_PER_INCH
CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN)


def ensure_reportlab() -> None:
    if REPORTLAB_IMPORT_ERROR is not None:
        raise ModuleNotFoundError(
            "Falta la dependencia 'reportlab'. Instalala con: pip install reportlab"
        ) from REPORTLAB_IMPORT_ERROR


def normalize_key(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def pick_first(record: dict[str, object], aliases: Iterable[str], default: str = "") -> str:
    for alias in aliases:
        key = normalize_key(alias)
        if key in record:
            value = record[key]
            if value is None:
                continue
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            return str(value).strip()
    return default


def to_float(value: str) -> float | None:
    if not value:
        return None
    try:
        return float(str(value).replace(",", "."))
    except ValueError:
        return None


def wrap_text(text: str, font_name: str, font_size: float, max_width: float) -> list[str]:
    words = (text or "").split()
    if not words:
        return []

    lines: list[str] = []
    current = words[0]

    for word in words[1:]:
        test_line = f"{current} {word}"
        if stringWidth(test_line, font_name, font_size) <= max_width:
            current = test_line
            continue
        lines.append(current)
        current = word

    lines.append(current)
    return lines


def draw_wrapped(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    width: float,
    font_name: str = "Helvetica",
    font_size: float = 8.5,
    leading: float = 10.0,
    max_lines: int | None = None,
) -> float:
    lines = wrap_text(text, font_name, font_size, width)
    if max_lines is not None and len(lines) > max_lines:
        lines = lines[:max_lines]
        if lines:
            trimmed = lines[-1]
            while trimmed and stringWidth(trimmed + "...", font_name, font_size) > width:
                trimmed = trimmed[:-1]
            lines[-1] = (trimmed.rstrip() + "...") if trimmed else "..."

    pdf.setFont(font_name, font_size)
    cursor = y
    for line in lines:
        pdf.drawString(x, cursor, line)
        cursor -= leading
    return cursor


@dataclass
class LabelRow:
    shipment_number: str
    shipment_date: str
    sender_name: str
    sender_address: str
    sender_phone: str
    sender_city: str
    sender_state: str
    recipient_name: str
    recipient_address: str
    recipient_phone: str
    recipient_city: str
    recipient_state: str
    content: str
    pieces: str
    weight_lb: str
    weight_kg: str
    declared_value: str
    tariff_code: str
    manifest: str
    instructions: str
    cost: str

    @classmethod
    def from_record(cls, record: dict[str, object]) -> "LabelRow":
        return cls(
            shipment_number=pick_first(record, ["envio", "guia", "numero de guia", "numero guia"]),
            shipment_date=pick_first(record, ["fecha guia", "fecha", "fecha envio"]),
            sender_name=pick_first(record, ["compania remitente", "remitente", "remitente nombre"]),
            sender_address=pick_first(record, ["remitente direccion", "direccion remitente"]),
            sender_phone=pick_first(record, ["remitente telefono", "telefono remitente"]),
            sender_city=pick_first(record, ["remitente ciudad", "ciudad remitente"]),
            sender_state=pick_first(record, ["remitente estado", "estado remitente"]),
            recipient_name=pick_first(record, ["nombre destino", "destinatario", "nombre destinatario"]),
            recipient_address=pick_first(record, ["destino direccion", "direccion destino", "direccion destinatario"]),
            recipient_phone=pick_first(record, ["destino telefono", "telefono destino", "telefono destinatario"]),
            recipient_city=pick_first(record, ["destino ciudad", "ciudad destino"]),
            recipient_state=pick_first(record, ["destino estado", "estado destino"]),
            content=pick_first(record, ["contenido", "descripcion", "producto"]),
            pieces=pick_first(record, ["piezas", "cantidad piezas"], "1"),
            weight_lb=pick_first(record, ["peso libras", "peso lb"]),
            weight_kg=pick_first(record, ["peso kilos", "peso kg"]),
            declared_value=pick_first(record, ["valor declarado", "valor"]),
            tariff_code=pick_first(record, ["posicion arancelaria", "partida arancelaria", "hs code"]),
            manifest=pick_first(record, ["manifiesto"]),
            instructions=pick_first(record, ["instrucciones", "observaciones"]),
            cost=pick_first(record, ["costo", "coste"]),
        )


def read_rows(workbook_source: str | Path | BinaryIO, sheet_name: str | None) -> list[LabelRow]:
    workbook = load_workbook(workbook_source, data_only=True)
    worksheet = workbook[sheet_name] if sheet_name else workbook[workbook.sheetnames[0]]

    headers = [normalize_key(cell.value) for cell in worksheet[1]]
    rows: list[LabelRow] = []

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if all(value in (None, "") for value in row):
            continue
        record = {
            headers[index]: value
            for index, value in enumerate(row)
            if index < len(headers) and headers[index]
        }
        label = LabelRow.from_record(record)
        if not label.shipment_number:
            continue
        rows.append(label)

    return rows


def format_weight(lb: str, kg: str) -> str:
    lb_value = to_float(lb)
    kg_value = to_float(kg)
    parts: list[str] = []
    if lb_value is not None:
        parts.append(f"{lb_value:.2f} Lb")
    elif lb:
        parts.append(f"{lb} Lb")
    if kg_value is not None:
        parts.append(f"{kg_value:.2f} Kg")
    elif kg:
        parts.append(f"{kg} Kg")
    return " - ".join(parts)


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", value.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:120] or "etiqueta"


def draw_centered_block(
    pdf: canvas.Canvas,
    center_x: float,
    top_y: float,
    lines: list[tuple[str, str, float, str]],
    leading: float,
) -> float:
    y = top_y
    for text, font_name, font_size, color in lines:
        pdf.setFillColor(HexColor(color))
        pdf.setFont(font_name, font_size)
        pdf.drawCentredString(center_x, y, text)
        y -= leading
    return y


def draw_box_icon(pdf: canvas.Canvas, center_x: float, center_y: float, size: float) -> None:
    half = size / 2
    left = center_x - half
    bottom = center_y - half
    pdf.setStrokeColor(HexColor("#6B7280"))
    pdf.setLineWidth(1)
    pdf.rect(left, bottom, size * 0.72, size * 0.62, stroke=1, fill=0)
    pdf.line(left, bottom + size * 0.62, left + size * 0.22, bottom + size * 0.78)
    pdf.line(left + size * 0.22, bottom + size * 0.78, left + size * 0.72, bottom + size * 0.62)
    pdf.line(left + size * 0.36, bottom + size * 0.70, left + size * 0.36, bottom + size * 0.08)
    pdf.line(left + size * 0.22, bottom + size * 0.78, left + size * 0.48, bottom + size * 0.60)
    pdf.line(left + size * 0.48, bottom + size * 0.60, left + size * 0.48, bottom + size * 0.42)
    pdf.line(left + size * 0.60, bottom + size * 0.68, left + size * 0.86, bottom + size * 0.94)
    pdf.line(left + size * 0.86, bottom + size * 0.94, left + size * 0.72, bottom + size * 0.62)
    pdf.line(left + size * 0.06, bottom + size * 0.16, left + size * 0.16, bottom + size * 0.12)
    pdf.line(left + size * 0.10, bottom + size * 0.22, left + size * 0.20, bottom + size * 0.18)


def draw_warning_icon(pdf: canvas.Canvas, center_x: float, center_y: float, size: float) -> None:
    pdf.setStrokeColor(HexColor("#111827"))
    pdf.setLineWidth(1.2)
    path = pdf.beginPath()
    path.moveTo(center_x, center_y + size * 0.55)
    path.lineTo(center_x - size * 0.52, center_y - size * 0.35)
    path.lineTo(center_x + size * 0.52, center_y - size * 0.35)
    path.close()
    pdf.drawPath(path, stroke=1, fill=0)
    pdf.setFont("Helvetica-Bold", size * 0.65)
    pdf.drawCentredString(center_x, center_y - size * 0.12, "!")


def draw_house_icon(pdf: canvas.Canvas, center_x: float, center_y: float, size: float) -> None:
    left = center_x - size * 0.36
    bottom = center_y - size * 0.34
    pdf.setStrokeColor(HexColor("#6B7280"))
    pdf.setLineWidth(1)
    path = pdf.beginPath()
    path.moveTo(center_x, center_y + size * 0.48)
    path.lineTo(center_x - size * 0.46, center_y)
    path.lineTo(center_x - size * 0.34, center_y)
    path.lineTo(center_x - size * 0.34, bottom)
    path.lineTo(center_x - size * 0.10, bottom)
    path.lineTo(center_x - size * 0.10, bottom + size * 0.24)
    path.lineTo(center_x + size * 0.10, bottom + size * 0.24)
    path.lineTo(center_x + size * 0.10, bottom)
    path.lineTo(center_x + size * 0.34, bottom)
    path.lineTo(center_x + size * 0.34, center_y)
    path.lineTo(center_x + size * 0.46, center_y)
    path.close()
    pdf.drawPath(path, stroke=1, fill=0)


def build_sender_lines(row: LabelRow) -> list[str]:
    lines = [row.sender_name or "N/D"]
    if row.sender_address:
        lines.append(row.sender_address)
    city_line = " ".join(part for part in [row.sender_city, row.sender_state] if part)
    if city_line:
        lines.append(f"{city_line} U.S.A.")
    return lines[:3]


def build_recipient_lines(row: LabelRow) -> list[str]:
    lines = [row.recipient_name or "N/D"]
    if row.recipient_address:
        lines.extend(wrap_text(row.recipient_address, "Helvetica", 7.4, 148)[:2])
    city_line = ", ".join(part for part in [row.recipient_city, row.recipient_state] if part)
    if city_line:
        lines.append(city_line)
    if row.recipient_phone:
        lines.append(f"Tel: {row.recipient_phone}")
    return lines[:4]


def draw_code128_barcode(
    pdf: canvas.Canvas,
    value: str,
    x: float,
    y: float,
    max_width: float,
    height: float,
) -> None:
    ensure_reportlab()
    drawing = createBarcodeDrawing(
        "Code128",
        value=value,
        barHeight=height,
        humanReadable=False,
    )
    scale = min(max_width / drawing.width, 1.5)
    drawing.scale(scale, 1)
    actual_width = drawing.width * scale
    drawing_x = x + (max_width - actual_width) / 2
    renderPDF.draw(drawing, pdf, drawing_x, y)


def draw_label(pdf: canvas.Canvas, row: LabelRow) -> None:
    ensure_reportlab()
    pdf.setFillColor(HexColor("#FFFFFF"))
    pdf.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, stroke=0, fill=1)

    outer_border = HexColor("#6B7280")
    divider = HexColor("#BFC5CD")
    split_x = PAGE_WIDTH * 0.505
    left_x0 = MARGIN
    left_x1 = split_x - 6
    right_x0 = split_x + 8
    right_x1 = PAGE_WIDTH - MARGIN
    left_width = left_x1 - left_x0
    right_width = right_x1 - right_x0

    pdf.setStrokeColor(outer_border)
    pdf.setLineWidth(1)
    pdf.rect(0.8, 0.8, PAGE_WIDTH - 1.6, PAGE_HEIGHT - 1.6, stroke=1, fill=0)
    pdf.line(split_x, 0, split_x, PAGE_HEIGHT)

    barcode_value = row.shipment_number.strip()
    draw_code128_barcode(
        pdf,
        barcode_value,
        left_x0 + 4,
        PAGE_HEIGHT - 56,
        left_width - 8,
        34,
    )
    pdf.setFillColor(HexColor("#111827"))
    pdf.setFont("Helvetica", 8.3)
    pdf.drawCentredString(left_x0 + left_width / 2, PAGE_HEIGHT - 62, barcode_value)

    y = PAGE_HEIGHT - 96
    y = draw_centered_block(
        pdf,
        left_x0 + left_width / 2,
        y,
        [
            ("!Hola!", "Helvetica-Bold", 20, "#F97316"),
            ("Somos tu mejor aliado en", "Helvetica", 9.5, "#111827"),
            ("Compras y Envios en el exterior", "Helvetica", 9.5, "#111827"),
            ("Ideal para Negocios", "Helvetica-Bold", 9.5, "#111827"),
        ],
        leading=11,
    )

    box_top = 116
    box_bottom = 68
    label_col_x = left_x0 + 8
    value_x = left_x0 + 82
    envio_x = left_x1 - 64

    pdf.setStrokeColor(divider)
    pdf.setLineWidth(0.8)

    sender_lines = build_sender_lines(row)
    recipient_lines = build_recipient_lines(row)

    pdf.setFillColor(HexColor("#111827"))
    pdf.setFont("Helvetica-Bold", 10)
    sender_label_y = box_top + 18
    recipient_label_y = box_bottom - 7
    pdf.drawString(label_col_x, sender_label_y, "Remitente:")
    pdf.drawString(label_col_x, recipient_label_y, "Destinatario:")

    pdf.setFont("Helvetica", 6.7)
    sender_text_y = box_top + 21
    sender_text_width = envio_x - value_x - 6
    current_sender_y = sender_text_y
    for line in sender_lines:
        current_sender_y = draw_wrapped(
            pdf,
            line,
            value_x,
            current_sender_y,
            sender_text_width,
            font_name="Helvetica",
            font_size=6.7,
            leading=7.2,
            max_lines=2,
        )
        current_sender_y -= 1.5

    recipient_y = box_bottom - 6
    for idx, line in enumerate(recipient_lines):
        pdf.drawString(value_x, recipient_y - (idx * 8.2), line)

    envio_center_x = envio_x + ((left_x1 - envio_x) / 2)
    pdf.setFont("Helvetica-Bold", 7.8)
    pdf.drawCentredString(envio_center_x, box_top + 14, "Envio")
    pdf.setFont("Helvetica-Bold", 15.5)
    pdf.drawCentredString(envio_center_x, box_top - 2, row.shipment_number)

    meta_y = 27
    pdf.setFillColor(HexColor("#111827"))
    draw_wrapped(
        pdf,
        f"Contenido de la caja: {row.content or 'N/D'}",
        left_x0 + 8,
        meta_y,
        left_width - 16,
        font_name="Helvetica-Bold",
        font_size=7.3,
        leading=8.2,
        max_lines=1,
    )
    draw_wrapped(
        pdf,
        f"Piezas: {row.pieces or '1'}   Peso: {format_weight(row.weight_lb, row.weight_kg) or 'N/D'}",
        left_x0 + 8,
        meta_y - 11,
        left_width - 16,
        font_name="Helvetica-Bold",
        font_size=7.3,
        leading=8.2,
        max_lines=1,
    )
    draw_wrapped(
        pdf,
        f"Valor declarado: {row.declared_value or 'N/D'}   Posicion arancelaria: {row.tariff_code or 'N/D'}",
        left_x0 + 8,
        meta_y - 22,
        left_width - 16,
        font_name="Helvetica-Bold",
        font_size=7.3,
        leading=8.2,
        max_lines=1,
    )

    right_center = right_x0 + right_width / 2
    draw_centered_block(
        pdf,
        right_center,
        PAGE_HEIGHT - 26,
        [
            ("Al momento de recibir tu", "Helvetica", 8.5, "#6B7280"),
            ("envio esta etiqueta no debe", "Helvetica", 8.5, "#6B7280"),
            ("estar rota o rasgada", "Helvetica", 8.5, "#6B7280"),
        ],
        leading=10,
    )

    draw_warning_icon(pdf, right_center, PAGE_HEIGHT - 72, 18)
    draw_box_icon(pdf, right_x0 + 42, PAGE_HEIGHT - 122, 34)
    draw_house_icon(pdf, right_x1 - 42, PAGE_HEIGHT - 122, 34)
    pdf.setStrokeColor(HexColor("#9CA3AF"))
    pdf.setLineWidth(0.8)
    pdf.line(right_x0 + 82, PAGE_HEIGHT - 122, right_x1 - 82, PAGE_HEIGHT - 122)

    draw_centered_block(
        pdf,
        right_center,
        PAGE_HEIGHT - 154,
        [
            ("Compra Online", "Helvetica-Bold", 8.4, "#4B5563"),
            ("Personal Shopper", "Helvetica-Bold", 8.4, "#4B5563"),
            ("Recogida en tienda (Pick up)", "Helvetica-Bold", 8.4, "#4B5563"),
        ],
        leading=15,
    )

    attention_y = 54
    pdf.setFillColor(HexColor("#6B7280"))
    pdf.setFont("Helvetica", 8)
    pdf.drawCentredString(right_center, attention_y, "Linea unica de Atencion")
    pdf.setFont("Helvetica", 8.4)
    pdf.drawCentredString(right_center, attention_y - 15, "(318) 242-8086")
    pdf.setFont("Helvetica-Bold", 9)
    pdf.drawCentredString(right_center, attention_y - 41, "www.encargomio.com")


def generate_pdf(rows: list[LabelRow], output_path: Path) -> None:
    ensure_reportlab()
    pdf = canvas.Canvas(str(output_path), pagesize=PAGE_SIZE)

    draw_label(pdf, rows[0])
    pdf.save()


def generate_pdf_bytes(row: LabelRow) -> bytes:
    ensure_reportlab()
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=PAGE_SIZE)
    draw_label(pdf, row)
    pdf.save()
    return buffer.getvalue()


def generate_pdfs(rows: list[LabelRow], output_dir: Path) -> list[Path]:
    ensure_reportlab()
    created_files: list[Path] = []
    used_names: dict[str, int] = {}

    for row in rows:
        base_name = sanitize_filename(row.shipment_number)
        duplicate_number = used_names.get(base_name, 0) + 1
        used_names[base_name] = duplicate_number

        if duplicate_number == 1:
            filename = f"{base_name}.pdf"
        else:
            filename = f"{base_name}_{duplicate_number}.pdf"

        output_path = output_dir / filename
        generate_pdf([row], output_path)
        created_files.append(output_path)

    return created_files


def generate_zip_bytes(rows: list[LabelRow]) -> bytes:
    ensure_reportlab()
    buffer = io.BytesIO()
    used_names: dict[str, int] = {}

    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for row in rows:
            base_name = sanitize_filename(row.shipment_number)
            duplicate_number = used_names.get(base_name, 0) + 1
            used_names[base_name] = duplicate_number

            if duplicate_number == 1:
                filename = f"{base_name}.pdf"
            else:
                filename = f"{base_name}_{duplicate_number}.pdf"

            archive.writestr(filename, generate_pdf_bytes(row))

    return buffer.getvalue()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Genera un PDF de etiquetas 4x6 pulgadas a partir de un archivo Excel."
    )
    parser.add_argument("input", type=Path, help="Ruta del archivo Excel de entrada.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("labels_4x6"),
        help="Carpeta de salida donde se guardara un PDF por fila. Default: labels_4x6",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        type=str,
        default=None,
        help="Nombre de la hoja a procesar. Si no se indica, usa la primera hoja.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    output_dir = args.output.expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"No existe el archivo: {input_path}")

    rows = read_rows(input_path, args.sheet)
    if not rows:
        raise ValueError("No se encontraron filas con numero de envio/guia.")

    output_dir.mkdir(parents=True, exist_ok=True)
    created_files = generate_pdfs(rows, output_dir)
    print(f"Carpeta de salida: {output_dir}")
    print(f"Etiquetas creadas: {len(created_files)}")


if __name__ == "__main__":
    main()
