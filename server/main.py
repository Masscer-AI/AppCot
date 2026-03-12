from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Annotated
from datetime import datetime
import json

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel, ConfigDict, Field
from starlette.background import BackgroundTask
import uvicorn


BASE_DIR = Path(__file__).resolve().parent
PRICES_PATH = BASE_DIR / "materiales" / "prices.json"
TARGET_SHEET_NAME = "Formato No. 1"
MAX_ITEMS = 4
COMMISSION_FACTOR = 1.15
SPANISH_MONTHS = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def remove_temp_file(path: str) -> None:
    try:
        Path(path).unlink(missing_ok=True)
    except OSError:
        pass


def get_today_date_spanish() -> str:
    today = datetime.now()
    month_name = SPANISH_MONTHS[today.month]
    return f"{today.day} de {month_name} de {today.year}"


def parse_excel_number(value: str | int | float | None) -> int | float | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else value

    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        parsed = float(text)
        return int(parsed) if parsed.is_integer() else parsed
    except ValueError:
        return None


def clear_product_row(sheet, row: int, start_col: str = "B", end_col: str = "M") -> None:
    for col_ord in range(ord(start_col), ord(end_col) + 1):
        sheet[f"{chr(col_ord)}{row}"] = None


def get_material_record(material_name: str) -> dict | None:
    if not PRICES_PATH.exists():
        return None
    try:
        with PRICES_PATH.open("r", encoding="utf-8") as file:
            prices_data = json.load(file)
    except Exception:  # noqa: BLE001
        return None

    tapas = prices_data.get("materiales", {}).get("tapas", [])
    return next(
        (item for item in tapas if str(item.get("name", "")).strip().lower() == material_name.strip().lower()),
        None,
    )


def get_milesimas_for_material(material_name: str, calibre_micras: str) -> str | None:
    material = get_material_record(material_name)
    if material is None:
        return None
    try:
        micras_key = str(int(float(calibre_micras.strip())))
    except ValueError:
        return None
    values = material.get("prices_by_micras", {}).get(micras_key)
    if values is None or values.get("espesor_milesimas") is None:
        return None
    return f"{float(values['espesor_milesimas']):.1f}"


def get_price_for_material(material_name: str, calibre_micras: str) -> float | None:
    material = get_material_record(material_name)
    if material is None:
        return None
    try:
        micras_key = str(int(float(calibre_micras.strip())))
    except ValueError:
        return None
    values = material.get("prices_by_micras", {}).get(micras_key)
    if values is None or values.get("price") is None:
        return None
    try:
        return float(values["price"])
    except (TypeError, ValueError):
        return None


class QuoteRequest(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    company_name: Annotated[str, Field(min_length=1, alias="companyName")]
    full_name: Annotated[str, Field(default="", alias="fullName")]
    product_name: Annotated[str, Field(default="AMILEN ML", alias="productName")]
    top_calibre: Annotated[str, Field(default="", alias="topCalibre")]
    product_side: Annotated[str, Field(default="TAPA", alias="productSide")]
    line_product: Annotated[str, Field(default="", alias="lineProduct")]
    barrier_type: Annotated[str, Field(default="alta", alias="barrierType")]
    seal_type: Annotated[str, Field(default="hermetico", alias="sealType")]
    item_type: Annotated[str, Field(default="", alias="itemType")]
    item_calibre: Annotated[str, Field(default="", alias="itemCalibre")]
    item_width: Annotated[str | int | float, Field(default="", alias="itemWidth")]
    monthly_scale: Annotated[str | int | float, Field(default="", alias="monthlyMeters")]
    items: list[dict] = Field(default_factory=list, alias="items", max_length=MAX_ITEMS)


app = FastAPI(title="Multivac Quote API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "http://localhost:3001",
        "http://localhost:3002",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/quotes/generate")
def generate_quote_excel(payload: QuoteRequest) -> FileResponse:
    with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        output_path = temp_file.name

    try:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = TARGET_SHEET_NAME

        thin = Side(border_style="thin", color="111111")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        dark_fill = PatternFill("solid", fgColor="2F2F2F")
        light_fill = PatternFill("solid", fgColor="E8E8E8")
        white_bold = Font(color="FFFFFF", bold=True)
        bold = Font(bold=True)

        column_widths = {
            "B": 25,
            "C": 10,
            "D": 14,
            "E": 47,
            "F": 8,
            "G": 10,
            "H": 11,
            "I": 10,
            "J": 11,
            "K": 11,
            "L": 11,
        }
        for col, width in column_widths.items():
            sheet.column_dimensions[col].width = width

        # Heading
        sheet.merge_cells("B2:C2")
        sheet["B2"] = "Version: 03"
        sheet.merge_cells("B3:C3")
        sheet["B3"] = "Pagina: 1"

        sheet.merge_cells("I2:M2")
        sheet["I2"] = "APPCOT"
        sheet["I2"].font = Font(bold=True, size=18)
        sheet.merge_cells("I3:M3")
        sheet["I3"] = "Codigo: AM-400-000"

        sheet.merge_cells("B5:D5")
        sheet["B5"] = payload.company_name
        sheet["B5"].fill = dark_fill
        sheet["B5"].font = white_bold
        sheet["B5"].alignment = left
        sheet["B5"].border = border

        attention_name = payload.full_name.strip() or "Nombre y Apellido"
        sheet.merge_cells("E5:H5")
        sheet["E5"] = f"Attn: {attention_name}"
        sheet["E5"].fill = dark_fill
        sheet["E5"].font = white_bold
        sheet["E5"].alignment = center
        sheet["E5"].border = border

        sheet.merge_cells("I5:J5")
        sheet["I5"] = "Fecha:"
        sheet["I5"].fill = dark_fill
        sheet["I5"].font = white_bold
        sheet["I5"].alignment = center
        sheet["I5"].border = border

        sheet.merge_cells("K5:M5")
        sheet["K5"] = get_today_date_spanish()
        sheet["K5"].fill = dark_fill
        sheet["K5"].font = Font(color="F28C28", bold=True)
        sheet["K5"].alignment = center
        sheet["K5"].border = border

        # Table header
        header_row = 7
        sheet.row_dimensions[header_row].height = 42
        headers = {
            "B": "Estructura",
            "C": "Tapa /\nFondo",
            "D": "Linea/Producto",
            "E": "Descripcion del Material",
            "F": "Ancho\n(mm)",
            "G": "Longitud\nde bobina\n(m)",
            "H": "Volumen\nanual\nproyectado\n(mts)",
            "I": "Escala\ncotizada\n(mts)",
            "J": "Precio metro\n(USD)",
            "K": "Precio bobina\n(USD)",
            "L": "Precio Km\n(USD)",
        }
        for col, label in headers.items():
            cell = sheet[f"{col}{header_row}"]
            cell.value = label
            cell.font = bold
            cell.fill = light_fill
            cell.alignment = center
            cell.border = border

        monthly_value = parse_excel_number(payload.monthly_scale)
        product_label = payload.product_name.strip() or "AMILEN ML"

        normalized_items = payload.items[:MAX_ITEMS]
        if not normalized_items:
            normalized_items = [
                {
                    "type": payload.item_type.strip() or payload.product_side.strip(),
                    "calibre": payload.item_calibre.strip() or payload.top_calibre.strip(),
                    "width": payload.item_width,
                    "barrierType": payload.barrier_type,
                    "sealType": payload.seal_type,
                }
            ]

        if not normalized_items:
            raise HTTPException(status_code=400, detail="At least one material item is required")

        for index, item in enumerate(normalized_items, start=1):
            calibre_value = str(item.get("calibre", "")).strip()
            width_value = parse_excel_number(item.get("width"))
            if not calibre_value or width_value is None:
                raise HTTPException(
                    status_code=400,
                    detail=f"Item #{index} must include both width and calibre",
                )

        data_start_row = 8
        for index, item in enumerate(normalized_items[:MAX_ITEMS]):
            row = data_start_row + index
            item_type = str(item.get("type", "")).strip().upper()
            item_type = item_type if item_type in {"TAPA", "FONDO"} else "TAPA"
            calibre = str(item.get("calibre", "")).strip()
            width_value = parse_excel_number(item.get("width"))
            barrier = str(item.get("barrierType", "alta")).strip().lower()
            seal = str(item.get("sealType", "hermetico")).strip().lower()
            barrier_text = "mediana barrera" if barrier == "mediana" else "alta barrera"
            seal_text = "pelable" if seal == "pelable" else "hermético"
            milesimas = get_milesimas_for_material(product_label, calibre)
            material_price = get_price_for_material(product_label, calibre)

            sheet[f"B{row}"] = f"{product_label} {calibre}".strip()
            sheet[f"C{row}"] = item_type
            sheet[f"D{row}"] = payload.line_product.strip()
            if milesimas:
                sheet[f"E{row}"] = (
                    f"Material coextruido y laminado, {barrier_text}, "
                    f"sello {seal_text} {milesimas} mil de espesor"
                )
            else:
                sheet[f"E{row}"] = f"Material coextruido y laminado, {barrier_text}, sello {seal_text}"

            sheet[f"F{row}"] = width_value if width_value is not None else ""
            sheet[f"G{row}"] = 914
            sheet[f"H{row}"] = "TBD"
            sheet[f"I{row}"] = monthly_value if monthly_value is not None else ""

            qmil = None
            pmil = None
            pbase = None
            price_km = None
            price_m = None
            price_bobina = None
            if material_price is not None and width_value is not None:
                qmil = 100000 / float(width_value)
                if qmil > 0:
                    pmil = material_price / qmil
                    pbase = pmil * COMMISSION_FACTOR
                    price_km = pbase * 1000
                    price_m = round(price_km / 1000, 3)
                    price_bobina = round(price_m * 914, 2)
                    sheet[f"J{row}"] = price_m
                    sheet[f"K{row}"] = price_bobina
                    sheet[f"L{row}"] = round(price_km, 2)
                else:
                    sheet[f"J{row}"] = ""
                    sheet[f"K{row}"] = ""
                    sheet[f"L{row}"] = ""
            else:
                sheet[f"J{row}"] = ""
                sheet[f"K{row}"] = ""
                sheet[f"L{row}"] = ""

            for col in headers.keys():
                cell = sheet[f"{col}{row}"]
                cell.border = border
                cell.alignment = left if col in {"B", "D", "E"} else center

            sheet[f"I{row}"].number_format = "#,##0"
            sheet[f"J{row}"].number_format = "$#,##0.000"
            sheet[f"K{row}"].number_format = "$#,##0.00"
            sheet[f"L{row}"].number_format = "$#,##0.00"

        used_count = len(normalized_items[:MAX_ITEMS])
        for index in range(used_count, MAX_ITEMS):
            clear_product_row(sheet, data_start_row + index, "B", "L")

        # Footer
        footer_start = data_start_row + max(used_count, 1) + 2
        footer_end = footer_start + 4
        sheet.merge_cells(f"B{footer_start}:I{footer_start + 1}")
        sheet[f"B{footer_start}"] = (
            "Los precios anteriores son en USD(*), al tipo de cambio del dia de la facturacion, "
            "no incluye IVA y son DDP, credito: Por definir."
        )
        sheet[f"B{footer_start}"].alignment = left
        sheet[f"B{footer_start}"].border = border

        sheet.merge_cells(f"B{footer_start + 2}:I{footer_end}")
        sheet[f"B{footer_start + 2}"] = (
            "Consulte los terminos y condiciones en la siguiente liga:\n"
            "Terms and Conditions | APPCOT\n\n"
            "Debido a la situacion actual de las materias primas y a las considerables "
            "fluctuaciones de las mismas, nos reservamos el derecho de ajustar los precios.\n"
            "La presente cancela cualquier cotizacion anterior y los precios son vigentes "
            "durante un periodo de 15 dias."
        )
        sheet[f"B{footer_start + 2}"].alignment = left
        sheet[f"B{footer_start + 2}"].border = border

        sheet.merge_cells(f"J{footer_start}:L{footer_end}")
        sheet[f"J{footer_start}"] = (
            "APPCOT\n"
            "Aldo Manzur Coronel\n"
            "aldo.manzur@mx.multivac.com\n"
            "Celular 55 3232 7977\n\n"
            "[QR Placeholder]"
        )
        sheet[f"J{footer_start}"].alignment = center
        sheet[f"J{footer_start}"].font = bold
        sheet[f"J{footer_start}"].border = border

        workbook.save(output_path)
        workbook.close()
    except HTTPException:
        remove_temp_file(output_path)
        raise
    except Exception as exc:  # noqa: BLE001
        remove_temp_file(output_path)
        raise HTTPException(status_code=500, detail=f"Failed to generate Excel: {exc}") from exc

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="cotizacion-generada.xlsx",
        background=BackgroundTask(remove_temp_file, output_path),
    )


@app.get("/api/materiales/tapas/calibres")
def get_tapa_calibres(
    material_name: Annotated[str, Query(alias="materialName")] = "Amilen ML",
) -> dict:
    material = get_material_record(material_name)
    if material is None:
        raise HTTPException(status_code=404, detail=f'Material "{material_name}" not found')

    prices_by_micras = material.get("prices_by_micras", {})
    calibres = []
    for micras_key, values in prices_by_micras.items():
        micras = int(micras_key)
        milesimas_raw = values.get("espesor_milesimas")
        milesimas = f"{float(milesimas_raw):.1f}"
        calibres.append({"micras": micras, "milesimas": milesimas})

    calibres.sort(key=lambda item: item["micras"])
    return {"material": material_name, "calibres": calibres}


if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=8009, reload=True)
