from pathlib import Path
from shutil import copy2
from tempfile import NamedTemporaryFile
from typing import Annotated
from datetime import datetime
import json
import re
import random

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from pydantic import BaseModel, ConfigDict, Field
from starlette.background import BackgroundTask
import uvicorn


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "materiales" / "template.xlsx"
PRICES_PATH = BASE_DIR / "materiales" / "prices.json"
TARGET_SHEET_NAME = "Fromato No. 1"
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
        # If cleanup fails we do not fail the API response.
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


def clear_product_block(sheet, start_row: int, end_row: int, start_col: str = "B", end_col: str = "W") -> None:
    for row in range(start_row, end_row + 1):
        for col_ord in range(ord(start_col), ord(end_col) + 1):
            sheet[f"{chr(col_ord)}{row}"] = None


def get_milesimas_for_material(material_name: str, calibre_micras: str) -> str | None:
    if not PRICES_PATH.exists():
        return None

    micras_key = calibre_micras.strip()
    if not micras_key:
        return None

    try:
        micras_key = str(int(float(micras_key)))
    except ValueError:
        return None

    try:
        with PRICES_PATH.open("r", encoding="utf-8") as file:
            prices_data = json.load(file)
    except Exception:  # noqa: BLE001
        return None

    tapas = prices_data.get("materiales", {}).get("tapas", [])
    material = next(
        (item for item in tapas if str(item.get("name", "")).strip().lower() == material_name.strip().lower()),
        None,
    )
    if material is None:
        return None

    values = material.get("prices_by_micras", {}).get(micras_key)
    if values is None:
        return None

    milesimas_raw = values.get("espesor_milesimas")
    if milesimas_raw is None:
        return None

    return f"{float(milesimas_raw):.1f}"


def get_price_for_material(material_name: str, calibre_micras: str) -> float | None:
    if not PRICES_PATH.exists():
        return None

    micras_key = calibre_micras.strip()
    if not micras_key:
        return None

    try:
        micras_key = str(int(float(micras_key)))
    except ValueError:
        return None

    try:
        with PRICES_PATH.open("r", encoding="utf-8") as file:
            prices_data = json.load(file)
    except Exception:  # noqa: BLE001
        return None

    tapas = prices_data.get("materiales", {}).get("tapas", [])
    material = next(
        (item for item in tapas if str(item.get("name", "")).strip().lower() == material_name.strip().lower()),
        None,
    )
    if material is None:
        return None

    values = material.get("prices_by_micras", {}).get(micras_key)
    if values is None:
        return None

    price_raw = values.get("price")
    if price_raw is None:
        return None

    try:
        return float(price_raw)
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
    items: list[dict] = Field(default_factory=list, alias="items", max_length=4)


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
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="Excel template not found")

    with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        output_path = temp_file.name

    copy2(TEMPLATE_PATH, output_path)

    try:
        workbook = load_workbook(output_path)

        if TARGET_SHEET_NAME not in workbook.sheetnames:
            available = ", ".join(workbook.sheetnames)
            raise HTTPException(
                status_code=500,
                detail=(
                    f'Sheet "{TARGET_SHEET_NAME}" was not found in template. '
                    f"Available sheets: {available}"
                ),
            )

        sheet = workbook[TARGET_SHEET_NAME]
        sheet["B8"] = payload.company_name
        attention_name = payload.full_name.strip() or "Nombre y Apellido"
        sheet["C8"] = f"Attn: {attention_name}"
        sheet["K8"] = get_today_date_spanish()
        monthly_value = parse_excel_number(payload.monthly_scale)

        product_label = payload.product_name.strip() or "AMILEN ML"
        barrier = payload.barrier_type.strip().lower()
        barrier_label = "Mediana barrera" if barrier == "mediana" else "Alta barrera"
        seal_label = "pelable" if payload.seal_type.strip().lower() == "pelable" else "hermético"

        normalized_items = payload.items[:4]
        if not normalized_items:
            normalized_items = [
                {
                    "type": payload.item_type.strip() or payload.product_side.strip(),
                    "calibre": payload.item_calibre.strip() or payload.top_calibre.strip(),
                    "width": payload.item_width,
                }
            ]
        if not normalized_items:
            raise HTTPException(status_code=400, detail="At least one material item is required")

        for index, item in enumerate(normalized_items[:4], start=1):
            calibre_value = str(item.get("calibre", "")).strip()
            width_value = parse_excel_number(item.get("width"))
            if not calibre_value or width_value is None:
                raise HTTPException(
                    status_code=400,
                    detail=f"Item #{index} must include both width and calibre",
                )

        base_row = 11
        row_step = 4
        max_items = 4
        for index, item in enumerate(normalized_items[:max_items]):
            row = base_row + index * row_step
            detail_row = row + 2

            item_type = str(item.get("type", "")).strip().upper()
            item_type = item_type if item_type in {"TAPA", "FONDO"} else "TAPA"
            calibre = str(item.get("calibre", "")).strip()

            width_value = parse_excel_number(item.get("width"))
            item_barrier = str(item.get("barrierType", payload.barrier_type)).strip().lower()
            item_barrier_label = "Mediana barrera" if item_barrier == "mediana" else "Alta barrera"
            item_seal = str(item.get("sealType", payload.seal_type)).strip().lower()
            item_seal_label = "pelable" if item_seal == "pelable" else "hermético"

            sheet[f"B{row}"] = f"{product_label} {calibre}".strip()
            sheet[f"C{row}"] = item_type
            sheet[f"D{row}"] = payload.line_product.strip()
            sheet[f"F{row}"] = width_value if width_value is not None else ""
            sheet[f"I{row}"] = monthly_value if monthly_value is not None else ""
            sheet[f"D{detail_row}"] = item_barrier_label

            milesimas = get_milesimas_for_material(product_label, calibre)
            material_price = get_price_for_material(product_label, calibre)
            if milesimas:
                sheet[f"E{detail_row}"] = (
                    f"al alto vacío o MAP, sello {item_seal_label} {milesimas} mil de espesor"
                )
            else:
                current_desc = str(sheet[f"E{detail_row}"].value or "")
                sheet[f"E{detail_row}"] = re.sub(
                    r"hermético|pelable", item_seal_label, current_desc, flags=re.IGNORECASE
                )

            if material_price is not None:
                sheet[f"R{row}"] = material_price
            commission_factor = random.uniform(1.10, 1.25)
            sheet[f"U{row}"] = f"=T{row}*{commission_factor:.2f}"

        used_count = len(normalized_items[:max_items])
        for index in range(used_count, max_items):
            row = base_row + index * row_step
            detail_row = row + 2
            clear_product_block(sheet, row, detail_row, "B", "W")
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
    if not PRICES_PATH.exists():
        raise HTTPException(status_code=500, detail="Prices file not found")

    try:
        with PRICES_PATH.open("r", encoding="utf-8") as file:
            prices_data = json.load(file)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Failed to read prices: {exc}") from exc

    tapas = prices_data.get("materiales", {}).get("tapas", [])
    material = next((item for item in tapas if item.get("name") == material_name), None)
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
