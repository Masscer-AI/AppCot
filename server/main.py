from pathlib import Path
from shutil import copy2
from tempfile import NamedTemporaryFile
from typing import Annotated
from datetime import datetime
import json

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


class QuoteRequest(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    company_name: Annotated[str, Field(min_length=1, alias="companyName")]
    full_name: Annotated[str, Field(default="", alias="fullName")]


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
