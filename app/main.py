"""Local web app: drop a NetCHB manifest report, get an AAMS/ECCF .xls back."""

from __future__ import annotations

import base64
import re
from pathlib import Path

from fastapi import FastAPI, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates

from .aams import FlightInfo, convert, write_checklist
from .netchb import NetchbError, read_manifest

app = FastAPI(title="Bombino Document Converter")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

# Kept small on purpose: this runs on a staff laptop, not a server farm.
MAX_UPLOAD_BYTES = 25 * 1024 * 1024

_SAFE_NAME = re.compile(r"[^A-Za-z0-9 _.&-]")


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {})


@app.post("/convert", response_class=HTMLResponse)
async def convert_upload(
    request: Request,
    manifest_file: UploadFile,
    man_code: str = Form(""),
    flight_info: str = Form(""),
    man_origin: str = Form(""),
    man_dest: str = Form(""),
):
    contents = await manifest_file.read()

    def fail(message: str):
        return templates.TemplateResponse(
            request, "index.html", {"error": message}, status_code=400)

    if not contents:
        return fail("No file was uploaded.")
    if len(contents) > MAX_UPLOAD_BYTES:
        return fail("That file is larger than 25 MB - it is probably not a "
                    "NetCHB manifest report.")
    if not (manifest_file.filename or "").lower().endswith(".xls"):
        return fail("Please upload the NetCHB report as a .xls file. If yours "
                    "is .xlsx, open it in Excel and re-save as "
                    "'Excel 97-2003 Workbook (*.xls)'.")

    try:
        manifest = read_manifest(contents)
    except NetchbError as exc:
        return fail(str(exc))

    # Every field is optional; blanks stay blank and land in the manual-fill list.
    flight = FlightInfo(
        man_code=man_code.strip().upper(),
        flight_info=flight_info.strip().upper(),
        man_origin=man_origin.strip().upper(),
        man_dest=man_dest.strip().upper(),
    )

    payload, result = convert(manifest, flight)
    checklist = write_checklist(result)

    # Name the file after the manifest code when given, else after the MAWB.
    suffix = f" {flight.man_code or manifest.mawb}".rstrip()
    filename = _SAFE_NAME.sub("", f"aams-data-bombino-express-eccf{suffix}.xls")

    return templates.TemplateResponse(request, "index.html", {
        "download_name": filename,
        "download_data": base64.b64encode(payload).decode("ascii"),
        "stats": {
            "shipments": result.shipment_count,
            "commodities": result.commodity_count,
            "rows": len(result.rows),
            "mawb": manifest.mawb,
            "man_date": int(manifest.man_date) if manifest.man_date else "",
        },
        "manual_fill": result.manual_fill,
        "adjustments": result.adjustments,
        "checklist_name": filename.replace(".xls", " - to do.xls"),
        "checklist_data": base64.b64encode(checklist).decode("ascii"),
    })
