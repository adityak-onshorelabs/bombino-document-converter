"""Reader for the NetCHB manifest report (legacy .xls).

The report is a flat table with one row per HTS entry line. Shipments are
recovered by grouping consecutive rows on ``House AWB``.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import xlrd


class NetchbError(Exception):
    """The uploaded file is not a usable NetCHB manifest report."""


# Columns the converter cannot work without.
REQUIRED_COLUMNS = (
    "House AWB",
    "GroupIdentifier",
    "Manifest Qty Piece count",
    "ConsigneeName",
    "ManufacturerName",
)


def norm(value) -> str:
    """Render a cell as the string a human would expect.

    xlrd hands back every number as a float, so ``4307923`` arrives as
    ``4307923.0``. Integral floats become plain integer strings; everything
    else is stripped text.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return repr(value).rstrip("0").rstrip(".") if "e" not in repr(value) else str(value)
    return str(value).strip()


def as_number(value):
    """Return a float for numeric-looking input, else ``None``."""
    if isinstance(value, (int, float)):
        return float(value)
    text = norm(value)
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


@dataclass
class Shipment:
    """One House AWB and the entry lines filed against it."""

    hawb: str
    group: str
    rows: list[dict] = field(default_factory=list)

    @property
    def first(self) -> dict:
        return self.rows[0]

    @property
    def has_commodities(self) -> bool:
        """Multi-line shipments carry a commodity block; single-line ones don't.

        Matches the sample, where 103 of 108 commodity-free shipments are
        single-line and every commodity-bearing shipment has at least two.
        """
        return len(self.rows) > 1


@dataclass
class Manifest:
    """A parsed NetCHB report."""

    shipments: list[Shipment]
    master_bill: str
    airline_code: str
    export_date: str  # DD-MM-YYYY as it appears in the report

    @property
    def mawb(self) -> str:
        if self.airline_code and self.master_bill:
            return f"{self.airline_code}-{self.master_bill}"
        return self.master_bill or ""

    @property
    def man_date(self) -> float | None:
        """Export date as the YYYYMMDD number the AAMS mawb record uses."""
        parts = self.export_date.split("-")
        if len(parts) != 3:
            return None
        day, month, year = parts
        try:
            return float(f"{int(year):04d}{int(month):02d}{int(day):02d}")
        except ValueError:
            return None


def read_manifest(contents: bytes) -> Manifest:
    """Parse NetCHB workbook bytes into shipments plus flight-level values."""
    try:
        book = xlrd.open_workbook(file_contents=contents)
    except Exception as exc:  # xlrd raises a grab-bag of exception types
        raise NetchbError(
            "Could not read the file as a legacy .xls workbook. "
            "NetCHB exports a .xls - if you have an .xlsx, re-save it as "
            "'Excel 97-2003 Workbook (*.xls)' first."
        ) from exc

    sheet = book.sheet_by_index(0)
    if sheet.nrows < 2:
        raise NetchbError("The workbook has no data rows.")

    headers = [norm(sheet.cell_value(0, c)) for c in range(sheet.ncols)]
    index = {name: i for i, name in enumerate(headers) if name}

    missing = [c for c in REQUIRED_COLUMNS if c not in index]
    if missing:
        raise NetchbError(
            "This does not look like a NetCHB manifest report - missing "
            f"column(s): {', '.join(missing)}."
        )

    shipments: list[Shipment] = []
    by_hawb: dict[str, Shipment] = {}
    master_bill = airline_code = export_date = ""

    for r in range(1, sheet.nrows):
        row = {name: norm(sheet.cell_value(r, c)) for name, c in index.items()}
        hawb = row["House AWB"]
        if not hawb:
            continue  # trailing blank / totals rows

        shipment = by_hawb.get(hawb)
        if shipment is None:
            shipment = Shipment(hawb=hawb, group=row["GroupIdentifier"] or "DDP")
            by_hawb[hawb] = shipment
            shipments.append(shipment)
        shipment.rows.append(row)

        master_bill = master_bill or row.get("Master Bill Number", "")
        airline_code = airline_code or row.get("Airline 3 digit code", "")
        export_date = export_date or row.get("Date of Export", "")

    if not shipments:
        raise NetchbError("No rows with a House AWB were found.")

    return Manifest(
        shipments=shipments,
        master_bill=master_bill,
        airline_code=airline_code,
        export_date=export_date,
    )
