"""Reader for the NetCHB manifest report (legacy .xls).

The report is a flat table with one row per HTS entry line. Shipments are
recovered by grouping consecutive rows on ``House AWB``.
"""

from __future__ import annotations

import datetime
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


def parse_export_date(value, ctype, datemode) -> datetime.date | None:
    """Read 'Date of Export' from whatever shape the export happens to use.

    NetCHB writes it as the text 'DD-MM-YYYY', but an export saved through
    Excel can carry a real date cell instead. Anything that does not parse
    cleanly as day-month-year returns None rather than a guess: reading
    '07-29-2026' as day 7 month 29 would silently produce a nonsense manifest
    date, which is worse than an empty one.
    """
    if ctype == xlrd.XL_CELL_DATE:
        try:
            year, month, day = xlrd.xldate_as_tuple(value, datemode)[:3]
            return datetime.date(year, month, day)
        except (ValueError, xlrd.XLDateError):
            return None

    text = norm(value)
    if not text:
        return None
    for separator in ("-", "/", "."):
        parts = text.split(separator)
        if len(parts) != 3:
            continue
        try:
            day, month, year = (int(p) for p in parts)
        except ValueError:
            return None
        if year < 100:
            year += 2000
        try:
            return datetime.date(year, month, day)
        except ValueError:
            # Day and month out of range - most likely a month-first export.
            return None
    return None


@dataclass
class Manifest:
    """A parsed NetCHB report."""

    shipments: list[Shipment]
    master_bill: str
    airline_code: str
    export_date: datetime.date | None
    export_date_raw: str = ""

    @property
    def mawb(self) -> str:
        if self.airline_code and self.master_bill:
            return f"{self.airline_code}-{self.master_bill}"
        return self.master_bill or ""

    @property
    def man_date(self) -> float | None:
        """Export date as the YYYYMMDD number the AAMS mawb record uses."""
        if self.export_date is None:
            return None
        return float(self.export_date.strftime("%Y%m%d"))


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
    master_bills: set[str] = set()
    airline_codes: set[str] = set()
    export_date = None
    export_date_raw = ""

    for r in range(1, sheet.nrows):
        row = {name: norm(sheet.cell_value(r, c)) for name, c in index.items()}
        hawb = row["House AWB"]
        if not hawb:
            continue  # trailing blank / totals rows

        shipment = by_hawb.get(hawb)
        if shipment is None:
            shipment = Shipment(hawb=hawb, group=row["GroupIdentifier"])
            by_hawb[hawb] = shipment
            shipments.append(shipment)
        shipment.rows.append(row)

        if row.get("Master Bill Number"):
            master_bills.add(row["Master Bill Number"])
        if row.get("Airline 3 digit code"):
            airline_codes.add(row["Airline 3 digit code"])

        if export_date is None and "Date of Export" in index:
            c = index["Date of Export"]
            export_date_raw = export_date_raw or norm(sheet.cell_value(r, c))
            export_date = parse_export_date(
                sheet.cell_value(r, c), sheet.cell_type(r, c), book.datemode)

    if not shipments:
        raise NetchbError("No rows with a House AWB were found.")

    # An AAMS file describes one master air waybill. Silently filing a second
    # flight's shipments under the first one's MAWB would be far worse than
    # refusing the file.
    if len(master_bills) > 1:
        listed = ", ".join(sorted(master_bills))
        raise NetchbError(
            f"This report covers {len(master_bills)} master bills ({listed}). "
            "An AAMS file describes a single flight, so please export one "
            "master bill at a time."
        )
    if len(airline_codes) > 1:
        listed = ", ".join(sorted(airline_codes))
        raise NetchbError(
            f"This report mixes airline codes ({listed}). Please export one "
            "flight at a time."
        )

    return Manifest(
        shipments=shipments,
        master_bill=next(iter(master_bills), ""),
        airline_code=next(iter(airline_codes), ""),
        export_date=export_date,
        export_date_raw=export_date_raw,
    )
