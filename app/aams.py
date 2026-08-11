"""Builds the AAMS/ECCF record stream and writes it as a legacy .xls.

The output is not a table. Each row is::

    [line_seq, shipment_index, group, ...record payload]

and every record type is preceded by its own ``# record_type`` header row.
Row building is kept separate from workbook writing so the structure can be
asserted on in tests without touching the filesystem.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass

import xlwt

from . import mapping as M
from .netchb import Manifest, Shipment, as_number, norm

# A cell is str (text), float (number) or None (blank).
Cell = str | float | None
Row = list[Cell]

_PRIOR_NOTICE = re.compile(M.PRIOR_NOTICE_PATTERN)
_WHITESPACE = re.compile(r"\s+")


@dataclass
class FlightInfo:
    """Flight-level values the NetCHB export does not carry.

    All four are optional. Anything left blank is written as an empty cell and
    reported in the manual-fill list rather than being guessed at.
    """

    man_code: str = ""
    flight_info: str = ""
    man_origin: str = ""
    man_dest: str = ""


@dataclass
class Warning:
    hawb: str
    issue: str
    detail: str


# Plain names for the AAMS fields staff have to fill in. The raw field names
# are meaningless to them.
FIELD_LABELS = {
    "man_code": "Manifest code",
    "flight_info": "Flight",
    "man_origin": "Origin airport",
    "man_dest": "Destination airport",
    "origin": "Origin airport",
    "final_destination": "Destination airport",
    "weight": "Weight",
    "declared_value": "Declared value",
    "hts_code": "HTS code",
}

RECORD_LABELS = {
    "mawb": "the top flight row",
    "hawb": "the shipment rows",
    "commodity": "the item rows",
}


@dataclass
class ManualFill:
    """An AAMS field the converter leaves blank for a human to complete."""

    record: str
    field: str
    blank: int
    total: int
    reason: str

    @property
    def column(self) -> str:
        return M.column_letter(self.record, self.field)

    @property
    def label(self) -> str:
        return FIELD_LABELS.get(self.field, self.field.replace("_", " "))

    @property
    def where(self) -> str:
        return RECORD_LABELS.get(self.record, self.record)


# Plain-English wording for each warning, and whether the operator must act.
# The people using this are not customs-systems people; "line HTS padded"
# means nothing to them.
WARNING_WORDING = {
    "line HTS missing": ("check", "HTS code missing - you need to add it"),
    "shipment HTS missing": ("check", "HTS code missing - you need to add it"),
    "line HTS leading zero restored": (
        "check", "HTS code was missing its leading zero - we put it back"),
    "shipment HTS leading zero restored": (
        "check", "HTS code was missing its leading zero - we put it back"),
    "shipper address overflow": (
        "check", "Shipper address too long - the end was cut off"),
    "line HTS padded": (
        "auto", "Short HTS code padded with zeros to 10 digits"),
    "shipment HTS padded": (
        "auto", "Short HTS code padded with zeros to 10 digits"),
    "commodity description truncated": (
        "auto", "Long item description shortened to fit"),
    "description truncated": (
        "auto", "Long shipment description shortened to fit"),
    "consignee street 1 truncated": (
        "auto", "Long consignee address shortened to fit"),
    "consignee street 2 truncated": (
        "auto", "Long consignee address shortened to fit"),
    "shipper name truncated": (
        "auto", "Long shipper name shortened to fit"),
}


@dataclass
class Adjustment:
    """A group of warnings, worded for someone who does not know the format."""

    text: str
    shipments: int
    values: int
    needs_check: bool


@dataclass
class Conversion:
    rows: list[Row]
    warnings: list[Warning]
    manual_fill: list[ManualFill]
    shipment_count: int
    commodity_count: int

    @property
    def adjustments(self) -> list[Adjustment]:
        """Warnings collapsed into a handful of plain-English lines."""
        grouped: dict[str, tuple[set, int, bool]] = {}
        for w in self.warnings:
            kind, text = WARNING_WORDING.get(w.issue, ("auto", w.issue))
            hawbs, count, _ = grouped.get(text, (set(), 0, False))
            hawbs.add(w.hawb)
            grouped[text] = (hawbs, count + 1, kind == "check")
        out = [Adjustment(text, len(h), n, check)
               for text, (h, n, check) in grouped.items()]
        # Things needing a human first, then by how widespread they are.
        out.sort(key=lambda a: (not a.needs_check, -a.values))
        return out


def _truncate(value: str, width: int, hawb: str, label: str,
              warnings: list[Warning]) -> str:
    if len(value) <= width:
        return value
    warnings.append(Warning(hawb, f"{label} truncated",
                            f"{len(value)} chars cut to {width}: {value!r}"))
    # Cutting mid-word can leave a trailing space; don't ship that.
    return value[:width].rstrip()


def _split_address(street: str, hawb: str,
                   warnings: list[Warning]) -> tuple[str, str]:
    """Fit a single NetCHB street string into two 25-character AAMS lines.

    NetCHB stores the street as one flat string; AAMS wants two fixed 25-char
    lines. The sample shows the carrier's origin system held these as separate
    fields all along, so no exact re-split rule is recoverable - neither word
    wrapping nor hard slicing reproduces the sample better than chance.

    Until Bombino confirms the real rule, split on word boundaries only, so
    every line the file carries is at least readable. Whatever does not fit in
    two lines is dropped and reported rather than cut mid-word.
    """
    words = street.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if len(candidate) <= M.WIDTH_ADDRESS_LINE:
            current = candidate
            continue
        if current:
            lines.append(current)
        while len(word) > M.WIDTH_ADDRESS_LINE:
            lines.append(word[: M.WIDTH_ADDRESS_LINE])
            word = word[M.WIDTH_ADDRESS_LINE:]
        current = word
    if current:
        lines.append(current)

    if len(lines) > 2:
        warnings.append(Warning(
            hawb, "shipper address overflow",
            f"dropped {' '.join(lines[2:])!r} - the AAMS format has only two "
            f"{M.WIDTH_ADDRESS_LINE}-character address lines"))
    return (lines[0] if lines else "",
            lines[1] if len(lines) > 1 else "")


def _prior_notice(street: str) -> tuple[str, Cell]:
    """Pull the FDA prior-notice number out of the shipper address.

    Returns the address with the Z-token removed and the number as a numeric
    cell (or ``None`` when the address carries no prior notice).
    """
    match = _PRIOR_NOTICE.search(street)
    if not match:
        return _WHITESPACE.sub(" ", street).strip(), None
    cleaned = _WHITESPACE.sub(" ", street.replace(match.group(0), " ")).strip()
    return cleaned, float(match.group(1))


def _hts(raw: str, hawb: str, label: str, warnings: list[Warning]) -> Cell:
    """Normalise an HTS code to exactly 10 digits.

    NOT CURRENTLY CALLED. Both HTS columns are written blank, because NetCHB
    files 8-digit Indian export codes and AAMS wants 10-digit US codes -
    padding is not that conversion. Kept, with its unit test, because the
    leading-zero handling below is verified and will be needed the moment
    Bombino confirms how the two code sets map to each other.

    Written as text when the code has a leading zero and as a number
    otherwise, matching the reference file exactly.
    """
    digits = re.sub(r"\D", "", raw)
    if not digits:
        warnings.append(Warning(hawb, f"{label} missing",
                                "NetCHB left the HTS blank; the AAMS field is "
                                "empty and must be filled before filing"))
        return None

    # NetCHB stores the HTS column as a number, so a leading zero is lost on
    # the way in: '09024020' (tea) arrives as 9024020. HTS codes only come in
    # 4, 6, 8 or 10 digits, so an odd 7 or 9 digits means exactly that. Without
    # this, tea lands in chapter 90 (optical instruments) and dal in chapter 71
    # (jewellery).
    if len(digits) in (M.HTS_LENGTH - 1, M.HTS_LENGTH - 3):
        warnings.append(Warning(hawb, f"{label} leading zero restored",
                                f"{digits} has an odd digit count; read as "
                                f"0{digits} (Excel dropped the leading zero)"))
        digits = "0" + digits

    if len(digits) < M.HTS_LENGTH:
        warnings.append(Warning(hawb, f"{label} padded",
                                f"{digits} is {len(digits)} digits, "
                                f"right-padded to {M.HTS_LENGTH}"))
        digits = digits.ljust(M.HTS_LENGTH, "0")
    elif len(digits) > M.HTS_LENGTH:
        warnings.append(Warning(hawb, f"{label} truncated",
                                f"{digits} cut to {M.HTS_LENGTH} digits"))
        digits = digits[: M.HTS_LENGTH]
    return digits if digits.startswith("0") else float(digits)


def _postal_code(raw: str) -> str:
    """US zips stay text so leading zeros survive (e.g. 08234)."""
    if raw.isdigit() and len(raw) < 5:
        return raw.rjust(5, "0")
    return raw


def _record(index: int, group: str, fields, values: dict) -> Row:
    row: Row = [None] * M.TOTAL_COLS
    row[1] = float(index)
    row[2] = group
    for offset, name in enumerate(fields):
        row[M.PAYLOAD_START_COL + offset] = values.get(name)
    return row


def _header_row(index: int, group: str, fields) -> Row:
    row: Row = [None] * M.TOTAL_COLS
    row[1] = float(index)
    row[2] = group
    for offset, name in enumerate(fields):
        row[M.PAYLOAD_START_COL + offset] = name
    return row


def _hawb_values(shipment: Shipment, flight: FlightInfo,
                 warnings: list[Warning]) -> dict:
    first = shipment.first
    hawb = shipment.hawb

    street, prior_notice = _prior_notice(first.get("ManufacturerStreetAddress", ""))
    address1, address2 = _split_address(street, hawb, warnings)

    shipper_zip = as_number(first.get("ManufacturerPostalCode", ""))
    if shipper_zip is None:
        shipper_zip = first.get("ManufacturerPostalCode", "") or None

    return {
        "# record_type": "hawb",
        "record_version": float(M.HAWB_VERSION),
        "hawb": as_number(hawb) if hawb.isdigit() else hawb,
        "origin": flight.man_origin or None,
        "final_destination": flight.man_dest or None,
        "num_pieces": as_number(first.get("Manifest Qty Piece count", "")),
        # No NetCHB column carries weight; the unit is 'KG' on every sample record.
        "weight": None,
        "weight_unit": M.WEIGHT_UNIT,
        "contents": M.CONTENTS_CODE,
        "currency_code": M.CURRENCY,
        # Deliberately blank. The NetCHB InvoiceValue is not the AAMS declared
        # value - do not derive this field from it.
        "declared_value": None,
        "description": _truncate(first.get("DescOfMerchandise", ""),
                                 M.WIDTH_HAWB_DESCRIPTION, hawb,
                                 "description", warnings),
        # Blank for the same reason as the item rows: NetCHB's codes are not
        # the codes AAMS wants, and a shipment can carry several of them.
        "hts_code": None,
        "fda_prior_notice": prior_notice,
        "packaging": M.PACKAGING_CODE,
        "shipper_name": _truncate(first.get("ManufacturerName", ""),
                                  M.WIDTH_SHIPPER_NAME, hawb,
                                  "shipper name", warnings),
        "shipper_address1": address1,
        "shipper_address2": address2 or None,
        "shipper_city": first.get("ManufacturerCity", ""),
        "shipper_zip": shipper_zip,
        "shipper_country": first.get("ManufacturerCountry", "") or None,
        "consignee_person": _truncate(first.get("ConsigneeName", ""),
                                      M.WIDTH_CONSIGNEE_PERSON, hawb,
                                      "consignee name", warnings),
        "consignee_street_1": _truncate(first.get("ConsigneeAddress1", ""),
                                        M.WIDTH_ADDRESS_LINE, hawb,
                                        "consignee street 1", warnings),
        "consignee_street_2": _truncate(first.get("ConsigneeAddress2", ""),
                                        M.WIDTH_ADDRESS_LINE, hawb,
                                        "consignee street 2", warnings) or None,
        "consignee_city": first.get("ConsigneeCity", ""),
        "consignee_state": first.get("ConsigneeState", "") or None,
        "consignee_postal_code": _postal_code(first.get("ConsigneeZip", "")),
        "consignee_country": M.CONSIGNEE_COUNTRY,
        "goods_coo": first.get("Country Of Origin", "") or None,
    }


def _commodity_values(row: dict, hawb: str, warnings: list[Warning]) -> dict:
    description = row.get("Description") or row.get("HTS-1", "")
    return {
        "# record_type": "commodity",
        "record_version": float(M.COMMODITY_VERSION),
        "quantity": as_number(row.get("HTSQty", "")),
        "description": _truncate(description, M.WIDTH_COMMODITY_DESCRIPTION,
                                 hawb, "commodity description", warnings),
        # Deliberately blank. NetCHB files 8-digit Indian export codes and AAMS
        # wants 10-digit US codes; padding with zeros is not that conversion.
        # Left for manual classification until Bombino confirms a rule.
        "hts_code": None,
        "origin_country": row.get("Country Of Origin", "") or None,
        "declared_value": as_number(row.get("HTSValue", "")),
        "declared_value_currency": M.CURRENCY,
    }


def build_rows(manifest: Manifest,
               flight: FlightInfo | None = None) -> Conversion:
    """Turn a parsed NetCHB manifest into the AAMS row stream."""
    flight = flight or FlightInfo()
    rows: list[Row] = []
    warnings: list[Warning] = []

    for address in M.NOTIFY_EMAILS:
        rows.append(_record(0, M.HEADER_GROUP,
                            ("record", "version", "scope", "address"),
                            {"record": "email",
                             "version": float(M.EMAIL_VERSION),
                             "scope": "all",
                             "address": address}))

    rows.append(_header_row(0, M.HEADER_GROUP, M.MAWB_FIELDS))
    rows.append(_record(0, M.HEADER_GROUP, M.MAWB_FIELDS, {
        "#  record_type": "mawb",
        "record_version": float(M.MAWB_VERSION),
        # All four are optional: blank unless the operator typed something.
        "man_code": flight.man_code or None,
        "man_date": manifest.man_date,
        "man_origin": flight.man_origin or None,
        "man_dest": flight.man_dest or None,
        "flight_info": flight.flight_info or None,
        "mawb": manifest.mawb,
    }))

    commodity_count = 0
    hawb_hts_blank = commodity_hts_blank = 0
    for index, shipment in enumerate(manifest.shipments, start=1):
        group = shipment.group
        values = _hawb_values(shipment, flight, warnings)
        if values["hts_code"] is None:
            hawb_hts_blank += 1
        rows.append(_header_row(index, group, M.HAWB_FIELDS))
        rows.append(_record(index, group, M.HAWB_FIELDS, values))

        if not shipment.has_commodities:
            continue

        rows.append(_header_row(index, group, M.COMMODITY_FIELDS))
        for line in shipment.rows:
            line_values = _commodity_values(line, shipment.hawb, warnings)
            if line_values["hts_code"] is None:
                commodity_hts_blank += 1
            rows.append(_record(index, group, M.COMMODITY_FIELDS, line_values))
            commodity_count += 1

    for seq, row in enumerate(rows, start=1):
        row[0] = float(seq)

    shipments = len(manifest.shipments)
    not_entered = "not in the NetCHB export and not entered on the form"
    no_group = sum(1 for s in manifest.shipments if not s.group)
    manual_fill = [
        ManualFill("mawb", "man_date", 0 if manifest.man_date else 1, 1,
                   f"the Date of Export ({manifest.export_date_raw!r}) could "
                   "not be read as a day-month-year date"
                   if manifest.export_date_raw
                   else "the NetCHB report has no Date of Export"),
        ManualFill("mawb", "man_code", 0 if flight.man_code else 1, 1,
                   not_entered),
        ManualFill("mawb", "flight_info", 0 if flight.flight_info else 1, 1,
                   not_entered),
        ManualFill("mawb", "man_origin", 0 if flight.man_origin else 1, 1,
                   not_entered),
        ManualFill("mawb", "man_dest", 0 if flight.man_dest else 1, 1,
                   not_entered),
        ManualFill("hawb", "origin",
                   0 if flight.man_origin else shipments, shipments,
                   "mirrors man_origin, which was left blank"),
        ManualFill("hawb", "final_destination",
                   0 if flight.man_dest else shipments, shipments,
                   "mirrors man_dest, which was left blank"),
        ManualFill("hawb", "weight", shipments, shipments,
                   "the NetCHB report has no weight column at all"),
        ManualFill("hawb", "declared_value", shipments, shipments,
                   "no NetCHB column holds the AAMS declared value - "
                   "InvoiceValue is not it"),
        ManualFill("hawb", "hts_code", hawb_hts_blank, shipments,
                   "NetCHB's 8-digit codes are not the 10-digit codes AAMS "
                   "wants, so nothing is filled in automatically"),
        ManualFill("commodity", "hts_code", commodity_hts_blank, commodity_count,
                   "NetCHB's 8-digit codes are not the 10-digit codes AAMS "
                   "wants, so nothing is filled in automatically"),
        ManualFill("hawb", "service_type", no_group, shipments,
                   "NetCHB left GroupIdentifier blank, so DDP/DDU is unknown"),
    ]

    return Conversion(rows=rows, warnings=warnings,
                      manual_fill=[m for m in manual_fill if m.blank],
                      shipment_count=shipments,
                      commodity_count=commodity_count)


def write_xls(rows: list[Row]) -> bytes:
    """Render the row stream as a legacy BIFF .xls workbook."""
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet(M.SHEET_NAME)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if value is None or value == "":
                continue
            sheet.write(r, c, value)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def write_checklist(result: Conversion) -> bytes:
    """A one-sheet .xls listing everything needing a human, for use in Excel.

    Staff can sort and filter this far more easily than a web page, and tick
    items off as they work through the converted file.
    """
    rows: list[Row] = [["House AWB", "What to do", "Details", "Done?"]]

    for item in result.manual_fill:
        rows.append([
            "(all shipments)" if item.blank > 1 else "(top flight row)",
            f"Fill in {item.label} - column {item.column}",
            f"{item.blank} of {item.total} "
            f"{'row is' if item.total == 1 else 'rows are'} blank. "
            f"{item.reason[0].upper()}{item.reason[1:]}.",
            "",
        ])

    for w in result.warnings:
        kind, text = WARNING_WORDING.get(w.issue, ("auto", w.issue))
        if kind != "check":
            continue
        rows.append([w.hawb, text, w.detail, ""])

    for w in result.warnings:
        kind, text = WARNING_WORDING.get(w.issue, ("auto", w.issue))
        if kind == "check":
            continue
        rows.append([w.hawb, f"FYI - {text}", w.detail, ""])

    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet("checklist")
    bold = xlwt.easyxf("font: bold on")
    for c, width in enumerate((5000, 12000, 18000, 2500)):
        sheet.col(c).width = width
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if value == "":
                continue
            sheet.write(r, c, value, bold if r == 0 else xlwt.Style.default_style)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def convert(manifest: Manifest,
            flight: FlightInfo | None = None) -> tuple[bytes, Conversion]:
    result = build_rows(manifest, flight)
    return write_xls(result.rows), result
