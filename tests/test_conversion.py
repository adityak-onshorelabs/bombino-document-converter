"""Structural and differential tests for the NetCHB -> AAMS converter.

The differential test compares against the real sample pair in ``sample_docs``.
It only asserts on fields that are genuinely derivable from the NetCHB export.
Commodity descriptions, HTS codes, line values and weight legitimately differ
between the two samples - the broker reclassified them downstream - so those
are deliberately not asserted.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
import xlrd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import mapping as M  # noqa: E402
from app.aams import FlightInfo, build_rows, write_xls  # noqa: E402
from app.netchb import NetchbError, read_manifest  # noqa: E402

SAMPLES = ROOT / "sample_docs"
NETCHB = SAMPLES / "manifest_netchb_report.xls"
AAMS = SAMPLES / "aams-data-bombino-express-eccf TY523 & TT523V MODI.xls"

# The flight fields are optional, so the default run supplies none of them.
FILLED = FlightInfo(man_code="TY523", flight_info="BA 0117",
                    man_origin="BOM", man_dest="JFK")


@pytest.fixture(scope="module")
def manifest():
    return read_manifest(NETCHB.read_bytes())


@pytest.fixture(scope="module")
def conversion(manifest):
    """The default run: no flight fields entered."""
    return build_rows(manifest)


@pytest.fixture(scope="module")
def output_sheet(conversion):
    book = xlrd.open_workbook(file_contents=write_xls(conversion.rows))
    return book.sheet_by_index(0)


@pytest.fixture(scope="module")
def filled_sheet(manifest):
    """A run with all four optional flight fields supplied."""
    book = xlrd.open_workbook(
        file_contents=write_xls(build_rows(manifest, FILLED).rows))
    return book.sheet_by_index(0)


def _payload(row, offset):
    return row[M.PAYLOAD_START_COL + offset]


def _records(sheet):
    """Walk a rendered AAMS sheet into (group, kind, {field: value}) tuples."""
    headers, out = {}, []
    for r in range(sheet.nrows):
        kind = sheet.cell_value(r, M.PAYLOAD_START_COL)
        if isinstance(kind, str) and kind.strip().startswith("#"):
            following = sheet.cell_value(r + 1, M.PAYLOAD_START_COL)
            headers[following] = [
                sheet.cell_value(r, c)
                for c in range(M.PAYLOAD_START_COL, sheet.ncols)
            ]
            continue
        fields = headers.get(kind)
        if not fields:
            continue  # email rows carry no header
        values = {}
        for i, name in enumerate(fields):
            if not name:
                continue
            value = sheet.cell_value(r, M.PAYLOAD_START_COL + i)
            if isinstance(value, float) and value == int(value):
                value = str(int(value))
            values[name] = str(value).strip()
        out.append((sheet.cell_value(r, 2), kind, values))
    return out


# --- reader -----------------------------------------------------------------

def test_reader_groups_by_house_awb(manifest):
    assert len(manifest.shipments) == 187
    assert sum(len(s.rows) for s in manifest.shipments) == 723
    assert manifest.mawb == "125-10391485"
    assert manifest.man_date == 20260729.0


def test_reader_rejects_junk():
    with pytest.raises(NetchbError):
        read_manifest(b"not a workbook")


def _rebuild(mutate, style_date=None, limit=40):
    """Rewrite the sample NetCHB report, applying mutate(row_dict, row_index).

    Used to feed the reader the kinds of file a user could plausibly upload:
    a different flight, a differently formatted date, two manifests at once.
    """
    import io

    import xlwt

    src = xlrd.open_workbook(NETCHB).sheet_by_index(0)
    headers = [src.cell_value(0, c) for c in range(src.ncols)]
    index = {str(h).strip(): i for i, h in enumerate(headers) if str(h).strip()}

    book = xlwt.Workbook()
    sheet = book.add_sheet("Worksheet")
    for c, h in enumerate(headers):
        sheet.write(0, c, h)

    out = 1
    for r in range(1, min(limit, src.nrows)):
        row = {h: src.cell_value(r, i) for h, i in index.items()}
        row = mutate(row, r)
        if row is None:
            continue
        for h, i in index.items():
            value = row[h]
            if value == "" or value is None:
                continue
            if style_date and h == "Date of Export":
                sheet.write(out, i, value, style_date)
            else:
                sheet.write(out, i, value)
        out += 1

    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def test_reader_refuses_a_file_covering_two_flights():
    """Two master bills in one export must not be merged under the first.

    Filing a second flight's shipments under the wrong MAWB would be far worse
    than refusing the upload.
    """
    def two_flights(row, r):
        if r > 20:
            row["Master Bill Number"] = 99999999.0
            row["House AWB"] = str(int(row["House AWB"])) + "X"
        return row

    with pytest.raises(NetchbError, match="master bills"):
        read_manifest(_rebuild(two_flights))


def test_reader_refuses_mixed_airline_codes():
    def two_airlines(row, r):
        if r > 20:
            row["Airline 3 digit code"] = 176.0
        return row

    with pytest.raises(NetchbError, match="airline codes"):
        read_manifest(_rebuild(two_airlines))


def test_reader_accepts_a_different_flight():
    """Nothing may be hardcoded to the one sample flight."""
    def other_flight(row, r):
        row["Airline 3 digit code"] = 176.0
        row["Master Bill Number"] = 12345678.0
        return row

    manifest = read_manifest(_rebuild(other_flight))
    assert manifest.mawb == "176-12345678"


def test_reader_handles_a_real_excel_date_cell():
    """An export saved through Excel can carry a date cell, not text."""
    import datetime

    import xlwt

    style = xlwt.easyxf(num_format_str="DD-MM-YYYY")
    manifest = read_manifest(_rebuild(
        lambda row, r: {**row, "Date of Export": datetime.datetime(2026, 7, 29)},
        style_date=style))
    assert manifest.man_date == 20260729.0


def test_reader_refuses_to_guess_at_a_month_first_date():
    """'07-29-2026' must not be read as day 7 of month 29."""
    manifest = read_manifest(_rebuild(
        lambda row, r: {**row, "Date of Export": "07-29-2026"}))
    assert manifest.man_date is None
    assert any(m.field == "man_date" for m in build_rows(manifest).manual_fill)


def test_reader_rejects_wrong_workbook():
    """A valid .xls that is not a NetCHB report is refused by name, not crash."""
    with pytest.raises(NetchbError, match="missing column"):
        read_manifest(AAMS.read_bytes())


# --- structure --------------------------------------------------------------

def test_sequence_column_is_contiguous(output_sheet):
    seq = [output_sheet.cell_value(r, 0) for r in range(output_sheet.nrows)]
    assert seq == [float(i) for i in range(1, output_sheet.nrows + 1)]


def test_shipment_index_and_group_columns(output_sheet):
    indexes = [output_sheet.cell_value(r, 1) for r in range(output_sheet.nrows)]
    assert indexes[:4] == [0.0, 0.0, 0.0, 0.0]  # 2 emails + mawb header + mawb
    assert all(b >= a for a, b in zip(indexes, indexes[1:]))
    assert indexes[-1] == 187.0

    groups = {output_sheet.cell_value(r, 2) for r in range(output_sheet.nrows)}
    assert groups == {"DUTY", "DDP", "DDU"}


def test_every_record_block_opens_with_its_header(output_sheet):
    """A record is preceded either by its '#' header or by its own kind.

    Commodity blocks emit one header then N commodity rows, so only the first
    row of a block sits directly under a header.
    """
    previous = None
    for r in range(output_sheet.nrows):
        kind = output_sheet.cell_value(r, M.PAYLOAD_START_COL)
        if kind in ("mawb", "hawb", "commodity"):
            opened = (isinstance(previous, str)
                      and previous.strip().startswith("#"))
            assert opened or previous == kind, (
                f"row {r}: {kind} record follows {previous!r}, "
                "not its header or another record of the same kind")
        previous = kind


def test_header_rows_match_the_sample_format(output_sheet):
    """Field lists must match byte-for-byte, misspellings included."""
    sample = xlrd.open_workbook(AAMS).sheet_by_index(0)

    def header_for(sheet, kind):
        for r in range(sheet.nrows):
            cell = sheet.cell_value(r, M.PAYLOAD_START_COL)
            if (isinstance(cell, str) and cell.strip().startswith("#")
                    and sheet.cell_value(r + 1, M.PAYLOAD_START_COL) == kind):
                return [sheet.cell_value(r, c)
                        for c in range(M.PAYLOAD_START_COL, sheet.ncols)]
        raise AssertionError(f"no header for {kind}")

    for kind in ("mawb", "hawb", "commodity"):
        assert header_for(output_sheet, kind) == header_for(sample, kind)


def test_header_block(output_sheet):
    assert _payload(output_sheet.row_values(0), 0) == "email"
    assert _payload(output_sheet.row_values(0), 3) == M.NOTIFY_EMAILS[0]
    assert _payload(output_sheet.row_values(1), 3) == M.NOTIFY_EMAILS[1]

    mawb = output_sheet.row_values(3)
    assert _payload(mawb, 0) == "mawb"
    assert _payload(mawb, 3) == 20260729.0  # man_date, from the NetCHB data
    assert _payload(mawb, 7) == "125-10391485"


def test_flight_fields_are_optional_and_blank_by_default(output_sheet):
    """Nothing entered on the form means empty cells, not invented values."""
    fields = list(M.MAWB_FIELDS)
    for name in M.OPERATOR_FIELDS:
        col = M.PAYLOAD_START_COL + fields.index(name)
        assert output_sheet.cell_type(3, col) == xlrd.XL_CELL_EMPTY, name

    # The route also mirrors onto every hawb record.
    hawb = list(M.HAWB_FIELDS)
    for name in ("origin", "final_destination"):
        col = M.PAYLOAD_START_COL + hawb.index(name)
        for r in range(output_sheet.nrows):
            if output_sheet.cell_value(r, M.PAYLOAD_START_COL) == "hawb":
                assert output_sheet.cell_type(r, col) == xlrd.XL_CELL_EMPTY


def test_flight_fields_are_written_through_when_supplied(filled_sheet):
    mawb = filled_sheet.row_values(3)
    assert _payload(mawb, 2) == "TY523"
    assert _payload(mawb, 4) == "BOM"
    assert _payload(mawb, 5) == "JFK"
    assert _payload(mawb, 6) == "BA 0117"

    hawb = list(M.HAWB_FIELDS)
    origin = M.PAYLOAD_START_COL + hawb.index("origin")
    dest = M.PAYLOAD_START_COL + hawb.index("final_destination")
    seen = 0
    for r in range(filled_sheet.nrows):
        if filled_sheet.cell_value(r, M.PAYLOAD_START_COL) == "hawb":
            seen += 1
            assert filled_sheet.cell_value(r, origin) == "BOM"
            assert filled_sheet.cell_value(r, dest) == "JFK"
    assert seen == 187


def test_commodity_block_only_for_multi_line_shipments(manifest, conversion):
    single = {s.hawb for s in manifest.shipments if len(s.rows) == 1}
    seen: dict[float, int] = {}
    for row in conversion.rows:
        if row[M.PAYLOAD_START_COL] == "commodity":
            seen[row[1]] = seen.get(row[1], 0) + 1
    multi = [s for s in manifest.shipments if len(s.rows) > 1]
    assert len(seen) == len(multi)
    assert conversion.commodity_count == sum(len(s.rows) for s in multi)
    assert len(single) == 187 - len(multi)


def test_cell_types_match_the_sample_convention(output_sheet):
    """hawb/zip/values numeric, postal code text, HTS text iff leading zero."""
    fields = list(M.HAWB_FIELDS)
    hawb_col = M.PAYLOAD_START_COL + fields.index("hawb")
    postal_col = M.PAYLOAD_START_COL + fields.index("consignee_postal_code")

    checked = 0
    for r in range(output_sheet.nrows):
        if output_sheet.cell_value(r, M.PAYLOAD_START_COL) != "hawb":
            continue
        checked += 1
        assert output_sheet.cell_type(r, hawb_col) == xlrd.XL_CELL_NUMBER
        assert output_sheet.cell_type(r, postal_col) == xlrd.XL_CELL_TEXT
    assert checked == 187

    com_hts = M.PAYLOAD_START_COL + list(M.COMMODITY_FIELDS).index("hts_code")
    for r in range(output_sheet.nrows):
        if output_sheet.cell_value(r, M.PAYLOAD_START_COL) != "commodity":
            continue
        cell = output_sheet.cell(r, com_hts)
        if cell.ctype == xlrd.XL_CELL_TEXT:
            assert cell.value.startswith("0")
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            assert not str(int(cell.value)).startswith("0")


def test_hts_codes_are_ten_digits(output_sheet):
    for kind_fields, kind in ((M.HAWB_FIELDS, "hawb"),
                              (M.COMMODITY_FIELDS, "commodity")):
        col = M.PAYLOAD_START_COL + list(kind_fields).index("hts_code")
        for r in range(output_sheet.nrows):
            if output_sheet.cell_value(r, M.PAYLOAD_START_COL) != kind:
                continue
            cell = output_sheet.cell(r, col)
            if cell.ctype == xlrd.XL_CELL_EMPTY:
                continue  # NetCHB left it blank; reported as a warning
            text = (cell.value if cell.ctype == xlrd.XL_CELL_TEXT
                    else str(int(cell.value)))
            assert len(text) == 10, f"row {r}: {text!r}"


def test_declared_value_is_never_derived(output_sheet):
    """hawb declared_value stays empty - InvoiceValue is not the AAMS value."""
    col = M.PAYLOAD_START_COL + list(M.HAWB_FIELDS).index("declared_value")
    for r in range(output_sheet.nrows):
        if output_sheet.cell_value(r, M.PAYLOAD_START_COL) == "hawb":
            assert output_sheet.cell_type(r, col) == xlrd.XL_CELL_EMPTY


def test_weight_is_blank_but_carries_its_unit(output_sheet):
    """No NetCHB column holds weight; the unit is 'KG' on every sample record."""
    col = M.PAYLOAD_START_COL + list(M.HAWB_FIELDS).index("weight")
    unit = M.PAYLOAD_START_COL + list(M.HAWB_FIELDS).index("weight_unit")
    for r in range(output_sheet.nrows):
        if output_sheet.cell_value(r, M.PAYLOAD_START_COL) == "hawb":
            assert output_sheet.cell_type(r, col) == xlrd.XL_CELL_EMPTY
            assert output_sheet.cell_value(r, unit) == M.WEIGHT_UNIT


def test_manual_fill_report_covers_every_blank_field(conversion, output_sheet):
    """Anything the operator must complete is reported, with real counts."""
    reported = {(m.record, m.field): m for m in conversion.manual_fill}
    assert set(reported) == {
        ("mawb", "man_code"), ("mawb", "flight_info"),
        ("mawb", "man_origin"), ("mawb", "man_dest"),
        ("hawb", "origin"), ("hawb", "final_destination"),
        ("hawb", "weight"), ("hawb", "declared_value"),
        ("hawb", "hts_code"), ("commodity", "hts_code"),
    }
    assert reported[("hawb", "weight")].blank == conversion.shipment_count
    assert reported[("hawb", "declared_value")].blank == conversion.shipment_count

    # The counts must match what actually landed in the workbook.
    for record, field, fields in (("hawb", "hts_code", M.HAWB_FIELDS),
                                  ("commodity", "hts_code", M.COMMODITY_FIELDS)):
        col = M.PAYLOAD_START_COL + list(fields).index(field)
        blank = sum(
            1 for r in range(output_sheet.nrows)
            if output_sheet.cell_value(r, M.PAYLOAD_START_COL) == record
            and output_sheet.cell_type(r, col) == xlrd.XL_CELL_EMPTY)
        assert reported[(record, field)].blank == blank


def test_manual_fill_drops_fields_the_operator_supplied(manifest):
    """Entering the flight details removes them from the manual-fill list."""
    reported = {(m.record, m.field)
                for m in build_rows(manifest, FILLED).manual_fill}
    assert set(M.OPERATOR_FIELDS).isdisjoint({f for _, f in reported})
    assert ("hawb", "origin") not in reported
    assert ("hawb", "final_destination") not in reported
    # The ones with no possible source stay.
    assert ("hawb", "weight") in reported
    assert ("commodity", "hts_code") in reported


def test_item_row_hts_is_always_blank(output_sheet, conversion):
    """Column I is left empty pending Bombino's answer on code conversion.

    NetCHB files 8-digit Indian export codes; AAMS wants 10-digit US codes.
    Padding with zeros is not that conversion, so nothing is written.
    """
    col = M.PAYLOAD_START_COL + list(M.COMMODITY_FIELDS).index("hts_code")
    rows = 0
    for r in range(output_sheet.nrows):
        if output_sheet.cell_value(r, M.PAYLOAD_START_COL) != "commodity":
            continue
        rows += 1
        assert output_sheet.cell_type(r, col) == xlrd.XL_CELL_EMPTY

    reported = {(m.record, m.field): m for m in conversion.manual_fill}
    assert reported[("commodity", "hts_code")].blank == rows


def test_hts_leading_zero_is_restored():
    """Excel drops the leading zero on chapter 01-09 codes; put it back.

    'TEA' is filed as 09024020 but reaches us as the number 9024020. Padding
    that on the right would land it in chapter 90, optical instruments. Still
    exercised because the shipment row (column Y) uses the same routine.
    """
    from app.aams import _hts

    def code(raw):
        cell = _hts(raw, "TEST", "shipment HTS", [])
        return cell if isinstance(cell, str) else str(int(cell))

    assert code("9024020") == "0902402000"    # tea
    assert code("7132020") == "0713202000"    # dal
    assert code("9109920") == "0910992000"    # spices
    # Genuine 4-, 6- and 8-digit codes must NOT be shifted.
    assert code("961519") == "9615190000"
    assert code("4420") == "4420000000"
    assert code("62114210") == "6211421000"
    assert _hts("", "TEST", "shipment HTS", []) is None


# --- differential against the real sample pair ------------------------------

@pytest.fixture(scope="module")
def paired(output_sheet):
    """hawb -> (ours, theirs) for the 185 shipments present in both files."""
    sample = xlrd.open_workbook(AAMS).sheet_by_index(0)
    ours = {v["hawb"]: (g, v) for g, k, v in _records(output_sheet) if k == "hawb"}
    theirs = {v["hawb"]: (g, v) for g, k, v in _records(sample) if k == "hawb"}
    common = ours.keys() & theirs.keys()
    assert len(common) == 185, "expected 185 shared House AWBs"
    return {h: (ours[h], theirs[h]) for h in common}


def test_group_matches_sample(paired):
    assert all(a[0] == b[0] for a, b in paired.values())


def test_pieces_match_sample(paired):
    mismatched = [h for h, (a, b) in paired.items()
                  if a[1]["num_pieces"] != b[1]["num_pieces"]]
    assert not mismatched


def test_prior_notice_matches_sample(paired):
    """The Z-number pulled out of the shipper address must match exactly."""
    ours = {h: a[1]["fda_prior_notice"] for h, (a, _) in paired.items()}
    theirs = {h: b[1]["fda_prior_notice"] for h, (_, b) in paired.items()}
    assert sum(1 for v in theirs.values() if v) == 50
    assert ours == theirs


def test_consignee_matches_sample(paired):
    for field in ("consignee_person", "consignee_city",
                  "consignee_state", "consignee_postal_code"):
        bad = [h for h, (a, b) in paired.items()
               if a[1][field].upper() != b[1][field].upper()]
        # The sample was hand-edited; allow a small number of divergences.
        assert len(bad) <= 8, f"{field}: {len(bad)} mismatches, e.g. {bad[:5]}"


def test_shipper_matches_sample(paired):
    for field in ("shipper_city", "shipper_zip"):
        bad = [h for h, (a, b) in paired.items()
               if a[1][field].upper() != b[1][field].upper()]
        assert len(bad) <= 8, f"{field}: {len(bad)} mismatches, e.g. {bad[:5]}"


def test_shipper_name_matches_sample_where_it_was_not_abbreviated(paired):
    """Long names diverge because a human abbreviated them, not truncated.

    The sample shows names like 'EXAMPLE TRADING PRIVATE LIMITED' rewritten as
    'Example Trading Pvt Ltd'. The converter cannot invent those
    abbreviations, so it truncates to the field width and raises a warning
    instead. Only names that fit the width unchanged are compared here.
    """
    bad = [h for h, (a, b) in paired.items()
           if len(a[1]["shipper_name"]) < M.WIDTH_SHIPPER_NAME
           and a[1]["shipper_name"].upper() != b[1]["shipper_name"].upper()]
    assert len(bad) <= 8, f"{len(bad)} mismatches, e.g. {bad[:5]}"


def test_no_cell_has_stray_whitespace(output_sheet):
    """Truncating mid-word must not leave a trailing space in the cell."""
    offenders = [
        (r, c, output_sheet.cell_value(r, c))
        for r in range(output_sheet.nrows)
        for c in range(output_sheet.ncols)
        if isinstance(output_sheet.cell_value(r, c), str)
        and output_sheet.cell_value(r, c) != output_sheet.cell_value(r, c).strip()
    ]
    assert not offenders, f"{len(offenders)} padded cells, e.g. {offenders[:3]}"


def test_shipper_address_lines_fit(output_sheet):
    fields = list(M.HAWB_FIELDS)
    for name in ("shipper_address1", "shipper_address2",
                 "consignee_street_1", "consignee_street_2"):
        col = M.PAYLOAD_START_COL + fields.index(name)
        for r in range(output_sheet.nrows):
            if output_sheet.cell_value(r, M.PAYLOAD_START_COL) != "hawb":
                continue
            value = output_sheet.cell_value(r, col)
            if isinstance(value, str):
                assert len(value) <= M.WIDTH_ADDRESS_LINE
