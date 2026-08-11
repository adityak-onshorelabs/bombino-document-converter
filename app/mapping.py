"""Constants describing the AAMS/ECCF file layout.

Every field list here was lifted verbatim from the ``# record_type`` header
rows of ``sample_docs/aams-data-bombino-express-eccf TY523 & TT523V MODI.xls``.
Do not "fix" the spelling or spacing - the receiving system matches on it.
"""

# Sheet name of the produced workbook (30 chars, under Excel's 31 limit).
SHEET_NAME = "aams-data-bombino-express-eccf"

# Notification addresses emitted in the header block.
NOTIFY_EMAILS = ("bombino@bombinoexp.com", "salim@bombinoexp.com")

# Group tag in column 2 for the header block.
HEADER_GROUP = "DUTY"

# The mawb header row uses two spaces after the '#'; hawb/commodity use one.
MAWB_FIELDS = (
    "#  record_type",
    "record_version",
    "man_code",
    "man_date",
    "man_origin",
    "man_dest",
    "flight_info",
    "mawb",
)

HAWB_FIELDS = (
    "# record_type",
    "record_version",
    "profile_key",
    "hawb",
    "ship_ref_num",
    "internal_reference",
    "vend_ref_num",
    "origin",
    "final_destination",
    "outlying",
    "service_provider",
    "dls_station",
    "dls_final_destination",
    "num_pieces",
    "weight",
    "weight_unit",
    "contents",
    "currency_code",
    "declared_value",
    "insurance_amount",
    "description",
    "hts_code",
    "fda_prior_notice",
    "terms",
    "packaging",
    "service_type",
    "collect_amount",
    "cust_key",
    "ship_acct_num",
    "dls_acct_num",
    "ext_cust_acct",
    "shipper_name",
    "shipper_address1",
    "shipper_address2",
    "shipper_city",
    "shipper_state",
    "shipper_zip",
    "shipper_country",
    "shipper_phone",
    "consignee_person",
    "consignee_company",
    "consignee_street_1",
    "consignee_street_2",
    "consignee_city",
    "consignee_state",
    "consignee_postal_code",
    "consignee_country",
    "consignee_phone",
    "consignee_email",
    "consignee_tax_id",
    "comment",
    "goods_coo",
    "incoming_container",
    "mid",
)

# 'manufactuer_id' is misspelled in the format itself.
COMMODITY_FIELDS = (
    "# record_type",
    "record_version",
    "quantity",
    "description",
    "commodity_url",
    "hts_code",
    "origin_country",
    "declared_value",
    "declared_value_currency",
    "weight_per_q",
    "weight_units",
    "order_id",
    "order_date",
    "list_price",
    "list_price_currency",
    "retail_price",
    "retail_price_currency",
    "manufactuer_id",
)

# Payload starts at column 3; the widest record (hawb) fixes the sheet width.
PAYLOAD_START_COL = 3
TOTAL_COLS = PAYLOAD_START_COL + len(HAWB_FIELDS)  # 57

RECORD_FIELDS = {
    "mawb": MAWB_FIELDS,
    "hawb": HAWB_FIELDS,
    "commodity": COMMODITY_FIELDS,
}


def column_letter(record: str, field: str) -> str:
    """Excel column letter for a field, e.g. hawb declared_value -> 'V'.

    Staff work in Excel, so telling them 'column V' beats naming the field.
    """
    index = PAYLOAD_START_COL + list(RECORD_FIELDS[record]).index(field)
    letters = ""
    while True:
        index, remainder = divmod(index, 26)
        letters = chr(ord("A") + remainder) + letters
        if index == 0:
            break
        index -= 1
    return letters

# Record version numbers observed in the sample.
EMAIL_VERSION = 1
MAWB_VERSION = 1
HAWB_VERSION = 14
COMMODITY_VERSION = 4

# Per-shipment constants the NetCHB export does not carry. Each of these is a
# single value across all 185 shipments in the sample.
CONTENTS_CODE = "APX"
PACKAGING_CODE = "B"
CURRENCY = "USD"
CONSIGNEE_COUNTRY = "US"
WEIGHT_UNIT = "KG"

# Flight-level fields the NetCHB export does not carry. All four are optional
# on the form: whatever the operator types is written through, and anything
# left blank stays blank and is listed in the manual-fill report.
OPERATOR_FIELDS = ("man_code", "flight_info", "man_origin", "man_dest")

# Field widths. Taken from the longest value seen in the sample, which shows
# clear truncation at these lengths (95 shipper_address1 values sit exactly at
# 25 characters). Values longer than these are cut and reported as warnings.
WIDTH_SHIPPER_NAME = 35
WIDTH_ADDRESS_LINE = 25
WIDTH_CONSIGNEE_PERSON = 35
WIDTH_HAWB_DESCRIPTION = 35
WIDTH_COMMODITY_DESCRIPTION = 28

# HTS codes are always exactly 10 digits in the AAMS file.
HTS_LENGTH = 10

# The FDA prior-notice number is embedded in the NetCHB shipper address as a
# 'Z' followed by 9-15 digits, e.g. '... DH CG ROAD Z260626032612'.
PRIOR_NOTICE_PATTERN = r"\bZ(\d{9,15})\b"
