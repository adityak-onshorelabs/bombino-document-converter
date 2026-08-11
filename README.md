# Bombino Document Converter

Converts a **NetCHB manifest report** (`.xls`) into an **AAMS / ECCF transmission
file** (`.xls`) for a single flight.

## Running it

```bash
pip install -r requirements.txt
python -m uvicorn app.main:app --port 8000
```

Open <http://127.0.0.1:8000>, drop the NetCHB report on the page, and download
the result. The four flight fields are optional — fill in what you know.

```bash
python -m pytest tests/          # 27 tests, run against the real sample pair
```

> **The sample files are not in this repository.** They contain consignee names
> and home addresses for around 187 people, so `sample_docs/` is gitignored. Drop
> `manifest_netchb_report.xls` and the AAMS sample into `sample_docs/` to run the
> test suite; the app itself runs without them.

## What the two formats are

**NetCHB manifest report** — a flat table, one sheet, **one row per HTS entry
line**. The sample has 723 rows across 133 columns, of which only 43 carry data;
the rest are unused FDA / Lacey / steel-aluminium compliance slots.

**AAMS / ECCF file** — not a table. A hierarchical record stream rendered into a
spreadsheet. Every row is:

```
[line_seq, shipment_index, group, ...record payload]
```

- **col 0** — line sequence, contiguous from 1
- **col 1** — `0` for the header block, then `1..N`, one per House AWB
- **col 2** — `DUTY` for the header block, else `DDP` / `DDU`
- **col 3+** — the record payload, preceded by its own `# record_type` header row

```
email     x2            notification addresses
# header
mawb      x1            flight-level record
  # header
  hawb    x1            one per House AWB
    # header            emitted only when commodities follow
    commodity x0..n
```

Shipments are recovered by grouping NetCHB rows on `House AWB`. A shipment with
a single entry line gets no commodity block (its detail lives on the `hawb`
record); one with two or more lines gets a commodity block.

The field lists in `app/mapping.py` are copied verbatim from the sample's header
rows, **including the `manufactuer_id` misspelling and the two-space
`#  record_type` on the mawb header**. Do not tidy them up — the receiving
system matches on them.

## Columns that must be filled in by hand

The NetCHB export is a downstream, broker-side view of the shipment, so some
AAMS columns have no source in it. The converter leaves these blank rather than
writing a plausible-looking guess, and the results page lists each one with a
live count. **These are the only columns that need manual work.**

**Never derivable** — no source exists, so these always need a human:

| Record | Column | Blank | Why |
|---|---|---|---|
| `hawb` | `weight` (R) | every row | The NetCHB report has no weight column at all — all 133 were checked. |
| `hawb` | `declared_value` (V) | every row | No NetCHB column holds the AAMS declared value. See below. |
| `hawb` | `hts_code` (Y) | every row | NetCHB's codes are not the codes AAMS wants. See below. |
| `commodity` | `hts_code` (I) | every row | Same. |

**Optional on the form** — typed in per run, blank if you skip them:

| Record | Column | Also fills |
|---|---|---|
| `mawb` | `man_code` | — |
| `mawb` | `flight_info` | — |
| `mawb` | `man_origin` | `origin` on all 187 `hawb` records |
| `mawb` | `man_dest` | `final_destination` on all 187 `hawb` records |

None of the four is in the NetCHB export. Whatever you type is written through;
whatever you leave blank stays blank and appears in the manual-fill table on the
results page. Note that the sample's `flight_info` (`BA 0117`) contradicts the
NetCHB `Carrier Code` column (`QR` on all 723 rows); the MAWB prefix `125` is
British Airways', so the NetCHB value looks like the unreliable one. The
converter reads neither — the flight only ever comes from the form.

Everything else is either derived from the NetCHB data or is a constant the
sample confirms (`contents=APX`, `packaging=B`, `currency_code=USD`,
`weight_unit=KG`, `consignee_country=US`). Columns such as `profile_key`,
`ship_ref_num`, `shipper_state` and `mid` stay empty because they are empty in
all 185 sample records too — the format does not use them.

### Why `declared_value` is blank

**Do not derive this from `InvoiceValue`.** It is not the same figure, and an
earlier build of this converter got it wrong.

For House AWB `4307923` the AAMS file shows a declared value of `8`. NetCHB
shows `InvoiceValue = 7400` for the same shipment, built from `1000 × 7` plus
`200 × 2` — note that NetCHB's `HTSValue` is a **unit price**, so
`InvoiceValue = Σ(HTSValue × HTSQty)`, which holds for 184 of 187 shipments.
That roll-up is internally consistent, but it is not the AAMS declared value.

Every candidate rule was tested against the 185 shipments the two sample files
share, and none of them reproduces the AAMS figure:

| Rule | Matches |
|---|---|
| `InvoiceValue` | 7/185 |
| `sum(HTSValue × HTSQty)` | 7/185 |
| `sum(HTSValue)` | 4/185 |
| `first HTSValue` | 3/185 |

The NetCHB-to-AAMS value ratio ranges from 0.75 to 10,738 with a median of 93.5,
so it is not a currency conversion either. The two systems simply hold
independent valuations. The column is therefore left empty rather than filled
with a number that would be wrong on a customs declaration.

`commodity.declared_value` is still mapped from `HTSValue`. If that turns out to
be wrong too, blank it the same way.

### Why the HTS columns are blank

AAMS wants exactly 10 digits. NetCHB files 8 digits on most lines, 10 on some,
and nothing at all on 225 of 723. An earlier build right-padded the short ones
with zeros; that has been removed.

NetCHB appears to carry Indian ITC-HS export codes while AAMS wants US HTSUS
codes, and appending zeros does not convert between the two. Measured against
the reference AAMS file, padded codes matched on **3%** of lines exactly, 4% at
8 digits, and 16% even at the 6-digit HS base — so the goods are being
reclassified outright, not merely given a statistical suffix.

Both HTS columns are therefore written blank pending a rule from Bombino.

`_hts()` in `app/aams.py` is deliberately kept although nothing calls it. It
holds one piece of verified handling worth not losing: Excel stores the NetCHB
HTS column as a *number*, so a leading zero disappears on the way in —
`09024020` arrives as `9024020`. Codes only come in 4, 6, 8 or 10 digits, so an
odd 7 or 9 means exactly that, and the zero is put back. Without it, tea files
under chapter 90 (optical instruments) and dal under chapter 71 (jewellery).
Its unit test still runs, so the logic stays honest until it is reconnected.

## Where the conversion loses fidelity

These are reported as warnings on the results page. They are not blank columns —
a value is written, but it may need review.

**Field widths.** AAMS address lines hold 25 characters and descriptions 28–35;
NetCHB street addresses run to 105 characters and descriptions to 64. Overlong
values are cut and reported. Note the sample shows a human *abbreviating* these
instead (`EXAMPLE TRADING PRIVATE LIMITED` → `Example Trading Pvt Ltd`) — the
converter cannot invent abbreviations, so review the warning list and shorten
anything that matters.

**Shipper address splitting.** NetCHB stores the street as one flat string; AAMS
wants two 25-character lines. In the sample these were separate fields in the
carrier's own system all along, so no exact re-split rule is recoverable —
neither word wrapping nor hard slicing reproduces the sample better than chance.
Until Bombino confirms the rule, the converter splits on **word boundaries
only**, so both lines stay readable, and drops whatever does not fit in two
lines rather than cutting mid-word. The drop is reported per shipment.

**Route.** Leaving `man_origin` / `man_dest` blank also blanks `origin` and
`final_destination` on every `hawb` record, since those mirror the mawb values.
The manual-fill table says so explicitly when it happens.

## A note on the sample pair

`sample_docs/` holds both files for the same flight (MAWB `125-10391485`,
29-07-2026). They are *not* a clean input/output pair — the AAMS one is marked
`MODI` and was hand-edited. 185 of the 187 House AWBs appear in both, and the
identity fields agree exactly (pieces 185/185, DDP/DDU 185/185, and all 50 FDA
prior-notice numbers). Commodity lines, HTS codes, descriptions and values do
**not** agree, because the broker reclassified them downstream. The tests assert
only on the derivable fields for that reason.

## Layout

| Path | What it holds |
|---|---|
| `app/mapping.py` | AAMS field lists, constants, field widths |
| `app/netchb.py` | NetCHB reader, `norm()`, grouping by House AWB |
| `app/aams.py` | record-stream builder and the `.xls` writer |
| `app/main.py` | FastAPI routes |
| `app/templates/index.html` | the single page |
| `tests/test_conversion.py` | structural + differential tests |

Cell types follow the sample: `hawb`, `shipper_zip` and `num_pieces` are
numeric; `consignee_postal_code` is text so leading zeros survive. The same
rule applies to `hts_code` — text when it starts with a zero, numeric
otherwise — for when that column is switched back on.
