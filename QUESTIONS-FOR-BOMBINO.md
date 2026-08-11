# Questions for the Bombino team

**About:** the tool that converts the NetCHB manifest report into the AAMS / ECCF file.

We have the converter working. It reads the NetCHB report and produces the AAMS
file with the correct structure — 187 shipments, 620 item lines, all the shipper
and consignee details carried across.

There are a few columns we cannot fill in reliably, because the NetCHB report
either does not contain that information at all, or contains something that does
not match your existing AAMS file. Rather than guess at these, we have left them
blank and listed them below.

Please reply against each question. Anything you confirm, we will automate.
Anything that has no source, we will leave blank for your team to fill in.

**A note on column letters:** where we say "column V", that is the column you
see when you open the converted file in Excel.

---

## The important ones

### 1. HTS codes — item rows, column I

**What happens now:** NetCHB gives us 8-digit codes, like `62114210`. The AAMS
file needs 10 digits. We add two zeros on the end, making `6211421000`.

**Why we are unsure:** we think NetCHB is giving us the Indian export codes,
while the AAMS file needs the US codes. Adding zeros does not convert one into
the other. When we compared our codes against your existing AAMS file, only
about 3 in every 100 matched.

**What we need from you:** should we keep adding zeros, or should we leave these
lines blank so someone classifies them properly? This affects 387 of the 723
lines, so it is a big decision either way.

**Your answer:**

---

### 2. Declared value — shipment rows, column V

**What happens now:** left blank.

**Why we are unsure:** we were originally taking NetCHB's `InvoiceValue`, but
that turned out to be wrong. For House AWB `4307923`, your AAMS file shows a
declared value of **8**, while NetCHB shows **7400** for the same shipment. We
tried every way of calculating it from the NetCHB numbers and none of them
matched your file.

**What we need from you:** where should this number come from? If it is not in
any export, we will leave it as a manual entry.

**Your answer:**

---

### 3. Weight — shipment rows, column R

**What happens now:** left blank.

**Why we are unsure:** there is no weight column anywhere in the NetCHB report —
we checked all 133 columns. But your existing AAMS file has a weight for every
single shipment, so it must come from somewhere else.

**What we need from you:** is there another export or system that has the weight
for each House AWB? If you can send us a sample, we can read the weight from
there automatically. Otherwise it stays a manual entry.

**Your answer:**

---

### 4. Shipper address — shipment rows, columns AJ and AK

**What happens now:** NetCHB keeps the whole street address in one box, up to
105 characters long. The AAMS file has two separate address boxes, which we
believe hold 25 characters each. We split the address at a word break and cut
off anything that does not fit.

**Why we are unsure:** on 66 shipments, the end of the address is lost
completely. We also do not know for certain that 25 characters is the real
limit — we worked it out by looking at the longest address in your sample file.

**What we need from you:** how long are those two address boxes really? And is
losing the end of long addresses acceptable, or should we handle it differently?

**Your answer:**

---

### 5. Which row wins — shipment rows, columns X and Y

**What happens now:** a single shipment often has several rows in NetCHB, one
for each type of goods. The AAMS shipment row only has space for one description
and one HTS code. At the moment we take whichever appeared first in the file.

**Why we are unsure:** "first in the file" is an arbitrary choice, not a rule.
80 of the 187 shipments have more than one description, and 49 have more than
one HTS code. So for those, we may be picking the wrong one.

**What we need from you:** which should we use — the first line, the
highest-value line, a general description like "MIXED GOODS", or leave it blank?

**Your answer:**

---

## Smaller ones

### 6. Three columns where NetCHB already has data we are not using

These would be quick to fix — we just need a yes or no on each.

| NetCHB column | What it contains | Where it could go |
|---|---|---|
| `ManufacturerProvince` | The shipper's state — MH, GJ, DL and so on. Filled on 647 of 723 rows. | `shipper_state`, column AM, which we currently leave blank |
| `Unit` | `CTN` on every row | `packaging`, column AB, where we currently always write `B` |
| `InvoiceNumber` | Numbers like `127-EI/EXP/26-27` | `ship_ref_num`, column H, which is currently blank |

We left these out because your existing AAMS file has those columns either blank
or set to something else. But it seems wasteful to throw away data NetCHB is
already giving us.

**Should we carry any of these across?**

**Your answer:**

---

### 7. Which carrier is correct?

**What happens now:** the flight is typed in by hand each time.

**Why we are unsure:** NetCHB says the carrier is **QR** on all 723 rows. Your
existing AAMS file says the flight is **BA 0117**. The master bill number starts
with **125**, which is British Airways' number, so that agrees with the AAMS file
and not with NetCHB.

**What we need from you:** which one is right? If NetCHB's carrier code is
reliable, we could fill the flight in automatically instead of asking someone to
type it.

**Your answer:**

---

### 8. Are the NetCHB amounts in rupees or dollars?

**What happens now:** we write `USD` as the currency, because your existing AAMS
file uses `USD` on every row.

**Why we are unsure:** the values in NetCHB look like they could be rupees. If
they are, then any amount we copy across would need converting first.

**What we need from you:** which currency are the NetCHB amounts in?

**Your answer:**

---

### 9. Can one file ever cover two manifests?

**What happens now:** one conversion produces one file covering one manifest.

**Why we are unsure:** the sample file you gave us is named
`TY523 & TT523V`, which looks like two manifest codes. But inside the file,
`TT523V` does not appear anywhere and there is only one flight row.

**What we need from you:** is it always one manifest per file? If two are ever
combined, please send us an example so we can match the format exactly.

**Your answer:**

---

### 10. A few fixed values we copied from your sample file

Every one of these is the same on all 185 rows of the AAMS file you sent us, so
we write the same value on every row. Please confirm they are always correct and
not just correct for that one flight.

| Column | We always write |
|---|---|
| `contents`, column T | `APX` |
| `weight_unit`, column S | `KG` |
| `currency_code`, column U | `USD` |
| `consignee_country`, column AX | `US` |
| Notification emails at the top of the file | bombino@bombinoexp.com and salim@bombinoexp.com |

**Your answer:**

---

## What is already working

For context, these are all confirmed correct — we checked them against the 185
shipments that appear in both your files, and they matched every time:

- House AWB, piece counts, and the DDP / DDU service type
- Consignee name, address, city, state, postal code and country
- Shipper name, city, postal code and country
- The master bill number and the manifest date
- The FDA prior notice number, which we pull out of the shipper address
- Item quantities and values, which add up exactly against the source

We also found and fixed one problem worth mentioning: Excel was silently dropping
the leading zero from HTS codes that start with `0`. This meant tea was being
filed under chapter 90 (optical instruments) and dal under chapter 71
(jewellery). That is now corrected.

---

*Prepared by Onshore Labs. All figures measured from `manifest_netchb_report.xls`
and `aams-data-bombino-express-eccf TY523 & TT523V MODI.xls`.*
