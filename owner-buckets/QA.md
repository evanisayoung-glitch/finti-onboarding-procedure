# QA — Finti Owner Buckets

Checked against the source workbook `Merchant_Onboarding_Manager_live.xlsx` (sheets Add Merchant, Dealer Master, Team Priorities, and the rest). This environment has no Excel 365, so `FILTER` / `SORTBY` spills and the Office Script were not executed inside Excel. What was checked is the workbook XML, an openpyxl reload, formula paren balance, and a Python replay of the Shaw Hold rule on the 48 dealer rows.

Run the checklist below in Excel for Microsoft 365 on the web before promoting the WIP.

## Column letters found

Unchanged lender and status columns:

| Letter | Header |
| --- | --- |
| A | Legal Business Name |
| K | Priority |
| L | Overall Status |
| P | Service Finance (SF) |
| T | Sunlight (SL) |
| X | Foundation (FF) |
| AB | Chowder (CH) |
| AF | US Bank (USB) |
| AJ | Merchant To-Do |
| AK | To-Do Done |

New columns, appended after `BY` so nothing above shifted:

| Letter | Header |
| --- | --- |
| BZ | Finti Owner |
| CA | Ready for rate sheets |
| CB | Shaw Hold |
| CC | Handoff note |

Dealer Master headers are row 4. Named dealers are rows 5–52 (48 dealers). JAAR BUILDERS is row 52 and sits outside the `DealerMaster` table in the source file. It is included in the buckets. The table range was not resized.

## Lender status vocabulary found

Dropdown on Service Finance, Sunlight, Foundation, Chowder, and US Bank for the data rows (`P5:P1021`, `T5:T1021`, `X5:X1021`, `AB5:AB1021`, `AF5:AF1021`):

`Waiting on Dealer`, `Ready 2 Submit`, `Submitted`, `Approved`, `Active`, `Declined`, `No Min RQ`, `N/A`

A second list exists only on row 1022 of those columns (`Not Started`, `Submitted`, `Approved`, `Active`, `Declined`, `N/A`, `Pending - verify`). None of those extra tokens (`Not Started`, `Pending - verify`) appear on rows 5–52.

Counted on rows 5–52 across the five lender columns (240 cells), after resolving the "Ready 2 Submit / Waiting on Dealer" formulas that follow To-Do Done:

| Status | Cells |
| --- | --- |
| N/A | 95 |
| Waiting on Dealer | 45 |
| Submitted | 29 |
| Active | 26 |
| Declined | 21 |
| Approved | 11 |
| No Min RQ | 8 |
| Ready 2 Submit | 5 |

Overall Status formula outputs (column L): `LIVE DLR`, `Training`, `Active`, `Dead`, `Waiting on Dealer`, `Waiting on Lender`, `Rate Sheet`, `Application Submitted`.

`Ready for rate sheets = Yes` returns `Rate Sheet` from that formula unless the dealer is already `LIVE DLR`, `Training`, `Active`, or `Dead`.

The Overall Status **dropdown** (column L, mostly below the data block) is a different list: `Not Started`, `In Process`, `Waiting on Dealer`, `Waiting on Lender`, `Active`, `Complete`, `Declined`, `Dead/Inactive`, `Applications`. It does **not** contain `Rate Sheet`. The script uses `Rate Sheet` because that is the token the formula and the Team Priorities queue already use.

## Owners

- Dropdown is exactly `Evan Young`, `Brittany Shaw`, `Brittany Coutu`, `Unassigned` (named range `FintiOwnerList` on hidden sheet OwnerLists).
- All 48 named dealers are `Unassigned`. Ready is `No`. No other owner text was written.
- Rep = Evan Young on Roof Boyz LLC, Supreme Home Services LLC, and DANIEL FREUND. Finti Owner on those rows is Unassigned. They belong on Evan All, not Evan Mine, until someone assigns them.

## Shaw Hold replay (expected once Excel calculates column CB)

| Shaw Hold | Dealers |
| --- | --- |
| Stays with Shaw - not all lenders Approved or Declined | 24 |
| Stays with Shaw - lender still Submitted | 18 |
| Clear - all applicable lenders Approved or Declined | 5 |
| No applicable lenders | 1 |

Spot checks:

| Row | Dealer | Lenders | Expected Shaw Hold |
| --- | --- | --- | --- |
| 5 | Vivid Windows LLC | SF/SL/FF Active, CH/USB N/A | Stays with Shaw - not all lenders Approved or Declined |
| 7 | Dirty Fence and Concrete | USB Submitted (FF Active) | Stays with Shaw - lender still Submitted |
| 13 | Vision Builders | SF/SL/FF Declined, CH/USB N/A | Clear - all applicable lenders Approved or Declined |
| 17 | Bluebird Turf Co | FF Approved, others N/A | Clear - all applicable lenders Approved or Declined |
| 20 | Supreme Home Services LLC | SF and FF Declined. Rep is Evan Young | Clear - all applicable lenders Approved or Declined. Owner stays Unassigned |
| 32 | Iron Ridge Group LLC | No Min RQ or N/A on every lender | No applicable lenders |

With every owner still Unassigned, Excel should show:

- Evan All: 48 dealers
- Evan Mine, Shaw, Coutu: empty (`No dealers for …`)
- Team Priorities J15:J17 = 0, J18 = 48, J19 = 48

## Add Merchant and macros

- Source file is `.xlsx`. No `vbaProject.bin`. Button shapes say `macro=""`.
- Three buttons remain on Add Merchant and still point at the original Office Script share links: Run Add Merchant, Run Copy Chowder Email, Run Clear Chowder Fields. The drawing part is byte-for-byte the same as the source.
- Those script bodies are not in the file. They do not write Finti Owner. Add Merchant `B29` is the new dropdown and defaults to Unassigned.
- If a different live copy has Windows VBA, that VBA stays desktop-Windows-only. This WIP does not use it for the handoff.

## Checklist in Excel

- [ ] File opens in Excel on the web without a repair dialog.
- [ ] Dealer Master `BZ5` dropdown offers only the four owner names. `CA5` offers Yes and No.
- [ ] `BZ5:BZ52` are Unassigned. `CA5:CA52` are No. No dealer was given Shaw, Coutu, or Evan.
- [ ] Evan All lists all 48 names, including Unassigned. Evan Mine, Shaw, and Coutu are empty.
- [ ] Set one test row's Finti Owner to Brittany Shaw. That dealer leaves the empty Shaw message and appears only on Shaw and Evan All. Set it back to Unassigned.
- [ ] Repeat for Brittany Coutu (Coutu + Evan All) and Evan Young (Evan Mine + Evan All).
- [ ] Shaw Hold on row 13 (Vision Builders) reads Clear. Row 7 (Dirty Fence and Concrete) reads Stays with Shaw - lender still Submitted. Row 5 (Vivid Windows LLC) is not clear, because Active is not Approved or Declined.
- [ ] On Ready for rate sheets, type `Vision Builders` in B3. Matched row is 13. Run `ReadyForRateSheets` (see INSTALL.md). Owner becomes Brittany Coutu, Ready becomes Yes, the dealer shows on Coutu and Evan All, and not on Shaw.
- [ ] B10 describes the handoff. If you run it on Dirty Fence and Concrete, B10 includes the Shaw Hold warning and the owner still becomes Brittany Coutu.
- [ ] Put the test rows back to Unassigned and Ready = No before anyone treats this as live data.
- [ ] On a phone, confirm the dropdown edits and that Evan All shows the same owner you just set. Do not expect Automate on mobile.
- [ ] Add Merchant `B29` shows the same four names. Existing Run Add Merchant button is still on the sheet.
- [ ] `index.html` and `dealer/index.html` are unchanged.

## Structural checks already run

- Workbook XML parses. openpyxl reloads all 20 sheets.
- 48 inline `Unassigned` owner values and 48 `No` ready values. Zero `Yes` values.
- Overall Status formulas on L5:L52, the table calculated column, and the totals formula all reference the Ready cell (`$CA` on that row). L5 matches the table formula.
- Evan All, Evan Mine, Shaw, and Coutu each contain `_xlfn._xlws.FILTER` and `_xlfn.SORTBY`.
- Team Priorities queue array now spills through column N (`Finti Owner`).
- Add Merchant drawing hash matches the source file.
- Onboarding HTML pages are not in this change.
