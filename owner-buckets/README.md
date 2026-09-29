# Finti Owner Buckets

Working copy of Merchant Onboarding Manager with Finti Owner assignment and four live views. This folder is the WIP. It does not edit the live OneDrive file.

Open `Merchant_Onboarding_Manager_OWNER_WIP.xlsx` in Excel for Microsoft 365.

## What changed

Dealer Master gained four columns **after** the existing table, so the lender columns stayed on their current letters.

| Column | Header | What it is |
| --- | --- | --- |
| A | Legal Business Name | Dealer name. Buckets ignore rows with a blank name. |
| K | Priority | 1, 2, or 3. Blank priority sorts last on the bucket sheets. |
| L | Overall Status | Existing calculated formula, plus a Ready check. |
| P | Service Finance | Lender status (SF). |
| T | Sunlight | Lender status (SL). |
| X | Foundation | Lender status (FF). |
| AB | Chowder | Lender status (CH). |
| AF | US Bank | Lender status (USB). |
| AJ | Merchant To-Do | Existing to-do text. |
| AK | To-Do Done | Existing done flag. |
| BZ | Finti Owner | Dropdown. |
| CA | Ready for rate sheets | Yes or No. |
| CB | Shaw Hold | Formula. Warning only. It does not assign an owner. |
| CC | Handoff note | Written by the Office Script when a handoff runs. |

Finti Owner dropdown, on Dealer Master `BZ5:BZ1021` and on Add Merchant `B29`:

`Evan Young` | `Brittany Shaw` | `Brittany Coutu` | `Unassigned`

Every existing dealer (48 named rows, Dealer Master rows 5–52) is **Unassigned**. Rep is not Finti Owner. Roof Boyz LLC, Supreme Home Services LLC, and DANIEL FREUND list Evan Young as Rep. Their Finti Owner is still Unassigned.

A blank owner on a named row is shown as Unassigned on the bucket sheets.

## Sheets

| Sheet | Who it shows | Sort |
| --- | --- | --- |
| Evan All | Every named dealer, including Unassigned | Finti Owner, then Priority, then name |
| Evan Mine | Finti Owner = Evan Young | Priority, then name |
| Shaw | Finti Owner = Brittany Shaw | Priority, then name |
| Coutu | Finti Owner = Brittany Coutu | Priority, then name |
| Ready for rate sheets | Handoff preview and instructions | |
| OwnerLists | Hidden list behind the dropdown | |

The four views are Excel 365 `FILTER` and `SORTBY` formulas. They recalculate on the web, on desktop, and in the Excel mobile app. They do not use VBA.

Team Priorities keeps the weekly focus and the existing queue. The queue now has a **Finti Owner** column, and columns I–J count how many named dealers sit with each owner.

## Shaw rule

Shaw keeps a dealer until every **applicable** lender is `Approved` or `Declined`.

Applicable means the lender status is not blank, `N/A`, or `No Min RQ`.

If any applicable lender is `Submitted`, Shaw Hold says `Stays with Shaw - lender still Submitted`.

`Active`, `Waiting on Dealer`, and `Ready 2 Submit` are not `Approved` or `Declined`, so Shaw Hold does not call those dealers clear.

Shaw Hold does not move the owner. Only a person, or the Ready for rate sheets handoff, changes Finti Owner.

## Ready for rate sheets

Marking a dealer ready switches Finti Owner to **Brittany Coutu**.

- Where Office Scripts run (Excel on the web, Windows, Mac): follow `INSTALL.md` and run `ReadyForRateSheets`.
- On a phone: set `BZ` to `Brittany Coutu` and `CA` to `Yes` on that Dealer Master row.

`CA = Yes` makes Overall Status return `Rate Sheet` unless the dealer is already `LIVE DLR`, `Training`, `Active`, or `Dead`. The script writes the words `Rate Sheet` only when Overall Status is a typed value with no formula.

## Add Merchant

Add Merchant cell `B29` is the same Finti Owner dropdown. The form starts at `Unassigned`.

The green **Run Add Merchant**, **Run Copy Chowder Email**, and **Run Clear Chowder Fields** buttons are the original Office Script buttons. Their script bodies are not stored in this file. They link to scripts in the SharePoint library that was on the source workbook. There is **no VBA project** in the source export (`vbaProject.bin` is absent, and the button macro field is empty).

That existing Add Merchant script does not know about column `BZ`. After it adds a row, set Finti Owner on Dealer Master. A blank owner shows as Unassigned and stays off the Shaw, Coutu, and Evan Mine sheets.

If a later live copy is a macro-enabled workbook, those macros run in desktop Excel for Windows only. They do not run in Excel on the web, on Mac, or on a phone. This pack does not add any. Online and mobile users still change Finti Owner on Dealer Master, and the FILTER sheets update.

## Device matrix

Microsoft 365 is required. The views use dynamic arrays (`FILTER`, `SORTBY`, `HSTACK`, `LET`).

| Feature | Excel on the web | Desktop Windows | Desktop Mac | Excel mobile |
| --- | --- | --- | --- | --- |
| Finti Owner dropdown | Yes | Yes | Yes | Yes |
| Ready Yes/No dropdown | Yes | Yes | Yes | Yes |
| Evan All, Evan Mine, Shaw, Coutu | Yes, formulas refresh | Yes | Yes | Yes, formulas refresh |
| Shaw Hold warning | Yes | Yes | Yes | Yes |
| ReadyForRateSheets Office Script | Yes, after INSTALL.md | Yes, Microsoft 365, file in OneDrive or SharePoint | Yes, same | No. Set BZ and CA by hand |
| Add Merchant script buttons | Yes, if this account can open the original SharePoint scripts | Same | Same | Usually no |
| VBA / macros | None in this file | None in this file | None in this file | None in this file |

## Promote to live

Test this WIP in Excel on the web before anyone copies it over the live workbook.

Prefer copying these pieces onto the live file, instead of replacing the live file:

- Dealer Master columns `BZ:CC`, the owner and Ready dropdowns, the Shaw Hold formulas, and the Overall Status formula change
- Sheets Evan All, Evan Mine, Shaw, Coutu, Ready for rate sheets, and OwnerLists
- The Finti Owner field on Add Merchant (`A29:B29` and its dropdown)
- The Team Priorities owner column and the counts in columns I–J

Replacing the live file can drop coauthoring history. The Add Merchant buttons in this WIP already carry the source file's script links. They keep working only for accounts that can open those SharePoint scripts.

Do not type owners for the current 48 dealers during the copy. Leave them Unassigned.
