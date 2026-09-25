# Install ReadyForRateSheets

This script is the handoff button for Excel on the web, Excel for Windows, and Excel for Mac. It is not a macro. Excel mobile cannot run Office Scripts; use the manual steps on the **Ready for rate sheets** sheet instead.

The script source is `ReadyForRateSheets.ts`. `ReadyForRateSheets.osts` is the same source in a JSON wrapper. Paste the TypeScript. Excel does not need the `.osts` file.

## Save the script (Excel on the web, Windows, or Mac)

1. Open `Merchant_Onboarding_Manager_OWNER_WIP.xlsx` from OneDrive or SharePoint in Excel for Microsoft 365. Office Scripts do not run from a file that only lives on the laptop with no cloud copy.
2. Go to **Automate** > **New Script**.
3. Delete the sample code. Paste the full contents of `ReadyForRateSheets.ts`.
4. Rename the script to `ReadyForRateSheets`.
5. Click **Save script**.

## Run it

1. Open the sheet **Ready for rate sheets**.
2. Type the dealer's **Legal Business Name** in `B3` (exact name, any capitalization). Or leave `B3` blank and type the Dealer Master row number in `B4`. Or leave both blank, select that dealer's row on **Dealer Master**, and run the script from there.
3. Read **Shaw Hold** in `B7` before you run it. If it says a lender is still Submitted, Brittany Shaw keeps that dealer under the team rule. The script still switches the owner when you run it. That is deliberate.
4. **Automate** > **ReadyForRateSheets** > **Run**.
5. `B10` reports what changed. On Dealer Master the row now has:
   - **Finti Owner** (`BZ`) = `Brittany Coutu`
   - **Ready for rate sheets** (`CA`) = `Yes`
   - **Handoff note** (`CC`) = a short note
   - **Overall Status** (`L`) shows `Rate Sheet` when the calculated formula is still there, unless the dealer is already `LIVE DLR`, `Training`, `Active`, or `Dead`. If `L` was a typed value with no formula, the script writes `Rate Sheet`.

Running it again on the same dealer is safe. It sets the same values.

## Optional button

After the script is saved, open it and choose **Add button** (wording may be **More options** > **Add button**). Put the button on **Ready for rate sheets**. The button only works in that workbook for people who can open the saved script. This pack cannot pre-bind the button: Office Script buttons point at a script in your OneDrive, and that link does not exist until you save the script.

## Mobile

Excel for iPhone and Android can edit the dropdowns and can show the FILTER buckets. It cannot run this script.

On the phone:

1. On **Ready for rate sheets**, type the legal name in `B3` and confirm the matched row.
2. On that Dealer Master row, set **Finti Owner** to `Brittany Coutu` and **Ready for rate sheets** to `Yes`.

The Coutu sheet updates from those two cells. You do not need the script for the buckets to move.

## If Run is missing

- The file must be in OneDrive or SharePoint, opened with a Microsoft 365 work or school account that is allowed to use Office Scripts.
- Desktop Excel must be Microsoft 365, not a perpetual license without Automate.
- A file opened from an email attachment will not run the script until it is saved to OneDrive.
