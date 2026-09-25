/**
 * Ready for rate sheets handoff for Merchant Onboarding Manager.
 *
 * Paste this into Excel Automate > New Script and name it ReadyForRateSheets.
 * Install steps: INSTALL.md.
 *
 * Columns on Dealer Master (headers are row 4, dealers start at row 5):
 *   A  Legal Business Name
 *   L  Overall Status
 *   BZ Finti Owner
 *   CA Ready for rate sheets
 *   CB Shaw Hold (formula, read-only warning)
 *   CC Handoff note
 *
 * Owner values: Evan Young | Brittany Shaw | Brittany Coutu | Unassigned
 * Ready values: Yes | No
 *
 * Overall Status token written by this script: "Rate Sheet"
 * That token matches the Dealer Master formula output and the Team Priorities
 * queue test (_xlpm.s="Rate Sheet"). It is not in the Overall Status dropdown
 * list (that list is Not Started, In Process, Waiting on Dealer, Waiting on
 * Lender, Active, Complete, Declined, Dead/Inactive, Applications).
 *
 * When L still contains the calculated formula, this script does not replace
 * it. Ready=Yes makes that formula return Rate Sheet unless the dealer is
 * already LIVE DLR, Training, Active, or Dead.
 *
 * Shaw Hold is a warning. A lender still Submitted means Shaw keeps the
 * dealer under the team rule. Running this script still switches the owner,
 * because the handoff is an explicit action.
 */
function main(workbook: ExcelScript.Workbook): string {
  const masterName = "Dealer Master";
  const handoffName = "Ready for rate sheets";
  const ownerCol = "BZ";
  const readyCol = "CA";
  const statusCol = "L";
  const holdCol = "CB";
  const noteCol = "CC";
  const nameCol = "A";
  const firstDataRow = 5;
  const lastDataRow = 1021;
  const coutu = "Brittany Coutu";
  const readyYes = "Yes";
  const rateSheet = "Rate Sheet";

  const master = workbook.getWorksheet(masterName);
  const handoff = workbook.getWorksheet(handoffName);
  if (!master || !handoff) {
    return "Missing the Dealer Master sheet or the Ready for rate sheets sheet. Nothing was changed.";
  }

  const resultCell = handoff.getRange("B10");
  const nameQuery = textOf(handoff.getRange("B3").getValue());
  const rowQuery = handoff.getRange("B4").getValue();

  let row = 0;
  let how = "";

  if (nameQuery) {
    const values = master.getRange(`${nameCol}${firstDataRow}:${nameCol}${lastDataRow}`).getValues();
    const hits: number[] = [];
    const wanted = nameQuery.toLowerCase();
    for (let i = 0; i < values.length; i++) {
      const name = textOf(values[i][0]);
      if (name && name.toLowerCase() === wanted) {
        hits.push(firstDataRow + i);
      }
    }
    if (hits.length === 0) {
      return finish(resultCell, `No dealer named "${nameQuery}" on Dealer Master. Nothing was changed.`);
    }
    if (hits.length > 1) {
      return finish(resultCell, `More than one dealer matches "${nameQuery}" (rows ${hits.join(", ")}). Nothing was changed.`);
    }
    row = hits[0];
    how = `name "${nameQuery}"`;
  } else if (rowQuery !== null && rowQuery !== "") {
    const parsed = typeof rowQuery === "number" ? rowQuery : parseInt(textOf(rowQuery), 10);
    if (!parsed || parsed < firstDataRow || parsed > lastDataRow) {
      return finish(resultCell, `Row "${rowQuery}" is outside Dealer Master rows ${firstDataRow}-${lastDataRow}. Nothing was changed.`);
    }
    row = parsed;
    how = `row ${row}`;
  } else if (workbook.getActiveWorksheet().getName() === masterName) {
    const selected = workbook.getActiveCell().getRowIndex() + 1;
    if (selected >= firstDataRow && selected <= lastDataRow) {
      row = selected;
      how = `selected Dealer Master row ${row}`;
    }
  }

  if (!row) {
    return finish(
      resultCell,
      "Type the Legal Business Name in B3, or a Dealer Master row number in B4, or select the dealer row on Dealer Master. Then run ReadyForRateSheets again. Nothing was changed."
    );
  }

  const legalName = textOf(master.getRange(`${nameCol}${row}`).getValue());
  if (!legalName) {
    return finish(resultCell, `Row ${row} has no Legal Business Name. Nothing was changed.`);
  }

  const hold = textOf(master.getRange(`${holdCol}${row}`).getValue());
  master.getRange(`${readyCol}${row}`).setValue(readyYes);
  master.getRange(`${ownerCol}${row}`).setValue(coutu);

  const statusCell = master.getRange(`${statusCol}${row}`);
  const formula = (statusCell.getFormula() || "").trim();
  const hasFormula = formula.startsWith("=") || formula.startsWith("IF") || formula.startsWith("_xlfn");
  let statusNote: string;
  if (hasFormula) {
    statusNote = "Overall Status formula left in place. Ready=Yes returns Rate Sheet unless the dealer is already LIVE DLR, Training, Active, or Dead.";
  } else {
    statusCell.setValue(rateSheet);
    statusNote = "Overall Status set to Rate Sheet.";
  }

  master.getRange(`${noteCol}${row}`).setValue(`Ready for rate sheets handoff to ${coutu}.`);

  let warning = "";
  if (hold.indexOf("Stays with Shaw") === 0) {
    warning = ` Warning: ${hold} Owner was still switched because this handoff was run on purpose.`;
  }

  return finish(
    resultCell,
    `Row ${row} (${legalName}) via ${how}: Finti Owner = ${coutu}, Ready for rate sheets = ${readyYes}. ${statusNote}${warning}`
  );
}

function textOf(value: string | number | boolean | undefined | null): string {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value).trim();
}

function finish(resultCell: ExcelScript.Range, message: string): string {
  resultCell.setValue(message);
  return message;
}
