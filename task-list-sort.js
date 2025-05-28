function TaskListSort(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Task List") return;

  const inputRow = 2;
  const dataStartRow = 3;
  const lastColumn = 4;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // ---- CASE 1: User is entering a new task in row 2 ----
  if (row === inputRow) {
    const inputValues = sheet.getRange(inputRow, 1, 1, lastColumn).getValues()[0];
    const [taskName, dateValue, dropdownValue, notes] = inputValues;

    // Only proceed if BOTH date and dropdown are filled
    if (!dateValue || !dropdownValue) return;

    // Copy row 3 formatting to the new row at bottom
    const lastRow = sheet.getLastRow();
    sheet.insertRowsAfter(lastRow, 1);
    const templateRange = sheet.getRange(dataStartRow, 1, 1, lastColumn);
    const newRowRange = sheet.getRange(lastRow + 1, 1, 1, lastColumn);
    templateRange.copyTo(newRowRange, { formatOnly: true });

    // Set values into new row
    newRowRange.setValues([[taskName, dateValue, dropdownValue, notes]]);

    // Clear A2, B2, D2; reset C2 dropdown to "n/a"
    sheet.getRange(inputRow, 1, 1, 1).clearContent(); // A2
    sheet.getRange(inputRow, 2, 1, 1).clearContent(); // B2
    sheet.getRange(inputRow, 4, 1, 1).clearContent(); // D2
    sheet.getRange(inputRow, 3).setValue("n/a"); // Reset dropdown

    // Sort rows 3+ by column B, descending
    const updatedLastRow = sheet.getLastRow();
    const sortRange = sheet.getRange(dataStartRow, 1, updatedLastRow - dataStartRow + 1, lastColumn);
    sortRange.sort({ column: 2, ascending: false });

    return;
  }

  // ---- CASE 2: User edits an existing date in column B from row 3 down ----
  if (col === 2 && row >= dataStartRow) {
    const lastRow = sheet.getLastRow();
    const sortRange = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastColumn);
    sortRange.sort({ column: 2, ascending: false });
  }
}
