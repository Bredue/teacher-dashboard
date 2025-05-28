function CFADataSort() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "CFA Data") return;
  
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, 1000, lastCol).getValues(); // Row 1–1000
  const headerRow = data[0];

  for (let col = 0; col < lastCol; col++) {
    if (headerRow[col] === "Student Roster") {
      const namesCol = col;
      const scoresCol = col + 1;
      const outputCol = col - 1;
      if (outputCol < 0) continue;

      // Clear old output starting from row 4 to prevent data retention
      sheet.getRange(4, outputCol + 1, 997).clearContent();

      const studentData = [];

      for (let row = 1; row < 1000; row++) {
        const nameCell = data[row][namesCol];
        const scoreCell = data[row][scoresCol];
        if (!nameCell || typeof nameCell !== 'string' || !nameCell.includes("-")) continue;
        if (typeof scoreCell !== 'number') continue;

        const [periodRaw, nameRaw] = nameCell.split("-");
        const period = periodRaw.trim();
        const name = nameRaw.trim();
        const score = scoreCell;

        studentData.push({ period, name, score });
      }

      const totalScore = studentData.reduce((sum, s) => sum + s.score, 0);
      const avgScore = studentData.length ? totalScore / studentData.length : 0;
      const avgCell = sheet.getRange(4, outputCol + 1);
      avgCell.setValue(avgScore.toFixed(2))
             .setFontWeight("bold")
             .setFontSize(12)
             .setHorizontalAlignment("center");

      const periodGroups = {};
      for (const { period, score } of studentData) {
        if (!periodGroups[period]) periodGroups[period] = [];
        periodGroups[period].push(score);
      }

      const periodAvgs = Object.entries(periodGroups)
        .map(([period, scores]) => {
          const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
          return { period, avg };
        })
        .sort((a, b) => a.period.localeCompare(b.period));

      const sortedStudents = [...studentData].sort((a, b) => b.score - a.score);
      const numStudents = sortedStudents.length;
      const topCount = Math.max(1, Math.floor(numStudents * 0.1));
      const bottomCount = Math.max(1, Math.floor(numStudents * 0.1));

      const topStudents = sortedStudents.slice(0, topCount);
      const bottomStudents = sortedStudents.slice(-bottomCount);

      const summaryLines = [];
      summaryLines.push("Period Averages");
      summaryLines.push(...periodAvgs.map(p => `${p.period}: ${p.avg.toFixed(2)}`));
      summaryLines.push("");

      summaryLines.push("Top 10%");
      summaryLines.push(...topStudents.map(s => `${s.name} (${s.score.toFixed(2)})`));
      summaryLines.push("");

      summaryLines.push("Bottom 10%");
      summaryLines.push(...bottomStudents.map(s => `${s.name} (${s.score.toFixed(2)})`));

      const startRow = 6;
      const range = sheet.getRange(startRow, outputCol + 1, summaryLines.length, 1);
      const values = summaryLines.map(line => [line]);
      range.setValues(values);

      for (let i = 0; i < summaryLines.length; i++) {
        const line = summaryLines[i];
        if (["Period Averages", "Top 10%", "Bottom 10%"].includes(line)) {
          const cell = sheet.getRange(startRow + i, outputCol + 1);
          cell.setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
        }
      }
    }
  }
}
