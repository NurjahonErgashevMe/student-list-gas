function getOrderColumnIndex(headerRow) {
  return headerRow.findIndex(
    (col) =>
      col.toString().trim().toLowerCase().includes("order") ||
      col.toString().trim().toLowerCase().includes("№")
  );
}

function collectRowsInfo(data) {
  return data.map((row, index) => ({
    isStudent: isStudentRow(row),
    data: row,
    originalIndex: index,
  }));
}

function sortStudentRows(students) {
  return students.sort((a, b) => {
    var classA = parseClass(a.data[2]);
    var classB = parseClass(b.data[2]);

    if (classA.grade !== classB.grade) return classA.grade - classB.grade;
    return classA.letter.localeCompare(classB.letter);
  });
}

function buildSortedResult(studentsToSort, nonStudentRows) {
  var result = [];
  var studentIndex = 0;
  var nonStudentIndex = 0;

  for (var i = 0; i < studentsToSort.length + nonStudentRows.length; i++) {
    if (nonStudentRows.some((r) => r.originalIndex === i)) {
      result.push(nonStudentRows[nonStudentIndex].data);
      nonStudentIndex++;
    } else {
      var student = studentsToSort[studentIndex];
      student.data[0] = studentIndex + 1;
      result.push(student.data);
      studentIndex++;
    }
  }

  return result;
}

function collectOrderUpdates(data, orderCol) {
  var updates = [];
  var currentOrder = 1;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (isStudentRow(row) && row[orderCol] !== currentOrder) {
      updates.push({
        row: i + 1,
        col: orderCol + 1,
        value: currentOrder,
      });
    }
    if (isStudentRow(row)) currentOrder++;
  }

  return updates;
}

function applyUpdates(sheet, updates) {
  updates.forEach((update) => {
    sheet.getRange(update.row, update.col).setValue(update.value);
  });
}

function findClassColumn(headers) {
  const possibleNames = ["sinfi", "класс", "class"];
  return headers.findIndex((header) =>
    possibleNames.some((name) => header.toString().toLowerCase().includes(name))
  );
}

function getUniqueClassesFromStudents(data, classColumn) {
  const classes = new Set();

  data.forEach((row) => {
    if (isStudentRow(row)) {
      const classValue = (row[classColumn] || "").toString().trim();
      if (classValue) classes.add(classValue);
    }
  });

  return Array.from(classes).sort();
}

function isStudentRow(row) {
  return (
    row.length >= 4 &&
    !isNaN(row[0]) &&
    typeof row[1] === "string" &&
    typeof row[2] === "string" &&
    typeof row[3] === "string"
  );
}

function deleteSelectedStudents(
  selectedClasses,
  startRow,
  startCol,
  classColumn
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(
    startRow,
    startCol,
    sheet.getLastRow() - startRow + 1,
    sheet.getLastColumn() - startCol + 1
  );
  const data = range.getValues();

  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (isStudentRow(row)) {
      const classValue = (row[classColumn] || "").toString().trim();
      if (selectedClasses.includes(classValue)) {
        rowsToDelete.push(startRow + i);
      }
    }
  }

  let deletedCount = 0;
  rowsToDelete.forEach((row) => {
    sheet.deleteRow(row);
    deletedCount++;
  });

  updateOrderNumbers(sheet);
  return deletedCount;
}

function updateOrderNumbers(sheet) {
  const data = sheet.getDataRange().getValues();
  let order = 1;

  for (let i = 1; i < data.length; i++) {
    if (isStudentRow(data[i])) {
      sheet.getRange(i + 1, 1).setValue(order);
      order++;
    }
  }
}

function buildDeleteDialogHtml(classes, range, classColumn) {
  const options = classes
    .map((cls) => `<option value="${cls}" selected>${cls}</option>`)
    .join("\n");
  const size = Math.min(classes.length, 10);

  const html = HtmlService.createHtmlOutput(
    `
      <html>
        <head>
          <style>
            body {
              font-family: Roboto, Arial, sans-serif;
              padding: 16px;
              background-color: #f5f5f5;
              margin: 0;
            }
            label {
              font-size: 14px;
              color: #212121;
              display: block;
              margin-bottom: 8px;
            }
            select[multiple] {
              width: 100%;
              padding: 8px;
              border: 1px solid #dadce0;
              border-radius: 4px;
              font-size: 14px;
              background-color: #fff;
              max-height: 150px;
              overflow-y: auto;
            }
            select[multiple]:focus {
              outline: none;
              border-color: #6200ea;
              box-shadow: 0 0 0 2px rgba(98, 0, 234, 0.2);
            }
            .button-container {
              display: flex;
              justify-content: space-between;
              margin-top: 16px;
            }
            button {
              background-color: #6200ea;
              color: white;
              padding: 8px 16px;
              border: none;
              border-radius: 4px;
              font-size: 14px;
              cursor: pointer;
              text-transform: uppercase;
            }
            button:hover {
              background-color: #3700b3;
            }
            button.secondary {
              background-color: #f5f5f5;
              color: #6200ea;
              border: 1px solid #dadce0;
            }
          </style>
        </head>
        <body>
          <label>Выберите классы для удаления:</label>
          <select multiple id="removeClasses" size="${size}">
            ${options}
          </select>
          <div class="button-container">
            <button onclick="google.script.host.close()" class="secondary">Отмена</button>
            <button onclick="submitOptions()">Удалить</button>
          </div>
          <script>
            function submitOptions() {
              const select = document.getElementById('removeClasses');
              const selectedClasses = Array.from(select.selectedOptions).map(opt => opt.value);
              if (selectedClasses.length === 0) {
                alert("Пожалуйста, выберите хотя бы один класс.");
                return;
              }
              google.script.run.withSuccessHandler(closeDialog)
                              .withFailureHandler(showError)
                              .deleteSelectedStudents(selectedClasses, ${range.getRow()}, ${range.getColumn()}, ${classColumn});
            }
            function closeDialog(count) {
              alert("Успешно удалено " + count + " учеников.");
              google.script.host.close();
            }
            function showError(error) {
              alert("Ошибка: " + error.message);
            }
          </script>
        </body>
      </html>
    `
  )
    .setWidth(350)
    .setHeight(350);

  return html;
}
