function showDeleteDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert("Пожалуйста, выделите диапазон с учениками.");
    return;
  }

  const data = range.getValues();
  const headerRow = data[0];
  let classColumn = findClassColumn(headerRow);

  if (classColumn === -1) {
    classColumn = headerRow.length >= 3 ? 2 : 0;
  }

  const uniqueClasses = getUniqueClassesFromStudents(data, classColumn);

  if (uniqueClasses.length === 0) {
    SpreadsheetApp.getUi().alert("Не найдено классов для удаления.");
    return;
  }

  const html = buildDeleteDialogHtml(uniqueClasses, range, classColumn);
  SpreadsheetApp.getUi().showModalDialog(html, "Удаление учеников по классам");
}
