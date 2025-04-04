function sortStudentsKeepingOtherRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selectedRange = sheet.getActiveRange();

  if (!selectedRange) {
    SpreadsheetApp.getUi().alert(
      "Пожалуйста, выделите диапазон для сортировки!"
    );
    return;
  }

  var data = selectedRange.getValues();
  var rowsInfo = collectRowsInfo(data);

  var students = rowsInfo.filter((r) => r.isStudent);
  var others = rowsInfo.filter((r) => !r.isStudent);

  var sortedStudents = sortStudentRows(students);
  var sortedResult = buildSortedResult(sortedStudents, others);

  selectedRange.setValues(sortedResult);

  SpreadsheetApp.getUi().alert(
    "Сортировка завершена!\n" +
      "Отсортировано учеников: " +
      sortedStudents.length +
      "\n" +
      "Строк оставлено на местах: " +
      others.length
  );
}
