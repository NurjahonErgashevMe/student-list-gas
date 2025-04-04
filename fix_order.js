function updateStudentsOrder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var header = data[0] || [];
  var orderCol = getOrderColumnIndex(header);
  if (orderCol === -1) orderCol = 0;

  var updates = collectOrderUpdates(data, orderCol);

  if (updates.length > 0) {
    applyUpdates(sheet, updates);
    SpreadsheetApp.getUi().alert(
      "Порядковые номера успешно обновлены!\nОбновлено записей: " +
        updates.length
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "Все порядковые номера уже актуальны. Изменений не требуется."
    );
  }
}
