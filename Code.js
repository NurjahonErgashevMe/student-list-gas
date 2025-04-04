function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Расширения от @NurjahonErgashevMe")
    .addItem("Удалить учеников по классам", "showDeleteDialog")
    .addItem("Сортировать выбранное", "sortStudentsKeepingOtherRows")
    .addItem("Обновить порядковые номера", "updateStudentsOrder")
    .addToUi();
}
