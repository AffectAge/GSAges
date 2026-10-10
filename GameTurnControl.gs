/**
 * Change Управление!B2 to TRUE (or check its checkbox) to request one turn.
 * Default onEdit needs no setup, but Google limits it to 30 seconds.
 * Run INSTALL_GAME_TURN_EDIT_TRIGGER once for longer, authorized executions.
 */
const GAME_TURN_CONTROL_CONFIG = {
  sheetName: 'Управление',
  requestCell: 'B2',
  statusCell: 'B3',
  finishedCell: 'B4',
};

function onEdit(event) {
  if (!event || !event.source) return;
  const key = 'GAME_TURN_EDIT_TRIGGER_' + event.source.getId();
  if (PropertiesService.getScriptProperties().getProperty(key)) return;
  HANDLE_GAME_TURN_EDIT(event);
}

function HANDLE_GAME_TURN_EDIT(event) {
  if (!event || !event.range || !event.source) return;
  const edited = event.range;
  const config = GAME_TURN_CONTROL_CONFIG;
  if (edited.getSheet().getName() !== config.sheetName ||
      edited.getNumRows() !== 1 || edited.getNumColumns() !== 1 ||
      edited.getA1Notation() !== config.requestCell ||
      String(event.value).toUpperCase() !== 'TRUE') return;
  // Ignore formulas: this is an explicit human request, not a calculation.
  if (edited.getFormula()) return;

  const lock = LockService.getDocumentLock();
  // Never queue a second turn behind an already running turn.
  if (!lock || !lock.tryLock(100)) return;
  let claimed = false;
  try {
    const sheet = edited.getSheet();
    const request = sheet.getRange(config.requestCell);
    const current = request.getValue();
    if (current !== true && String(current).trim().toUpperCase() !== 'TRUE') return;
    // Claim and reset before processing. Duplicate edit events see FALSE.
    request.setValue(false);
    claimed = true;
    sheet.getRange(config.statusCell).setValue('Выполняется ход…');
    sheet.getRange(config.finishedCell).clearContent();
    SpreadsheetApp.flush();
    try {
      // Already holding the engine's document lock; do not acquire it twice.
      const result = GameTurnEngine._processUnlocked(event.source);
      sheet.getRange(config.statusCell).setValue(
        'Ход завершён. Следующий ход: ' + result.nextTurn);
    } catch (error) {
      sheet.getRange(config.statusCell).setValue(
        'Ошибка: ' + String(error && error.message ? error.message : error).slice(0, 1000));
      throw error;
    } finally {
      request.setValue(false);
      sheet.getRange(config.finishedCell).setValue(new Date());
    }
  } finally {
    try {
      if (claimed) SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

/** Creates a separate UI sheet only; never overwrites an existing sheet. */
function INITIALIZE_GAME_TURN_CONTROL() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const config = GAME_TURN_CONTROL_CONFIG;
  if (spreadsheet.getSheetByName(config.sheetName)) return;
  const sheet = spreadsheet.insertSheet(config.sheetName);
  sheet.getRange('A1:B4').setValues([
    ['Управление ходом', 'Значение'],
    ['Запустить один ход', false],
    ['Результат', 'Готов к запуску'],
    ['Время завершения', ''],
  ]);
  sheet.getRange(config.requestCell).insertCheckboxes();
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 420);
  sheet.getRange(config.finishedCell).setNumberFormat('dd.MM.yyyy HH:mm:ss');
}

/** Install once as the owner. Future deployments preserve this trigger. */
function INSTALL_GAME_TURN_EDIT_TRIGGER() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    INITIALIZE_GAME_TURN_CONTROL();
    const key = 'GAME_TURN_EDIT_TRIGGER_' + spreadsheet.getId();
    const properties = PropertiesService.getScriptProperties();
    const triggers = ScriptApp.getProjectTriggers().filter(function (trigger) {
      return trigger.getHandlerFunction() === 'HANDLE_GAME_TURN_EDIT' &&
        trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
        trigger.getTriggerSourceId() === spreadsheet.getId();
    });
    if (properties.getProperty(key) && !triggers.length) {
      throw new Error('Триггер уже зарегистрирован другим владельцем или был удалён. ' +
        'Проверьте триггеры проекта и свойство ' + key + ' перед повторной установкой.');
    }
    const trigger = triggers[0] || ScriptApp.newTrigger('HANDLE_GAME_TURN_EDIT')
      .forSpreadsheet(spreadsheet).onEdit().create();
    properties.setProperty(key, trigger.getUniqueId());
    triggers.slice(1).forEach(function (extra) { ScriptApp.deleteTrigger(extra); });
  } finally {
    lock.releaseLock();
  }
}
