/**
 * Universal order infrastructure. OrderSystem owns the lifecycle, batching,
 * dependencies and routing; game systems own validation and execution rules.
 * Core processing uses only ctx.data and never calls SpreadsheetApp.
 */

const ORDER_SYSTEM_CONFIG = {
  rangeName: 'NR_ORDERS',
  activeSubrange: 'ACTIVE',
  historySubrange: 'HISTORY',
  countryBooksRange: 'NR_GAME_CORE',
  countryBooksSubrange: 'COUNTRY_BOOKS',
  maxOrdersPerCell: 20,
  category: 'ORDERS',
};

const OrderSystem = {
  process: function (ctx) {
    const active = this._getSubrange(ctx, ORDER_SYSTEM_CONFIG.rangeName, ORDER_SYSTEM_CONFIG.activeSubrange);
    const history = this._getSubrange(ctx, ORDER_SYSTEM_CONFIG.rangeName, ORDER_SYSTEM_CONFIG.historySubrange);
    if (!active || !history) {
      throw new Error('NR_ORDERS must contain ACTIVE and HISTORY columns.');
    }

    const activeOrders = this._readBatches(active, 'NR_ORDERS.ACTIVE');
    const historyOrders = this._readBatches(history, 'NR_ORDERS.HISTORY');
    const registry = this._buildRegistry();
    const historyById = this._indexById(historyOrders, 'NR_ORDERS.HISTORY');
    // This full index lets an order wait for a dependency that appears later
    // in the same batch. `processedActiveIds` still detects a true duplicate.
    const activeById = Object.create(null);
    activeOrders.forEach(function (order) {
      if (!OrderSystem._isObject(order)) return;
      const id = OrderSystem._text(order.id);
      if (id && !activeById[id]) activeById[id] = order;
    });
    const processedActiveIds = Object.create(null);
    const completed = [];
    const retained = [];
    const report = { received: activeOrders.length, executed: 0, rejected: 0, failed: 0, queued: 0, cancelled: 0 };
    const self = this;

    activeOrders.sort(function (left, right) {
      const leftDefinition = registry[left && left.type];
      const rightDefinition = registry[right && right.type];
      const leftPriority = leftDefinition ? leftDefinition.priority : Number.MAX_SAFE_INTEGER;
      const rightPriority = rightDefinition ? rightDefinition.priority : Number.MAX_SAFE_INTEGER;
      if (leftPriority !== rightPriority) return leftPriority - rightPriority;
      return Number(left.sequence || 0) - Number(right.sequence || 0);
    });

    activeOrders.forEach(function (order) {
      if (!self._isObject(order)) {
        self._emitMalformedOrder(ctx, 'Приказ должен быть JSON-объектом.');
        report.rejected += 1;
        return;
      }
      const knownOrderId = self._text(order.id);
      const duplicateKnown = !!(knownOrderId &&
        (processedActiveIds[knownOrderId] || historyById[knownOrderId]));
      const generalError = self._validateGeneralOrder(order, processedActiveIds, historyById);
      if (generalError) {
        self._finishOrder(ctx, order, 'REJECTED', generalError, null);
        // A second copy must be removed but cannot be put into HISTORY: that
        // would create the very duplicate ID the next turn must reject.
        if (!duplicateKnown && self._text(order.id)) completed.push(order);
        report.rejected += 1;
        return;
      }
      processedActiveIds[order.id] = order;

      const definition = registry[order.type];
      if (!definition) {
        self._finishOrder(ctx, order, 'REJECTED', 'Такой вид распоряжения пока не поддерживается.', null);
        completed.push(order);
        report.rejected += 1;
        return;
      }

      order.category = definition.category;
      order.phase = definition.phase;
      order.priority = definition.priority;

      const dependency = self._checkDependencies(order, activeById, historyById);
      if (dependency.state === 'WAITING') {
        order.status = 'QUEUED';
        retained.push(order);
        report.queued += 1;
        return;
      }
      if (dependency.state === 'CANCELLED') {
        self._finishOrder(ctx, order, 'CANCELLED', dependency.message, definition);
        completed.push(order);
        report.cancelled += 1;
        return;
      }
      if (dependency.state === 'REJECTED') {
        self._finishOrder(ctx, order, 'REJECTED', dependency.message, definition);
        completed.push(order);
        report.rejected += 1;
        return;
      }

      const validation = self._readOperationResult(definition.validate(ctx, order));
      if (!validation.ok) {
        self._finishOrder(ctx, order, 'REJECTED', validation.message, definition);
        completed.push(order);
        report.rejected += 1;
        return;
      }

      const execution = self._readOperationResult(definition.execute(ctx, order));
      if (!execution.ok) {
        self._finishOrder(ctx, order, 'FAILED', execution.message, definition);
        completed.push(order);
        report.failed += 1;
        return;
      }

      self._finishOrder(ctx, order, 'EXECUTED', execution.message || 'Приказ выполнен.', definition);
      completed.push(order);
      report.executed += 1;
    });

    this._writeBatches(active, retained, 'NR_ORDERS.ACTIVE');
    this._writeBatches(history, historyOrders.concat(completed), 'NR_ORDERS.HISTORY');
    ctx.runtime.orderReport = report;
    return report;
  },

  /** Converts a player workbook entry into the authoritative central format. */
  normalizeClientOrder: function (countryId, clientOrder, turn) {
    if (!this._isObject(clientOrder)) throw new Error('Приказ игрока должен быть JSON-объектом.');
    const clientOrderId = this._text(clientOrder.clientOrderId || clientOrder.id);
    const type = this._text(clientOrder.type);
    if (!clientOrderId) throw new Error('В приказе отсутствует clientOrderId.');
    if (!type) throw new Error('В приказе ' + clientOrderId + ' отсутствует type.');
    if (!this._isObject(clientOrder.payload)) {
      throw new Error('В приказе ' + clientOrderId + ' отсутствует объект payload.');
    }

    const issuer = this._text(countryId);
    if (!issuer) throw new Error('В реестре книг не указан ID страны.');
    const dependsOn = this._normalizeDependencies(clientOrder.dependsOn, issuer);
    return {
      id: issuer + ':' + clientOrderId,
      clientOrderId: clientOrderId,
      issuer: issuer,
      type: type.toUpperCase(),
      payload: cloneJsonValue(clientOrder.payload),
      dependsOn: dependsOn,
      submittedTurn: turn,
      sequence: this._number(clientOrder.sequence, 0),
      status: 'SUBMITTED',
    };
  },

  /** Appends externally imported orders once; replayed client batches are safe. */
  importExternalOrders: function (ctx, orders) {
    if (!Array.isArray(orders) || !orders.length) return { imported: 0, duplicates: 0 };
    const active = this._getSubrange(ctx, ORDER_SYSTEM_CONFIG.rangeName, ORDER_SYSTEM_CONFIG.activeSubrange);
    const history = this._getSubrange(ctx, ORDER_SYSTEM_CONFIG.rangeName, ORDER_SYSTEM_CONFIG.historySubrange);
    if (!active || !history) throw new Error('NR_ORDERS must contain ACTIVE and HISTORY columns.');

    const activeOrders = this._readBatches(active, 'NR_ORDERS.ACTIVE');
    const historyOrders = this._readBatches(history, 'NR_ORDERS.HISTORY');
    const known = this._indexById(historyOrders, 'NR_ORDERS.HISTORY');
    let imported = 0;
    let duplicates = 0;
    activeOrders.forEach(function (order) {
      if (order && order.id) known[order.id] = true;
    });

    orders.forEach(function (order) {
      if (known[order.id]) {
        duplicates += 1;
        return;
      }
      known[order.id] = true;
      activeOrders.push(order);
      imported += 1;
    });
    this._writeBatches(active, activeOrders, 'NR_ORDERS.ACTIVE');
    return { imported: imported, duplicates: duplicates };
  },

  _buildRegistry: function () {
    const registry = Object.create(null);
    const definitions = [];
    if (typeof FactorySystem !== 'undefined' && FactorySystem.getOrderDefinitions) {
      FactorySystem.getOrderDefinitions().forEach(function (definition) { definitions.push(definition); });
    }

    definitions.forEach(function (definition) {
      if (!definition || !OrderSystem._text(definition.type) ||
          typeof definition.validate !== 'function' || typeof definition.execute !== 'function') {
        throw new Error('Некорректное определение типа приказа.');
      }
      const type = OrderSystem._text(definition.type).toUpperCase();
      if (registry[type]) throw new Error('Тип приказа зарегистрирован дважды: ' + type + '.');
      registry[type] = {
        type: type,
        category: this._text(definition.category || ORDER_SYSTEM_CONFIG.category).toUpperCase(),
        phase: this._text(definition.phase || 'TURN').toUpperCase(),
        priority: Math.max(0, Math.floor(this._number(definition.priority, 1000))),
        validate: definition.validate,
        execute: definition.execute,
      };
    }.bind(this));
    return registry;
  },

  _validateGeneralOrder: function (order, processedActiveIds, historyById) {
    const id = this._text(order.id);
    if (!id) return 'В распоряжении отсутствует служебный номер.';
    order.id = id;
    if (processedActiveIds[order.id] || historyById[order.id]) return 'Это распоряжение уже было принято канцелярией.';
    const issuer = this._text(order.issuer);
    if (!issuer) return 'Не удалось определить страну-отправителя распоряжения.';
    order.issuer = issuer;
    const type = this._text(order.type);
    if (!type) return 'В распоряжении не выбран вид действия.';
    order.type = type.toUpperCase();
    if (!this._isObject(order.payload)) return 'В распоряжении отсутствуют необходимые сведения.';
    const status = this._text(order.status || 'SUBMITTED').toUpperCase();
    if (status !== 'SUBMITTED' && status !== 'QUEUED') {
      return 'Распоряжение находится в состоянии, в котором его нельзя рассмотреть.';
    }
    order.status = status;
    order.dependsOn = this._normalizeDependencies(order.dependsOn, order.issuer);
    return null;
  },

  _checkDependencies: function (order, activeById, historyById) {
    const dependencies = order.dependsOn || [];
    for (let index = 0; index < dependencies.length; index += 1) {
      const id = dependencies[index];
      if (id === order.id) return { state: 'REJECTED', message: 'Распоряжение не может ожидать исполнения самого себя.' };
      const dependency = activeById[id] || historyById[id];
      if (!dependency) return { state: 'REJECTED', message: 'Не найдено предыдущее распоряжение, от которого зависит это решение.' };
      const status = this._text(dependency.status).toUpperCase();
      if (status === 'REJECTED' || status === 'FAILED' || status === 'CANCELLED') {
        return { state: 'CANCELLED', message: 'Предыдущее распоряжение не было исполнено.' };
      }
      if (status !== 'EXECUTED') return { state: 'WAITING' };
    }
    return { state: 'READY' };
  },

  _finishOrder: function (ctx, order, status, message, definition) {
    order.status = status;
    order.completedTurn = ctx.turn;
    order.result = { message: String(message || ''), turn: ctx.turn };
    const priority = {
      EXECUTED: 'SUCCESS',
      REJECTED: 'HIGH',
      CANCELLED: 'HIGH',
      FAILED: 'CRITICAL',
    }[status] || 'NORMAL';
    const visible = this._text(order.issuer)
      ? { type: 'COUNTRY', targets: [order.issuer] }
      : { type: 'DEBUG', targets: [] };
    ctx.journal.emit({
      country: order.issuer || null,
      category: definition ? definition.category : ORDER_SYSTEM_CONFIG.category,
      subCategory: order.type || 'GENERAL',
      priority: priority,
      visibility: visible,
      message: this._getOrderMessage(order, status, message),
      ttlTurns: status === 'FAILED' ? null : 3,
    });
  },

  _getOrderMessage: function (order, status, message) {
    const orderTitle = this._getOrderTitle(order.type);
    const labels = {
      EXECUTED: 'Канцелярия исполнила распоряжение: ' + orderTitle,
      REJECTED: 'Канцелярия отклонила распоряжение: ' + orderTitle,
      CANCELLED: 'Канцелярия отменила распоряжение: ' + orderTitle,
      FAILED: 'При исполнении распоряжения возникла неполадка: ' + orderTitle,
    };
    return (labels[status] || 'Канцелярия рассмотрела распоряжение: ' + orderTitle) +
      '. ' + String(message || 'Дополнительных сведений нет.');
  },

  _getOrderTitle: function (type) {
    return {
      BUILD_FACTORY: 'строительство фабрики',
    }[this._text(type) ? type.toUpperCase() : ''] || 'особое распоряжение';
  },

  _emitMalformedOrder: function (ctx, message) {
    ctx.journal.emit({
      category: ORDER_SYSTEM_CONFIG.category,
      subCategory: 'VALIDATION',
      priority: 'HIGH',
      visibility: { type: 'DEBUG', targets: [] },
      message: 'Канцелярия отклонила некорректно оформленное распоряжение. ' + message,
      ttlTurns: 3,
    });
  },

  _readOperationResult: function (value) {
    if (value === undefined || value === true || value === null) return { ok: true, message: null };
    if (value === false) return { ok: false, message: 'Операция отклонена обработчиком.' };
    if (typeof value === 'string') return { ok: false, message: value };
    if (this._isObject(value)) {
      return {
        ok: value.ok !== false,
        message: value.message === undefined || value.message === null ? null : String(value.message),
      };
    }
    return { ok: false, message: 'Обработчик вернул недопустимый результат.' };
  },

  _readBatches: function (matrix, label) {
    const orders = [];
    matrix.forEach(function (row, rowIndex) {
      row.forEach(function (batch, columnIndex) {
        if (batch === null || batch === undefined) return;
        if (!Array.isArray(batch)) {
          throw new Error(label + '[' + rowIndex + ',' + columnIndex + '] должен содержать JSON-массив приказов.');
        }
        if (batch.length > ORDER_SYSTEM_CONFIG.maxOrdersPerCell) {
          throw new Error(label + '[' + rowIndex + ',' + columnIndex + '] содержит больше ' + ORDER_SYSTEM_CONFIG.maxOrdersPerCell + ' приказов.');
        }
        batch.forEach(function (order) { orders.push(order); });
      });
    });
    return orders;
  },

  _writeBatches: function (matrix, orders, label) {
    const capacity = matrix.length * (matrix[0] ? matrix[0].length : 0) * ORDER_SYSTEM_CONFIG.maxOrdersPerCell;
    if (orders.length > capacity) {
      throw new Error(label + ' переполнен: доступно ' + capacity + ' приказов.');
    }
    for (let row = 0; row < matrix.length; row += 1) {
      for (let column = 0; column < matrix[row].length; column += 1) matrix[row][column] = null;
    }
    let orderIndex = 0;
    for (let row = 0; row < matrix.length && orderIndex < orders.length; row += 1) {
      for (let column = 0; column < matrix[row].length && orderIndex < orders.length; column += 1) {
        matrix[row][column] = orders.slice(orderIndex, orderIndex + ORDER_SYSTEM_CONFIG.maxOrdersPerCell);
        orderIndex += ORDER_SYSTEM_CONFIG.maxOrdersPerCell;
      }
    }
  },

  _indexById: function (orders, label) {
    const index = Object.create(null);
    orders.forEach(function (order) {
      if (!order || !OrderSystem._text(order.id)) return;
      if (index[order.id]) throw new Error('Повторяющийся ID «' + order.id + '» в ' + label + '.');
      index[order.id] = order;
    });
    return index;
  },

  _normalizeDependencies: function (value, issuer) {
    if (value === null || value === undefined || value === '') return [];
    const source = Array.isArray(value) ? value : [value];
    const result = [];
    source.forEach(function (dependency) {
      const text = OrderSystem._text(dependency);
      if (!text) throw new Error('dependsOn должен содержать ID приказа.');
      result.push(text.indexOf(':') === -1 ? issuer + ':' + text : text);
    });
    return result;
  },

  _getSubrange: function (ctx, rangeName, subrangeName) {
    const container = ctx.data[rangeName];
    return container && container[subrangeName];
  },

  _isObject: function (value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
  },

  _text: function (value) {
    return typeof value === 'string' && value.trim() !== '' ? value.trim() : null;
  },

  _number: function (value, fallback) {
    return typeof value === 'number' && isFinite(value) ? value : fallback;
  },
};

/**
 * The only cross-workbook part of the order flow. It is called before systems
 * run, injects normalized orders into ctx.data, and never belongs to a system.
 */
const CountryOrderGateway = {
  importSubmittedOrders: function (ctx) {
    const books = this._readRegisteredBooks(ctx);
    const imported = [];
    const acknowledgements = [];
    const self = this;

    books.forEach(function (book) {
      const countryId = book.countryId;
      const settings = book.settings;
      try {
        const batches = self._readCountryOutbox(settings);
        const normalizedOrders = [];
        const clientOrderIds = [];
        batches.forEach(function (clientOrder) {
          const normalized = OrderSystem.normalizeClientOrder(countryId, clientOrder, ctx.turn);
          normalizedOrders.push(normalized);
          clientOrderIds.push(normalized.clientOrderId);
        });
        imported.push.apply(imported, normalizedOrders);
        if (clientOrderIds.length) {
          acknowledgements.push({ countryId: countryId, settings: settings, clientOrderIds: clientOrderIds });
        }
      } catch (error) {
        ctx.journal.emit({
          country: countryId,
          category: ORDER_SYSTEM_CONFIG.category,
          subCategory: 'IMPORT',
          priority: 'HIGH',
          visibility: { type: 'COUNTRY', targets: [countryId] },
          message: 'Канцелярия не смогла принять новые распоряжения. Проверьте их оформление и доступ к книге страны.',
          ttlTurns: 3,
        });
        writeServerLog('[WARNING][ORDERS] ' + countryId + ': ' + String(error.message || error));
      }
    });
    const report = OrderSystem.importExternalOrders(ctx, imported);
    ctx.runtime.countryOrderAcknowledgements = acknowledgements;
    return report;
  },

  /**
   * Runs only after the central turn was fully saved. If this step fails, the
   * player order remains in the source book and replay protection handles it.
   */
  acknowledgeImportedOrders: function (ctx) {
    const acknowledgements = ctx.runtime.countryOrderAcknowledgements || [];
    const report = { countries: 0, removed: 0, failed: 0 };
    acknowledgements.forEach(function (acknowledgement) {
      try {
        const removed = CountryOrderGateway._removeAcknowledgedOrders(
          acknowledgement.settings,
          acknowledgement.clientOrderIds
        );
        report.countries += 1;
        report.removed += removed;
      } catch (error) {
        report.failed += 1;
        writeServerLog(
          '[WARNING][ORDERS_ACK] ' + acknowledgement.countryId + ': ' + String(error.message || error)
        );
      }
    });
    return report;
  },

  initializeRegisteredBooks: function (ctx) {
    const books = this._readRegisteredBooks(ctx);
    const created = [];
    books.forEach(function (book) {
      const countryId = book.countryId;
      const settings = book.settings;
      const spreadsheet = SpreadsheetApp.openById(settings.spreadsheetId);
      if (CountryOrderGateway._ensureCountryOutbox(spreadsheet, settings)) created.push(countryId);
    });
    return created;
  },

  /** Reads one JSON registration record per COUNTRY_BOOKS cell. */
  _readRegisteredBooks: function (ctx) {
    const matrix = this._getOrCreateRegistry(ctx);
    const books = [];
    const seenCountryIds = Object.create(null);
    const seenSpreadsheetIds = Object.create(null);

    matrix.forEach(function (row, rowIndex) {
      row.forEach(function (record, columnIndex) {
        if (record === null || record === undefined) return;
        if (!OrderSystem._isObject(record)) {
          throw new Error('COUNTRY_BOOKS[' + rowIndex + ',' + columnIndex + '] должен содержать JSON-объект.');
        }
        const countryId = OrderSystem._text(record.id || record.countryId);
        const spreadsheetId = OrderSystem._text(record.spreadsheetId);
        if (!countryId) {
          throw new Error('В COUNTRY_BOOKS[' + rowIndex + ',' + columnIndex + '] отсутствует id страны.');
        }
        if (!spreadsheetId) {
          throw new Error('У страны «' + countryId + '» отсутствует spreadsheetId.');
        }
        if (seenCountryIds[countryId]) throw new Error('Страна «' + countryId + '» дважды записана в COUNTRY_BOOKS.');
        if (seenSpreadsheetIds[spreadsheetId]) throw new Error('Одна книга зарегистрирована для нескольких стран.');
        seenCountryIds[countryId] = true;
        seenSpreadsheetIds[spreadsheetId] = true;
        books.push({
          countryId: countryId,
          settings: {
            spreadsheetId: spreadsheetId,
            ordersRange: OrderSystem._text(record.ordersRange) || ORDER_SYSTEM_CONFIG.rangeName,
            activeHeader: OrderSystem._text(record.activeHeader) || ORDER_SYSTEM_CONFIG.activeSubrange,
            ordersSheetName: OrderSystem._text(record.ordersSheetName) || '_ORDERS',
          },
        });
      });
    });
    return books;
  },

  _getOrCreateRegistry: function (ctx) {
    const core = ctx.data[ORDER_SYSTEM_CONFIG.countryBooksRange];
    if (!core) throw new Error('Не найден контейнер ' + ORDER_SYSTEM_CONFIG.countryBooksRange + '.');
    const existing = core[ORDER_SYSTEM_CONFIG.countryBooksSubrange];
    if (existing) return existing;

    let rows = 0;
    Object.keys(core).some(function (key) {
      const matrix = core[key];
      if (!Array.isArray(matrix)) return false;
      rows = matrix.length;
      return rows > 0;
    });
    if (!rows) throw new Error('Невозможно определить размер COUNTRY_BOOKS.');
    core[ORDER_SYSTEM_CONFIG.countryBooksSubrange] = createEmptyColumnMatrix(rows);
    return core[ORDER_SYSTEM_CONFIG.countryBooksSubrange];
  },

  _readCountryOutbox: function (settings) {
    const spreadsheet = SpreadsheetApp.openById(settings.spreadsheetId);
    const rangeName = settings.ordersRange || ORDER_SYSTEM_CONFIG.rangeName;
    const range = this._ensureCountryOutbox(spreadsheet, settings) || spreadsheet.getRangeByName(rangeName);
    const values = range.getValues().map(function (row) {
      return row.map(CellCodec.decode);
    });
    const header = String(settings.activeHeader || ORDER_SYSTEM_CONFIG.activeSubrange).trim().toUpperCase();
    const activeColumn = (values[0] || []).map(readContainerKey).indexOf(header);
    if (activeColumn === -1) throw new Error('В контейнере ' + rangeName + ' нет колонки ' + header + '.');

    const orders = [];
    for (let row = 1; row < values.length; row += 1) {
      const batch = values[row][activeColumn];
      if (batch === null || batch === undefined) continue;
      if (!Array.isArray(batch)) throw new Error('ACTIVE[' + (row - 1) + '] должен содержать JSON-массив.');
      if (batch.length > ORDER_SYSTEM_CONFIG.maxOrdersPerCell) {
        throw new Error('ACTIVE[' + (row - 1) + '] содержит больше ' + ORDER_SYSTEM_CONFIG.maxOrdersPerCell + ' приказов.');
      }
      batch.forEach(function (order) { orders.push(order); });
    }
    return orders;
  },

  _removeAcknowledgedOrders: function (settings, clientOrderIds) {
    const acknowledged = Object.create(null);
    clientOrderIds.forEach(function (clientOrderId) { acknowledged[clientOrderId] = true; });
    if (!Object.keys(acknowledged).length) return 0;

    const spreadsheet = SpreadsheetApp.openById(settings.spreadsheetId);
    const rangeName = settings.ordersRange || ORDER_SYSTEM_CONFIG.rangeName;
    const range = spreadsheet.getRangeByName(rangeName);
    if (!range) throw new Error('Контейнер ' + rangeName + ' больше не существует.');
    const values = range.getValues().map(function (row) {
      return row.map(CellCodec.decode);
    });
    const header = String(settings.activeHeader || ORDER_SYSTEM_CONFIG.activeSubrange).trim().toUpperCase();
    const activeColumn = (values[0] || []).map(readContainerKey).indexOf(header);
    if (activeColumn === -1) throw new Error('В контейнере ' + rangeName + ' нет колонки ' + header + '.');

    let removed = 0;
    let changed = false;
    for (let row = 1; row < values.length; row += 1) {
      const batch = values[row][activeColumn];
      if (!Array.isArray(batch)) continue;
      const retained = batch.filter(function (clientOrder) {
        if (!OrderSystem._isObject(clientOrder)) return true;
        const clientOrderId = OrderSystem._text(clientOrder.clientOrderId || clientOrder.id);
        if (!clientOrderId || !acknowledged[clientOrderId]) return true;
        removed += 1;
        return false;
      });
      if (retained.length === batch.length) continue;
      values[row][activeColumn] = retained.length ? retained : null;
      changed = true;
    }
    if (changed) {
      range.setValues(values.map(function (row) {
        return row.map(CellCodec.encode);
      }));
    }
    return removed;
  },

  /** Creates the player outbox on its own sheet if the workbook has none yet. */
  _ensureCountryOutbox: function (spreadsheet, settings) {
    const rangeName = settings.ordersRange || ORDER_SYSTEM_CONFIG.rangeName;
    if (spreadsheet.getRangeByName(rangeName)) return null;

    const sheetName = settings.ordersSheetName || '_ORDERS';
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
    ensureSheetCapacity(sheet, 1001, 2);
    const range = sheet.getRange(1, 1, 1001, 2);
    spreadsheet.setNamedRange(rangeName, range);
    range.setValues(createOrdersInitialValues(1001, 2).map(function (row) {
      return row.map(CellCodec.encode);
    }));
    return range;
  },
};
