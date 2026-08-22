/**
 * Google Sheets turn engine (Apps Script, V8).
 *
 * One processing cycle is always: LOAD -> SYSTEMS -> VALIDATE -> SAVE.
 * Game systems work only with ctx.data and must never call SpreadsheetApp.
 */

const GAME_ENGINE_CONFIG = {
  // Add every game range here.  Or set loadAllNamedRanges: true.
  namedRanges: [
    'NR_GAME_CORE',
    'NR_WORLD',
    'NR_JOURNAL',
    'NR_COUNTRIES',
    // 'NR_PLAYERS',
    // 'NR_CITIES',
    // 'NR_UNITS',
    // 'NR_ORDERS',
    // 'NR_RULES',
  ],

  // With true every named range in the spreadsheet is exposed as data[rangeName].
  // Keep false when the spreadsheet contains service/UI ranges that must not be saved.
  loadAllNamedRanges: false,

  // Fallback sheet for a range without an explicit sheetName.
  // Automatically discovered ranges (loadAllNamedRanges) cannot be created because
  // their names are not known until they already exist.
  autoCreateMissingRanges: true,
  autoCreateSheetName: '_GAME_DATA',
  defaultNewRange: { rows: 1, columns: 1 },
  rangeDefaults: {
    NR_GAME_CORE: {
      rows: 10,
      columns: 10,
      sheetName: '_GAME_CORE',
      initialValuesFactory: createGameCoreInitialValues,
    },
    // Increase rows before adding many provinces, units, or buildings.
    NR_WORLD: {
      rows: 2,
      columns: 100,
      sheetName: '_GAME_WORLD',
      initialValuesFactory: createWorldInitialValues,
    },
    NR_JOURNAL: { rows: 500, columns: 1, sheetName: '_GAME_JOURNAL' },
    // First row is reserved for technical country IDs, such as RUS or FRA.
    NR_COUNTRIES: { rows: 25, columns: 10, sheetName: '_GAME_COUNTRIES' },
    // Example when adding a range:
    // NR_PLAYERS: { rows: 100, columns: 1 },
  },

  // These ranges are loaded but never written by the engine (for example, rules).
  readOnlyRanges: [
    // 'NR_RULES',
  ],

  // A container is read and (when changed) written in one batch. Containers may
  // expose their columns as virtual named ranges or as country record collections.
  containers: {
    NR_GAME_CORE: {
      mode: 'COLUMN_MATRICES',
      headerRow: 0,
      dataStartRow: 1,
    },
    NR_WORLD: {
      mode: 'COLUMN_MATRICES',
      headerRow: 0,
      dataStartRow: 1,
    },
    NR_COUNTRIES: {
      mode: 'KEYED_JSON_ROWS',
      headerRow: 0,
      dataStartRow: 1,
      recordKey: 'key',
    },
  },

  gameMetaRange: 'NR_GAME_CORE',
  gameMetaSubrange: 'MAIN',
  // Zero-based position of an object such as {"turn": 45, "status": "WAITING"}.
  gameMetaCell: { row: 0, column: 0 },

  // Every cell of this range can hold one journal entry. It may have several columns.
  journalRange: 'NR_JOURNAL',
  journal: {
    defaultTTLTurns: 1,
    defaultVisibility: { type: 'PUBLIC', targets: [] },
    overflow: 'DROP_OLDEST', // DROP_OLDEST or THROW
  },

  // System messages are stored in the same NR_JOURNAL range as game messages.
  // DEBUG visibility keeps stack traces out of ordinary player-facing views.
  systemJournal: {
    turnChange: {
      enabled: true,
      category: 'SYSTEM',
      priority: 'NORMAL',
      visibility: { type: 'PUBLIC', targets: [] },
      ttlTurns: 1,
    },
    log: {
      enabled: true,
      category: 'SYSTEM',
      visibility: { type: 'DEBUG', targets: [] },
      infoTTLTurns: 1,
      warningTTLTurns: 3,
      errorTTLTurns: null,
      includeStack: true,
      maxStackLength: 2000,
    },
  },

  lockTimeoutMs: 30000,
  incrementTurn: true,

  // Register systems here. Smaller priority runs first.
  systems: [
    // { id: 'ORDERS', priority: 100, handler: processOrders },
    // { id: 'ECONOMY', priority: 500, handler: processEconomy },
  ],

  // A validator may return false, a string, or an array of error strings to abort saving.
  validators: [
    validateCoreGameState,
  ],
};

/**
 * Edit this object to extend the province schema. Missing fields are added on
 * every generator run; existing values, including 0, false, and null, stay intact.
 */
const PROVINCE_GENERATOR_CONFIG = {
  rangeName: 'NR_WORLD',
  subrange: 'PROVINCES',
  idPrefix: 'prov_',
  namePrefix: 'Провинция ',
  // When absent, the first generated province receives this ID. Every other
  // province receives it as a default neighbour, but never as its own neighbour.
  defaultNeighborId: 'prov_1',
  defaults: {
    owner: null,
    terrain: 'PLAINS',
    elevation: 0,
    radiation: 0,
    pollution: 0,
    temperature: 0,
    humidity: 0,
    area: 0,
    // Soil quality/yield potential. This is distinct from fertileLandPercent.
    soilFertility: 0,
    fertileLandPercent: 0,
    resources: {
      water: { amount: 0 },
    },
    neighbors: [],
    // Future fields may be freely added here, including nested objects:
    // infrastructure: { roads: 0, railways: 0 },
  },
};

/** Public entry point for a button, menu item, or time trigger. */
function PROCESS_TURN() {
  return GameEngine.processTurn();
}

/** Optional menu for a spreadsheet-bound Apps Script project. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Game engine')
    .addItem('Process turn', 'PROCESS_TURN')
    .addItem('Generate province data', 'GENERATE_PROVINCES')
    .addToUi();
}

/** Public entry point: fills blank/partial cells of NR_WORLD.PROVINCES with JSON data. */
function GENERATE_PROVINCES() {
  return ProvinceGenerator.generate();
}

const GameEngine = {
  processTurn: function () {
    const lock = LockService.getDocumentLock();
    lock.waitLock(GAME_ENGINE_CONFIG.lockTimeoutMs);

    try {
      return this._processTurnUnlocked();
    } finally {
      lock.releaseLock();
    }
  },

  _processTurnUnlocked: function () {
    const startedAt = new Date();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let context = null;

    try {
      const loaded = NamedRangeStorage.load(spreadsheet, GAME_ENGINE_CONFIG);
      context = createTurnContext(loaded.data, startedAt);

      // An entry created on turn 50 with ttlTurns: 1 expires when turn 51 starts.
      context.journal.expireForTurn(context.turn);
      runTurnSystems(context, GAME_ENGINE_CONFIG.systems);
      runValidators(context, GAME_ENGINE_CONFIG.validators);

      if (GAME_ENGINE_CONFIG.incrementTurn) {
        context.meta.turn = context.turn + 1;
      }
      context.meta.lastProcessedAt = new Date().toISOString();
      emitTurnChangeMessage(context);

      NamedRangeStorage.save(loaded, GAME_ENGINE_CONFIG);

      return {
        processedTurn: context.turn,
        nextTurn: context.meta.turn,
        emittedMessages: context.journal.emittedCount,
        createdRanges: loaded.createdRanges,
        durationMs: new Date().getTime() - startedAt.getTime(),
      };
    } catch (error) {
      // Game state is not saved after an error, but the diagnostic is persisted
      // separately so it is visible in the next run and in a journal UI.
      SystemJournal.recordFailure(spreadsheet, context, error);
      throw error;
    }
  },
};

/** Reads/writes matrices without losing a cell's position in its named range. */
const NamedRangeStorage = {
  load: function (spreadsheet, config, selectedNames) {
    const names = selectedNames || this._getManagedNames(spreadsheet, config);
    const createdRanges = this._createMissingRanges(spreadsheet, names, config);
    const data = Object.create(null);
    const bindings = Object.create(null);

    names.forEach(function (name) {
      const range = spreadsheet.getRangeByName(name);
      if (!range) {
        throw new Error('Named range not found: ' + name);
      }

      const values = range.getValues();
      const matrix = values.map(function (row) {
        return row.map(CellCodec.decode);
      });
      const containerDefinition = config.containers && config.containers[name];
      const binding = {
        range: range,
        rows: range.getNumRows(),
        columns: range.getNumColumns(),
        columnCount: range.getNumColumns(),
        type: containerDefinition ? 'CONTAINER' : 'MATRIX',
        matrix: matrix,
        snapshot: matrixSignature(matrix),
      };

      if (containerDefinition) {
        binding.definition = containerDefinition;
        data[name] = ContainerStorage.createView(binding, name);
      } else {
        data[name] = matrix;
      }
      bindings[name] = binding;
    });

    return { data: data, bindings: bindings, createdRanges: createdRanges };
  },

  save: function (loaded, config) {
    const readOnly = indexByValue(config.readOnlyRanges || []);

    Object.keys(loaded.bindings).forEach(function (name) {
      if (readOnly[name]) return;

      const binding = loaded.bindings[name];
      let matrix = loaded.data[name];
      if (binding.type === 'CONTAINER') {
        ContainerStorage.sync(binding, loaded.data[name], name);
        matrix = binding.matrix;
      }
      assertMatrixShape(matrix, binding.rows, binding.columns, name);

      if (matrixSignature(matrix) === binding.snapshot) return;

      const encoded = matrix.map(function (row) {
        return row.map(CellCodec.encode);
      });
      binding.range.setValues(encoded);
      binding.snapshot = matrixSignature(matrix);
    });
  },

  _getManagedNames: function (spreadsheet, config) {
    const seen = Object.create(null);
    const names = [];
    const add = function (name) {
      if (!name || seen[name]) return;
      seen[name] = true;
      names.push(name);
    };

    (config.namedRanges || []).forEach(add);
    if (config.loadAllNamedRanges) {
      spreadsheet.getNamedRanges().forEach(function (namedRange) {
        add(namedRange.getName());
      });
    }
    add(config.gameMetaRange);
    add(config.journalRange);

    return names;
  },

  _createMissingRanges: function (spreadsheet, names, config) {
    if (!config.autoCreateMissingRanges) return [];

    const created = [];
    const nextFreeRowBySheet = Object.create(null);
    const getNextFreeRow = function (sheet) {
      const key = sheet.getSheetId();
      if (nextFreeRowBySheet[key]) return nextFreeRowBySheet[key];

      let nextRow = Math.max(1, sheet.getLastRow() + 1);
      spreadsheet.getNamedRanges().forEach(function (namedRange) {
        const range = namedRange.getRange();
        if (range.getSheet().getSheetId() === key) {
          nextRow = Math.max(nextRow, range.getLastRow() + 1);
        }
      });
      nextFreeRowBySheet[key] = nextRow;
      return nextRow;
    };

    names.forEach(function (name) {
      if (spreadsheet.getRangeByName(name)) return;

      const definition = config.rangeDefaults[name] || config.defaultNewRange || {};
      const rows = Number(definition.rows || 1);
      const columns = Number(definition.columns || 1);
      if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(columns) || columns < 1) {
        throw new Error('Invalid automatic size for named range ' + name + '.');
      }

      const sheetName = definition.sheetName || config.autoCreateSheetName;
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

      const startRow = definition.startRow || getNextFreeRow(sheet);
      const startColumn = definition.startColumn || 1;
      ensureSheetCapacity(sheet, startRow + rows - 1, startColumn + columns - 1);

      const range = sheet.getRange(startRow, startColumn, rows, columns);
      spreadsheet.setNamedRange(name, range);

      const initialValues = definition.initialValuesFactory
        ? definition.initialValuesFactory(rows, columns)
        : definition.initialValues;
      if (initialValues !== undefined) {
        assertMatrixShape(initialValues, rows, columns, 'initialValues for ' + name);
        range.setValues(initialValues.map(function (row) {
          return row.map(CellCodec.encode);
        }));
      }
      const key = sheet.getSheetId();
      nextFreeRowBySheet[key] = Math.max(nextFreeRowBySheet[key] || 1, startRow + rows);
      created.push(name);
    });
    return created;
  },
};

/**
 * Converts a physical named range into virtual subranges without extra Sheets API
 * calls. COLUMN_MATRICES exposes each headed column as a one-column matrix.
 * KEYED_JSON_ROWS additionally indexes JSON records in a country column by .key.
 */
const ContainerStorage = {
  createView: function (binding, rangeName) {
    const definition = binding.definition || {};
    const headerRow = Number(definition.headerRow);
    const dataStartRow = Number(definition.dataStartRow);
    if (!Number.isInteger(headerRow) || headerRow < 0 || headerRow >= binding.rows) {
      throw new Error('Invalid headerRow for container ' + rangeName + '.');
    }
    if (!Number.isInteger(dataStartRow) || dataStartRow <= headerRow || dataStartRow >= binding.rows) {
      throw new Error('Invalid dataStartRow for container ' + rangeName + '.');
    }

    binding.headerRow = headerRow;
    binding.dataStartRow = dataStartRow;
    binding.columnBindings = Object.create(null);

    const mode = definition.mode || 'COLUMN_MATRICES';
    if (mode === 'COLUMN_MATRICES') {
      return this._createColumnMatrixView(binding, rangeName);
    }
    if (mode === 'KEYED_JSON_ROWS') {
      return this._createKeyedRecordView(binding, rangeName);
    }
    throw new Error('Unknown container mode ' + mode + ' for ' + rangeName + '.');
  },

  sync: function (binding, view, rangeName) {
    const mode = binding.definition.mode || 'COLUMN_MATRICES';
    if (mode === 'COLUMN_MATRICES') {
      this._syncColumnMatrixView(binding, view, rangeName);
      return;
    }
    if (mode === 'KEYED_JSON_ROWS') {
      this._syncKeyedRecordView(binding, view, rangeName);
      return;
    }
    throw new Error('Unknown container mode ' + mode + ' for ' + rangeName + '.');
  },

  _createColumnMatrixView: function (binding, rangeName) {
    const view = Object.create(null);
    for (let column = 0; column < binding.columnCount; column += 1) {
      const key = readContainerKey(binding.matrix[binding.headerRow][column]);
      if (!key) continue;
      this._assertUniqueColumn(binding, key, rangeName);

      const columnMatrix = [];
      for (let row = binding.dataStartRow; row < binding.rows; row += 1) {
        columnMatrix.push([binding.matrix[row][column]]);
      }
      view[key] = columnMatrix;
      binding.columnBindings[key] = { column: column };
    }
    return view;
  },

  _syncColumnMatrixView: function (binding, view, rangeName) {
    assertContainerView(view, rangeName);
    const dataRows = binding.rows - binding.dataStartRow;
    const self = this;

    Object.keys(view).forEach(function (key) {
      let columnBinding = binding.columnBindings[key];
      if (!columnBinding) {
        const column = self._findEmptyColumn(binding);
        if (column === -1) {
          throw new Error('No blank column is available for subrange ' + key + ' in ' + rangeName + '.');
        }
        binding.matrix[binding.headerRow][column] = key;
        columnBinding = { column: column };
        binding.columnBindings[key] = columnBinding;
      }

      const columnMatrix = view[key];
      assertMatrixShape(columnMatrix, dataRows, 1, rangeName + '.' + key);
      for (let index = 0; index < dataRows; index += 1) {
        binding.matrix[binding.dataStartRow + index][columnBinding.column] = columnMatrix[index][0];
      }
    });
  },

  _createKeyedRecordView: function (binding, rangeName) {
    const view = Object.create(null);
    const recordKeyField = binding.definition.recordKey || 'key';
    binding.recordKeyField = recordKeyField;

    for (let column = 0; column < binding.columnCount; column += 1) {
      const countryId = readContainerKey(binding.matrix[binding.headerRow][column]);
      if (!countryId) continue;
      this._assertUniqueColumn(binding, countryId, rangeName);

      const country = Object.create(null);
      const columnBinding = { column: column, records: Object.create(null) };
      binding.columnBindings[countryId] = columnBinding;
      view[countryId] = country;

      for (let row = binding.dataStartRow; row < binding.rows; row += 1) {
        const record = binding.matrix[row][column];
        if (record === null || record === undefined) continue;
        const recordId = this._validateRecord(record, recordKeyField, rangeName, countryId, row);
        if (columnBinding.records[recordId]) {
          throw new Error('Duplicate record key ' + recordId + ' in ' + rangeName + '.' + countryId + '.');
        }
        columnBinding.records[recordId] = { row: row };
        this._defineRecordProperty(country, recordId, binding, column, row);
      }
    }
    return view;
  },

  _syncKeyedRecordView: function (binding, view, rangeName) {
    assertContainerView(view, rangeName);
    const self = this;

    Object.keys(view).forEach(function (countryId) {
      let columnBinding = binding.columnBindings[countryId];
      const country = view[countryId];
      if (!isPlainObject(country)) {
        throw new Error(rangeName + '.' + countryId + ' must be an object of JSON records.');
      }

      if (!columnBinding) {
        const column = self._findEmptyColumn(binding);
        if (column === -1) {
          throw new Error('No blank column is available for country ' + countryId + ' in ' + rangeName + '.');
        }
        binding.matrix[binding.headerRow][column] = countryId;
        columnBinding = { column: column, records: Object.create(null) };
        binding.columnBindings[countryId] = columnBinding;
      }

      Object.keys(columnBinding.records).forEach(function (recordId) {
        if (!Object.prototype.hasOwnProperty.call(country, recordId)) {
          binding.matrix[columnBinding.records[recordId].row][columnBinding.column] = null;
          delete columnBinding.records[recordId];
        }
      });

      Object.keys(country).forEach(function (recordId) {
        let recordBinding = columnBinding.records[recordId];
        if (!recordBinding) {
          const row = self._findEmptyRow(binding, columnBinding.column);
          if (row === -1) {
            throw new Error('No blank row is available for ' + recordId + ' in ' + rangeName + '.' + countryId + '.');
          }
          recordBinding = { row: row };
          columnBinding.records[recordId] = recordBinding;
        }

        const normalized = self._normalizeRecord(country[recordId], binding.recordKeyField, recordId, rangeName, countryId);
        binding.matrix[recordBinding.row][columnBinding.column] = normalized;
        self._defineRecordProperty(country, recordId, binding, columnBinding.column, recordBinding.row);
      });
    });
  },

  _defineRecordProperty: function (country, recordId, binding, column, row) {
    Object.defineProperty(country, recordId, {
      enumerable: true,
      configurable: true,
      get: function () {
        return binding.matrix[row][column];
      },
      set: function (value) {
        binding.matrix[row][column] = ContainerStorage._normalizeRecord(
          value,
          binding.recordKeyField,
          recordId,
          'container',
          'record'
        );
      },
    });
  },

  _normalizeRecord: function (record, recordKeyField, recordId, rangeName, countryId) {
    if (!isPlainObject(record)) {
      throw new Error(rangeName + '.' + countryId + '.' + recordId + ' must be a JSON object.');
    }
    const existingId = readContainerKey(record[recordKeyField]);
    if (existingId && existingId !== recordId) {
      throw new Error('Record key mismatch in ' + rangeName + '.' + countryId + ': ' + existingId + ' != ' + recordId + '.');
    }
    record[recordKeyField] = recordId;
    return record;
  },

  _validateRecord: function (record, recordKeyField, rangeName, countryId, row) {
    if (!isPlainObject(record)) {
      throw new Error(rangeName + '.' + countryId + ' row ' + row + ' must contain a JSON object.');
    }
    const recordId = readContainerKey(record[recordKeyField]);
    if (!recordId) {
      throw new Error(rangeName + '.' + countryId + ' row ' + row + ' has no ' + recordKeyField + '.');
    }
    return recordId;
  },

  _assertUniqueColumn: function (binding, key, rangeName) {
    if (binding.columnBindings[key]) {
      throw new Error('Duplicate technical header ' + key + ' in ' + rangeName + '.');
    }
  },

  _findEmptyColumn: function (binding) {
    for (let column = 0; column < binding.columnCount; column += 1) {
      if (readContainerKey(binding.matrix[binding.headerRow][column])) continue;
      let isEmpty = true;
      for (let row = binding.dataStartRow; row < binding.rows; row += 1) {
        if (binding.matrix[row][column] !== null && binding.matrix[row][column] !== undefined) {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) return column;
    }
    return -1;
  },

  _findEmptyRow: function (binding, column) {
    for (let row = binding.dataStartRow; row < binding.rows; row += 1) {
      if (binding.matrix[row][column] === null || binding.matrix[row][column] === undefined) {
        return row;
      }
    }
    return -1;
  },
};

function readContainerKey(value) {
  return typeof value === 'string' && value.trim() !== '' ? value.trim() : null;
}

function assertContainerView(view, rangeName) {
  if (!isPlainObject(view)) {
    throw new Error('Container ' + rangeName + ' was replaced with a non-object value.');
  }
}

function matrixSignature(matrix) {
  return JSON.stringify(matrix);
}

function createBlankMatrix(rows, columns) {
  const matrix = [];
  for (let row = 0; row < rows; row += 1) {
    const values = [];
    for (let column = 0; column < columns; column += 1) values.push('');
    matrix.push(values);
  }
  return matrix;
}

function createGameCoreInitialValues(rows, columns) {
  const matrix = createBlankMatrix(rows, columns);
  matrix[0][0] = 'MAIN';
  matrix[1][0] = { turn: 1, status: 'WAITING' };
  return matrix;
}

function createWorldInitialValues(rows, columns) {
  const matrix = createBlankMatrix(rows, columns);
  matrix[0][0] = 'PROVINCES';
  return matrix;
}

function createEmptyColumnMatrix(rows) {
  const matrix = [];
  for (let row = 0; row < rows; row += 1) matrix.push([null]);
  return matrix;
}

const ProvinceGenerator = {
  generate: function () {
    const lock = LockService.getDocumentLock();
    lock.waitLock(GAME_ENGINE_CONFIG.lockTimeoutMs);

    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const loaded = NamedRangeStorage.load(
        spreadsheet,
        GAME_ENGINE_CONFIG,
        [PROVINCE_GENERATOR_CONFIG.rangeName]
      );
      const container = loaded.data[PROVINCE_GENERATOR_CONFIG.rangeName];
      let matrix = container[PROVINCE_GENERATOR_CONFIG.subrange];
      if (!matrix) {
        const binding = loaded.bindings[PROVINCE_GENERATOR_CONFIG.rangeName];
        matrix = createEmptyColumnMatrix(binding.rows - binding.dataStartRow);
        container[PROVINCE_GENERATOR_CONFIG.subrange] = matrix;
      }
      const report = this._fillMissingProvinceData(matrix);
      NamedRangeStorage.save(loaded, GAME_ENGINE_CONFIG);

      writeServerLog(
        '[INFO][PROVINCE_GENERATOR] Created: ' + report.created + ', updated: ' + report.updated + '.'
      );
      return report;
    } finally {
      lock.releaseLock();
    }
  },

  _fillMissingProvinceData: function (matrix) {
    const config = PROVINCE_GENERATOR_CONFIG;
    const records = [];
    const usedIds = Object.create(null);

    GameHelpers.forEachCell(matrix, function (value, row, column) {
      let province;
      const wasBlank = value === null || value === undefined;
      if (wasBlank) {
        province = {};
      } else if (isPlainObject(value)) {
        province = value;
      } else {
        throw new Error(
          config.rangeName + '[' + row + '][' + column + '] must be blank or contain a JSON object.'
        );
      }

      if (hasProvinceId(province)) {
        if (usedIds[province.id]) {
          throw new Error('Duplicate province id: ' + province.id + '.');
        }
        usedIds[province.id] = true;
      }
      records.push({ province: province, row: row, column: column, wasBlank: wasBlank });
    });

    let needsDefaultNeighbor = !usedIds[config.defaultNeighborId];
    let nextNumber = 1;
    const getNextId = function () {
      let candidate;
      do {
        candidate = config.idPrefix + nextNumber;
        nextNumber += 1;
      } while (usedIds[candidate] || (needsDefaultNeighbor && candidate === config.defaultNeighborId));
      usedIds[candidate] = true;
      return candidate;
    };

    let created = 0;
    let updated = 0;
    records.forEach(function (record) {
      const province = record.province;
      const before = JSON.stringify(province);

      if (!hasProvinceId(province)) {
        province.id = needsDefaultNeighbor ? config.defaultNeighborId : getNextId();
        usedIds[province.id] = true;
        needsDefaultNeighbor = false;
      }
      if (!hasNonEmptyText(province.name)) {
        province.name = makeProvinceName(province.id, config);
      }

      fillMissingProperties(province, config.defaults);
      ensureDefaultNeighbor(province, config.defaultNeighborId);

      if (record.wasBlank) created += 1;
      else if (JSON.stringify(province) !== before) updated += 1;
      matrix[record.row][record.column] = province;
    });

    if (needsDefaultNeighbor) {
      throw new Error(
        'Default neighbour ' + config.defaultNeighborId +
        ' is missing. Add a blank province cell or create that province first.'
      );
    }

    return { created: created, updated: updated, total: records.length };
  },
};

function hasProvinceId(province) {
  return typeof province.id === 'string' && province.id.trim() !== '';
}

function hasNonEmptyText(value) {
  return typeof value === 'string' && value.trim() !== '';
}

function makeProvinceName(id, config) {
  const suffix = id.indexOf(config.idPrefix) === 0 ? id.slice(config.idPrefix.length) : id;
  return config.namePrefix + suffix;
}

function ensureDefaultNeighbor(province, defaultNeighborId) {
  if (!Array.isArray(province.neighbors)) {
    throw new Error('Province ' + province.id + ' has a non-array neighbors field.');
  }
  if (province.id !== defaultNeighborId && province.neighbors.indexOf(defaultNeighborId) === -1) {
    province.neighbors.push(defaultNeighborId);
  }
}

function fillMissingProperties(target, defaults) {
  Object.keys(defaults).forEach(function (key) {
    const defaultValue = defaults[key];
    if (!Object.prototype.hasOwnProperty.call(target, key) || target[key] === undefined) {
      target[key] = cloneJsonValue(defaultValue);
    } else if (isPlainObject(target[key]) && isPlainObject(defaultValue)) {
      fillMissingProperties(target[key], defaultValue);
    }
  });
  return target;
}

function cloneJsonValue(value) {
  return JSON.parse(JSON.stringify(value));
}

function isPlainObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

/** Converts only JSON objects/arrays. Text such as "Russia" remains plain text. */
const CellCodec = {
  decode: function (value) {
    if (value === '' || value === null || value === undefined) return null;
    if (typeof value !== 'string') return value;

    const trimmed = value.trim();
    const looksLikeJson =
      (trimmed.charAt(0) === '{' && trimmed.charAt(trimmed.length - 1) === '}') ||
      (trimmed.charAt(0) === '[' && trimmed.charAt(trimmed.length - 1) === ']');

    if (!looksLikeJson) return value;
    try {
      return JSON.parse(trimmed);
    } catch (error) {
      // Invalid JSON is still legitimate text, so do not destroy it.
      return value;
    }
  },

  encode: function (value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value;
    if (typeof value === 'object') return JSON.stringify(value);
    return value;
  },
};

function createTurnContext(data, startedAt) {
  const meta = getGameMeta(data);
  const turn = Number(meta.turn);
  if (!Number.isInteger(turn) || turn < 0) {
    throw new Error('NR_GAME_CORE.MAIN.turn must be a non-negative integer.');
  }

  const context = {
    data: data,
    meta: meta,
    turn: turn,
    startedAt: startedAt,
    runtime: Object.create(null), // Temporary, never persisted state for systems.
    helpers: GameHelpers,
    journal: new GameJournal(data[GAME_ENGINE_CONFIG.journalRange], turn, GAME_ENGINE_CONFIG.journal),
  };
  context.log = new GameLog(context, GAME_ENGINE_CONFIG.systemJournal.log);
  return context;
}

function getGameMeta(data) {
  const rangeName = GAME_ENGINE_CONFIG.gameMetaRange;
  const container = data[rangeName];
  const subrange = GAME_ENGINE_CONFIG.gameMetaSubrange;
  const matrix = subrange ? container && container[subrange] : container;
  const position = GAME_ENGINE_CONFIG.gameMetaCell;

  if (!matrix || !matrix[position.row] || matrix[position.row][position.column] === undefined) {
    throw new Error('Game meta cell is outside ' + rangeName + (subrange ? '.' + subrange : '') + '.');
  }

  const meta = matrix[position.row][position.column];
  if (!meta || typeof meta !== 'object' || Array.isArray(meta) || meta instanceof Date) {
    throw new Error(
      rangeName + (subrange ? '.' + subrange : '') +
      ' at [' + position.row + ',' + position.column + '] must contain a JSON object, for example {"turn": 1}.'
    );
  }
  return meta;
}

function emitTurnChangeMessage(context) {
  const settings = GAME_ENGINE_CONFIG.systemJournal.turnChange || {};
  if (settings.enabled === false) return;

  const previousTurn = context.turn;
  const nextTurn = context.meta.turn;
  const changed = previousTurn !== nextTurn;
  context.journal.emit({
    category: settings.category || 'SYSTEM',
    priority: settings.priority || 'NORMAL',
    visibility: settings.visibility || { type: 'PUBLIC', targets: [] },
    message: changed
      ? 'Ход ' + previousTurn + ' завершён. Начат ход ' + nextTurn + '.'
      : 'Ход ' + previousTurn + ' обработан.',
    ttlTurns: Object.prototype.hasOwnProperty.call(settings, 'ttlTurns') ? settings.ttlTurns : 1,
    payload: {
      type: 'TURN_PROCESSED',
      previousTurn: previousTurn,
      nextTurn: nextTurn,
    },
  });
}

/** Use ctx.log inside a system for non-fatal server-console style diagnostics. */
function GameLog(context, config) {
  this.context = context;
  this.config = config || {};
}

GameLog.prototype.info = function (message, payload) {
  return this._emit('INFO', message, payload);
};

GameLog.prototype.warn = function (message, payload) {
  return this._emit('WARNING', message, payload);
};

GameLog.prototype.error = function (message, payload) {
  return this._emit('ERROR', message, payload);
};

GameLog.prototype._emit = function (level, message, payload) {
  if (this.config.enabled === false) return null;

  const source = this.context.runtime.currentSystem || 'ENGINE';
  const ttlByLevel = {
    INFO: this.config.infoTTLTurns,
    WARNING: this.config.warningTTLTurns,
    ERROR: this.config.errorTTLTurns,
  };
  const priorityByLevel = { INFO: 'LOW', WARNING: 'HIGH', ERROR: 'CRITICAL' };
  const text = '[' + level + '][' + source + '] ' + String(message);
  const ttlTurns = ttlByLevel[level] === undefined ? null : ttlByLevel[level];

  writeServerLog(text);
  return this.context.journal.emit({
    category: this.config.category || 'SYSTEM',
    priority: priorityByLevel[level],
    visibility: this.config.visibility || { type: 'DEBUG', targets: [] },
    message: text,
    ttlTurns: ttlTurns,
    payload: {
      type: 'SYSTEM_LOG',
      level: level,
      source: source,
      details: toJournalSafeValue(payload),
    },
  });
};

/** Persists fatal errors without saving partial game-state mutations. */
const SystemJournal = {
  recordFailure: function (spreadsheet, context, error) {
    try {
      const journalRange = spreadsheet.getRangeByName(GAME_ENGINE_CONFIG.journalRange);
      if (!journalRange) {
        writeServerLog('[ERROR][ENGINE] Journal range is unavailable: ' + describeError(error).message);
        return;
      }

      const turn = context ? context.turn : this._readCurrentTurn(spreadsheet);
      const values = journalRange.getValues();
      const matrix = values.map(function (row) {
        return row.map(CellCodec.decode);
      });
      const journal = new GameJournal(matrix, turn, GAME_ENGINE_CONFIG.journal);
      const details = describeError(error, GAME_ENGINE_CONFIG.systemJournal.log);
      const source = context && context.runtime.currentSystem ? context.runtime.currentSystem : 'ENGINE';
      const logConfig = GAME_ENGINE_CONFIG.systemJournal.log || {};
      const text = '[ERROR][' + source + '] ' + details.name + ': ' + details.message;

      writeServerLog(text);
      journal.emit({
        category: logConfig.category || 'SYSTEM',
        priority: 'CRITICAL',
        visibility: logConfig.visibility || { type: 'DEBUG', targets: [] },
        message: text,
        ttlTurns: Object.prototype.hasOwnProperty.call(logConfig, 'errorTTLTurns')
          ? logConfig.errorTTLTurns
          : null,
        payload: {
          type: 'ENGINE_FAILURE',
          level: 'ERROR',
          source: source,
          error: details,
        },
      });

      journalRange.setValues(matrix.map(function (row) {
        return row.map(CellCodec.encode);
      }));
    } catch (journalError) {
      // Never hide the original error behind a diagnostics failure.
      writeServerLog('[ERROR][ENGINE] Could not save diagnostic: ' + describeError(journalError).message);
    }
  },

  _readCurrentTurn: function (spreadsheet) {
    const metaRange = spreadsheet.getRangeByName(GAME_ENGINE_CONFIG.gameMetaRange);
    if (!metaRange) return 0;
    const position = GAME_ENGINE_CONFIG.gameMetaCell;
    const values = metaRange.getValues();
    let raw;
    if (GAME_ENGINE_CONFIG.gameMetaSubrange) {
      const definition = GAME_ENGINE_CONFIG.containers[GAME_ENGINE_CONFIG.gameMetaRange];
      const headerRow = definition && definition.headerRow;
      const dataStartRow = definition && definition.dataStartRow;
      const headers = values[headerRow] || [];
      const column = headers.map(readContainerKey).indexOf(GAME_ENGINE_CONFIG.gameMetaSubrange);
      raw = column === -1 || !values[dataStartRow + position.row]
        ? null
        : values[dataStartRow + position.row][column];
    } else {
      raw = values[position.row] && values[position.row][position.column];
    }
    const meta = CellCodec.decode(raw);
    return meta && Number.isInteger(meta.turn) ? meta.turn : 0;
  },
};

function describeError(error, config) {
  const value = error || new Error('Unknown error');
  const name = String(value.name || 'Error');
  const message = String(value.message || value);
  const includeStack = !config || config.includeStack !== false;
  const maxStackLength = config && Number.isInteger(config.maxStackLength) ? config.maxStackLength : 2000;
  const stack = includeStack && value.stack ? String(value.stack).slice(0, maxStackLength) : null;
  return { name: name, message: message, stack: stack };
}

function toJournalSafeValue(value) {
  if (value === undefined) return null;
  if (value instanceof Error) return describeError(value, GAME_ENGINE_CONFIG.systemJournal.log);
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (error) {
    return String(value);
  }
}

function writeServerLog(message) {
  if (typeof Logger !== 'undefined' && Logger.log) {
    Logger.log(message);
  } else if (typeof console !== 'undefined' && console.log) {
    console.log(message);
  }
}

/**
 * Journal entries live directly in the decoded NR_JOURNAL matrix.
 * Existing messages never move: expiration clears their own cell and new messages
 * use a blank cell (or replace the oldest one only when the range is full).
 */
function GameJournal(matrix, currentTurn, config) {
  if (!Array.isArray(matrix) || !matrix.length || !Array.isArray(matrix[0])) {
    throw new Error('Journal range must contain at least one cell.');
  }
  this.matrix = matrix;
  this.currentTurn = currentTurn;
  this.config = config || {};
  this.emittedCount = 0;
}

GameJournal.prototype.emit = function (input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw new Error('journal.emit expects a message object.');
  }
  if (input.message === null || input.message === undefined || String(input.message).trim() === '') {
    throw new Error('journal.emit requires a non-empty message.');
  }

  const ttlTurns = this._resolveTTL(input);
  const slot = this._findEmptySlot() || this._resolveOverflowSlot();
  const entry = {
    createdTurn: this.currentTurn,
    country: input.country || null,
    category: String(input.category || 'GENERAL').toUpperCase(),
    priority: String(input.priority || 'NORMAL').toUpperCase(),
    visibility: this._normalizeVisibility(input.visibility),
    message: String(input.message),
    ttlTurns: ttlTurns,
    expiresAtTurn: ttlTurns === null ? null : this.currentTurn + ttlTurns,
  };

  // Temporary records do not need an ID. A permanent record gets a compact ID
  // only when it should later be addressable through journal.remove(id).
  if (typeof input.id === 'string' && input.id) {
    entry.id = input.id;
  } else if (ttlTurns === null && input.removable !== false) {
    entry.id = this._makeCompactId(slot);
  }

  // Optional structured details for a UI, without prescribing a game schema.
  if (input.payload !== undefined) entry.payload = input.payload;

  this.matrix[slot.row][slot.column] = entry;
  this.emittedCount += 1;
  return entry.id || null;
};

GameJournal.prototype.remove = function (messageId) {
  if (typeof messageId !== 'string' || !messageId) return false;
  const slot = this._findById(messageId);
  if (!slot) return false;
  this.matrix[slot.row][slot.column] = null;
  return true;
};

GameJournal.prototype.expireForTurn = function (turn) {
  let removed = 0;
  this._eachEntry(function (entry, row, column) {
    if (Number.isInteger(entry.expiresAtTurn) && entry.expiresAtTurn <= turn) {
      this.matrix[row][column] = null;
      removed += 1;
    }
  }.bind(this));
  return removed;
};

GameJournal.prototype.listVisibleTo = function (viewer) {
  const result = [];
  this._eachEntry(function (entry) {
    if (isEntryVisibleTo(entry, viewer || {})) result.push(entry);
  });
  return result;
};

GameJournal.prototype._resolveTTL = function (input) {
  const hasValue = Object.prototype.hasOwnProperty.call(input, 'ttlTurns');
  const ttl = hasValue ? input.ttlTurns : this.config.defaultTTLTurns;
  if (ttl === null) return null;
  if (!Number.isInteger(ttl) || ttl < 1) {
    throw new Error('ttlTurns must be a positive integer or null for a permanent message.');
  }
  return ttl;
};

GameJournal.prototype._normalizeVisibility = function (value) {
  const source = value || this.config.defaultVisibility || { type: 'PUBLIC', targets: [] };
  const targets = Array.isArray(source.targets) ? source.targets : [];
  return {
    type: String(source.type || 'PUBLIC').toUpperCase(),
    targets: targets.map(String),
  };
};

GameJournal.prototype._findEmptySlot = function () {
  for (let row = 0; row < this.matrix.length; row += 1) {
    for (let column = 0; column < this.matrix[row].length; column += 1) {
      if (this.matrix[row][column] === null || this.matrix[row][column] === undefined) {
        return { row: row, column: column };
      }
    }
  }
  return null;
};

GameJournal.prototype._resolveOverflowSlot = function () {
  if (String(this.config.overflow || 'DROP_OLDEST').toUpperCase() === 'THROW') {
    throw new Error('Journal range is full. Add cells to ' + GAME_ENGINE_CONFIG.journalRange + '.');
  }

  let oldest = null;
  this._eachEntry(function (entry, row, column) {
    if (!oldest || compareJournalAge(entry, oldest.entry) < 0) {
      oldest = { entry: entry, row: row, column: column };
    }
  });
  if (!oldest) {
    throw new Error('Journal range has no usable cell.');
  }
  return oldest;
};

GameJournal.prototype._findById = function (messageId) {
  let found = null;
  this._eachEntry(function (entry, row, column) {
    if (entry.id === messageId) found = { row: row, column: column };
  });
  return found;
};

GameJournal.prototype._makeCompactId = function (slot) {
  let flatIndex = 0;
  for (let row = 0; row < slot.row; row += 1) {
    flatIndex += this.matrix[row].length;
  }
  flatIndex += slot.column + 1;
  return 'm' + this.currentTurn + '_' + flatIndex;
};

GameJournal.prototype._eachEntry = function (callback) {
  for (let row = 0; row < this.matrix.length; row += 1) {
    for (let column = 0; column < this.matrix[row].length; column += 1) {
      const value = this.matrix[row][column];
      if (isJournalEntry(value)) {
        callback(value, row, column);
      }
    }
  }
};

function compareJournalAge(left, right) {
  const leftTurn = Number.isInteger(left.createdTurn) ? left.createdTurn : Number.MAX_SAFE_INTEGER;
  const rightTurn = Number.isInteger(right.createdTurn) ? right.createdTurn : Number.MAX_SAFE_INTEGER;
  return leftTurn - rightTurn;
}

function isJournalEntry(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  // `id` keeps old journal records compatible; new temporary messages have no ID.
  return Boolean(value.id) ||
    (typeof value.message === 'string' && Number.isInteger(value.createdTurn));
}

function isEntryVisibleTo(entry, viewer) {
  const visibility = entry.visibility || { type: 'PUBLIC', targets: [] };
  const type = String(visibility.type || 'PUBLIC').toUpperCase();
  const targets = visibility.targets || [];
  if (type === 'PUBLIC') return true;
  if (type === 'COUNTRY' || type === 'COUNTRIES') return targets.indexOf(String(viewer.country || '')) !== -1;
  if (type === 'PLAYER' || type === 'PLAYERS') return targets.indexOf(String(viewer.player || '')) !== -1;
  if (type === 'DEBUG') return Boolean(viewer.debug);
  return false; // Unknown types stay private until a rule for them is implemented.
}

function runTurnSystems(context, systems) {
  (systems || [])
    .map(function (system, index) {
      return { system: system, index: index };
    })
    .sort(function (left, right) {
      const priorityDifference = Number(left.system.priority || 0) - Number(right.system.priority || 0);
      return priorityDifference || left.index - right.index;
    })
    .forEach(function (item) {
      const system = item.system;
      if (!system || typeof system.handler !== 'function') {
        throw new Error('Every turn system needs a handler function.');
      }
      context.runtime.currentSystem = system.id || ('SYSTEM_' + item.index);
      system.handler(context);
    });
  delete context.runtime.currentSystem;
}

function runValidators(context, validators) {
  (validators || []).forEach(function (validator, index) {
    if (typeof validator !== 'function') {
      throw new Error('Validator at index ' + index + ' is not a function.');
    }
    const result = validator(context);
    const errors = result === false ? ['Validation failed.'] :
      typeof result === 'string' ? [result] : Array.isArray(result) ? result : [];
    if (errors.length) {
      throw new Error('Turn ' + context.turn + ' was not saved: ' + errors.join(' | '));
    }
  });
}

/** Basic invariant; add game-specific validators in GAME_ENGINE_CONFIG.validators. */
function validateCoreGameState(context) {
  if (context.meta.turn !== context.turn) {
    return 'NR_GAME_CORE.MAIN.turn was changed by a system. Change it only through engine configuration.';
  }
  return true;
}

const GameHelpers = {
  /** Iterates all cells, including nulls, while preserving matrix coordinates. */
  forEachCell: function (matrix, callback) {
    matrix.forEach(function (row, rowIndex) {
      row.forEach(function (value, columnIndex) {
        callback(value, rowIndex, columnIndex);
      });
    });
  },

  /** Finds an object by .id in any cell of a named-range matrix. */
  findById: function (matrix, id) {
    let found = null;
    this.forEachCell(matrix, function (value, row, column) {
      if (!found && value && typeof value === 'object' && value.id === id) {
        found = { value: value, row: row, column: column };
      }
    });
    return found;
  },
};

function assertMatrixShape(matrix, expectedRows, expectedColumns, rangeName) {
  if (!Array.isArray(matrix) || matrix.length !== expectedRows) {
    throw new Error('The number of rows in ' + rangeName + ' was changed in memory.');
  }
  matrix.forEach(function (row, index) {
    if (!Array.isArray(row) || row.length !== expectedColumns) {
      throw new Error('The number of columns in row ' + index + ' of ' + rangeName + ' was changed in memory.');
    }
  });
}

function ensureSheetCapacity(sheet, requiredLastRow, requiredLastColumn) {
  const missingRows = requiredLastRow - sheet.getMaxRows();
  const missingColumns = requiredLastColumn - sheet.getMaxColumns();
  if (missingRows > 0) sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
  if (missingColumns > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), missingColumns);
}

function indexByValue(values) {
  const index = Object.create(null);
  values.forEach(function (value) { index[value] = true; });
  return index;
}

/*
 * Example system. Register it in GAME_ENGINE_CONFIG.systems to use it.
 *
 * function processEconomy(ctx) {
 *   ctx.helpers.forEachCell(ctx.data.NR_PLAYERS, function (player) {
 *     if (!player) return;
 *     player.gold = Number(player.gold || 0) + 100;
 *   });
 *
 *   ctx.journal.emit({
 *     country: 'RUS',
 *     category: 'ECONOMY',
 *     priority: 'NORMAL',
 *     visibility: { type: 'COUNTRY', targets: ['RUS'] },
 *     message: 'Государственный бюджет получил 100 золота.',
 *     ttlTurns: 1,
 *   });
 * }
 */
