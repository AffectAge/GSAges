/**
 * Standalone Google Sheets game-turn engine (Apps Script V8).
 *
 * Public functions:
 *   INITIALIZE_GAME_TURN_ENGINE()  - creates Игровой_журнал and Состояние_движка.
 *   PROCESS_GAME_TURN()            - runs one game turn.
 *
 * Every game range listed in GAME_TURN_CONFIG.dataRangeNames must have headers
 * in its first row. All reads, processors, validation, and serialization happen
 * in memory. No game range is written when loading, a processor, or validation
 * fails. Engine errors are appended to Игровой_журнал separately.
 */

const GAME_TURN_CONFIG = {
  // List only ranges owned by the game engine: do not include formula/UI ranges.
  dataRangeNames: [
    'Игровые_данные',
  ],

  journal: {
    rangeName: 'Игровой_журнал',
    sheetName: '_Игровой_журнал',
    rows: 10000, // One header + 500 reports.
    overflow: 'DROP_OLDEST', // DROP_OLDEST or THROW.
    defaultTtl: 1, // -1 means permanent.
  },

  state: {
    rangeName: 'Состояние_движка',
    sheetName: '_Состояние_движка',
  },

  lockTimeoutMs: 30000,
  maxJsonCellCharacters: 48000,

  // ctx.country('GER') takes a colour from here. A processor may override it.
  countryColors: {
    GER: '#DC2626',
    FRA: '#2563EB',
  },

  // Used only by GENERATE_EMPTY_PROVINCES(). Edit these defaults before
  // running the command if another initial province layout is needed.
  provinceTemplateGenerator: {
    rangeName: 'Игровые_данные',
    header: 'Провинции',
    firstId: 1,
    namesById: {
      1001: 'Ломбардия',
    },
    fallbackNamePrefix: 'Провинция ',
    template: {
      ownerId: 'NULL',
      areaKm2: 20000.0,
      temperatureC: 16.2,
      precipitationMm: 850,
      latitudeZone: 'Умеренная',
      season: 'Лето',
      climate: 'Умеренно-континентальный',
      terrainType: 'Суша',
      landscapeType: 'Равнина',
      hasSeaAccess: false,
      elevationM: 250,
      pollution: 20,
      radiation: 0,
      fertility: 0.5,
      resources: [
        { type: 'Железная руда', amount: 125000 },
        { type: 'Уголь', amount: 68000 },
      ],
      continent: 'Европа',
      planet: 'Земля',
      neighbors: [1000, 1002, 1003, 1045],
    },
  },

  // Processors run in ascending priority. Add game mechanics here.
  // Handler signature: function (data, ctx) { ...; return data; }
  processors: [
    // { id: 'FACTORIES', priority: 100, handler: processFactories },
    // { id: 'ECONOMY', priority: 200, handler: processEconomy },
  ],

  // A validator returns true/undefined for success, or a string / string array
  // to cancel the turn. Add game-specific invariants here.
  validators: [
    validateGameTurnData,
  ],

  // One technical GameLog row is added after every successful registered
  // processor. The measured duration includes only handler execution.
  processorTimingReports: {
    enabled: true,
    type: 'PERFORMANCE',
    category: 'ENGINE',
    priority: 'NORMAL',
    visibility: 'SYSTEM',
    ttl: 1,
  },
};

const GAME_TURN_JOURNAL_HEADERS = [
  'ID',
  'ХОД',
  'ТИП',
  'КАТЕГОРИЯ',
  'ПРИОРИТЕТ',
  'СТРАНА',
  'ВИДИМОСТЬ',
  'СРОК (ХОДОВ)',
  'СООБЩЕНИЕ',
  'ИСТОЧНИК',
];

const GAME_TURN_JOURNAL = {
  ID: 0,
  TURN: 1,
  TYPE: 2,
  CATEGORY: 3,
  PRIORITY: 4,
  COUNTRY: 5,
  VISIBILITY: 6,
  TTL: 7,
  MESSAGE: 8,
  SOURCE: 9,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Игровой движок')
    .addItem('Инициализировать журнал', 'INITIALIZE_GAME_TURN_ENGINE')
    .addItem('Заполнить пустые провинции', 'GENERATE_EMPTY_PROVINCES')
    .addItem('Обработать ход', 'PROCESS_GAME_TURN')
    .addToUi();
}

/** Creates only the engine-owned ranges. It never creates game-data ranges. */
function INITIALIZE_GAME_TURN_ENGINE() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const journal = GameTurnJournalStorage.ensure(spreadsheet, GAME_TURN_CONFIG);
  const state = GameTurnStateStorage.ensure(spreadsheet, GAME_TURN_CONFIG);
  return {
    journal: journal.getA1Notation(),
    state: state.getA1Notation(),
  };
}

/** Public entry point for a menu, a button, or a time-driven trigger. */
function PROCESS_GAME_TURN() {
  return GameTurnEngine.process();
}

/**
 * Manually fills only blank cells in Игровые_данные.Провинции from the
 * configured template. It is independent from turn processing.
 */
function GENERATE_EMPTY_PROVINCES() {
  return GameTurnProvinceGenerator.fillEmpty();
}

const GameTurnEngine = {
  process: function () {
    const lock = LockService.getDocumentLock();
    lock.waitLock(GAME_TURN_CONFIG.lockTimeoutMs);

    try {
      return this._processUnlocked(SpreadsheetApp.getActiveSpreadsheet());
    } finally {
      lock.releaseLock();
    }
  },

  _processUnlocked: function (spreadsheet) {
    const startedAt = new Date();
    let state = null;
    let turn = 0;
    let context = null;

    try {
      state = GameTurnStateStorage.load(spreadsheet, GAME_TURN_CONFIG);
      turn = state.turn;
      const journalBinding = GameTurnJournalStorage.load(spreadsheet, GAME_TURN_CONFIG);
      const loaded = GameTurnRangeStorage.loadAll(spreadsheet, GAME_TURN_CONFIG);
      context = GameTurnContext.create(loaded, journalBinding, turn, startedAt, GAME_TURN_CONFIG);

      context.journal.expire(turn);
      GameTurnProcessors.run(context, GAME_TURN_CONFIG.processors);
      GameTurnValidators.run(context, GAME_TURN_CONFIG.validators);

      // Encode every range before the first write. A non-serializable value
      // therefore cancels the turn without partially saving game data.
      const preparedRanges = GameTurnRangeStorage.prepare(loaded, context.data, GAME_TURN_CONFIG);
      GameTurnRangeStorage.savePrepared(preparedRanges);

      context.journal.commit();
      GameTurnStateStorage.save(state, turn + 1, 'READY', new Date());

      return {
        processedTurn: turn,
        nextTurn: turn + 1,
        processors: context.processedProcessors.slice(),
        processorTimings: context.processorTimings.slice(),
        messages: context.journal.pendingCount,
        durationMs: new Date().getTime() - startedAt.getTime(),
      };
    } catch (error) {
      // Deliberately independent from game-data persistence: invalid JSON and
      // system errors become durable technical reports even though the game turn
      // itself is discarded.
      GameTurnJournalStorage.appendFatal(spreadsheet, GAME_TURN_CONFIG, turn, error,
        context && context.runtime.currentProcessor);
      if (state) GameTurnStateStorage.save(state, turn, 'FAILED', new Date());
      throw error;
    }
  },
};

/** Manual one-column province-template generator; it never advances a turn. */
const GameTurnProvinceGenerator = {
  fillEmpty: function () {
    const lock = LockService.getDocumentLock();
    lock.waitLock(GAME_TURN_CONFIG.lockTimeoutMs);

    try {
      const config = GAME_TURN_CONFIG.provinceTemplateGenerator;
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const range = spreadsheet.getRangeByName(config.rangeName);
      if (!range) throw new Error('Не найден именованный диапазон «' + config.rangeName + '».');
      if (range.getNumRows() < 2) {
        throw new Error('Диапазон «' + config.rangeName + '» должен содержать заголовки и хотя бы одну строку данных.');
      }

      const headers = range.getValues()[0];
      const column = headers.map(function (header) {
        return typeof header === 'string' ? header.trim() : header;
      }).indexOf(config.header);
      if (column === -1) {
        throw new Error('В диапазоне «' + config.rangeName + '» нет колонки «' + config.header + '».');
      }

      const target = range.offset(1, column, range.getNumRows() - 1, 1);
      const values = target.getValues();
      const usedIds = this._collectExistingIds(values, config, range, column);
      let nextId = this._nextId(usedIds, Number(config.firstId));
      let created = 0;

      values.forEach(function (row) {
        if (!GameTurnProvinceGenerator._isBlank(row[0])) return;
        while (usedIds[nextId]) nextId += 1;
        const province = GameTurnProvinceGenerator._createProvince(nextId, config);
        row[0] = JSON.stringify(province);
        usedIds[nextId] = true;
        nextId += 1;
        created += 1;
      });

      if (created) target.setValues(values);
      return {
        rangeName: config.rangeName,
        header: config.header,
        created: created,
        nextId: nextId,
      };
    } finally {
      lock.releaseLock();
    }
  },

  _collectExistingIds: function (values, config, range, column) {
    const ids = Object.create(null);
    values.forEach(function (row, index) {
      const raw = row[0];
      if (GameTurnProvinceGenerator._isBlank(raw)) return;
      const location = {
        rangeName: config.rangeName,
        header: config.header,
        sheetRow: range.getRow() + index + 1,
        sheetColumn: range.getColumn() + column,
      };
      const province = GameTurnCodec.decode(raw, location);
      if (!GameTurnProvinceGenerator._isObject(province)) {
        throw new GameTurnDataError(GameTurnCodec._location(location) + ': провинция должна быть JSON-объектом.');
      }
      if (!Number.isInteger(province.id) || province.id < 1) {
        throw new GameTurnDataError(GameTurnCodec._location(location) + ': province.id должен быть целым числом не меньше 1.');
      }
      if (ids[province.id]) {
        throw new GameTurnDataError('Повторяющийся id провинции «' + province.id + '» в колонке «' + config.header + '».');
      }
      ids[province.id] = true;
    });
    return ids;
  },

  _nextId: function (usedIds, firstId) {
    if (!Number.isInteger(firstId) || firstId < 1) {
      throw new Error('provinceTemplateGenerator.firstId должен быть целым числом не меньше 1.');
    }
    let nextId = firstId;
    while (usedIds[nextId]) nextId += 1;
    return nextId;
  },

  _createProvince: function (id, config) {
    const province = this._clone(config.template);
    province.id = id;
    province.name = (config.namesById && config.namesById[id]) ||
      String(config.fallbackNamePrefix || 'Провинция ') + id;
    return province;
  },

  _clone: function (value) {
    return JSON.parse(JSON.stringify(value));
  },

  _isBlank: function (value) {
    return value === null || value === undefined || value === '';
  },

  _isObject: function (value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
  },
};

/** Reads named ranges into plain column arrays while retaining cell mapping. */
const GameTurnRangeStorage = {
  loadAll: function (spreadsheet, config) {
    const names = config.dataRangeNames || [];
    if (!names.length) throw new Error('GAME_TURN_CONFIG.dataRangeNames не содержит игровых диапазонов.');

    const seen = Object.create(null);
    const ranges = Object.create(null);
    const data = Object.create(null);

    names.forEach(function (name) {
      if (typeof name !== 'string' || !name.trim()) {
        throw new Error('Каждое имя в dataRangeNames должно быть непустой строкой.');
      }
      if (seen[name]) throw new Error('Диапазон «' + name + '» указан в dataRangeNames дважды.');
      seen[name] = true;

      const range = spreadsheet.getRangeByName(name);
      if (!range) throw new Error('Не найден именованный диапазон «' + name + '».');
      if (range.getNumRows() < 2) {
        throw new Error('Диапазон «' + name + '» должен содержать строку заголовков и минимум одну строку данных.');
      }

      const raw = range.getValues();
      const state = this._decodeRange(name, range, raw);
      ranges[name] = state;
      data[name] = state.columns;
    }.bind(this));

    return { ranges: ranges, data: data };
  },

  _decodeRange: function (name, range, raw) {
    const headers = raw[0].map(function (value, index) {
      if (typeof value !== 'string' || !value.trim()) {
        throw new Error('Диапазон «' + name + '»: заголовок в колонке ' + (index + 1) + ' пустой.');
      }
      return value.trim();
    });
    const headerIndex = Object.create(null);
    headers.forEach(function (header, index) {
      if (headerIndex[header] !== undefined) {
        throw new Error('Диапазон «' + name + '» содержит повторяющийся заголовок «' + header + '».');
      }
      headerIndex[header] = index;
    });

    const decoded = raw.map(function (row, rowIndex) {
      return row.map(function (value, columnIndex) {
        if (rowIndex === 0) return value;
        return GameTurnCodec.decode(value, {
          rangeName: name,
          header: headers[columnIndex],
          sheetRow: range.getRow() + rowIndex,
          sheetColumn: range.getColumn() + columnIndex,
        });
      });
    });

    const columns = Object.create(null);
    headers.forEach(function (header, columnIndex) {
      columns[header] = [];
      for (let rowIndex = 1; rowIndex < decoded.length; rowIndex += 1) {
        columns[header].push(decoded[rowIndex][columnIndex]);
      }
    });

    return {
      name: name,
      range: range,
      rows: range.getNumRows(),
      columnsCount: range.getNumColumns(),
      headers: headers,
      headerIndex: headerIndex,
      originalEncoded: this._encodeMatrix(decoded, name, range, GAME_TURN_CONFIG),
      columns: columns,
    };
  },

  prepare: function (loaded, data, config) {
    this._assertDataShape(loaded, data);
    return Object.keys(loaded.ranges).map(function (name) {
      const state = loaded.ranges[name];
      const matrix = this._matrixFromColumns(state, data[name]);
      const encoded = this._encodeMatrix(matrix, name, state.range, config);
      return {
        state: state,
        encoded: encoded,
        changed: GameTurnCodec.signature(encoded) !== GameTurnCodec.signature(state.originalEncoded),
      };
    }.bind(this));
  },

  savePrepared: function (preparedRanges) {
    preparedRanges.forEach(function (prepared) {
      if (!prepared.changed) return;
      prepared.state.range.setValues(prepared.encoded);
    });
  },

  _matrixFromColumns: function (state, columns) {
    const matrix = [state.headers.slice()];
    for (let row = 0; row < state.rows - 1; row += 1) {
      const targetRow = [];
      state.headers.forEach(function (header) {
        targetRow.push(columns[header][row]);
      });
      matrix.push(targetRow);
    }
    return matrix;
  },

  _encodeMatrix: function (matrix, rangeName, range, config) {
    return matrix.map(function (row, rowIndex) {
      return row.map(function (value, columnIndex) {
        if (rowIndex === 0) return value;
        return GameTurnCodec.encode(value, {
          rangeName: rangeName,
          sheetRow: range.getRow() + rowIndex,
          sheetColumn: range.getColumn() + columnIndex,
        }, config.maxJsonCellCharacters);
      });
    });
  },

  _assertDataShape: function (loaded, data) {
    if (!data || typeof data !== 'object' || Array.isArray(data)) {
      throw new Error('Процессор должен вернуть объект data с теми же диапазонами и колонками.');
    }
    Object.keys(loaded.ranges).forEach(function (name) {
      const state = loaded.ranges[name];
      const rangeData = data[name];
      if (!rangeData || typeof rangeData !== 'object' || Array.isArray(rangeData)) {
        throw new Error('В возвращённых данных отсутствует диапазон «' + name + '».');
      }
      state.headers.forEach(function (header) {
        const column = rangeData[header];
        if (!Array.isArray(column) || column.length !== state.rows - 1) {
          throw new Error('Диапазон «' + name + '», колонка «' + header + '» изменила размер.');
        }
      });
    });
  },
};

/** JSON convention: a string beginning with { or [ must be valid JSON. */
const GameTurnCodec = {
  decode: function (value, location) {
    if (value === '' || value === null || value === undefined) return null;
    if (typeof value !== 'string') return value;
    const trimmed = value.trim();
    if (trimmed.charAt(0) !== '{' && trimmed.charAt(0) !== '[') return value;

    try {
      return JSON.parse(trimmed);
    } catch (error) {
      throw new GameTurnDataError(this._formatJsonError(location, value, error));
    }
  },

  encode: function (value, location, maxCharacters) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value;
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
    if (typeof value !== 'object') {
      throw new GameTurnDataError(this._location(location) + ': недопустимый тип значения «' + typeof value + '».');
    }

    let serialized;
    try {
      serialized = JSON.stringify(value);
    } catch (error) {
      throw new GameTurnDataError(this._location(location) + ': значение нельзя сериализовать в JSON. ' + error.message);
    }
    if (serialized === undefined) {
      throw new GameTurnDataError(this._location(location) + ': значение нельзя сериализовать в JSON.');
    }
    if (serialized.length > maxCharacters) {
      throw new GameTurnDataError(this._location(location) + ': JSON длиннее ' + maxCharacters + ' символов.');
    }
    return serialized;
  },

  signature: function (value) {
    return JSON.stringify(value, function (key, item) {
      return item instanceof Date ? item.toISOString() : item;
    });
  },

  _formatJsonError: function (location, value, error) {
    return 'Некорректный JSON.\n' +
      'Диапазон: ' + location.rangeName + '\n' +
      'Столбец: ' + location.header + '\n' +
      'Строка листа: ' + location.sheetRow + '\n' +
      'Ячейка: ' + this._toA1(location.sheetColumn, location.sheetRow) + '\n' +
      'Значение: ' + String(value).slice(0, 500) + '\n' +
      'Причина: ' + String(error && error.message ? error.message : error);
  },

  _location: function (location) {
    return 'Диапазон «' + location.rangeName + '», ячейка ' +
      this._toA1(location.sheetColumn, location.sheetRow);
  },

  _toA1: function (column, row) {
    let current = Number(column);
    let letters = '';
    while (current > 0) {
      const remainder = (current - 1) % 26;
      letters = String.fromCharCode(65 + remainder) + letters;
      current = Math.floor((current - 1) / 26);
    }
    return (letters || '?') + String(row || '?');
  },
};

function GameTurnDataError(message) {
  this.name = 'GameTurnDataError';
  this.message = message;
  this.stack = (new Error(message)).stack;
}
GameTurnDataError.prototype = Object.create(Error.prototype);
GameTurnDataError.prototype.constructor = GameTurnDataError;

/** Context exposed to every processor. It has no SpreadsheetApp access. */
const GameTurnContext = {
  create: function (loaded, journalBinding, turn, startedAt, config) {
    const context = {
      data: loaded.data,
      config: config,
      turn: turn,
      startedAt: startedAt,
      runtime: { currentProcessor: null },
      processedProcessors: [],
      processorTimings: [],
      journal: new GameTurnJournal(journalBinding, turn, config),
    };

    context.getRange = function (rangeName) {
      if (!context.data[rangeName]) throw new Error('Не найден игровой диапазон «' + rangeName + '».');
      return context.data[rangeName];
    };
    context.getColumn = function (rangeName, header) {
      const range = context.getRange(rangeName);
      if (!Object.prototype.hasOwnProperty.call(range, header)) {
        throw new Error('В диапазоне «' + rangeName + '» нет колонки «' + header + '».');
      }
      return range[header];
    };
    context.forEach = function (rangeName, header, callback) {
      const column = context.getColumn(rangeName, header);
      column.forEach(function (value, index) {
        const cell = GameTurnContext._cell(context, loaded.ranges[rangeName], header, index);
        callback(cell, index);
      });
    };
    context.set = function (rangeName, header, rowIndex, value) {
      const column = context.getColumn(rangeName, header);
      if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= column.length) {
        throw new Error('Недопустимый индекс строки для «' + rangeName + '.' + header + '».');
      }
      column[rowIndex] = value;
    };

    // Rich-text helpers. They create data only; the journal renders it after a
    // successful turn has written all game ranges.
    context.text = function (text, options) { return GameTurnMessage.text(text, options); };
    context.country = function (country, options) { return GameTurnMessage.country(country, options, config); };
    context.number = function (number, options) { return GameTurnMessage.number(number, options); };
    context.positive = function (text, options) { return GameTurnMessage.positive(text, options); };
    context.negative = function (text, options) { return GameTurnMessage.negative(text, options); };
    context.warning = function (text, options) { return GameTurnMessage.warning(text, options); };
    context.errorText = function (text, options) { return GameTurnMessage.errorText(text, options); };
    context.log = function (entry) { return context.journal.emit(entry, context.runtime.currentProcessor); };
    context.warn = function (message, options) {
      const settings = GameTurnMessage.copy(options, {});
      settings.type = settings.type || 'GAME';
      settings.category = settings.category || 'SYSTEM';
      settings.priority = settings.priority || 'HIGH';
      settings.message = message;
      return context.log(settings);
    };
    return context;
  },

  _cell: function (context, state, header, dataRowIndex) {
    const column = state.headerIndex[header];
    const rangeName = state.name;
    const cell = {
      rangeName: rangeName,
      header: header,
      rowIndex: dataRowIndex,
      row: state.range.getRow() + dataRowIndex + 1,
      column: column + 1,
      a1: GameTurnCodec._toA1(state.range.getColumn() + column, state.range.getRow() + dataRowIndex + 1),
      empty: false,
    };
    Object.defineProperty(cell, 'value', {
      enumerable: true,
      get: function () { return context.data[rangeName][header][dataRowIndex]; },
      set: function (value) { context.data[rangeName][header][dataRowIndex] = value; },
    });
    cell.empty = cell.value === null || cell.value === undefined || cell.value === '';
    return cell;
  },
};

/** Executes priority-sorted handlers and accepts their returned data object. */
const GameTurnProcessors = {
  run: function (context, processors) {
    (processors || [])
      .map(function (processor, index) { return { processor: processor, index: index }; })
      .sort(function (left, right) {
        const priority = Number(left.processor && left.processor.priority || 0) -
          Number(right.processor && right.processor.priority || 0);
        return priority || left.index - right.index;
      })
      .forEach(function (item) {
        const processor = item.processor;
        if (!processor || typeof processor.handler !== 'function') {
          throw new Error('Каждый процессор должен иметь функцию handler.');
        }
        const id = processor.id || ('PROCESSOR_' + item.index);
        context.runtime.currentProcessor = id;
        const startedAt = new Date().getTime();
        try {
          const result = processor.handler(context.data, context);
          const durationMs = new Date().getTime() - startedAt;
          if (result !== undefined) context.data = result;
          context.processedProcessors.push(id);
          GameTurnProcessors._recordTiming(context, id, durationMs);
        } catch (error) {
          // The fatal error report is saved independently of the failed turn.
          // Put the duration on the error so it can identify a slow failure.
          const failure = error instanceof Error ? error : new Error(String(error));
          failure.gameTurnProcessorId = id;
          failure.gameTurnProcessorDurationMs = new Date().getTime() - startedAt;
          throw failure;
        }
      });
    context.runtime.currentProcessor = null;
  },

  _recordTiming: function (context, processorId, durationMs) {
    const settings = context.config.processorTimingReports || {};
    const report = { id: processorId, durationMs: durationMs };
    context.processorTimings.push(report);
    if (settings.enabled === false) return;

    context.journal.emit({
      type: settings.type || 'PERFORMANCE',
      category: settings.category || 'ENGINE',
      priority: settings.priority || 'NORMAL',
      visibility: settings.visibility || 'SYSTEM',
      ttl: settings.ttl === undefined ? 1 : settings.ttl,
      source: processorId,
      message: [
        GameTurnMessage.text('Процессор «'),
        GameTurnMessage.text(processorId, { bold: true }),
        GameTurnMessage.text('» выполнен за '),
        GameTurnMessage.positive(String(durationMs) + ' мс'),
        GameTurnMessage.text('.'),
      ],
    }, processorId);
  },
};

const GameTurnValidators = {
  run: function (context, validators) {
    (validators || []).forEach(function (validator, index) {
      if (typeof validator !== 'function') throw new Error('Валидатор #' + (index + 1) + ' не является функцией.');
      const result = validator(context.data, context);
      if (result === undefined || result === true) return;
      const errors = result === false ? ['Проверка данных не пройдена.'] :
        typeof result === 'string' ? [result] : Array.isArray(result) ? result :
          ['Валидатор вернул недопустимый результат.'];
      throw new GameTurnDataError('Ход ' + context.turn + ' отменён. ' + errors.join(' | '));
    });
  },
};

/** Basic structural validator; extend GAME_TURN_CONFIG.validators for game rules. */
function validateGameTurnData(data, context) {
  // The final structural check happens again in GameTurnRangeStorage.prepare.
  // This validator is intentionally a hook for domain rules such as balances.
  return true;
}

/** Message segments are rendered to RichTextValue only while saving the journal. */
const GameTurnMessage = {
  text: function (text, options) {
    const segment = this.copy(options, {});
    segment.text = String(text === null || text === undefined ? '' : text);
    return this.normalizeSegment(segment);
  },

  country: function (country, options, config) {
    const settings = this.copy(options, {});
    if (!settings.color) settings.color = (config.countryColors || {})[String(country)] || '#2563EB';
    if (settings.bold === undefined) settings.bold = true;
    return this.text(country, settings);
  },

  number: function (number, options) {
    return this.text(number, options);
  },

  positive: function (text, options) {
    return this.text(text, this.copy(options, { color: '#15803D', bold: true }));
  },

  negative: function (text, options) {
    return this.text(text, this.copy(options, { color: '#B91C1C', bold: true }));
  },

  warning: function (text, options) {
    return this.text(text, this.copy(options, { color: '#C2410C', bold: true }));
  },

  errorText: function (text, options) {
    return this.text(text, this.copy(options, { color: '#B91C1C', bold: true }));
  },

  normalize: function (message) {
    if (typeof message === 'string' || typeof message === 'number' || typeof message === 'boolean') {
      return { text: String(message), segments: null };
    }
    const input = Array.isArray(message) ? message : [message];
    const segments = input.map(function (item) {
      return GameTurnMessage.normalizeSegment(item);
    }).filter(function (segment) {
      return segment.text !== '';
    });
    return { text: segments.map(function (segment) { return segment.text; }).join(''), segments: segments };
  },

  normalizeSegment: function (input) {
    if (typeof input === 'string' || typeof input === 'number' || typeof input === 'boolean') {
      return { text: String(input) };
    }
    if (!input || typeof input !== 'object' || Array.isArray(input) || input.text === undefined || input.text === null) {
      throw new Error('Форматированный фрагмент сообщения должен содержать поле text.');
    }
    const segment = { text: String(input.text) };
    ['color', 'link'].forEach(function (field) {
      if (input[field] === undefined || input[field] === null || input[field] === '') return;
      if (typeof input[field] !== 'string') throw new Error('Свойство ' + field + ' фрагмента должно быть строкой.');
      segment[field] = input[field];
    });
    ['bold', 'italic', 'underline', 'strikethrough'].forEach(function (field) {
      if (input[field] === undefined || input[field] === null) return;
      if (typeof input[field] !== 'boolean') throw new Error('Свойство ' + field + ' фрагмента должно быть boolean.');
      segment[field] = input[field];
    });
    return segment;
  },

  copy: function (source, defaults) {
    const result = {};
    Object.keys(defaults || {}).forEach(function (key) { result[key] = defaults[key]; });
    if (source && typeof source === 'object' && !Array.isArray(source)) {
      Object.keys(source).forEach(function (key) { result[key] = source[key]; });
    }
    return result;
  },
};

/** In-memory GameLog; it performs no writes until commit() is called. */
function GameTurnJournal(binding, turn, config) {
  this.binding = binding;
  this.turn = turn;
  this.config = config;
  this.pendingSpecs = Object.create(null);
  this.pendingCount = 0;
}

GameTurnJournal.prototype.emit = function (input, source) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw new Error('ctx.log ожидает объект сообщения.');
  }
  const message = GameTurnMessage.normalize(input.message);
  if (!message.text.trim()) throw new Error('Сообщение журнала не может быть пустым.');
  const row = this._findEmptyRow() || this._resolveOverflowRow();
  const ttl = input.ttl === undefined ? this.config.journal.defaultTtl : Number(input.ttl);
  if (!Number.isInteger(ttl) || ttl === 0 || ttl < -1) {
    throw new Error('ttl сообщения должен быть положительным числом ходов или -1 для постоянного сообщения.');
  }

  const entry = {
    id: this._makeId(row),
    turn: this.turn,
    type: String(input.type || 'GAME').toUpperCase(),
    category: String(input.category || 'GENERAL').toUpperCase(),
    priority: String(input.priority || 'NORMAL').toUpperCase(),
    country: input.country === undefined || input.country === null ? '' : String(input.country),
    visibility: String(input.visibility || 'PUBLIC').toUpperCase(),
    ttl: ttl,
    message: message.text,
    source: String(input.source || source || 'ENGINE'),
  };
  GameTurnJournalStorage.writeEntry(this.binding.matrix, row, entry);
  if (message.segments) {
    this.pendingSpecs[row] = { text: message.text, segments: message.segments };
  } else {
    delete this.pendingSpecs[row];
  }
  this.pendingCount += 1;
  return entry.id;
};

GameTurnJournal.prototype.expire = function (turn) {
  for (let row = 1; row < this.binding.matrix.length; row += 1) {
    const entry = GameTurnJournalStorage.readEntry(this.binding.matrix[row]);
    if (!entry || entry.ttl === -1) continue;
    if (entry.turn + entry.ttl <= turn) {
      GameTurnJournalStorage.clearEntry(this.binding.matrix, row);
      delete this.pendingSpecs[row];
    }
  }
};

GameTurnJournal.prototype.commit = function () {
  const encoded = this.binding.matrix.map(function (row) { return row.slice(); });
  const signature = GameTurnCodec.signature(encoded);
  if (signature === this.binding.signature) return;
  this.binding.range.setValues(encoded);
  GameTurnJournalStorage.applyRichText(
    this.binding.range,
    this.binding.matrix,
    this.binding.richTextValues,
    this.pendingSpecs
  );
  this.binding.signature = signature;
};

GameTurnJournal.prototype._findEmptyRow = function () {
  for (let row = 1; row < this.binding.matrix.length; row += 1) {
    if (!GameTurnJournalStorage.readEntry(this.binding.matrix[row])) return row;
  }
  return null;
};

GameTurnJournal.prototype._resolveOverflowRow = function () {
  if (String(this.config.journal.overflow || 'DROP_OLDEST').toUpperCase() === 'THROW') {
    throw new Error('Игровой_журнал заполнен. Расширьте именованный диапазон.');
  }
  let oldest = null;
  for (let row = 1; row < this.binding.matrix.length; row += 1) {
    const entry = GameTurnJournalStorage.readEntry(this.binding.matrix[row]);
    if (!entry) return row;
    if (!oldest || entry.turn < oldest.entry.turn) oldest = { row: row, entry: entry };
  }
  return oldest ? oldest.row : null;
};

GameTurnJournal.prototype._makeId = function (row) {
  return 't' + this.turn + '_' + row + '_' + (this.pendingCount + 1);
};

/** Storage and rich-text renderer for the separate Игровой_журнал range. */
const GameTurnJournalStorage = {
  ensure: function (spreadsheet, config) {
    let range = spreadsheet.getRangeByName(config.journal.rangeName);
    if (range) {
      this._assertShape(range, config.journal.rangeName);
      return range;
    }

    const sheet = this._createEmptySheet(spreadsheet, config.journal.sheetName, config.journal.rangeName);
    this._ensureCapacity(sheet, config.journal.rows, GAME_TURN_JOURNAL_HEADERS.length);
    range = sheet.getRange(1, 1, config.journal.rows, GAME_TURN_JOURNAL_HEADERS.length);
    spreadsheet.setNamedRange(config.journal.rangeName, range);
    const matrix = this.createEmptyMatrix(config.journal.rows);
    range.setValues(matrix);
    return range;
  },

  load: function (spreadsheet, config) {
    const range = this.ensure(spreadsheet, config);
    const matrix = range.getValues();
    this._assertHeaders(matrix, config.journal.rangeName);
    return {
      range: range,
      matrix: matrix,
      signature: GameTurnCodec.signature(matrix),
      richTextValues: range.getRichTextValues(),
    };
  },

  appendFatal: function (spreadsheet, config, turn, error, source) {
    try {
      const binding = this.load(spreadsheet, config);
      const journal = new GameTurnJournal(binding, turn, config);
      const processorId = source || (error && error.gameTurnProcessorId) || 'ENGINE';
      const durationMs = error && Number.isFinite(error.gameTurnProcessorDurationMs)
        ? error.gameTurnProcessorDurationMs
        : null;
      const message = [GameTurnMessage.errorText('Ошибка хода')];
      if (processorId !== 'ENGINE') {
        message.push(GameTurnMessage.text(' в процессоре «'));
        message.push(GameTurnMessage.negative(processorId));
        message.push(GameTurnMessage.text('»'));
      }
      if (durationMs !== null) {
        message.push(GameTurnMessage.text(' через '));
        message.push(GameTurnMessage.warning(String(durationMs) + ' мс'));
      }
      message.push(GameTurnMessage.text(': '));
      message.push(GameTurnMessage.text(String(error && error.message ? error.message : error)));
      journal.emit({
        type: 'ERROR',
        category: 'ENGINE',
        priority: 'CRITICAL',
        visibility: 'SYSTEM',
        ttl: -1,
        source: processorId,
        message: message,
      }, processorId);
      journal.commit();
    } catch (journalError) {
      console.log('Не удалось записать ошибку в Игровой_журнал: ' + journalError.message);
    }
  },

  createEmptyMatrix: function (rows) {
    const matrix = [];
    for (let row = 0; row < rows; row += 1) matrix.push(Array(GAME_TURN_JOURNAL_HEADERS.length).fill(''));
    matrix[0] = GAME_TURN_JOURNAL_HEADERS.slice();
    return matrix;
  },

  readEntry: function (row) {
    if (!Array.isArray(row)) return null;
    const message = row[GAME_TURN_JOURNAL.MESSAGE];
    const turn = Number(row[GAME_TURN_JOURNAL.TURN]);
    if (typeof message !== 'string' || !message.trim() || !Number.isInteger(turn)) return null;
    return {
      id: row[GAME_TURN_JOURNAL.ID],
      turn: turn,
      ttl: Number(row[GAME_TURN_JOURNAL.TTL]),
      message: message,
    };
  },

  writeEntry: function (matrix, row, entry) {
    matrix[row] = [
      entry.id,
      entry.turn,
      entry.type,
      entry.category,
      entry.priority,
      entry.country,
      entry.visibility,
      entry.ttl,
      entry.message,
      entry.source,
    ];
  },

  clearEntry: function (matrix, row) {
    matrix[row] = Array(GAME_TURN_JOURNAL_HEADERS.length).fill('');
  },

  applyRichText: function (range, matrix, previousRichTextValues, pendingSpecs) {
    if (matrix.length < 2) return;
    const messageRange = range.offset(1, GAME_TURN_JOURNAL.MESSAGE, matrix.length - 1, 1);
    const values = matrix.slice(1).map(function (row, index) {
      const physicalRow = index + 1;
      const message = String(row[GAME_TURN_JOURNAL.MESSAGE] || '');
      const previous = previousRichTextValues && previousRichTextValues[physicalRow]
        ? previousRichTextValues[physicalRow][GAME_TURN_JOURNAL.MESSAGE]
        : null;
      const spec = pendingSpecs && pendingSpecs[physicalRow];
      if (!spec && previous && typeof previous.getText === 'function' && previous.getText() === message) {
        return [previous];
      }
      return [GameTurnJournalStorage._buildRichText(
        message,
        String(row[GAME_TURN_JOURNAL.PRIORITY] || 'NORMAL').toUpperCase(),
        spec && spec.text === message ? spec.segments : null
      )];
    });
    messageRange.setRichTextValues(values);
  },

  _buildRichText: function (message, priority, segments) {
    const builder = SpreadsheetApp.newRichTextValue().setText(message);
    const baseColor = this._priorityColor(priority);
    builder.setTextStyle(this._buildStyle({ color: baseColor }, baseColor));
    if (Array.isArray(segments) && segments.map(function (segment) { return segment.text; }).join('') === message) {
      let start = 0;
      segments.forEach(function (segment) {
        const end = start + segment.text.length;
        if (end > start) {
          builder.setTextStyle(start, end, GameTurnJournalStorage._buildStyle(segment, baseColor));
          if (segment.link) builder.setLinkUrl(start, end, segment.link);
        }
        start = end;
      });
    }
    return builder.build();
  },

  _buildStyle: function (settings, fallbackColor) {
    const builder = SpreadsheetApp.newTextStyle().setForegroundColor(settings.color || fallbackColor);
    if (settings.bold !== undefined) builder.setBold(settings.bold);
    if (settings.italic !== undefined) builder.setItalic(settings.italic);
    if (settings.underline !== undefined) builder.setUnderline(settings.underline);
    if (settings.strikethrough !== undefined) builder.setStrikethrough(settings.strikethrough);
    return builder.build();
  },

  _priorityColor: function (priority) {
    return {
      SUCCESS: '#15803D',
      HIGH: '#C2410C',
      CRITICAL: '#B91C1C',
      ERROR: '#B91C1C',
      NORMAL: '#1D4ED8',
    }[priority] || '#475569';
  },

  _assertShape: function (range, rangeName) {
    if (range.getNumRows() < 2 || range.getNumColumns() !== GAME_TURN_JOURNAL_HEADERS.length) {
      throw new Error('«' + rangeName + '» должен иметь минимум две строки и ' +
        GAME_TURN_JOURNAL_HEADERS.length + ' колонок.');
    }
    this._assertHeaders(range.getValues(), rangeName);
  },

  _assertHeaders: function (matrix, rangeName) {
    const headers = matrix[0] || [];
    GAME_TURN_JOURNAL_HEADERS.forEach(function (header, column) {
      if (headers[column] !== header) {
        throw new Error('«' + rangeName + '»: в колонке ' + (column + 1) + ' ожидается заголовок «' + header + '».');
      }
    });
  },

  _createEmptySheet: function (spreadsheet, sheetName, rangeName) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return spreadsheet.insertSheet(sheetName);
    if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
      throw new Error('Лист «' + sheetName + '» не пуст, но диапазон «' + rangeName + '» отсутствует.');
    }
    return sheet;
  },

  _ensureCapacity: function (sheet, rows, columns) {
    if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
    if (sheet.getMaxColumns() < columns) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
  },
};

/** Separate engine-owned range holding the number of the next game turn. */
const GameTurnStateStorage = {
  ensure: function (spreadsheet, config) {
    let range = spreadsheet.getRangeByName(config.state.rangeName);
    if (range) {
      this._assert(range, config.state.rangeName);
      return range;
    }

    let sheet = spreadsheet.getSheetByName(config.state.sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(config.state.sheetName);
    } else if (sheet.getLastRow() > 0 || sheet.getLastColumn() > 0) {
      throw new Error('Лист «' + config.state.sheetName + '» не пуст, но диапазон состояния отсутствует.');
    }
    if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(sheet.getMaxRows(), 2 - sheet.getMaxRows());
    if (sheet.getMaxColumns() < 3) sheet.insertColumnsAfter(sheet.getMaxColumns(), 3 - sheet.getMaxColumns());
    range = sheet.getRange(1, 1, 2, 3);
    spreadsheet.setNamedRange(config.state.rangeName, range);
    range.setValues([
      ['TURN', 'STATUS', 'LAST_RUN'],
      [1, 'READY', ''],
    ]);
    return range;
  },

  load: function (spreadsheet, config) {
    const range = this.ensure(spreadsheet, config);
    const values = range.getValues();
    const turn = Number(values[1][0]);
    if (!Number.isInteger(turn) || turn < 1) throw new Error('Состояние движка: TURN должен быть целым числом не меньше 1.');
    return { range: range, turn: turn };
  },

  save: function (state, turn, status, date) {
    state.range.setValues([
      ['TURN', 'STATUS', 'LAST_RUN'],
      [turn, status, date.toISOString()],
    ]);
  },

  _assert: function (range, name) {
    if (range.getNumRows() !== 2 || range.getNumColumns() !== 3) {
      throw new Error('«' + name + '» должен иметь размер 2 × 3.');
    }
    const headers = range.getValues()[0];
    if (headers[0] !== 'TURN' || headers[1] !== 'STATUS' || headers[2] !== 'LAST_RUN') {
      throw new Error('«' + name + '» имеет неверные заголовки.');
    }
  },
};
