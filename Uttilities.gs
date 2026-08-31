/**
 * Manual utility commands that do not run a game turn.
 *
 * GENERATE_EMPTY_PROVINCES() fills only blank cells in the province column.
 * Edit PROVINCE_TEMPLATE_GENERATOR_CONFIG before running it.
 */

const PROVINCE_TEMPLATE_GENERATOR_CONFIG = {
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
};

function GENERATE_EMPTY_PROVINCES() {
  return GameTurnProvinceGenerator.fillEmpty();
}

const GameTurnProvinceGenerator = {
  fillEmpty: function () {
    const lock = LockService.getDocumentLock();
    lock.waitLock(GAME_TURN_CONFIG.lockTimeoutMs);

    try {
      const config = PROVINCE_TEMPLATE_GENERATOR_CONFIG;
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
      throw new Error('PROVINCE_TEMPLATE_GENERATOR_CONFIG.firstId должен быть целым числом не меньше 1.');
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
