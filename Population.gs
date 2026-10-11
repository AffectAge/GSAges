/**
 * Population groups are arrays packed into Игровые_данные.Население cells.
 * This module does no spreadsheet writes during a turn.
 */
function processPopulation(data, ctx) {
  const rules = GAME_RULES.population;
  const columns = GamePopulation.columns(data, rules);
  const provinces = GamePopulation.provinces(columns[rules.provinceHeader]);
  const groups = GamePopulation.read(columns[rules.header], provinces);
  const years = rules.yearsPerTurn;
  if (typeof years !== 'number' || !Number.isFinite(years) || years <= 0) {
    throw new GameTurnDataError('GAME_RULES.population.yearsPerTurn должен быть положительным числом.');
  }
  const updated = groups.map(function (group) {
    const copy = JSON.parse(JSON.stringify(group));
    copy.men = GamePopulation.nextCount(group.men, group.birthRate, group.deathRate, years);
    copy.women = GamePopulation.nextCount(group.women, group.birthRate, group.deathRate, years);
    copy.total = copy.men + copy.women;
    GamePopulation.validate(copy, provinces);
    return copy;
  });
  columns[rules.header] = GamePopulation.pack(updated, columns[rules.header].length,
    ctx.config.maxJsonCellCharacters);
  return data;
}

/** Final check/repacking also covers population added by later processors. */
function validatePopulationData(data, ctx) {
  const rules = GAME_RULES.population;
  const columns = GamePopulation.columns(data, rules);
  const groups = GamePopulation.read(columns[rules.header],
    GamePopulation.provinces(columns[rules.provinceHeader]));
  columns[rules.header] = GamePopulation.pack(groups, columns[rules.header].length,
    ctx.config.maxJsonCellCharacters);
  return true;
}

const GamePopulation = {
  columns: function (data, rules) {
    const columns = data[rules.rangeName];
    if (!columns || !Array.isArray(columns[rules.provinceHeader]) ||
        !Array.isArray(columns[rules.header])) {
      throw new GameTurnDataError('В диапазоне «' + rules.rangeName +
        '» нужны колонки «' + rules.provinceHeader + '» и «' + rules.header + '».');
    }
    return columns;
  },

  provinces: function (cells) {
    const ids = Object.create(null);
    cells.forEach(function (province) {
      if (province === null || province === undefined || province === '') return;
      if (!province || typeof province !== 'object' || Array.isArray(province) ||
          !Number.isSafeInteger(province.id) || province.id < 1) {
        throw new GameTurnDataError('Для населения каждая провинция должна иметь положительный целый id.');
      }
      if (ids[province.id]) throw new GameTurnDataError('Повторяющийся id провинции: ' + province.id);
      ids[province.id] = true;
    });
    return ids;
  },

  read: function (cells, provinces) {
    const groups = [];
    const ids = Object.create(null);
    cells.forEach(function (cell, row) {
      if (cell === null || cell === undefined || cell === '') return;
      if (!Array.isArray(cell)) {
        throw new GameTurnDataError('Население: ячейка данных #' + (row + 1) +
          ' должна содержать JSON-массив групп.');
      }
      cell.forEach(function (group) {
        GamePopulation.validate(group, provinces);
        if (ids[group.id]) throw new GameTurnDataError('Повторяющийся id группы населения: ' + group.id);
        ids[group.id] = true;
        groups.push(group);
      });
    });
    return groups;
  },

  validate: function (group, provinces) {
    if (!group || typeof group !== 'object' || Array.isArray(group)) {
      throw new GameTurnDataError('Группа населения должна быть JSON-объектом.');
    }
    if (typeof group.id !== 'string' || !group.id.trim()) {
      throw new GameTurnDataError('Группа населения должна иметь непустой строковый id.');
    }
    if (!Number.isSafeInteger(group.provinceId) || !provinces[group.provinceId]) {
      throw new GameTurnDataError('Группа «' + group.id + '»: неизвестная провинция ' + group.provinceId);
    }
    ['culture', 'religion'].forEach(function (key) {
      if (typeof group[key] !== 'string' || !group[key].trim()) {
        throw new GameTurnDataError('Группа «' + group.id + '»: ' + key + ' должна быть непустой строкой.');
      }
    });
    ['men', 'women', 'total'].forEach(function (key) {
      if (!Number.isSafeInteger(group[key]) || group[key] < 0) {
        throw new GameTurnDataError('Группа «' + group.id + '»: ' + key +
          ' должно быть неотрицательным безопасным целым числом.');
      }
    });
    if (!Number.isSafeInteger(group.men + group.women) || group.total !== group.men + group.women) {
      throw new GameTurnDataError('Группа «' + group.id + '»: total должен равняться men + women.');
    }
    ['birthRate', 'deathRate'].forEach(function (key) {
      if (typeof group[key] !== 'number' || !Number.isFinite(group[key]) || group[key] < 0) {
        throw new GameTurnDataError('Группа «' + group.id + '»: ' + key +
          ' должен быть конечным неотрицательным числом.');
      }
    });
  },

  nextCount: function (count, births, deaths, years) {
    // Both rates apply to the starting count, separately for each sex.
    // Fractions are discarded once from the final count, never carried over.
    const result = Math.max(0, Math.floor(count + count * births * years - count * deaths * years));
    if (!Number.isSafeInteger(result)) {
      throw new GameTurnDataError('Расчёт населения вышел за предел безопасных целых чисел.');
    }
    return result;
  },

  pack: function (groups, capacity, limit) {
    if (!Number.isSafeInteger(capacity) || capacity < 0 ||
        !Number.isSafeInteger(limit) || limit < 2 || limit > 48000) {
      throw new GameTurnDataError('Недопустимые пределы упаковки населения (максимум 48000 символов).');
    }
    // First-fit decreasing: whole groups only, using exact JSON lengths.
    const entries = groups.map(function (group, index) {
      const json = JSON.stringify(group);
      if (json.length + 2 > limit) {
        throw new GameTurnDataError('Группа «' + group.id +
          '» не помещается в одну ячейку населения: ' + (json.length + 2) +
          ' символов, предел ' + limit + '. Ход отменён.');
      }
      return { group: group, length: json.length, index: index };
    }).sort(function (a, b) { return b.length - a.length || a.index - b.index; });
    const bins = [];
    entries.forEach(function (entry) {
      let bin = bins.find(function (candidate) {
        return candidate.length + entry.length + (candidate.groups.length ? 1 : 0) <= limit;
      });
      if (!bin) {
        bin = { groups: [], length: 2 };
        bins.push(bin);
      }
      bin.length += entry.length + (bin.groups.length ? 1 : 0);
      bin.groups.push(entry.group);
    });
    if (bins.length > capacity) {
      throw new GameTurnDataError('Недостаточно места в столбце «Население»: упаковка требует ' +
        bins.length + ' ячеек, доступно ' + capacity +
        '. Расширьте именованный диапазон «Игровые_данные». Ход отменён.');
    }
    const cells = bins.map(function (bin) { return bin.groups; });
    while (cells.length < capacity) cells.push(null);
    return cells;
  },

  generate: function (spreadsheet, config) {
    const rules = GAME_RULES.population;
    const range = spreadsheet.getRangeByName(rules.rangeName);
    if (!range) throw new GameTurnDataError('Не найден диапазон «' + rules.rangeName + '».');
    const binding = GameTurnRangeStorage._decodeRange(rules.rangeName, range, range.getValues());
    const data = {};
    data[rules.rangeName] = binding.columns;
    const columns = this.columns(data, rules);
    const provinces = this.provinces(columns[rules.provinceHeader]);
    const groups = this.read(columns[rules.header], provinces);
    const populated = Object.create(null);
    const ids = Object.create(null);
    groups.forEach(function (group) { populated[group.provinceId] = true; ids[group.id] = true; });
    if (!Number.isSafeInteger(config.firstId) || config.firstId < 1) {
      throw new GameTurnDataError('POPULATION_TEMPLATE_GENERATOR_CONFIG.firstId должен быть положительным целым.');
    }
    let nextId = config.firstId;
    let created = 0;
    columns[rules.provinceHeader].forEach(function (province) {
      if (!province || populated[province.id]) return;
      const group = JSON.parse(JSON.stringify(config.template));
      let id;
      do {
        if (!Number.isSafeInteger(nextId)) throw new GameTurnDataError('Исчерпаны id населения.');
        id = 'pop_' + nextId;
        nextId += 1;
      } while (ids[id]);
      group.id = id;
      group.provinceId = province.id;
      group.total = group.men + group.women;
      // Defaults apply only on creation, never during a turn.
      if (group.birthRate === undefined) group.birthRate = rules.defaultBirthRate;
      if (group.deathRate === undefined) group.deathRate = rules.defaultDeathRate;
      this.validate(group, provinces);
      groups.push(group);
      ids[id] = true;
      populated[province.id] = true;
      created += 1;
    }.bind(this));
    const packed = this.pack(groups, columns[rules.header].length, GAME_TURN_CONFIG.maxJsonCellCharacters);
    const target = range.offset(1, binding.headerIndex[rules.header], range.getNumRows() - 1, 1);
    if (target.getFormulas().some(function (row) { return !!row[0]; })) {
      throw new GameTurnDataError('Столбец «Население» содержит формулы: генерация отменена.');
    }
    // Encode everything before the only population-column write.
    const values = packed.map(function (cell) { return [cell === null ? '' : JSON.stringify(cell)]; });
    if (created) target.setValues(values);
    return { created: created, groups: groups.length, usedCells: packed.filter(Boolean).length };
  },
};

function GENERATE_INITIAL_POPULATION() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(GAME_TURN_CONFIG.lockTimeoutMs);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    return GamePopulation.generate(spreadsheet, POPULATION_TEMPLATE_GENERATOR_CONFIG);
  } catch (error) {
    let turn = 0;
    try { turn = GameTurnStateStorage.load(spreadsheet, GAME_TURN_CONFIG).turn; } catch (ignored) {}
    GameTurnJournalStorage.appendFatal(spreadsheet, GAME_TURN_CONFIG, turn, error, 'POPULATION_GENERATOR');
    throw error;
  } finally {
    lock.releaseLock();
  }
}
