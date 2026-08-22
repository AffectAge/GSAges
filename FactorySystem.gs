/**
 * Factory mechanics. This file is intentionally independent from GameEngine.gs;
 * it only receives the in-memory turn context and never calls SpreadsheetApp.
 */

const FACTORY_SYSTEM_CONFIG = {
  factoriesRange: 'NR_FACTORIES',
  factoriesSubrange: 'FACTORIES',
  templatesRange: 'NR_GAME_CORE',
  templatesSubrange: 'FACTORY_TEMPLATES',
  provincesRange: 'NR_WORLD',
  provincesSubrange: 'PROVINCES',
  category: 'INDUSTRY',
  seedDefaultTemplate: true,
};

/**
 * Factory record example:
 * {
 *   "id":"factory_rus_1", "templateId":"TEXTILE_MILL", "owner":"RUS",
 *   "provinceId":"prov_1", "level":1, "efficiency":1,
 *   "stockpile":{"cotton":20,"coal":3,"clothes":0}, "status":"ACTIVE"
 * }
 *
 * Template example:
 * {
 *   "id":"TEXTILE_MILL", "inputs":{"cotton":2,"coal":0.1},
 *   "outputs":{"clothes":1}, "productionPerLevel":1,
 *   "pollutionPerCycle":2
 * }
 */
const FactorySystem = {
  process: function (ctx) {
    const initialized = this._ensureStorage(ctx);
    const factories = initialized.factories;
    if (!factories) return { processed: 0, producedCycles: 0, pollution: 0 };

    const templates = this._indexTemplates(ctx);
    const provinces = this._indexProvinces(ctx);
    const report = { processed: 0, producedCycles: 0, pollution: 0 };
    const self = this;

    ctx.helpers.forEachCell(factories, function (factory, row, column) {
      if (factory === null || factory === undefined) return;
      if (!FactorySystem._isObject(factory)) {
        ctx.log.error('Factory cell contains a non-JSON value.', { row: row, column: column });
        return;
      }

      self._processFactory(ctx, factory, templates, provinces, report, row, column);
    });

    ctx.runtime.factoryReport = report;
    return report;
  },

  /**
   * Creates missing virtual columns in old as well as new containers. It never
   * removes or replaces user data. The first template is merely a usable
   * starting point and may be edited directly in the sheet afterwards.
   */
  _ensureStorage: function (ctx) {
    const templates = this._getOrCreateSubrange(
      ctx,
      FACTORY_SYSTEM_CONFIG.templatesRange,
      FACTORY_SYSTEM_CONFIG.templatesSubrange
    );
    const factories = this._getOrCreateSubrange(
      ctx,
      FACTORY_SYSTEM_CONFIG.factoriesRange,
      FACTORY_SYSTEM_CONFIG.factoriesSubrange
    );

    if (templates && FACTORY_SYSTEM_CONFIG.seedDefaultTemplate &&
        !this._containsId(templates, 'TEXTILE_MILL')) {
      this._putIntoFirstEmptyCell(templates, createDefaultFactoryTemplate(), 'factory template');
    }

    return { templates: templates, factories: factories };
  },

  _getOrCreateSubrange: function (ctx, rangeName, subrangeName) {
    const existing = this._getSubrange(ctx, rangeName, subrangeName);
    if (existing) return existing;

    const container = ctx.data[rangeName];
    if (!container) return null;
    const rows = this._getContainerDataRows(container, rangeName);
    const matrix = createEmptyColumnMatrix(rows);
    container[subrangeName] = matrix;
    return matrix;
  },

  _getContainerDataRows: function (container, rangeName) {
    const keys = Object.keys(container);
    for (let index = 0; index < keys.length; index += 1) {
      const candidate = container[keys[index]];
      if (Array.isArray(candidate)) return candidate.length;
    }

    const rangeDefinition = GAME_ENGINE_CONFIG.rangeDefaults[rangeName] || {};
    const containerDefinition = GAME_ENGINE_CONFIG.containers[rangeName] || {};
    const rows = Number(rangeDefinition.rows) - Number(containerDefinition.dataStartRow);
    if (!Number.isInteger(rows) || rows < 1) {
      throw new Error('Cannot determine the data-row count for ' + rangeName + '.');
    }
    return rows;
  },

  _containsId: function (matrix, id) {
    let found = false;
    GameHelpers.forEachCell(matrix, function (value) {
      if (FactorySystem._isObject(value) && value.id === id) found = true;
    });
    return found;
  },

  _putIntoFirstEmptyCell: function (matrix, value, label) {
    for (let row = 0; row < matrix.length; row += 1) {
      for (let column = 0; column < matrix[row].length; column += 1) {
        if (matrix[row][column] === null || matrix[row][column] === undefined) {
          matrix[row][column] = value;
          return;
        }
      }
    }
    throw new Error('No empty cell is available for the default ' + label + '.');
  },

  _processFactory: function (ctx, factory, templates, provinces, report, row, column) {
    if (!this._isText(factory.id)) {
      ctx.log.error('Factory has no id.', { row: row, column: column });
      return;
    }
    this._ensureFactoryDefaults(factory);
    if (!this._isText(factory.templateId)) {
      this._setOperationalStatus(ctx, factory, 'TEMPLATE_MISSING', 'Factory has no templateId.');
      return;
    }

    const template = templates[factory.templateId];
    if (!template) {
      this._setOperationalStatus(ctx, factory, 'TEMPLATE_MISSING', 'Factory template is missing: ' + factory.templateId + '.');
      return;
    }

    if (factory.status === 'CONSTRUCTING') {
      this._processConstruction(ctx, factory);
      return;
    }
    if (factory.status !== 'ACTIVE') return;

    report.processed += 1;
    const cycles = this._calculateCycles(factory, template);
    if (cycles <= 0) {
      factory.operationalStatus = 'IDLE';
      factory.lastProduction = { turn: ctx.turn, cycles: 0, outputs: {} };
      return;
    }

    const availableCycles = this._limitByInputs(factory.stockpile, template.inputs || {}, cycles);
    if (availableCycles <= 0) {
      this._setOperationalStatus(ctx, factory, 'INPUT_SHORTAGE', 'Factory has insufficient input goods.');
      factory.lastProduction = { turn: ctx.turn, cycles: 0, outputs: {} };
      return;
    }

    this._consumeInputs(factory.stockpile, template.inputs || {}, availableCycles);
    const outputs = this._produceOutputs(factory.stockpile, template.outputs || {}, availableCycles);
    factory.lastProduction = { turn: ctx.turn, cycles: availableCycles, outputs: outputs };
    report.producedCycles += availableCycles;

    const pollution = this._number(template.pollutionPerCycle, 0) * availableCycles;
    let provinceWasUpdated = true;
    if (pollution > 0) {
      report.pollution += pollution;
      provinceWasUpdated = this._applyPollution(ctx, factory, provinces, pollution);
    }
    if (provinceWasUpdated) factory.operationalStatus = 'RUNNING';
  },

  _processConstruction: function (ctx, factory) {
    const remaining = Math.max(0, Math.floor(this._number(factory.constructionTurnsRemaining, 0)));
    if (remaining > 1) {
      factory.constructionTurnsRemaining = remaining - 1;
      return;
    }

    factory.constructionTurnsRemaining = 0;
    factory.status = 'ACTIVE';
    factory.operationalStatus = 'RUNNING';
    this._emit(ctx, factory, 'NORMAL', 'Construction completed: ' + factory.id + '.', 2);
  },

  _calculateCycles: function (factory, template) {
    const level = Math.max(1, Math.floor(this._number(factory.level, 1)));
    const efficiency = Math.max(0, this._number(factory.efficiency, 1));
    return Math.max(0, this._number(template.productionPerLevel, 1)) * level * efficiency;
  },

  _limitByInputs: function (stockpile, inputs, requestedCycles) {
    let cycles = requestedCycles;
    Object.keys(inputs).forEach(function (goodId) {
      const perCycle = FactorySystem._number(inputs[goodId], 0);
      if (perCycle <= 0) return;
      cycles = Math.min(cycles, Math.max(0, FactorySystem._number(stockpile[goodId], 0)) / perCycle);
    });
    return Math.max(0, cycles);
  },

  _consumeInputs: function (stockpile, inputs, cycles) {
    Object.keys(inputs).forEach(function (goodId) {
      const amount = Math.max(0, FactorySystem._number(inputs[goodId], 0)) * cycles;
      if (amount > 0) stockpile[goodId] = Math.max(0, FactorySystem._number(stockpile[goodId], 0) - amount);
    });
  },

  _produceOutputs: function (stockpile, outputs, cycles) {
    const produced = Object.create(null);
    Object.keys(outputs).forEach(function (goodId) {
      const amount = Math.max(0, FactorySystem._number(outputs[goodId], 0)) * cycles;
      if (amount <= 0) return;
      stockpile[goodId] = FactorySystem._number(stockpile[goodId], 0) + amount;
      produced[goodId] = amount;
    });
    return produced;
  },

  _applyPollution: function (ctx, factory, provinces, amount) {
    const province = provinces[factory.provinceId];
    if (!province) {
      this._setOperationalStatus(ctx, factory, 'PROVINCE_MISSING', 'Factory province is missing: ' + factory.provinceId + '.');
      return false;
    }
    province.pollution = this._number(province.pollution, 0) + amount;
    return true;
  },

  _ensureFactoryDefaults: function (factory) {
    if (!this._isText(factory.status)) factory.status = 'ACTIVE';
    if (!this._isObject(factory.stockpile)) factory.stockpile = Object.create(null);
    if (!this._isFiniteNumber(factory.level) || factory.level < 1) factory.level = 1;
    if (!this._isFiniteNumber(factory.efficiency) || factory.efficiency < 0) factory.efficiency = 1;
  },

  _setOperationalStatus: function (ctx, factory, status, message) {
    const changed = factory.operationalStatus !== status;
    factory.operationalStatus = status;
    if (changed) this._emit(ctx, factory, 'HIGH', message, 2);
  },

  _emit: function (ctx, factory, priority, message, ttlTurns) {
    const owner = this._isText(factory.owner) ? factory.owner : null;
    ctx.journal.emit({
      country: owner,
      category: FACTORY_SYSTEM_CONFIG.category,
      priority: priority,
      visibility: owner
        ? { type: 'COUNTRY', targets: [owner] }
        : { type: 'DEBUG', targets: [] },
      message: message,
      ttlTurns: ttlTurns,
      payload: { type: 'FACTORY_EVENT', factoryId: factory.id, provinceId: factory.provinceId || null },
    });
  },

  _indexTemplates: function (ctx) {
    const matrix = this._getSubrange(ctx, FACTORY_SYSTEM_CONFIG.templatesRange, FACTORY_SYSTEM_CONFIG.templatesSubrange);
    const index = Object.create(null);
    if (!matrix) return index;

    ctx.helpers.forEachCell(matrix, function (template, row, column) {
      if (template === null || template === undefined) return;
      if (!FactorySystem._isObject(template) || !FactorySystem._isText(template.id)) {
        throw new Error('Invalid factory template at [' + row + ',' + column + '].');
      }
      if (index[template.id]) throw new Error('Duplicate factory template id: ' + template.id + '.');
      index[template.id] = template;
    });
    return index;
  },

  _indexProvinces: function (ctx) {
    const matrix = this._getSubrange(ctx, FACTORY_SYSTEM_CONFIG.provincesRange, FACTORY_SYSTEM_CONFIG.provincesSubrange);
    const index = Object.create(null);
    if (!matrix) return index;

    ctx.helpers.forEachCell(matrix, function (province) {
      if (!FactorySystem._isObject(province) || !FactorySystem._isText(province.id)) return;
      index[province.id] = province;
    });
    return index;
  },

  _getSubrange: function (ctx, rangeName, subrangeName) {
    const container = ctx.data[rangeName];
    return container && container[subrangeName];
  },

  _isObject: function (value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
  },

  _isText: function (value) {
    return typeof value === 'string' && value.trim() !== '';
  },

  _isFiniteNumber: function (value) {
    return typeof value === 'number' && isFinite(value);
  },

  _number: function (value, fallback) {
    return this._isFiniteNumber(value) ? value : fallback;
  },
};
