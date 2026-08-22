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
  /** Order types owned by this mechanic. OrderSystem invokes these handlers. */
  getOrderDefinitions: function () {
    return [{
      type: 'BUILD_FACTORY',
      category: 'CONSTRUCTION',
      phase: 'CONSTRUCTION',
      priority: 500,
      validate: function (ctx, order) {
        return FactorySystem.validateBuildFactoryOrder(ctx, order);
      },
      execute: function (ctx, order) {
        return FactorySystem.executeBuildFactoryOrder(ctx, order);
      },
    }];
  },

  /**
   * The first order implementation. Costs and construction materials are not
   * reserved yet: the order checks ownership, template and free factory slot.
   */
  validateBuildFactoryOrder: function (ctx, order) {
    const payload = order.payload || {};
    const templateId = this._isText(payload.templateId) ? payload.templateId.trim() : null;
    const provinceId = this._isText(payload.provinceId) ? payload.provinceId.trim() : null;
    const level = payload.level === undefined ? 1 : payload.level;
    if (!templateId) return 'Не указан templateId фабрики.';
    if (!provinceId) return 'Не указана provinceId для строительства.';
    if (!this._isFiniteNumber(level) || !Number.isInteger(level) || level < 1) {
      return 'Уровень фабрики должен быть целым числом не меньше 1.';
    }

    const templates = this._indexTemplates(ctx);
    if (!templates[templateId]) return 'Шаблон фабрики «' + templateId + '» не найден.';

    const provinces = this._indexProvinces(ctx);
    const province = provinces[provinceId];
    if (!province) return 'Провинция «' + provinceId + '» не найдена.';
    if (province.owner !== order.issuer) {
      return 'Провинция «' + provinceId + '» не принадлежит стране «' + order.issuer + '».';
    }

    const factories = this._getSubrange(ctx, FACTORY_SYSTEM_CONFIG.factoriesRange, FACTORY_SYSTEM_CONFIG.factoriesSubrange);
    if (!factories || !this._findFirstEmptyCell(factories)) return 'Нет свободного места в контейнере фабрик.';
    const factoryId = this._factoryIdForOrder(order);
    if (this._containsId(factories, factoryId)) {
      return 'Фабрика для приказа «' + order.id + '» уже существует.';
    }
    return { ok: true };
  },

  executeBuildFactoryOrder: function (ctx, order) {
    const factories = this._getSubrange(ctx, FACTORY_SYSTEM_CONFIG.factoriesRange, FACTORY_SYSTEM_CONFIG.factoriesSubrange);
    const target = factories && this._findFirstEmptyCell(factories);
    if (!target) return { ok: false, message: 'Нет свободного места в контейнере фабрик.' };

    const payload = order.payload;
    const template = this._indexTemplates(ctx)[payload.templateId];
    const constructionTurns = Math.max(1, Math.floor(this._number(template.constructionTurns, 1)));
    const factoryId = this._factoryIdForOrder(order);
    factories[target.row][target.column] = {
      id: factoryId,
      templateId: payload.templateId,
      owner: order.issuer,
      provinceId: payload.provinceId,
      level: payload.level === undefined ? 1 : payload.level,
      efficiency: 1,
      stockpile: {},
      status: 'CONSTRUCTING',
      constructionStartedTurn: ctx.turn,
      constructionTurnsRemaining: constructionTurns,
    };
    return {
      ok: true,
      message: 'Строительство фабрики «' + factoryId + '» начато. Срок: ' + constructionTurns + ' ход(а).',
    };
  },

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
    const target = this._findFirstEmptyCell(matrix);
    if (target) {
      matrix[target.row][target.column] = value;
      return;
    }
    throw new Error('No empty cell is available for the default ' + label + '.');
  },

  _findFirstEmptyCell: function (matrix) {
    for (let row = 0; row < matrix.length; row += 1) {
      for (let column = 0; column < matrix[row].length; column += 1) {
        if (matrix[row][column] === null || matrix[row][column] === undefined) {
          return {
            row: row,
            column: column,
          };
        }
      }
    }
    return null;
  },

  _factoryIdForOrder: function (order) {
    const country = String(order.issuer || 'country').toLowerCase().replace(/[^a-z0-9_-]+/g, '_');
    const source = String(order.clientOrderId || order.id || 'order').replace(/[^a-zA-Z0-9_-]+/g, '_');
    return 'factory_' + country + '_' + source;
  },

  _processFactory: function (ctx, factory, templates, provinces, report, row, column) {
    if (!this._isText(factory.id)) {
      ctx.log.error('Factory has no id.', { row: row, column: column });
      return;
    }
    this._ensureFactoryDefaults(factory);
    if (!this._isText(factory.templateId)) {
      this._setOperationalStatus(
        ctx,
        factory,
        'TEMPLATE_MISSING',
        'Для фабрики «' + factory.id + '» не выбран шаблон.',
        'TEMPLATE'
      );
      return;
    }

    const template = templates[factory.templateId];
    if (!template) {
      this._setOperationalStatus(
        ctx,
        factory,
        'TEMPLATE_MISSING',
        'Фабрика «' + factory.id + '» остановлена: шаблон «' + factory.templateId + '» не найден.',
        'TEMPLATE'
      );
      return;
    }

    if (factory.status === 'CONSTRUCTING') {
      if (factory.constructionStartedTurn === ctx.turn) return;
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
      this._setOperationalStatus(
        ctx,
        factory,
        'INPUT_SHORTAGE',
        'Фабрика «' + factory.id + '» остановлена: на её складе не хватает сырья.',
        'INPUTS'
      );
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
    this._emit(ctx, factory, 'SUCCESS', 'Строительство фабрики «' + factory.id + '» завершено.', 2, 'CONSTRUCTION');
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
      this._setOperationalStatus(
        ctx,
        factory,
        'PROVINCE_MISSING',
        'Фабрика «' + factory.id + '» не может работать: провинция «' + factory.provinceId + '» не найдена.',
        'PROVINCE'
      );
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

  _setOperationalStatus: function (ctx, factory, status, message, subCategory) {
    const changed = factory.operationalStatus !== status;
    factory.operationalStatus = status;
    if (changed) this._emit(ctx, factory, 'HIGH', message, 2, subCategory || 'FACTORY');
  },

  _emit: function (ctx, factory, priority, message, ttlTurns, subCategory) {
    const owner = this._isText(factory.owner) ? factory.owner : null;
    ctx.journal.emit({
      country: owner,
      category: FACTORY_SYSTEM_CONFIG.category,
      subCategory: subCategory || 'FACTORY',
      priority: priority,
      visibility: owner
        ? { type: 'COUNTRY', targets: [owner] }
        : { type: 'DEBUG', targets: [] },
      message: message,
      ttlTurns: ttlTurns,
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
