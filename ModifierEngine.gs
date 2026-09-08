/**
 * Generic extensible modifier engine for GSAges.
 * Entity types are not hard-coded: provinces, regions, states, countries,
 * markets and future entities are resolved through relations.
 */
const MODIFIER_ENGINE_CONFIG = {
  fields: { id: 'id', type: 'type', relations: 'relations', activeModifiers: 'activeModifiers' },
  relationScopes: [
    { id: 'SELF', priority: 100, relation: null },
    { id: 'COUNTRY', priority: 30, relation: 'country' },
    { id: 'REGION', priority: 50, relation: 'region' },
    { id: 'STATE', priority: 40, relation: 'state' },
    { id: 'MARKET', priority: 20, relation: 'market' },
  ],
};
const GAME_MODIFIERS = {};

const ModifierEngine = {
  create: function (data, config) {
    const cfg = config || MODIFIER_ENGINE_CONFIG;
    return {
      calculate: function (request) { return ModifierEngine.calculate(request, data, cfg); },
      explain: function (request) { return ModifierEngine.explain(request, data, cfg); },
      getModifiers: function (request) { return ModifierEngine.getApplicableEffects(request, data, cfg); },
      has: function (entity, modifierId) { return ModifierEngine.has(entity, modifierId, cfg); },
    };
  },

  calculate: function (request, data, config) {
    return this._calculate(request, data, config).value;
  },

  explain: function (request, data, config) {
    const result = this._calculate(request, data, config);
    return {
      stat: request.stat, base: request.baseValue, final: result.value,
      modifiers: result.effects.map(function (effect) {
        return { id: effect.modifierId, source: effect.source, operation: effect.operation,
          value: effect.value, priority: effect.priority };
      }),
    };
  },

  _calculate: function (request, data, config) {
    this._assertRequest(request);
    const effects = this.getApplicableEffects(request, data, config);
    let value = request.baseValue;
    effects.filter(function (e) { return e.operation === 'ADD'; }).forEach(function (e) { value += e.value; });
    const percent = effects.filter(function (e) { return e.operation === 'ADD_PERCENT'; });
    if (percent.length) value *= 1 + percent.reduce(function (sum, e) { return sum + e.value; }, 0);
    effects.filter(function (e) { return e.operation === 'MULTIPLY'; }).forEach(function (e) { value *= e.value; });

    const sets = effects.filter(function (e) { return e.operation === 'SET'; });
    if (sets.length) {
      sets.sort(function (a, b) { return b.priority - a.priority || b.scopePriority - a.scopePriority; });
      value = sets[0].value;
    }
    effects.filter(function (e) { return e.operation === 'MIN'; }).forEach(function (e) { value = Math.max(value, e.value); });
    effects.filter(function (e) { return e.operation === 'MAX'; }).forEach(function (e) { value = Math.min(value, e.value); });
    return { value: value, effects: effects };
  },

  getApplicableEffects: function (request, data, config) {
    this._assertRequest(request);
    const effects = [];
    this.resolveSources(request.entity, data, config).forEach(function (source) {
      ModifierEngine._modifierIds(source.entity, config).forEach(function (modifierId) {
        const definition = GAME_MODIFIERS[modifierId];
        if (!definition) throw new Error('Не найдено определение модификатора «' + modifierId + '».');
        (definition.effects || []).forEach(function (effect, index) {
          if (!effect || effect.stat !== request.stat) return;
          ModifierEngine._assertEffect(effect, modifierId, index);
          effects.push({
            modifierId: modifierId,
            source: { scope: source.scope, entityId: ModifierEngine._entityId(source.entity, config),
              entityType: source.entity[config.fields.type] || null },
            operation: effect.operation, value: effect.value,
            priority: Number(effect.priority || 0), scopePriority: source.priority,
          });
        });
      });
    });
    return effects;
  },

  resolveSources: function (entity, data, config) {
    if (!entity || typeof entity !== 'object') throw new Error('Modifier Engine: entity должен быть объектом.');
    const result = [];
    const seen = Object.create(null);
    const push = function (scope, priority, source) {
      if (!source || typeof source !== 'object') return;
      const key = scope + ':' + (ModifierEngine._entityId(source, config) || JSON.stringify(source));
      if (seen[key]) return;
      seen[key] = true;
      result.push({ scope: scope, priority: priority, entity: source });
    };
    config.relationScopes.forEach(function (scope) {
      if (scope.id === 'SELF') push(scope.id, scope.priority, entity);
      else {
        const relations = entity[config.fields.relations] || {};
        push(scope.id, scope.priority,
          ModifierEngine._resolveEntityReference(relations[scope.relation], data, config));
      }
    });
    return result;
  },

  _resolveEntityReference: function (reference, data, config) {
    if (!reference) return null;
    if (typeof reference === 'object') return reference;
    const id = String(reference);
    Object.keys(data || {}).some(function (rangeName) {
      return Object.keys(data[rangeName] || {}).some(function (header) {
        const values = data[rangeName][header] || [];
        const found = values.filter(function (candidate) {
          return candidate && typeof candidate === 'object' &&
            String(candidate[config.fields.id]) === id;
        })[0];
        if (found) { reference = found; return true; }
        return false;
      });
    });
    return typeof reference === 'object' ? reference : null;
  },

  has: function (entity, modifierId, config) {
    return this._modifierIds(entity, config).indexOf(modifierId) !== -1;
  },

  _modifierIds: function (entity, config) {
    const values = entity && entity[config.fields.activeModifiers];
    if (values === null || values === undefined) return [];
    if (!Array.isArray(values)) throw new Error('activeModifiers у сущности должен быть массивом ID модификаторов.');
    return values.map(function (value) { return typeof value === 'string' ? value : value && value.id; })
      .map(function (id) {
        if (typeof id !== 'string' || !id.trim()) throw new Error('Каждый activeModifiers должен содержать непустой ID.');
        return id;
      });
  },

  _entityId: function (entity, config) {
    return entity && entity[config.fields.id] !== undefined && entity[config.fields.id] !== null
      ? String(entity[config.fields.id]) : null;
  },

  _assertRequest: function (request) {
    if (!request || typeof request !== 'object' || !request.entity || typeof request.entity !== 'object')
      throw new Error('Modifier Engine: request.entity обязателен.');
    if (typeof request.stat !== 'string' || !request.stat.trim()) throw new Error('Modifier Engine: request.stat обязателен.');
    if (typeof request.baseValue !== 'number' || !isFinite(request.baseValue))
      throw new Error('Modifier Engine: request.baseValue должен быть конечным числом.');
  },

  _assertEffect: function (effect, modifierId, index) {
    if (['ADD', 'ADD_PERCENT', 'MULTIPLY', 'SET', 'MIN', 'MAX'].indexOf(effect.operation) === -1)
      throw new Error('Модификатор «' + modifierId + '», эффект #' + index + ': неизвестная operation.');
    if (typeof effect.value !== 'number' || !isFinite(effect.value))
      throw new Error('Модификатор «' + modifierId + '», эффект #' + index + ': value должен быть числом.');
  },
};