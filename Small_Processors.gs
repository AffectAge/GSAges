/** Sets one shared season for every province from the current game turn. */

function processProvinceSeasons(data, ctx) {
  const rules = GAME_RULES;
  const provinces = ctx.getColumn(rules.provinces.rangeName, rules.provinces.header);
  const season = getSeasonForTurn(ctx.turn, rules.seasons);

  provinces.forEach(function (province) {
    if (ProvinceRuleUtils.isBlank(province)) return;
    ProvinceRuleUtils.assertProvince(province, rules.provinces);
    province[rules.provinces.fields.season] = season;
  });

  return data;
}

function getSeasonForTurn(turn, settings) {
  const cycle = settings && settings.cycle;
  const turnsPerSeason = Number(settings && settings.turnsPerSeason);
  if (!Array.isArray(cycle) || !cycle.length || cycle.some(function (season) {
    return typeof season !== 'string' || !season.trim();
  })) {
    throw new Error('GAME_RULES.seasons.cycle должен содержать непустые названия сезонов.');
  }
  if (!Number.isInteger(turnsPerSeason) || turnsPerSeason < 1) {
    throw new Error('GAME_RULES.seasons.turnsPerSeason должен быть целым числом не меньше 1.');
  }
  if (!Number.isInteger(turn) || turn < 1) {
    throw new Error('Номер хода должен быть целым числом не меньше 1 для расчёта сезона.');
  }
  return cycle[Math.floor((turn - 1) / turnsPerSeason) % cycle.length];
}

/** Calculates province.fertility from the editable GAME_RULES.fertility rules. */

function processProvinceFertility(data, ctx) {
  const rules = GAME_RULES;
  const provinces = ctx.getColumn(rules.provinces.rangeName, rules.provinces.header);

  provinces.forEach(function (province) {
    if (ProvinceRuleUtils.isBlank(province)) return;
    ProvinceRuleUtils.assertProvince(province, rules.provinces);

    const fields = rules.provinces.fields;
    if (province[fields.fertilityLocked] === true) return;
    if (ProvinceRuleUtils.includes(rules.fertility.nonFertileTerrainTypes, province[fields.terrain])) {
      province[fields.fertility] = null;
      return;
    }

    province[fields.fertility] = calculateProvinceFertility(province, rules);
  });

  return data;
}

function calculateProvinceFertility(province, rules) {
  const fertilityRules = rules.fertility;
  const fields = rules.provinces.fields;
  const temperature = ProvinceRuleUtils.number(province[fields.temperature]);
  const precipitation = ProvinceRuleUtils.number(province[fields.precipitation]);
  const elevation = ProvinceRuleUtils.number(province[fields.elevation]);
  const pollution = ProvinceRuleUtils.number(province[fields.pollution]);
  const radiation = ProvinceRuleUtils.number(province[fields.radiation]);
  const landscape = province[fields.landscape];
  const season = province[fields.season];

  let result = Object.prototype.hasOwnProperty.call(fertilityRules.baseByLandscape, landscape) ?
    fertilityRules.baseByLandscape[landscape] : fertilityRules.defaultBase;

  result += calculateOptimalRangeAdjustment(temperature, fertilityRules.temperature, 1, 'penaltyPerDegreeOutside');
  result += calculateOptimalRangeAdjustment(precipitation, fertilityRules.precipitation, 100, 'penaltyPer100MmOutside');
  result += getSeasonFertilityBonus(season, fertilityRules.seasonBonus);

  if (elevation !== null && elevation > fertilityRules.elevation.freeElevationM) {
    result -= ((elevation - fertilityRules.elevation.freeElevationM) / 1000) *
      fertilityRules.elevation.penaltyPer1000MAboveFree;
  }
  if (pollution !== null) result -= pollution * fertilityRules.pollutionPenaltyPerUnit;
  if (radiation !== null) result -= radiation * fertilityRules.radiationPenaltyPerUnit;

  return ProvinceRuleUtils.clamp(result, fertilityRules.min, fertilityRules.max);
}

function getSeasonFertilityBonus(season, bonuses) {
  if (!bonuses || !Object.prototype.hasOwnProperty.call(bonuses, season)) return 0;
  const value = bonuses[season];
  if (typeof value !== 'number' || !isFinite(value)) {
    throw new Error('Значение GAME_RULES.fertility.seasonBonus для сезона «' + season + '» должно быть числом.');
  }
  return value;
}

function calculateOptimalRangeAdjustment(value, settings, unit, penaltyKey) {
  if (value === null) return 0;
  if (value >= settings.optimalMin && value <= settings.optimalMax) return settings.optimalBonus;
  const distance = value < settings.optimalMin ? settings.optimalMin - value : value - settings.optimalMax;
  return -((distance / unit) * settings[penaltyKey]);
}

/** Calculates province.climate from the editable GAME_RULES.climate rules. */

function processProvinceClimate(data, ctx) {
  const rules = GAME_RULES;
  const provinces = ctx.getColumn(rules.provinces.rangeName, rules.provinces.header);

  provinces.forEach(function (province) {
    if (ProvinceRuleUtils.isBlank(province)) return;
    ProvinceRuleUtils.assertProvince(province, rules.provinces);

    const fields = rules.provinces.fields;
    if (province[fields.climateLocked] === true) return;

    const terrain = province[fields.terrain];
    if (ProvinceRuleUtils.includes(rules.climate.waterTerrainTypes, terrain)) {
      province[fields.climate] = rules.climate.waterClimate;
      return;
    }

    const matchingRule = ProvinceRuleUtils.firstMatchingClimateRule(province, rules);
    province[fields.climate] = matchingRule ? matchingRule.title : rules.climate.defaultClimate;
  });

  return data;
}

const ProvinceRuleUtils = {
  isBlank: function (value) {
    return value === null || value === undefined || value === '';
  },

  assertProvince: function (province, provinceRules) {
    if (province === null || typeof province !== 'object' || Array.isArray(province) || province instanceof Date) {
      throw new Error('Каждая непустая ячейка «' + provinceRules.rangeName + '.' + provinceRules.header + '» должна содержать JSON-объект провинции.');
    }
  },

  includes: function (values, value) {
    return Array.isArray(values) && values.indexOf(value) !== -1;
  },

  number: function (value) {
    return typeof value === 'number' && isFinite(value) ? value : null;
  },

  firstMatchingClimateRule: function (province, rules) {
    return (rules.climate.rules || [])
      .map(function (rule, index) { return { rule: rule, index: index }; })
      .sort(function (left, right) {
        return Number(left.rule.priority || 0) - Number(right.rule.priority || 0) || left.index - right.index;
      })
      .map(function (item) { return item.rule; })
      .filter(function (rule) { return ProvinceRuleUtils.matchesClimateRule(province, rule.when || {}, rules.provinces.fields); })[0] || null;
  },

  matchesClimateRule: function (province, when, fields) {
    const temperature = ProvinceRuleUtils.number(province[fields.temperature]);
    const precipitation = ProvinceRuleUtils.number(province[fields.precipitation]);
    const elevation = ProvinceRuleUtils.number(province[fields.elevation]);

    if (!ProvinceRuleUtils.matchesNumber(temperature, when.temperatureMin, when.temperatureMax)) return false;
    if (!ProvinceRuleUtils.matchesNumber(precipitation, when.precipitationMin, when.precipitationMax)) return false;
    if (!ProvinceRuleUtils.matchesNumber(elevation, when.elevationMin, when.elevationMax)) return false;
    if (when.hasSeaAccess !== undefined && province[fields.seaAccess] !== when.hasSeaAccess) return false;
    if (when.terrainTypes && !ProvinceRuleUtils.includes(when.terrainTypes, province[fields.terrain])) return false;
    if (when.landscapeTypes && !ProvinceRuleUtils.includes(when.landscapeTypes, province[fields.landscape])) return false;
    if (when.latitudeZones && !ProvinceRuleUtils.includes(when.latitudeZones, province[fields.latitudeZone])) return false;
    if (when.seasons && !ProvinceRuleUtils.includes(when.seasons, province[fields.season])) return false;
    return true;
  },

  matchesNumber: function (value, min, max) {
    if (min === undefined && max === undefined) return true;
    if (value === null) return false;
    return (min === undefined || value >= min) && (max === undefined || value <= max);
  },

  clamp: function (value, min, max) {
    return Math.max(min, Math.min(max, value));
  },
};
