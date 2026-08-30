/**
 * Human-readable province information for GameTurnEngine.gs.
 *
 * Reads Игровые_данные.Провинции and writes readable text into the paired
 * Игровые_данные.Провинции информация column. It never edits province JSON.
 */

const PROVINCE_REPORT_CONFIG = {
  countryNames: {
    ITA: 'Италия',
    GER: 'Германия',
    FRA: 'Франция',
  },
};

function processProvinceReports(data, ctx) {
  const config = getProvinceReportConfig();
  const provinces = ctx.getColumn(config.rangeName, config.inputHeader);
  const provinceInformation = ctx.getColumn(config.rangeName, config.outputHeader);

  provinces.forEach(function (province, rowIndex) {
    // Removing a province also removes its obsolete human-readable card.
    if (province === null || province === undefined) {
      provinceInformation[rowIndex] = null;
      return;
    }
    if (!isProvinceReportObject(province)) {
      throw new Error('Каждая непустая ячейка «' + config.rangeName + '.' + config.inputHeader + '» должна содержать JSON-объект провинции.');
    }

    const ownerId = province.ownerId || '—';
    provinceInformation[rowIndex] = buildProvinceInformation(province, ownerId, config);
  });

  return data;
}

function getProvinceReportConfig() {
  return {
    rangeName: GAME_RULES.provinces.rangeName,
    inputHeader: GAME_RULES.provinces.header,
    outputHeader: GAME_RULES.provinces.informationHeader,
    countryNames: PROVINCE_REPORT_CONFIG.countryNames,
  };
}

function buildProvinceInformation(province, ownerId, config) {
  const ownerName = config.countryNames[ownerId] || ownerId;
  const resources = formatProvinceResources(province.resources);
  const neighbors = formatProvinceNeighbors(province.neighbors);
  const fertility = formatProvincePercent(province.fertility);
  const pollution = valueOrDash(province.pollution);
  const radiation = valueOrDash(province.radiation);
  const seaAccess = province.hasSeaAccess === true ? 'есть выход к морю' : 'нет выхода к морю';
  const isWater = ProvinceRuleUtils.includes(GAME_RULES.landscape.waterTerrainTypes, province.terrainType);
  const landscape = isWater ? 'не применяется' : valueOrDash(province.landscapeType);

  return '🗺️ ' + (province.name || ('Провинция ' + province.id)) + ' (ID ' + valueOrDash(province.id) + ')\n' +
    '🏛️ Владелец: ' + ownerName + ' [' + ownerId + ']\n' +
    '🌍 ' + valueOrDash(province.continent) + ', ' + valueOrDash(province.planet) +
      ' · ' + valueOrDash(province.latitudeZone) + ' · ' + valueOrDash(province.season) + '\n' +
    '🌦️ ' + valueOrDash(province.climate) + ' · 🌡️ ' + valueOrDash(province.temperatureC) +
      ' °C · 🌧️ ' + valueOrDash(province.precipitationMm) + ' мм/год\n' +
    '🏞️ ' + valueOrDash(province.terrainType) + ' · ' + landscape +
      ' · ⛰️ ' + valueOrDash(province.elevationM) + ' м · 🌊 ' + seaAccess + '\n' +
    '📐 Площадь: ' + valueOrDash(province.areaKm2) + ' км² · 🌾 Плодородие: ' + fertility + '\n' +
    '⛏️ Ресурсы: ' + resources + '\n' +
    '⚠️ Загрязнение: ' + pollution + ' · ☢️ Радиация: ' + radiation + '\n' +
    '🤝 Соседи: ' + neighbors;
}

function formatProvinceResources(resources) {
  if (!Array.isArray(resources) || !resources.length) return 'не указаны';
  return resources.map(function (resource) {
    if (!resource || typeof resource !== 'object') return 'неизвестный ресурс';
    return valueOrDash(resource.type) + ' — ' + valueOrDash(resource.amount);
  }).join('; ');
}

function formatProvinceNeighbors(neighbors) {
  if (!Array.isArray(neighbors) || !neighbors.length) return 'нет данных';
  return neighbors.map(function (id) { return String(id); }).join(', ');
}

function formatProvincePercent(value) {
  if (typeof value !== 'number' || !isFinite(value)) return 'не указано';
  return Math.round(value * 10000) / 100 + '%';
}

function valueOrDash(value) {
  return value === null || value === undefined || value === '' ? 'не указано' : String(value);
}

function isProvinceReportObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}
