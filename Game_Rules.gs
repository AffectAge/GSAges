/**
 * All editable province rules live here.
 *
 * Change values and add or remove entries in this file; the processors do not
 * contain climate thresholds, terrain types, or fertility coefficients.
 */

const GAME_RULES = {
  provinces: {
    rangeName: 'Игровые_данные',
    header: 'Провинции',
    informationHeader: 'Провинции информация',

    fields: {
      climate: 'climate',
      climateLocked: 'climateLocked',
      fertility: 'fertility',
      fertilityLocked: 'fertilityLocked',
      temperature: 'temperatureC',
      precipitation: 'precipitationMm',
      elevation: 'elevationM',
      seaAccess: 'hasSeaAccess',
      terrain: 'terrainType',
      landscape: 'landscapeType',
      pollution: 'pollution',
      radiation: 'radiation',
      latitudeZone: 'latitudeZone',
      season: 'season',
    },
  },

  // These lists are a reference for data entry. Add new names here before
  // using them in province JSON and, when necessary, in the rules below.
  types: {
    terrain: [
      'Суша',
      'Остров',
      'Озеро',
      'Море',
      'Океан',
      'Прибрежные воды',
    ],
    landscape: [
      'Равнина',
      'Холмы',
      'Низкогорье',
      'Горы',
    ],
  },

  // All provinces use the same season. Turn 1 is the first item of the
  // cycle; with turnsPerSeason: 1, every game turn advances the season.
  seasons: {
    turnsPerSeason: 1,
    cycle: ['Весна', 'Лето', 'Осень', 'Зима'],
  },

  climate: {
    // A water province receives this value instead of a land-climate rule.
    waterTerrainTypes: ['Озеро', 'Море', 'Океан', 'Прибрежные воды'],
    waterClimate: 'Водный',
    defaultClimate: 'Не определён',

    // Rules with a smaller priority are checked first. All specified conditions
    // in `when` must match. Omit any condition that should not matter.
    // Available conditions: temperatureMin/Max, precipitationMin/Max,
    // elevationMin/Max, hasSeaAccess, terrainTypes, landscapeTypes,
    // latitudeZones and seasons.
    rules: [
      {
        id: 'HOT_DESERT',
        title: 'Жаркий пустынный',
        priority: 10,
        when: { temperatureMin: 28, precipitationMax: 250 },
      },
      {
        id: 'COLD_DESERT',
        title: 'Холодный пустынный',
        priority: 20,
        when: { temperatureMax: 10, precipitationMax: 250 },
      },
      {
        id: 'ARCTIC',
        title: 'Арктический',
        priority: 30,
        when: { temperatureMax: -10 },
      },
      {
        id: 'SUBARCTIC',
        title: 'Субарктический',
        priority: 40,
        when: { temperatureMin: -10, temperatureMax: 8 },
      },
      {
        id: 'TROPICAL_RAIN',
        title: 'Экваториальный влажный',
        priority: 50,
        when: { temperatureMin: 24, precipitationMin: 1800 },
      },
      {
        id: 'TROPICAL_MONSOON',
        title: 'Тропический муссонный',
        priority: 60,
        when: { temperatureMin: 22, precipitationMin: 1200 },
      },
      {
        id: 'TROPICAL_SAVANNA',
        title: 'Тропический саванный',
        priority: 70,
        when: { temperatureMin: 22, precipitationMin: 500, precipitationMax: 1200 },
      },
      {
        id: 'SUBTROPICAL_HUMID',
        title: 'Субтропический влажный',
        priority: 80,
        when: { temperatureMin: 18, precipitationMin: 800 },
      },
      {
        id: 'MOUNTAIN',
        title: 'Горный',
        priority: 85,
        when: { elevationMin: 1500, temperatureMin: 8, temperatureMax: 20 },
      },
      {
        id: 'MEDITERRANEAN',
        title: 'Средиземноморский',
        priority: 90,
        when: {
          temperatureMin: 14,
          temperatureMax: 24,
          precipitationMin: 300,
          precipitationMax: 900,
          hasSeaAccess: true,
        },
      },
      {
        id: 'SUBTROPICAL_ARID',
        title: 'Субтропический засушливый',
        priority: 100,
        when: { temperatureMin: 18, precipitationMax: 500 },
      },
      {
        id: 'COLD_TEMPERATE',
        title: 'Холодный умеренный',
        priority: 110,
        when: { temperatureMin: 8, temperatureMax: 13 },
      },
      {
        id: 'TEMPERATE_MARITIME',
        title: 'Умеренно-морской',
        priority: 120,
        when: {
          temperatureMin: 8,
          temperatureMax: 18,
          precipitationMin: 400,
          hasSeaAccess: true,
        },
      },
      {
        id: 'TEMPERATE_CONTINENTAL',
        title: 'Умеренно-континентальный',
        priority: 130,
        when: {
          temperatureMin: 8,
          temperatureMax: 18,
          precipitationMin: 400,
          precipitationMax: 1200,
          hasSeaAccess: false,
        },
      },
      {
        id: 'SUBTROPICAL',
        title: 'Субтропический',
        priority: 140,
        when: { temperatureMin: 18 },
      },
      {
        id: 'TEMPERATE',
        title: 'Умеренный',
        priority: 150,
        when: { temperatureMin: 8, temperatureMax: 18 },
      },
    ],
  },

  fertility: {
    // Water provinces do not have agricultural fertility. Their fertility is
    // written as null; resource rules for water can be added separately later.
    nonFertileTerrainTypes: ['Озеро', 'Море', 'Океан', 'Прибрежные воды'],
    min: 0,
    max: 1,

    baseByLandscape: {
      'Равнина': 0.65,
      'Холмы': 0.45,
      'Низкогорье': 0.35,
      'Горы': 0.20,
    },
    defaultBase: 0.50,

    temperature: {
      optimalMin: 15,
      optimalMax: 30,
      optimalBonus: 10.0,
      penaltyPerDegreeOutside: 1.0,
    },
    precipitation: {
      optimalMin: 500,
      optimalMax: 1200,
      optimalBonus: 10.0,
      penaltyPer100MmOutside: 1.0,
    },
    elevation: {
      freeElevationM: 250,
      penaltyPer1000MAboveFree: 20.0,
    },
    seasonBonus: {
      'Весна': 0.08,
      'Лето': 0.12,
      'Осень': 0.03,
      'Зима': -0.18,
    },

    // Pollution and radiation are used as raw values. The coefficients below
    // define how much one raw unit lowers fertility.
    pollutionPenaltyPerUnit: 0.01,
    radiationPenaltyPerUnit: 0.05,
  },
};
