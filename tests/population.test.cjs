const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const context = { assert, console,
  processProvinceSeasons() {}, processProvinceLandscape() {}, processProvinceClimate() {},
  processProvinceFertility() {}, processProvinceReports() {}
};
vm.createContext(context);
vm.runInContext(['Game_Rules.gs', 'GameTurnEngine.gs', 'Uttilities.gs', 'Population.gs']
  .map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n'), context);
vm.runInContext(`
const make = (id = 'pop_1', overrides = {}) => Object.assign({
  id, provinceId: 1, race: 'Люди', culture: 'italian', religion: 'catholic',
  men: 1000, women: 500, total: 1500, birthRate: 0.035, deathRate: 0.025
}, overrides);
assert.equal(GAME_RULES.population.yearsPerTurn, 1);
const provinces = GamePopulation.provinces([{id: 1}, {id: 2}]);
assert.equal(GamePopulation.nextCount(1000, .035, 0, 1), 1035);
assert.equal(GamePopulation.nextCount(500, .035, 0, 1), 517);
assert.equal(GamePopulation.nextCount(1000, .035, .025, .25), 1002);
assert.equal(GamePopulation.nextCount(1, .035, 0, .25), 1);
assert.equal(GamePopulation.nextCount(1, 0, .01, 1), 0);
assert.equal(GamePopulation.nextCount(100, 0, 9, 1), 0);
assert.throws(() => GamePopulation.validate(make('a', {men: 1.5}), provinces), /целым/);
assert.throws(() => GamePopulation.validate(make('a', {total: 999}), provinces), /total/);
assert.throws(() => GamePopulation.validate(make('a', {provinceId: 3}), provinces), /провинция/);
assert.throws(() => GamePopulation.validate(make('a', {birthRate: -1}), provinces), /birthRate/);
assert.throws(() => GamePopulation.read([[make(), make()]], provinces), /Повторяющийся/);
assert.throws(() => GamePopulation.validate(make('a', {race: ''}), provinces), /race/);
assert.throws(() => GamePopulation.validate(make('a', {race: 42}), provinces), /race/);
const legacy = make('legacy');
delete legacy.race;
const migrated = GamePopulation.read([[legacy]], provinces)[0];
assert.equal(migrated.race, 'Люди');
assert.equal(legacy.race, undefined);
assert.equal(GamePopulation.read([[make('elf', {race:'Эльфы'})]], provinces)[0].race, 'Эльфы');
const g1 = make('a'), g2 = make('b');
const exact = JSON.stringify([g1, g2]).length;
assert.equal(GamePopulation.pack([g1, g2], 1, exact)[0].length, 2);
assert.throws(() => GamePopulation.pack([g1, g2], 1, exact - 1), /Недостаточно места/);
assert.throws(() => GamePopulation.pack([g1], 5, JSON.stringify([g1]).length - 1), /одну ячейку/);
assert.deepEqual(GamePopulation.pack([], 2, 48000), [null, null]);
const source = {Игровые_данные: {Провинции: [{id: 1}], Население: [[make()]]}};
const before = JSON.stringify(source);
processPopulation(source, {config: GAME_TURN_CONFIG});
assert.equal(source.Игровые_данные.Население[0][0].men, 1010);
assert.equal(source.Игровые_данные.Население[0][0].women, 505);
assert.equal(source.Игровые_данные.Население[0][0].total, 1515);
assert.equal(source.Игровые_данные.Население[0][0].birthRate, .035);
// Exact overflow after population growth leaves input column untouched.
const growing = make('g', {men: 99, women: 99, total: 198, birthRate: 1, deathRate: 0});
const originalCells = [[growing]];
const small = {Игровые_данные: {Провинции: [{id:1}], Население: originalCells}};
assert.throws(() => processPopulation(small, {config: {maxJsonCellCharacters: JSON.stringify(originalCells[0]).length}}), /одну ячейку/);
assert.strictEqual(small.Игровые_данные.Население, originalCells);
assert.equal(growing.men, 99);
// Generator uses real decoder, writes population column only, and is repeat-safe.
let writes = 0;
let matrix = [['Провинции','Население','Другие данные'],
  ['{"id":1}', '', 'keep'], ['{"id":2}', '', 'keep']];
const target = {getFormulas: () => [[''],['']], setValues: values => {
  writes++; values.forEach((row,i) => {matrix[i+1][1] = row[0];});
}};
const range = {getValues: () => matrix.map(row => row.slice()), getNumRows: () => 3,
  getNumColumns: () => 3, getRow: () => 1, getColumn: () => 1,
  offset: (row,col,rows,width) => {assert.equal(col,1);assert.equal(width,1);return target;}};
const spreadsheet = {getRangeByName: () => range};
assert.equal(GamePopulation.generate(spreadsheet, POPULATION_TEMPLATE_GENERATOR_CONFIG).created, 2);
assert.equal(writes, 1);
assert.equal(GamePopulation.generate(spreadsheet, POPULATION_TEMPLATE_GENERATOR_CONFIG).created, 0);
assert.equal(writes, 1);
assert.equal(matrix[1][2], 'keep');
const all = GamePopulation.read(matrix.slice(1).map(row => row[1] ? JSON.parse(row[1]) : null), provinces);
assert.equal(all.length, 2);assert.equal(new Set(all.map(g=>g.id)).size, 2);
assert.equal(all[0].birthRate, .035);
assert.equal(all[0].race, 'Люди');
// Generator overflow must fail before any write.
matrix = [['Провинции','Население','Другие данные'], ['{"id":1}', '', 'keep'], ['{"id":2}', '', 'keep']];
const oldLimit = GAME_TURN_CONFIG.maxJsonCellCharacters;
GAME_TURN_CONFIG.maxJsonCellCharacters = 10;
assert.throws(() => GamePopulation.generate(spreadsheet, POPULATION_TEMPLATE_GENERATOR_CONFIG), /одну ячейку/);
assert.equal(writes, 1);assert.equal(matrix[1][1], '');
GAME_TURN_CONFIG.maxJsonCellCharacters = oldLimit;
// Integration: failure reaches fatal journal; no game save, turn unchanged.
let gameWrites = 0, fatal = 0, stateTurn = 5;
GameTurnStateStorage.load = () => ({turn: stateTurn});
GameTurnStateStorage.save = (state, turn) => {stateTurn = turn;};
GameTurnJournalStorage.load = () => ({});
GameTurnJournalStorage.appendFatal = () => {fatal++;};
GameTurnRangeStorage.loadAll = () => ({data: small});
GameTurnRangeStorage.savePrepared = () => {gameWrites++;};
GameTurnContext.create = loaded => ({data: loaded.data, config: {maxJsonCellCharacters: 10},
  runtime: {}, journal: {expire() {},commit() {}}, processedProcessors: [], processorTimings: []});
GameTurnProcessors.run = ctx => processPopulation(ctx.data, ctx);
assert.throws(() => GameTurnEngine._processUnlocked({}), /одну ячейку/);
assert.equal(gameWrites, 0);assert.equal(fatal, 1);assert.equal(stateTurn, 5);
console.log('PASS: demographics, truncation, validation, exact packing boundaries, repeat-safe generation, no writes on overflow, fatal journal and unchanged turn');
`, context);

