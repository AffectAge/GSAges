# Google Sheets turn engine

`GameEngine.gs` is a ready-to-paste Google Apps Script core for a turn-based game stored in Google Sheets. It reads named ranges into `ctx.data`, runs all game systems in memory, validates the result, then writes the named ranges back in batches.

## Quick start

1. Open the target spreadsheet: **Extensions → Apps Script**.
2. Create script files named `GameEngine.gs`, `OrderSystem.gs`, and `FactorySystem.gs`, then paste in [GameEngine.gs](GameEngine.gs), [OrderSystem.gs](OrderSystem.gs), and [FactorySystem.gs](FactorySystem.gs).
3. By default, every missing active range is created on its own sheet at the start of `PROCESS_TURN`: `_GAME_CORE`, `_GAME_WORLD`, `_GAME_ORDERS`, `_GAME_FACTORIES`, `_GAME_COUNTRIES`, and `_GAME_JOURNAL`.
4. `NR_GAME_CORE` is initialized with technical headers `MAIN`, `FACTORY_TEMPLATES`, and `COUNTRY_BOOKS`. The template column contains the editable base template `TEXTILE_MILL`; an older container receives a missing technical column on the next `PROCESS_TURN`, without replacing existing data. `_GAME_FACTORIES` likewise receives the missing `FACTORIES` column, but starts with no factory instances.

   ```json
   {"turn":1,"status":"WAITING"}
   ```

5. Adjust `rangeDefaults` if `NR_JOURNAL` needs another capacity, or if a new game range needs a particular size/sheet/cell position.
6. Add your game ranges to `GAME_ENGINE_CONFIG.namedRanges`, add their automatic-creation settings to `rangeDefaults` if needed, define handlers and register them in `GAME_ENGINE_CONFIG.systems`.
7. Run `PROCESS_TURN` once from the Apps Script editor to grant permissions. After reloading the sheet, the **Game engine → Process turn** menu will be available.

## Mobile turn button

Once, under the central script owner account, run **Game engine → Enable mobile turn button** (or `INSTALL_MOBILE_TURN_TRIGGER()` in the Apps Script editor) and grant the requested permissions. The engine creates a checkbox in the second data cell of `NR_GAME_CORE.MAIN`: `MAIN[1][0]` (normally cell `A3` on `_GAME_CORE`).

From the Google Sheets phone app, tap that checkbox. It changes to `TRUE`, starts one game turn under the owner account, and automatically returns to `FALSE` only after the central game state was saved successfully. If the turn fails, the checkbox remains selected; after fixing the problem, tap it off and on again to retry. The owner must keep the installable edit trigger enabled.

## Data model

Every named range is a matrix, so cell coordinates cannot be lost:

```javascript
ctx.data.NR_UNITS[2][0] // always the third row, first column of NR_UNITS
```

Blank cells become `null`. JSON objects and arrays are decoded automatically; normal strings, numbers, booleans, and dates stay as their native values. During saving, objects/arrays are serialized back to JSON and `null` becomes a blank cell.

The engine currently reads values, not formula definitions. Do not place formulas in a writable engine range. Put calculated/formula ranges in `readOnlyRanges` instead.

## Containers and country data

A container is one physical named range read in one `getValues()` call and written in at most one `setValues()` call. Its technical headers create virtual data areas in `ctx.data`, avoiding a named range and a Sheets API call for every mechanic. The engine's active data layout is:

```text
NR_GAME_CORE: MAIN | FACTORY_TEMPLATES | COUNTRY_BOOKS
NR_WORLD:     PROVINCES
NR_ORDERS:    ACTIVE | HISTORY
NR_FACTORIES: FACTORIES
NR_COUNTRIES: RUS | FRA | GER
NR_JOURNAL:   ХОД | ВЕДОМСТВО | ТЕМА | СТРАНА | ВАЖНОСТЬ | ДОСТУП | СРОК (ХОДОВ) | СВОДКА | ИД
```

`NR_GAME_CORE.MAIN` contains game metadata, so its turn is available as `ctx.data.NR_GAME_CORE.MAIN[0][0].turn`. `NR_WORLD.PROVINCES` is the province matrix used by the province generator.

`NR_FACTORIES.FACTORIES` contains one factory JSON object per cell and has 10,000 data slots (plus its technical header) on `_GAME_FACTORIES`. A factory's input and output goods belong in its own `stockpile` object rather than in a country-level stockpile.

## Factory system

[FactorySystem.gs](FactorySystem.gs) is a separate turn-mechanics file and is already registered in `GAME_ENGINE_CONFIG.systems` as `FACTORIES`. Add both script files to the Apps Script project. The system uses only the factory's own `stockpile`: it consumes inputs, adds outputs to that stockpile, scales production by level and efficiency, handles construction completion, and applies factory pollution to its province. It does not use a country stockpile, workers, or market.

`NR_GAME_CORE.FACTORY_TEMPLATES` is created automatically and begins with this editable template:

```json
{"id":"TEXTILE_MILL","inputs":{"cotton":2,"coal":0.1},"outputs":{"clothes":1},"productionPerLevel":1,"constructionTurns":3,"pollutionPerCycle":2}
```

Put factory instances in `NR_FACTORIES.FACTORIES`:

```json
{"id":"factory_rus_1","templateId":"TEXTILE_MILL","owner":"RUS","provinceId":"prov_1","level":1,"efficiency":1,"stockpile":{"cotton":20,"coal":3,"clothes":0},"status":"ACTIVE"}
```

For every production cycle, input quantities are removed from the factory stockpile and output quantities are added to that same object. `CONSTRUCTING`, `ACTIVE`, and `PAUSED` are supported statuses. Warnings about missing templates or input goods and construction-completion events go to the game journal. Old `workers` and `workersPerLevel` fields, if present, are ignored and left unchanged.

## Orders and country workbooks

[OrderSystem.gs](OrderSystem.gs) holds the universal order queue in the central `_GAME_ORDERS` sheet. `ACTIVE` is the pending queue; `HISTORY` contains orders that have finished, been rejected, failed, or were cancelled. Each non-empty cell is a JSON array of at most 20 orders. There is deliberately no separate character-count limit in the engine.

For a country workbook, add one JSON record to a free cell in the central `NR_GAME_CORE.COUNTRY_BOOKS` column:

```json
{"id":"RUS","spreadsheetId":"GOOGLE_SPREADSHEET_ID"}
{"id":"FRA","spreadsheetId":"ANOTHER_GOOGLE_SPREADSHEET_ID"}
```

`id` is the country identifier; it must be unique. `spreadsheetId` is the ID from the country book URL; it must also be unique. Optional technical fields are `ordersRange`, `activeHeader`, and `ordersSheetName`. On the first `PROCESS_TURN`, the central engine opens every registered workbook. If its `NR_ORDERS` named range is absent, the engine automatically creates a separate `_ORDERS` sheet with a 1,001 × 2 `NR_ORDERS` container and headers `ACTIVE | HISTORY`. The menu item **Game engine → Create country order containers** creates the same missing containers immediately. The central script account needs editor access to each country workbook.

Players put only their client order into a cell of their workbook's `ACTIVE` column. The server obtains the country ID from `NR_GAME_CORE.COUNTRY_BOOKS`, builds the authoritative ID `<country>:<clientOrderId>`, and ignores duplicate deliveries safely:

```json
[
  {
    "clientOrderId": "build_textile_1",
    "type": "BUILD_FACTORY",
    "payload": {
      "templateId": "TEXTILE_MILL",
      "provinceId": "prov_1",
      "level": 1
    }
  }
]
```

`BUILD_FACTORY` checks the template, a free factory slot, and that the named province belongs to the issuing country. It creates a `CONSTRUCTING` factory with its own empty `stockpile`; the factory starts construction on the current turn and loses its first construction turn only on the following turn. Costs and construction materials are intentionally not reserved yet, so later economic mechanics can add those rules to the same order handler.

After the central game state has been saved successfully, the engine removes accepted orders from the source country's `ACTIVE` cells. Invalid orders stay in the player book so they can be corrected. If the central turn or source-book acknowledgement fails, the source order is left intact; its authoritative ID makes the next import safe from duplicate execution.

`NR_COUNTRIES` is configured as a country container. The first row holds technical country IDs; every following non-empty cell must be a small JSON object with a unique `key` in that country's column:

```text
RUS                                           | FRA
{"key":"CORE","name":"Russia"}           | {"key":"CORE","name":"France"}
{"key":"ECONOMY","treasury":1000}         | {"key":"ECONOMY","treasury":800}
{"key":"MILITARY","manpower":50000}       |
```

The engine exposes this as an indexed view, independent of physical row order:

```javascript
ctx.data.NR_COUNTRIES.RUS.ECONOMY.treasury += 100;
ctx.data.NR_COUNTRIES.FRA.MILITARY = { manpower: 42000 };
ctx.data.NR_COUNTRIES.GER = {
  CORE: { name: 'Germany' },
  ECONOMY: { treasury: 700 },
};
```

The engine automatically adds `key` to a newly assigned record, uses a blank row for a new record, and a blank column for a new country. Extend the physical named range if it has no room. Duplicate country headers or duplicate `key` values in one country cause an error instead of ambiguous data.

Both `NR_GAME_CORE` and `NR_WORLD` use `COLUMN_MATRICES`: a headed column such as `MAIN` appears as `ctx.data.NR_GAME_CORE.MAIN`, with the same one-column matrix structure as the old named-range data.

## Province generator

Run `GENERATE_PROVINCES()` to fill every blank cell in `NR_WORLD.PROVINCES` with a base JSON province and to add missing properties to existing JSON provinces. Existing values are never overwritten, including `0`, `false`, `null`, and nested resource quantities. The schema is editable in `PROVINCE_GENERATOR_CONFIG.defaults`; adding a field there and running the generator again performs a deep, missing-fields-only update.

The initial schema contains `id`, `name`, `owner`, `terrain`, `elevation`, radiation, pollution, temperature, humidity, area, `soilFertility` (soil quality/yield potential), `fertileLandPercent` (the percentage of province area suitable for agriculture), `resources.water.amount`, and `neighbors`. IDs and names are generated automatically. `prov_1` is the default neighbour for every other province; change `defaultNeighborId` in the generator config to use another anchor province.

The `NR_WORLD` container should have enough data rows for province slots and other world subranges. It starts with one `PROVINCES` data cell; increase its physical row count before first use when more provinces are needed. A non-blank province cell must contain a JSON object; ordinary text and JSON arrays are rejected to avoid silently destroying data.

## Adding a game system

```javascript
function processEconomy(ctx) {
  ctx.helpers.forEachCell(ctx.data.NR_PLAYERS, function (player) {
    if (!player) return;
    player.gold = Number(player.gold || 0) + 100;
  });

  ctx.journal.emit({
    country: 'RUS',
    category: 'ECONOMY',
    subCategory: 'BUDGET',
    priority: 'NORMAL',
    visibility: { type: 'COUNTRY', targets: ['RUS'] },
    message: 'Государственный бюджет получил 100 золота.',
    ttlTurns: 1,
  });
}
```

Then register it in `GAME_ENGINE_CONFIG.systems`:

```javascript
systems: [
  { id: 'ECONOMY', priority: 500, handler: processEconomy },
],
```

Systems run in ascending `priority` order. They only receive `ctx` and should not access `SpreadsheetApp`; persistence happens after every system and validator completes successfully. Unchanged ranges/containers are not written. On an error, game state is not saved, though a separate diagnostic journal entry is persisted.

## Journal format and TTL

`NR_JOURNAL` is a readable player-facing table: one row is one report. Its headers are `ХОД`, `ВЕДОМСТВО`, `ТЕМА`, `СТРАНА`, `ВАЖНОСТЬ`, `ДОСТУП`, `СРОК (ХОДОВ)`, `СВОДКА`, and `ИД`.

Game systems still use internal values such as `category: 'INDUSTRY'` and `priority: 'HIGH'`, but players never see those codes. The journal renders them as ministries and clear Russian importance levels. Every message receives a short report-style introduction, an emoji, and a colour: green for good news, orange when attention is needed, red for urgent news, and blue for ordinary reports. For example:

```text
12 | Министерство промышленности | Снабжение | RUS | Требует внимания | Странам: RUS | 2 | 🏭 ⚠️ Требуется внимание: Предприятие временно остановлено: на его складе не хватает сырья. |
```

`СРОК (ХОДОВ): 1` exists during its creation turn and is cleared at the beginning of the next turn. A permanent entry displays `Постоянно`; create it by passing `ttlTurns: null`. Temporary rows always have an empty `ИД`; a removable permanent row gets a short ID such as `m28_3` so it can later be removed with `ctx.journal.remove(messageId)`. Pass `removable: false` to omit the ID as well.

Supported viewer filtering helper:

```javascript
ctx.journal.listVisibleTo({ country: 'RUS' });
ctx.journal.listVisibleTo({ player: 'player_15' });
ctx.journal.listVisibleTo({ debug: true });
```

The visibility types currently handled are `PUBLIC`, `COUNTRY`, `COUNTRIES`, `PLAYER`, `PLAYERS`, and `DEBUG`. Unknown types are treated as private until explicitly implemented.

When the journal range is full, its configuration uses `DROP_OLDEST` by default, replacing only the oldest message's cell. Use `overflow: 'THROW'` to abort the whole turn instead.

## System journal and errors

After every successful turn, the government chancellery adds a public report such as `Ход 4 завершён. Правительство приступило к ходу 5.`. It exists for one turn by default; change its privacy, priority, or TTL in `GAME_ENGINE_CONFIG.systemJournal.turnChange`.

Every uncaught exception from the engine, a registered game system, or any function called by that system aborts the game-state save and is added separately to `NR_JOURNAL` as a clear urgent report. It is `DEBUG`-only and permanent by default; technical error names and stack details remain in the Apps Script execution log.

For non-fatal problems and server-style messages, use the context logger in a game system:

```javascript
ctx.log.info('Orders have been loaded.', { orders: 12 });
ctx.log.warn('Province has no owner.', { provinceId: 'moscow' });
ctx.log.error('Unsupported unit type.', { unitId: 'unit_42' });
```

These entries are also written to the Apps Script execution log. Warning/error TTL and stack-trace settings are in `GAME_ENGINE_CONFIG.systemJournal.log`.

## Important rules

- Do not add/remove rows or columns from ordinary matrices in `ctx.data`; mutate cells and the objects stored in them only. A container may add a record/country only into a physical blank row/column already reserved in its named range.
- Keep `NR_GAME_CORE.MAIN[0][0].turn` under engine control. It is increased after a successful pipeline.
- Add game-specific validation functions to `validators`. Returning `false`, an error string, or an array of error strings cancels saving.
- Use `readOnlyRanges` for rules, UI data, and formula ranges that the engine should be able to read but never overwrite.
- `PROCESS_TURN` uses a document lock, preventing two simultaneous turn calculations.
