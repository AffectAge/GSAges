# Google Sheets turn engine

`GameEngine.gs` is a ready-to-paste Google Apps Script core for a turn-based game stored in Google Sheets. It reads named ranges into `ctx.data`, runs all game systems in memory, validates the result, then writes the named ranges back in batches.

## Quick start

1. Open the target spreadsheet: **Extensions → Apps Script**.
2. Create a script file named `GameEngine.gs` and paste in [GameEngine.gs](GameEngine.gs).
3. By default, missing explicitly configured named ranges are created automatically on `_GAME_DATA` at the start of `PROCESS_TURN`. This creates `NR_GAME_META` with turn `1` and a 500-cell `NR_JOURNAL`.
4. To create a range manually instead, create it with the same name. In the first cell of `NR_GAME_META`, put:

   ```json
   {"turn":1,"status":"WAITING"}
   ```

5. Adjust `rangeDefaults` if `NR_JOURNAL` needs another capacity, or if a new game range needs a particular size/sheet/cell position.
6. Add your game ranges to `GAME_ENGINE_CONFIG.namedRanges`, add their automatic-creation settings to `rangeDefaults` if needed, define handlers and register them in `GAME_ENGINE_CONFIG.systems`.
7. Run `PROCESS_TURN` once from the Apps Script editor to grant permissions. After reloading the sheet, the **Game engine → Process turn** menu will be available.

## Data model

Every named range is a matrix, so cell coordinates cannot be lost:

```javascript
ctx.data.NR_UNITS[2][0] // always the third row, first column of NR_UNITS
```

Blank cells become `null`. JSON objects and arrays are decoded automatically; normal strings, numbers, booleans, and dates stay as their native values. During saving, objects/arrays are serialized back to JSON and `null` becomes a blank cell.

The engine currently reads values, not formula definitions. Do not place formulas in a writable engine range. Put calculated/formula ranges in `readOnlyRanges` instead.

## Province generator

Run `GENERATE_PROVINCES()` to fill every blank cell in `NR_PROVINCES` with a base JSON province and to add missing properties to existing JSON provinces. Existing values are never overwritten, including `0`, `false`, `null`, and nested resource quantities. The schema is editable in `PROVINCE_GENERATOR_CONFIG.defaults`; adding a field there and running the generator again performs a deep, missing-fields-only update.

The initial schema contains `id`, `name`, `owner`, `terrain`, `elevation`, radiation, pollution, temperature, humidity, area, `soilFertility` (soil quality/yield potential), `fertileLandPercent` (the percentage of province area suitable for agriculture), `resources.water.amount`, and `neighbors`. IDs and names are generated automatically. `prov_1` is the default neighbour for every other province; change `defaultNeighborId` in the generator config to use another anchor province.

The named range should cover exactly the number of province slots required. If it does not exist, the core creates a single-cell `NR_PROVINCES` range; change its configured size before first use when more provinces are needed. A non-blank cell must contain a JSON object; ordinary text and JSON arrays are rejected to avoid silently destroying data.

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

Systems run in ascending `priority` order. They only receive `ctx` and should not access `SpreadsheetApp`; all persistence happens once after every system and validator completes successfully. An error means nothing is saved.

## Journal format and TTL

One `NR_JOURNAL` cell stores one JSON message. A temporary message stores only game-relevant data: `createdTurn` and `expiresAtTurn` are added automatically, while `id` and real-time timestamps are omitted.

```json
{
  "createdTurn":1,
  "country":"RUS",
  "category":"ECONOMY",
  "priority":"NORMAL",
  "visibility":{"type":"COUNTRY","targets":["RUS"]},
  "message":"Государственный бюджет получил 100 золота.",
  "ttlTurns":1,
  "expiresAtTurn":2
}
```

`ttlTurns: 1` exists during its creation turn and is cleared at the beginning of the next turn. Use `ttlTurns: 3` for three turns, or `ttlTurns: null` for a permanent entry. A permanent message gets a short ID such as `m28_3` by default, so it can be removed later with `ctx.journal.remove(messageId)`. Pass `removable: false` to a permanent message to omit its ID as well.

Supported viewer filtering helper:

```javascript
ctx.journal.listVisibleTo({ country: 'RUS' });
ctx.journal.listVisibleTo({ player: 'player_15' });
ctx.journal.listVisibleTo({ debug: true });
```

The visibility types currently handled are `PUBLIC`, `COUNTRY`, `COUNTRIES`, `PLAYER`, `PLAYERS`, and `DEBUG`. Unknown types are treated as private until explicitly implemented.

When the journal range is full, its configuration uses `DROP_OLDEST` by default, replacing only the oldest message's cell. Use `overflow: 'THROW'` to abort the whole turn instead.

## System journal and errors

After every successful turn, the engine adds a `SYSTEM` message such as `Ход 4 завершён. Начат ход 5.`. It is public by default and exists for one turn; change its privacy, priority, or TTL in `GAME_ENGINE_CONFIG.systemJournal.turnChange`.

Every uncaught exception from the engine, a registered game system, or any function called by that system aborts the game-state save and is added separately to `NR_JOURNAL`. The error record has `category: 'SYSTEM'`, `priority: 'CRITICAL'`, source-system ID, message, and (by default) a stack trace in `payload`. It is `DEBUG`-only and permanent by default, so a player UI should show it only to a debug/admin viewer.

For non-fatal problems and server-style messages, use the context logger in a game system:

```javascript
ctx.log.info('Orders have been loaded.', { orders: 12 });
ctx.log.warn('Province has no owner.', { provinceId: 'moscow' });
ctx.log.error('Unsupported unit type.', { unitId: 'unit_42' });
```

These entries are also written to the Apps Script execution log. Warning/error TTL and stack-trace settings are in `GAME_ENGINE_CONFIG.systemJournal.log`.

## Important rules

- Do not add/remove rows or columns from `ctx.data`; mutate cells and the objects stored in them only.
- Keep `NR_GAME_META.turn` under engine control. It is increased after a successful pipeline.
- Add game-specific validation functions to `validators`. Returning `false`, an error string, or an array of error strings cancels saving.
- Use `readOnlyRanges` for rules, UI data, and formula ranges that the engine should be able to read but never overwrite.
- `PROCESS_TURN` uses a document lock, preventing two simultaneous turn calculations.
