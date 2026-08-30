# Новый движок игрового хода

Используйте только `GameTurnEngine.gs`. Он не обращается к прежним `GameEngine.gs`, `TurnEngine.gs` или игровым системам.

## Настройка

1. В Google Sheets создайте каждый игровой именованный диапазон, указанный в `GAME_TURN_CONFIG.dataRangeNames`. Первая строка диапазона — уникальные заголовки.
2. В `GameTurnEngine.gs` укажите их имена и добавьте обработчики в `processors`.
3. Один раз запустите `INITIALIZE_GAME_TURN_ENGINE()`. Движок создаст независимые диапазоны `Игровой_журнал` и `Состояние_движка`.
4. Запускайте `PROCESS_GAME_TURN()` через меню, кнопку или триггер.

Диапазоны, которыми управляет движок, не должны содержать формулы или ручные поля: при сохранении движок записывает их целиком одним `setValues()`.

## Формат данных

Пустая ячейка становится `null`. Строка, начинающаяся с `{` или `[`, считается JSON. Некорректный JSON отменяет ход до запуска процессоров и пишет в журнал диапазон, заголовок и точную ячейку. Обычные текст, числа, `true`/`false` и даты остаются исходными значениями.

Пример диапазона `Игровые_данные`:

| Фабрики | Страны |
| --- | --- |
| `{"id":"ger_1","owner":"GER","status":"ACTIVE"}` | `{"id":"GER","treasury":100}` |
|  | `{"id":"FRA","treasury":80}` |

## Ручное заполнение провинций

`GENERATE_EMPTY_PROVINCES()` — отдельная команда, не являющаяся процессором и не запускающая ход. Она берёт колонку `Провинции` в именованном диапазоне `Игровые_данные` и заполняет **только пустые ячейки** JSON-объектами провинций.

Параметры находятся в `GAME_TURN_CONFIG.provinceTemplateGenerator`. Туда уже внесён переданный шаблон Ломбардии. Каждая созданная провинция получает уникальный числовой `id`: первая — `1001`, затем `1002`, `1003` и так далее. Для `1001` имя — `Ломбардия`; у остальных по умолчанию `Провинция <id>`. Измените `namesById`, `template` или `firstId`, если нужны другие стартовые данные.

Существующие непустые ячейки не меняются. Но они должны содержать корректные JSON-объекты с уникальным числовым `id`, иначе команда остановится без записи новых провинций.

## Процессоры

Процессор получает `data` и `ctx`, работает с ними в памяти и возвращает обновлённый `data`. Очерёдность определяется меньшим `priority`.

```javascript
const GAME_TURN_CONFIG = {
  // ...
  processors: [
    { id: 'FACTORIES', priority: 100, handler: processFactories },
    { id: 'ECONOMY', priority: 200, handler: processEconomy },
  ],
};

function processFactories(data, ctx) {
  const factories = ctx.getColumn('Игровые_данные', 'Фабрики');

  factories.forEach(function (factory, row) {
    if (!factory) return;
    factory.status = 'PROCESSED';
    ctx.log({
      category: 'INDUSTRY',
      country: factory.owner,
      priority: 'SUCCESS',
      ttl: 1,
      message: [
        ctx.country(factory.owner),
        ctx.text(': фабрика «'),
        ctx.positive(factory.id),
        ctx.text('» обработана.'),
      ],
    });
  });
  return data;
}
```

### Время выполнения процессоров

После каждого успешно завершённого обработчика движок автоматически создаёт в `Игровой_журнал` технический отчёт вида: `Процессор «FACTORIES» выполнен за 38 мс.` Время относится только к телу функции `handler`, без создания самой записи в журнале.

Настройка находится в `GAME_TURN_CONFIG.processorTimingReports`:

```javascript
processorTimingReports: {
  enabled: true,
  type: 'PERFORMANCE',
  category: 'ENGINE',
  priority: 'NORMAL',
  visibility: 'SYSTEM', // Только служебный журнал.
  ttl: 1,
},
```

Поставьте `enabled: false`, если такие записи временно не нужны. Если процессор завершился ошибкой, постоянная запись `ERROR` всё равно содержит его имя и время до ошибки.

`ctx.forEach('Игровые_данные', 'Фабрики', callback)` передаёт объект ячейки с `cell.value`, `cell.empty`, `cell.a1`, `cell.row`, `cell.header`. Это удобно, когда нужно заменить примитивное значение и увидеть координаты:

```javascript
ctx.forEach('Игровые_данные', 'Число ходов', function (cell) {
  if (cell.empty) return;
  cell.value = Number(cell.value) + 1;
});
```

## Игровой журнал

`Игровой_журнал` имеет отдельные колонки для хода, типа, категории, важности, видимости, срока, источника и текста. Сообщения от процессоров записываются только после успешного сохранения игровых данных. Ошибки JSON, процессоров и валидаторов не сохраняют игровые данные, но добавляются в журнал как постоянные технические сообщения.

В `message` можно передать обычную строку либо массив фрагментов. Доступны `ctx.text`, `ctx.country`, `ctx.number`, `ctx.positive`, `ctx.negative`, `ctx.warning`, `ctx.errorText`. Каждый фрагмент может иметь `color`, `bold`, `italic`, `underline`, `strikethrough` и `link`:

```javascript
ctx.log({
  category: 'DIPLOMACY',
  priority: 'HIGH',
  visibility: 'PUBLIC',
  ttl: 3,
  message: [
    ctx.country('GER', { color: '#DC2626' }),
    ctx.text(' объявила войну государству '),
    ctx.country('FRA', { color: '#2563EB' }),
    ctx.text('!'),
  ],
});
```

`ttl: 1` удаляет запись при начале следующего хода. `ttl: -1` хранит её постоянно.
