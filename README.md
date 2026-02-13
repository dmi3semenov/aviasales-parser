# Aviasales Parser

Парсер Aviasales для поиска билетов на сложные маршруты (multi-city). Автоматизирует поиск, фильтрует по критериям и выгружает в Excel.

## Возможности

- Парсинг multi-city маршрутов (например, MOW → DXB → MRU → MOW)
- Фильтрация по критериям (прямые рейсы, количество пересадок)
- Поддержка нескольких дат в одной сессии браузера (одна капча)
- Выгрузка в Excel с двумя вкладками: все билеты + сводка подходящих
- Автоматическое объединение сводок из нескольких поисков
- Генератор комбинаций дат

## Установка

```bash
npm install
npx playwright install chromium
```

## Использование

### 1. Сгенерировать комбинации дат

```bash
node generate-trips.js
```

Настройки в файле `generate-trips.js`:
```javascript
const config = {
    minDeparture: '20.02',      // Самая ранняя дата вылета
    maxReturn: '10.03',         // Самая поздняя дата возврата
    dubaiNightsMin: 3,          // Мин. ночей в Дубае
    dubaiNightsMax: 4,          // Макс. ночей в Дубае
    mauritiusNightsMin: 7,      // Мин. ночей на Маврикии
    mauritiusNightsMax: 9,      // Макс. ночей на Маврикии
    route: ['MOW', 'DXB', 'MRU', 'MOW']
};
```

### 2. Запустить парсинг

```bash
# Один URL
node aviasales-parser.js "https://www.aviasales.ru/search/MOW2002DXB2402MRU0303MOW2"

# Несколько URL (одна капча на все)
node aviasales-parser.js \
  "https://www.aviasales.ru/search/MOW2002DXB2402MRU0303MOW2" \
  "https://www.aviasales.ru/search/MOW2102DXB2502MRU0403MOW2" \
  "https://www.aviasales.ru/search/MOW2202DXB2602MRU0503MOW2"
```

### 3. Объединить сводки (опционально)

```bash
node merge-summaries.js
```

## Формат URL Aviasales

```
MOW2002DXB2402MRU0303MOW2
│  │   │  │   │  │   │
│  │   │  │   │  │   └── количество пассажиров
│  │   │  │   │  └────── дата вылета из MRU (DDMM)
│  │   │  │   └───────── город 3
│  │   │  └───────────── дата вылета из DXB (DDMM)
│  │   └──────────────── город 2
│  └──────────────────── дата вылета из MOW (DDMM)
└─────────────────────── город 1
```

## Выходные файлы

Все файлы сохраняются в папку `output/`:

```
as_2002-2402-0303_MOW-DXB-MRU_02-13_19-26.xlsx  # Excel с двумя вкладками
as_2002-2402-0303_MOW-DXB-MRU_02-13_19-26.json  # JSON с данными
as_2002-2402-0303_MOW-DXB-MRU_02-13_19-26.png   # Скриншот страницы
as_СВОДКИ_02-13_19-30.xlsx                      # Объединённая сводка
```

## Критерии фильтрации

По умолчанию настроено для маршрута MOW → DXB → MRU → MOW:
- MOW → DXB: только прямой рейс
- DXB → MRU: только прямой рейс
- MRU → MOW: прямой или 1 пересадка

## Адаптация под другой маршрут

Парсер легко адаптируется под любой multi-city маршрут. Измените:
1. Маршрут в `generate-trips.js` и `trip-config.json`
2. Критерии фильтрации в `aviasales-parser.js` (функция `checkFlightCriteria`)

## Зависимости

- [Playwright](https://playwright.dev/) — автоматизация браузера
- [xlsx](https://www.npmjs.com/package/xlsx) — генерация Excel файлов

## Лицензия

MIT
