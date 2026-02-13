/**
 * Генератор комбинаций дат для поиска билетов
 *
 * Использование:
 *   node generate-trips.js
 *
 * Параметры задаются в конфигурации ниже
 */

const { execSync, spawn } = require('child_process');

// ══════════════════════════════════════════════════════════════════
// КОНФИГУРАЦИЯ — ИЗМЕНИ ПОД СВОЙ ЗАПРОС
// ══════════════════════════════════════════════════════════════════

const config = {
    // Границы поездки
    minDeparture: '20.02',   // Самая ранняя дата вылета из Москвы
    maxReturn: '10.03',      // Самая поздняя дата возврата в Москву

    // Сколько ночей проводим
    dubaiNightsMin: 3,       // Минимум ночей в Дубае
    dubaiNightsMax: 4,       // Максимум ночей в Дубае
    mauritiusNightsMin: 7,   // Минимум ночей на Маврикии
    mauritiusNightsMax: 9,   // Максимум ночей на Маврикии

    // Маршрут (не меняй если летишь MOW→DXB→MRU→MOW)
    route: ['MOW', 'DXB', 'MRU', 'MOW']
};

// ══════════════════════════════════════════════════════════════════

// Парсим дату DD.MM в объект
function parseDate(str) {
    const [day, month] = str.split('.').map(Number);
    return { day, month };
}

// Преобразуем в формат для URL (DDMM)
function toUrlFormat(day, month) {
    return day.toString().padStart(2, '0') + month.toString().padStart(2, '0');
}

// День года (для вычисления разницы)
function dayOfYear(day, month) {
    const daysInMonth = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let total = day;
    for (let i = 1; i < month; i++) total += daysInMonth[i];
    return total;
}

// Добавить дни к дате
function addDays(day, month, days) {
    const daysInMonth = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let d = day + days;
    let m = month;
    while (d > daysInMonth[m]) {
        d -= daysInMonth[m];
        m++;
        if (m > 12) m = 1;
    }
    return { day: d, month: m };
}

// Генерируем все комбинации
function generateCombinations() {
    const minDep = parseDate(config.minDeparture);
    const maxRet = parseDate(config.maxReturn);
    const minDepDOY = dayOfYear(minDep.day, minDep.month);
    const maxRetDOY = dayOfYear(maxRet.day, maxRet.month);

    const combinations = [];

    // Перебираем даты вылета из Москвы
    for (let depDOY = minDepDOY; depDOY <= maxRetDOY - config.dubaiNightsMin - config.mauritiusNightsMin; depDOY++) {
        // Перебираем количество ночей в Дубае
        for (let dxbNights = config.dubaiNightsMin; dxbNights <= config.dubaiNightsMax; dxbNights++) {
            // Перебираем количество ночей на Маврикии
            for (let mruNights = config.mauritiusNightsMin; mruNights <= config.mauritiusNightsMax; mruNights++) {

                const returnDOY = depDOY + dxbNights + mruNights;

                // Проверяем что возврат не позже максимальной даты
                if (returnDOY > maxRetDOY) continue;

                // Вычисляем даты
                let currentDay = minDep.day;
                let currentMonth = minDep.month;

                // Сдвигаем на нужное количество дней от минимальной даты
                const offset = depDOY - minDepDOY;
                const dep = addDays(currentDay, currentMonth, offset);
                const dxb = addDays(dep.day, dep.month, dxbNights);
                const ret = addDays(dxb.day, dxb.month, mruNights);

                combinations.push({
                    departure: dep,
                    dubai: dxb,
                    return: ret,
                    dxbNights,
                    mruNights,
                    url: `https://www.aviasales.ru/search/${config.route[0]}${toUrlFormat(dep.day, dep.month)}${config.route[1]}${toUrlFormat(dxb.day, dxb.month)}${config.route[2]}${toUrlFormat(ret.day, ret.month)}${config.route[3]}2`
                });
            }
        }
    }

    return combinations;
}

// Главная функция
function main() {
    console.log('\n' + '█'.repeat(70));
    console.log('  ГЕНЕРАТОР КОМБИНАЦИЙ ДАТ');
    console.log('█'.repeat(70));

    console.log('\nПараметры:');
    console.log(`  Вылет из Москвы: ${config.minDeparture} - ...`);
    console.log(`  Возврат не позже: ${config.maxReturn}`);
    console.log(`  Дубай: ${config.dubaiNightsMin}-${config.dubaiNightsMax} ночей`);
    console.log(`  Маврикий: ${config.mauritiusNightsMin}-${config.mauritiusNightsMax} ночей`);

    const combinations = generateCombinations();

    console.log(`\n✓ Сгенерировано комбинаций: ${combinations.length}\n`);

    combinations.forEach((c, i) => {
        const d1 = `${c.departure.day}.${c.departure.month.toString().padStart(2, '0')}`;
        const d2 = `${c.dubai.day}.${c.dubai.month.toString().padStart(2, '0')}`;
        const d3 = `${c.return.day}.${c.return.month.toString().padStart(2, '0')}`;
        console.log(`  ${(i + 1).toString().padStart(2)}. ${d1} → ${d2} → ${d3}  (Дубай ${c.dxbNights}н, Маврикий ${c.mruNights}н)`);
    });

    if (combinations.length === 0) {
        console.log('\n⚠️ Нет подходящих комбинаций. Проверь параметры.');
        return;
    }

    // Формируем команду для запуска парсера
    const urls = combinations.map(c => `"${c.url}"`).join(' \\\n  ');

    console.log('\n' + '─'.repeat(70));
    console.log('Команда для запуска парсера:\n');
    console.log(`node aviasales-parser.js \\`);
    console.log(`  ${urls}`);
    console.log('\n' + '─'.repeat(70));

    // Спрашиваем, запустить ли сразу
    console.log('\nДля запуска парсера скопируй команду выше или запусти:');
    console.log(`  node generate-trips.js --run`);

    // Если передан флаг --run, запускаем парсер
    if (process.argv.includes('--run')) {
        console.log('\n🚀 Запускаем парсер...\n');
        const urlsArray = combinations.map(c => c.url);
        const child = spawn('node', ['aviasales-parser.js', ...urlsArray], {
            stdio: 'inherit'
        });
    }
}

main();
