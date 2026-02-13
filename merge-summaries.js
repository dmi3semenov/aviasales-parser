const XLSX = require('xlsx');
const fs = require('fs');
const { execSync } = require('child_process');

// Закрываем Excel
console.log('Закрываем Excel...');
try {
    execSync('pkill -f "Microsoft Excel" 2>/dev/null || true', { stdio: 'ignore' });
} catch (e) {}

// Находим все файлы as_*.xlsx в папке output
const outputDir = 'output';
const files = fs.readdirSync(outputDir)
    .filter(f => f.startsWith('as_') && f.endsWith('.xlsx') && !f.includes('СВОДКИ'))
    .sort();

console.log(`\nНайдено ${files.length} файлов:\n`);
files.forEach((f, i) => console.log(`  ${i + 1}. ${f}`));

// Собираем все сводки
const allData = [];

files.forEach((file, idx) => {
    console.log(`\nЧитаю: ${file}`);

    const wb = XLSX.readFile(`${outputDir}/${file}`);

    if (!wb.Sheets['Сводка']) {
        console.log('  ⚠️ Нет вкладки "Сводка"');
        return;
    }

    const data = XLSX.utils.sheet_to_json(wb.Sheets['Сводка']);
    console.log(`  ✓ Строк в сводке: ${data.length}`);

    // Извлекаем даты из имени файла для подзаголовка
    const match = file.match(/as_(\d{4})-(\d{4})-(\d{4})_/);
    let label = file;
    if (match) {
        // Парсим даты (DDMM)
        const d1_day = parseInt(match[1].slice(0,2));
        const d1_mon = parseInt(match[1].slice(2));
        const d2_day = parseInt(match[2].slice(0,2));
        const d2_mon = parseInt(match[2].slice(2));
        const d3_day = parseInt(match[3].slice(0,2));
        const d3_mon = parseInt(match[3].slice(2));

        // Считаем ночи (упрощённо, в рамках одного года)
        function dayOfYear(day, mon) {
            const daysInMonth = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
            let total = day;
            for (let i = 1; i < mon; i++) total += daysInMonth[i];
            return total;
        }

        const dxbNights = dayOfYear(d2_day, d2_mon) - dayOfYear(d1_day, d1_mon);
        const mruNights = dayOfYear(d3_day, d3_mon) - dayOfYear(d2_day, d2_mon);

        const d1 = match[1].slice(0,2) + '.' + match[1].slice(2);
        const d2 = match[2].slice(0,2) + '.' + match[2].slice(2);
        const d3 = match[3].slice(0,2) + '.' + match[3].slice(2);
        label = `${d1} → ${d2} → ${d3}  |  Дубай ${dxbNights}н, Маврикий ${mruNights}н`;
    }

    // Добавляем заголовок с датами
    allData.push({ 'Цена': `═══ ${label} ═══` });

    // Добавляем данные
    data.forEach(row => allData.push(row));

    // 1 пустая строка между таблицами
    allData.push({});
});

// Создаём новый Excel
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.json_to_sheet(allData);

// Ширина колонок
ws['!cols'] = [
    { wch: 20 },  // Цена (или заголовок)
    { wch: 5 },   // Кол-во
    { wch: 10 },  // Р1
    { wch: 8 },   // Р1 Дата
    { wch: 14 },  // Р1 Вылет
    { wch: 10 },  // Р2
    { wch: 8 },   // Р2 Дата
    { wch: 14 },  // Р2 Вылет
    { wch: 10 },  // Р3
    { wch: 8 },   // Р3 Дата
    { wch: 14 }   // Р3 Вылет
];

XLSX.utils.book_append_sheet(wb, ws, 'Все сводки');

// Сохраняем (с датой и временем)
const now = new Date();
const dateStr = (now.getMonth() + 1).toString().padStart(2, '0') + '-' + now.getDate().toString().padStart(2, '0');
const timeStr = now.toTimeString().slice(0, 5).replace(':', '-');
const outputFile = `${outputDir}/as_СВОДКИ_${dateStr}_${timeStr}.xlsx`;
XLSX.writeFile(wb, outputFile);
console.log(`\n✓ Объединённый файл: ${outputFile}`);

// Открываем
execSync(`open "${outputFile}"`);
console.log('✓ Файл открыт!');
