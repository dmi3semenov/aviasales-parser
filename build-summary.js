/**
 * build-summary.js — сборка итоговой сводки из output/*.xlsx
 * Использование: node build-summary.js
 */

const XLSX    = require('xlsx');
const ExcelJS = require('exceljs');
const fs      = require('fs');
const path    = require('path');
const { execSync } = require('child_process');

const OUTPUT_DIR  = path.join(__dirname, 'output');
const tripConfig  = JSON.parse(fs.readFileSync(path.join(__dirname, 'trip-config.json'), 'utf-8'));
const windowDefs  = (tripConfig['даты'] && tripConfig['даты']['окна_вылета']) || [];

// ─── Утилиты дат ─────────────────────────────────────────────────────────────

const _DIM = [0,31,28,31,30,31,30,31,31,30,31,30,31];
function doy(d, m) { let n = d; for (let i = 1; i < m; i++) n += _DIM[i]; return n; }

function addDaysTo(d, m, days) {
    let nd = d + days, nm = m;
    while (nd > _DIM[nm]) { nd -= _DIM[nm]; nm++; if (nm > 12) nm = 1; }
    return { d: nd, m: nm };
}

// Определяем год по тому, в будущем ли дата относительно сегодня
function flightYear(day, month) {
    const today = new Date();
    const tDOY  = doy(today.getDate(), today.getMonth() + 1);
    const fDOY  = doy(day, month);
    return today.getFullYear() + (fDOY < tDOY ? 1 : 0);
}

const DOW_RU = ['вс','пн','вт','ср','чт','пт','сб'];
function dowStr(day, month) {
    const d = new Date(flightYear(day, month), month - 1, day);
    return DOW_RU[d.getDay()];
}

// "29.04" → "29.04 пт"
function fmtDateDow(dk) {
    const [d, m] = dk.split('.').map(Number);
    return `${dk} ${dowStr(d, m)}`;
}

// ─── Утилиты билетов ─────────────────────────────────────────────────────────

const SLOTS = ['Ночь 00–06', 'Утро 06–12', 'День 12–18', 'Вечер 18–24'];

function timeSlot(t) {
    const m = String(t || '').match(/(\d{1,2}):(\d{2})/);
    if (!m) return null;
    const h = +m[1];
    if (h < 6)  return 'Ночь 00–06';
    if (h < 12) return 'Утро 06–12';
    if (h < 18) return 'День 12–18';
    return 'Вечер 18–24';
}

function parsePrice(ps) {
    if (!(ps || '').includes('₽')) return null;
    const v = parseInt(ps.replace(/\D/g, ''));
    return (v >= 5000 && v <= 300000) ? v : null;
}

function firstTime(s) { return (s || '').split(',')[0].trim(); }

function fmtRub(v) { return v != null ? v.toLocaleString('ru-RU') + ' ₽' : '—'; }

// ─── Загрузка данных ─────────────────────────────────────────────────────────

// Для каждой комбинации дат берём только самый свежий файл (по алфавиту = по времени)
const _allFiles = fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.match(/^as_\d{4}-\d{4}_MOW-MSQ/) && f.endsWith('.xlsx'))
    .sort();
const _latestByKey = {};
_allFiles.forEach(f => {
    const m = f.match(/^as_(\d{4}-\d{4})_/);
    if (m) _latestByKey[m[1]] = f;
});
const files = Object.values(_latestByKey).sort();

console.log(`Файлов: ${files.length}`);

const createdFiles = [];
// grid[dk][nights][outSlot][retSlot] = { minPrice, outTime, retTime }
const grid        = {};
const allOptions  = [];
const ticketCount = {}; // ticketCount[dk][nights] = кол-во валидных строк

files.forEach(file => {
    const m = file.match(/^as_(\d{2})(\d{2})-(\d{2})(\d{2})_(\w+-\w+-\w+)/);
    if (!m) return;
    const dD=+m[1], dM=+m[2], rD=+m[3], rM=+m[4];
    const nights = doy(rD, rM) - doy(dD, dM);
    const dk = m[1] + '.' + m[2];
    createdFiles.push({ file, dD, dM, rD, rM, nights, dk });

    if (!grid[dk])              grid[dk] = {};
    if (!grid[dk][nights])      grid[dk][nights] = {};
    if (!ticketCount[dk])       ticketCount[dk] = {};
    if (!ticketCount[dk][nights]) ticketCount[dk][nights] = 0;

    const wb = XLSX.readFile(path.join(OUTPUT_DIR, file));
    const sheet = wb.Sheets['Сводка'];
    if (!sheet) return;

    XLSX.utils.sheet_to_json(sheet).forEach(r => {
        const pv = parsePrice(r['Цена'] || '');
        if (!pv) return;
        ticketCount[dk][nights]++;
        const outTime = firstTime(r['Р1 Вылет']);
        const retTime = firstTime(r['Р2 Вылет']);
        const outSlot = timeSlot(outTime);
        const retSlot = timeSlot(retTime);
        if (!outSlot || !retSlot) return;

        if (!grid[dk][nights][outSlot])          grid[dk][nights][outSlot] = {};
        const cell = grid[dk][nights][outSlot][retSlot];
        if (!cell || pv < cell.minPrice) {
            grid[dk][nights][outSlot][retSlot] = { minPrice: pv, outTime, retTime };
        }

        allOptions.push({
            '₽': pv, 'Вылет': dk, 'Ноч.': nights,
            'Туда дата': r['Р1 Дата']||'', 'Туда время': r['Р1 Вылет']||'',
            'Прилёт MSQ': r['Р1 Прилёт']||'',
            'Обратно дата': r['Р2 Дата']||'', 'Обратно время': r['Р2 Вылет']||'',
            'Прилёт MOW': r['Р2 Прилёт']||'',
        });
    });
});

allOptions.sort((a, b) => a['₽'] - b['₽']);

const sortedDates = [...new Set(createdFiles.map(f => f.dk))].sort((a, b) => {
    const [da,ma] = a.split('.').map(Number);
    const [db,mb] = b.split('.').map(Number);
    return doy(da,ma) - doy(db,mb);
});
const sortedNights = [...new Set(createdFiles.map(f => f.nights))].sort((a,b) => a-b);

// ─── Аналитика для Рекомендаций ───────────────────────────────────────────────

function analyzeGroup(dk, nights) {
    const g = grid[dk] && grid[dk][nights];
    if (!g) return null;

    let globalMin = Infinity, globalMax = 0;
    let bestCell = null, worstCell = null;

    // outSlotMins[slot] = min price across all return slots for this outbound slot
    const outSlotMins = {};
    const outSlotBest = {}; // { minPrice, outTime, retTime, retSlot }

    SLOTS.forEach(outSlot => {
        if (!g[outSlot]) return;
        Object.entries(g[outSlot]).forEach(([retSlot, cell]) => {
            if (cell.minPrice < globalMin) { globalMin = cell.minPrice; bestCell = { ...cell, outSlot, retSlot }; }
            if (cell.minPrice > globalMax) { globalMax = cell.minPrice; worstCell = { ...cell, outSlot, retSlot }; }
            if (!outSlotMins[outSlot] || cell.minPrice < outSlotMins[outSlot]) {
                outSlotMins[outSlot] = cell.minPrice;
                outSlotBest[outSlot] = { ...cell, retSlot };
            }
        });
    });

    if (!bestCell) return null;

    const spread = globalMax / globalMin;
    const spreadLabel = spread < 1.20 ? 'небольшой — летите в любое время' :
                        spread < 1.50 ? `умеренный (×${spread.toFixed(1)}) — есть разница по времени` :
                                        `значительный (×${spread.toFixed(1)}) — время вылета сильно влияет на цену`;

    // Classify outbound slots
    const GOOD_THRESH = globalMin * 1.20;
    const OK_THRESH   = globalMin * 1.45;
    const slotRating  = {};
    Object.entries(outSlotMins).forEach(([slot, price]) => {
        slotRating[slot] = price <= GOOD_THRESH ? 'good' : price <= OK_THRESH ? 'ok' : 'bad';
    });

    const goodSlots = SLOTS.filter(s => slotRating[s] === 'good');
    const okSlots   = SLOTS.filter(s => slotRating[s] === 'ok');
    const badSlots  = SLOTS.filter(s => slotRating[s] === 'bad');

    // Narrative advice
    const parts = [];
    if (spread < 1.20) {
        parts.push(`Летите в любое время — разброс небольшой (${fmtRub(globalMin)} — ${fmtRub(globalMax)}).`);
    } else {
        if (goodSlots.length) {
            const list = goodSlots.map(s => `${s.slice(0,5)} (от ${fmtRub(outSlotMins[s])})`).join(', ');
            parts.push(`Выгодно: ${list}.`);
        }
        if (okSlots.length) {
            const list = okSlots.map(s => `${s.slice(0,5)} (от ${fmtRub(outSlotMins[s])})`).join(', ');
            parts.push(`Приемлемо: ${list}.`);
        }
        if (badSlots.length) {
            const list = badSlots.map(s => {
                const pct = Math.round((outSlotMins[s] / globalMin - 1) * 100);
                return `${s.slice(0,5)} (от ${fmtRub(outSlotMins[s])}, +${pct}%)`;
            }).join(', ');
            parts.push(`Избегать: ${list}.`);
        }
    }
    const advice = parts.join(' ');

    return { globalMin, globalMax, bestCell, worstCell, outSlotMins, outSlotBest,
             slotRating, goodSlots, okSlots, badSlots, spread, spreadLabel, advice };
}

function buildRecommendations() {
    // Все группы отсортированные по globalMin
    const groups = [];
    sortedDates.forEach(dk => {
        sortedNights.forEach(nights => {
            const a = analyzeGroup(dk, nights);
            if (!a) return;
            const [d, m] = dk.split('.').map(Number);
            const ret = addDaysTo(d, m, nights);
            const retDk = `${String(ret.d).padStart(2,'0')}.${String(ret.m).padStart(2,'0')}`;
            groups.push({ dk, nights, retDk, ...a });
        });
    });
    groups.sort((a, b) => a.globalMin - b.globalMin);
    return groups;
}

// ─── Сборка Excel ─────────────────────────────────────────────────────────────

const CF_RULE = {
    type: 'colorScale', priority: 1,
    cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
    color: [{ argb: 'FF63BE7B' }, { argb: 'FFFFEB84' }, { argb: 'FFF8696B' }]
};

async function build() {
    const wb = new ExcelJS.Workbook();

    const styleHeader = { font:{bold:true,size:11}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFD9E1F2'}}, alignment:{vertical:'middle',horizontal:'center'} };
    const styleGroup  = { font:{bold:true,size:11,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF2F5496'}}, alignment:{vertical:'middle'} };
    const styleSummR  = { font:{bold:true,size:10,italic:true}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}}, alignment:{horizontal:'center'} };
    const styleMiss   = { alignment:{horizontal:'center'}, font:{color:{argb:'FFBBBBBB'}} };
    const styleNum    = { numFmt:'# ##0', alignment:{horizontal:'center'} };
    const sCard  = { font:{bold:true,size:11,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF2F5496'}} };
    const sGood  = { font:{color:{argb:'FF1F6B35'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EFDA'}} };
    const sOk    = { font:{color:{argb:'FF7D6608'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}} };
    const sBad   = { font:{color:{argb:'FF9C0006'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFFFC7CE'}} };
    const sLabel = { font:{bold:true,color:{argb:'FF44546A'}} };
    const sTip   = { font:{italic:true,color:{argb:'FF203864'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFDCE6F1'}}, alignment:{wrapText:true} };

    function addCardRow(ws, label, value, style) {
        const r = ws.addRow([label, value]);
        if (style) { Object.assign(r.getCell(1), style); Object.assign(r.getCell(2), style); }
        r.getCell(2).alignment = { wrapText: true };
        return r;
    }

    const groups = buildRecommendations();

    // Вкладки — создаём все сразу, чтобы задать порядок и цвет
    const wsItog    = wb.addWorksheet('Итог');              wsItog.properties.tabColor    = { argb: 'FFED7D31' };
    const wsTime    = wb.addWorksheet('По времени вылета'); wsTime.properties.tabColor    = { argb: 'FFFFC000' };
    const wsRec     = wb.addWorksheet('Рек - Карточки');    wsRec.properties.tabColor     = { argb: 'FF92D050' };
    const wsRecT    = wb.addWorksheet('Рек - Таблица');     wsRecT.properties.tabColor    = { argb: 'FF92D050' };
    const wsRecTD   = wb.addWorksheet('Рек - По дате');     wsRecTD.properties.tabColor   = { argb: 'FF92D050' };
    const wsRecW    = wb.addWorksheet('Рек - По окнам');    wsRecW.properties.tabColor    = { argb: 'FF92D050' };
    const wsDates   = wb.addWorksheet('По датам');
    const wsAll     = wb.addWorksheet('Все варианты');
    const wsDetails = wb.addWorksheet('Все сводки');
    const wsHelp    = wb.addWorksheet('Пояснения');

    // ── 1. По времени вылета ──────────────────────────────────────────────────

    wsTime.getColumn(1).width = 20;
    SLOTS.forEach((_, i) => wsTime.getColumn(i + 2).width = 15);
    wsTime.getColumn(SLOTS.length + 2).width = 46;  // Лучший (строка)
    wsTime.getColumn(SLOTS.length + 3).width = 46;  // Лучший (столбец)

    // Заголовок: вылет / обратно / итог
    const colHeaders = ['Вылет МСК  ╲  Обратно MSQ →', ...SLOTS, 'Лучший → (дешевле всего вернуться)', 'Лучший ↓ (дешевле всего вылететь)'];
    const hRow = wsTime.addRow(colHeaders);
    hRow.eachCell(c => Object.assign(c, styleHeader));
    hRow.height = 22;
    wsTime.views = [{ state: 'frozen', ySplit: 1 }];

    const cfRanges = [];

    sortedDates.forEach(dk => {
        sortedNights.forEach(nights => {
            if (!grid[dk] || !grid[dk][nights]) return;
            const g = grid[dk][nights];

            // Заголовок группы
            const [d, m] = dk.split('.').map(Number);
            const ret = addDaysTo(d, m, nights);
            const retDk = `${String(ret.d).padStart(2,'0')}.${String(ret.m).padStart(2,'0')}`;
            const groupLabel = `${fmtDateDow(dk)}  →  +${nights} ноч. (обратно ${fmtDateDow(retDk)})`;
            const gRow = wsTime.addRow([groupLabel]);
            wsTime.mergeCells(gRow.number, 1, gRow.number, colHeaders.length);
            Object.assign(gRow.getCell(1), styleGroup);
            gRow.height = 18;

            const dataStart = wsTime.lastRow.number + 1;

            // Вычислить лучшего по СТОЛБЦУ (для каждого retSlot — мин по всем outSlot)
            const colBest = {}; // retSlot → { minPrice, outTime, retTime, outSlot }
            SLOTS.forEach(retSlot => {
                SLOTS.forEach(outSlot => {
                    const cell = g[outSlot] && g[outSlot][retSlot];
                    if (!cell) return;
                    if (!colBest[retSlot] || cell.minPrice < colBest[retSlot].minPrice) {
                        colBest[retSlot] = { ...cell, outSlot };
                    }
                });
            });

            // Строки данных
            SLOTS.forEach(outSlot => {
                const rowVals = [outSlot];
                const numericVals = [];

                SLOTS.forEach(retSlot => {
                    const cell = g[outSlot] && g[outSlot][retSlot];
                    rowVals.push(cell ? cell.minPrice : null);
                    numericVals.push(cell ? cell.minPrice : null);
                });

                // Лучший по строке (мин среди return-слотов)
                const rowCells = Object.entries(g[outSlot] || {});
                let rowVerdict = '';
                if (rowCells.length > 0) {
                    const best = rowCells.reduce((a, b) => a[1].minPrice <= b[1].minPrice ? a : b);
                    rowVerdict = `${fmtRub(best[1].minPrice)}  ·  туда ${best[1].outTime}, обратно ${best[1].retTime} (${best[0].slice(0,5)})`;
                }

                // Лучший по столбцу — пустой для обычных строк; заполняем только в итоговой строке
                rowVals.push(rowVerdict);
                rowVals.push('');

                const dr = wsTime.addRow(rowVals);
                dr.getCell(1).font = { italic: true };
                dr.getCell(1).alignment = { horizontal: 'left' };

                numericVals.forEach((v, i) => {
                    const c = dr.getCell(i + 2);
                    if (v !== null) { c.value = v; Object.assign(c, styleNum); }
                    else { c.value = '—'; Object.assign(c, styleMiss); }
                });
                dr.getCell(SLOTS.length + 2).alignment = { wrapText: false };
                dr.getCell(SLOTS.length + 3).value = '';
            });

            // Итоговая строка: лучший по каждому СТОЛБЦУ (обратный рейс)
            const sumVals = ['Лучший ↓'];
            SLOTS.forEach(retSlot => {
                const cb = colBest[retSlot];
                sumVals.push(cb ? cb.minPrice : null);
            });
            // Глобальный мин для всей группы
            const globalBest = Object.values(colBest).reduce((a, b) => (!a || b.minPrice < a.minPrice) ? b : a, null);
            const globalVerdict = globalBest
                ? `${fmtRub(globalBest.minPrice)}  ·  туда ${globalBest.outTime} (${globalBest.outSlot.slice(0,5)}), обратно ${globalBest.retTime}`
                : '';
            sumVals.push('');
            sumVals.push(globalVerdict);

            const sumRow = wsTime.addRow(sumVals);
            sumRow.getCell(1).value = 'Лучший ↓';
            Object.assign(sumRow.getCell(1), styleSummR);
            SLOTS.forEach((retSlot, i) => {
                const cb = colBest[retSlot];
                const c = sumRow.getCell(i + 2);
                if (cb) { c.value = cb.minPrice; Object.assign(c, { ...styleNum, font:{bold:true,italic:true} }); }
                else { c.value = '—'; Object.assign(c, styleMiss); }
            });
            sumRow.getCell(SLOTS.length + 2).value = '';
            sumRow.getCell(SLOTS.length + 3).value = globalVerdict;
            sumRow.getCell(SLOTS.length + 3).font = { bold: true, italic: true };
            sumRow.getCell(SLOTS.length + 3).alignment = { wrapText: false };

            const dataEnd = wsTime.lastRow.number;
            if (dataEnd >= dataStart) {
                const s = wsTime.getCell(dataStart, 2).address;
                const e = wsTime.getCell(dataEnd - 1, SLOTS.length + 1).address; // без строки "Лучший↓"
                cfRanges.push(`${s}:${e}`);
            }

            wsTime.addRow([]);
        });
    });

    if (cfRanges.length > 0) {
        wsTime.addConditionalFormatting({ ref: cfRanges.join(' '), rules: [CF_RULE] });
    }

    // ── 2. По датам ───────────────────────────────────────────────────────────
    const dHeaders = ['Дата вылета', ...sortedNights.map(n => `${n} ноч.`)];
    const dhRow = wsDates.addRow(dHeaders);
    dhRow.eachCell(c => Object.assign(c, styleHeader));
    wsDates.getColumn(1).width = 16;
    sortedNights.forEach((_, i) => wsDates.getColumn(i + 2).width = 14);

    sortedDates.forEach(dk => {
        const row = [fmtDateDow(dk)];
        sortedNights.forEach(nights => {
            const g = grid[dk] && grid[dk][nights];
            let minP = null;
            if (g) Object.values(g).forEach(rs => Object.values(rs).forEach(c => { if (minP===null||c.minPrice<minP) minP=c.minPrice; }));
            row.push(minP);
        });
        const r = wsDates.addRow(row);
        sortedNights.forEach((_, i) => {
            const c = r.getCell(i + 2);
            if (c.value !== null) { c.numFmt='# ##0'; c.alignment={horizontal:'center'}; }
            else { c.value='—'; c.alignment={horizontal:'center'}; c.font={color:{argb:'FFBBBBBB'}}; }
        });
    });
    wsDates.addConditionalFormatting({
        ref: `B2:${String.fromCharCode(65+sortedNights.length)}${sortedDates.length+1}`,
        rules: [CF_RULE]
    });

    // ── 3. Все варианты ───────────────────────────────────────────────────────
    const aHeaders = ['₽','Вылет','Ноч.','Туда дата','Туда время','Прилёт MSQ','Обратно дата','Обратно время','Прилёт MOW'];
    const ahRow = wsAll.addRow(aHeaders);
    ahRow.eachCell(c => Object.assign(c, styleHeader));
    [12,10,5,10,10,10,12,12,10].forEach((w,i) => wsAll.getColumn(i+1).width = w);
    allOptions.forEach(r => {
        const row = wsAll.addRow(aHeaders.map(k => r[k]));
        row.getCell(1).numFmt = '# ##0';
    });

    // ── 4. Все сводки ─────────────────────────────────────────────────────────
    const detH = ['Цена','Кол-во','Пересадка','Р1','Р1 Дата','Р1 Вылет','Р2','Р2 Дата','Р2 Вылет','Р3','Р3 Дата','Р3 Вылет'];
    const deRow = wsDetails.addRow(detH);
    deRow.eachCell(c => Object.assign(c, styleHeader));
    [20,5,12,10,8,14,10,8,14,10,8,14].forEach((w,i) => wsDetails.getColumn(i+1).width = w);
    createdFiles.forEach(f => {
        const label = `${fmtDateDow(f.dk)} → ${String(f.rD).padStart(2,'0')}.${String(f.rM).padStart(2,'0')} ${dowStr(f.rD,f.rM)}  (${f.nights}н)`;
        const sr = wsDetails.addRow([`═══ ${label} ═══`]);
        wsDetails.mergeCells(sr.number, 1, sr.number, detH.length);
        Object.assign(sr.getCell(1), { font:{bold:true}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EFDA'}} });
        const wb2 = XLSX.readFile(path.join(OUTPUT_DIR, f.file));
        const s = wb2.Sheets['Сводка'];
        if (!s) return;
        XLSX.utils.sheet_to_json(s).forEach(r => wsDetails.addRow(detH.map(k => r[k]||'')));
        wsDetails.addRow([]);
    });

    // ── 5а. Рек - Карточки ─────────────────────────────────────────────────────
    wsRec.getColumn(1).width = 30;
    wsRec.getColumn(2).width = 88;

    // Итоговый вывод вверху
    const best = groups[0];
    const bestText =
        `Самый выгодный вариант: ${fmtDateDow(best.dk)}, ${best.nights}н ` +
        `(обратно ${fmtDateDow(best.retDk)}) — от ${fmtRub(best.globalMin)}. ` +
        `Вылетать ${best.bestCell.outSlot.slice(0,5)} (${best.bestCell.outTime}), ` +
        `обратно ${best.bestCell.retSlot.slice(0,5)} (${best.bestCell.retTime}). ` +
        best.advice;
    const topR = wsRec.addRow(['💡 Лучший вариант', bestText]);
    wsRec.mergeCells(topR.number, 1, topR.number, 2);
    topR.getCell(1).value = `💡 Лучший вариант — ${fmtDateDow(best.dk)}, ${best.nights}н, от ${fmtRub(best.globalMin)}`;
    topR.getCell(1).font  = { bold:true, size:13, color:{argb:'FFFFFFFF'} };
    topR.getCell(1).fill  = { type:'pattern', pattern:'solid', fgColor:{argb:'FF1F6B35'} };
    topR.height = 22;
    wsRec.addRow(['', bestText]).getCell(2).alignment = { wrapText:true };
    wsRec.addRow([]);

    // Карточка на каждую (дата, ночи) — в порядке от дешёвой к дорогой
    groups.forEach((grp, idx) => {
        const { dk, nights, retDk, globalMin, globalMax, bestCell, worstCell,
                outSlotMins, slotRating, spread, spreadLabel, advice } = grp;

        // Заголовок карточки
        const rankLabel = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `${idx+1}.`;
        const cardTitle = `${rankLabel} ${fmtDateDow(dk)} → ${fmtDateDow(retDk)}  (${nights} ноч.)   мин. ${fmtRub(globalMin)}`;
        const cr = wsRec.addRow([cardTitle]);
        wsRec.mergeCells(cr.number, 1, cr.number, 2);
        Object.assign(cr.getCell(1), sCard);
        cr.height = 18;

        // Мин / макс
        const worstPct = Math.round((globalMax / globalMin - 1) * 100);
        addCardRow(wsRec, 'Диапазон цен',
            `от ${fmtRub(globalMin)} (${bestCell.outSlot.slice(0,5)} → ${bestCell.retSlot.slice(0,5)})  до  ${fmtRub(globalMax)} (${worstCell.outSlot.slice(0,5)} → ${worstCell.retSlot.slice(0,5)}, +${worstPct}%)`,
            sLabel);

        // Разброс
        addCardRow(wsRec, 'Разброс', spreadLabel, sLabel);

        // Строка по каждому времени вылета
        SLOTS.forEach(slot => {
            if (!outSlotMins[slot]) return;
            const pct = Math.round((outSlotMins[slot] / globalMin - 1) * 100);
            const label = slot;
            const val   = pct === 0
                ? `${fmtRub(outSlotMins[slot])}  ← минимум`
                : `${fmtRub(outSlotMins[slot])}  (+${pct}% к минимуму)`;
            const rating = slotRating[slot];
            const style  = rating === 'good' ? sGood : rating === 'ok' ? sOk : sBad;
            addCardRow(wsRec, label, val, style);
        });

        // Совет
        addCardRow(wsRec, '💡 Совет', advice, sTip);
        wsRec.addRow([]);
    });

    // ── 5б. Рек - Таблица ──────────────────────────────────────────────────────
    wsRecT.getColumn(1).width = 5;
    wsRecT.getColumn(2).width = 14;
    wsRecT.getColumn(3).width = 14;
    wsRecT.getColumn(4).width = 6;
    wsRecT.getColumn(5).width = 18;
    wsRecT.getColumn(6).width = 9;
    wsRecT.getColumn(7).width = 18;
    wsRecT.getColumn(8).width = 26;
    wsRecT.getColumn(9).width = 64;

    const tHeaders = ['#', 'Вылет', 'Возврат', 'Ноч.', 'Мин цена (2 пас.)', 'Кол-во', 'Лучший слот вылета', 'Разброс', '💡 Совет'];
    const thRow = wsRecT.addRow(tHeaders);
    thRow.eachCell(c => Object.assign(c, styleHeader));
    wsRecT.views = [{ state: 'frozen', ySplit: 1 }];

    groups.forEach((grp, idx) => {
        const { dk, nights, retDk, globalMin, bestCell, spread, advice } = grp;
        const rankLabel = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `${idx+1}.`;
        const spreadPct   = Math.round((spread - 1) * 100);
        const spreadShort = spread < 1.20 ? `небольшой (+${spreadPct}%)`
                          : spread < 1.50 ? `умеренный (+${spreadPct}%)`
                          :                 `большой (+${spreadPct}%)`;
        const cnt = (ticketCount[dk] && ticketCount[dk][nights]) || 0;
        const r = wsRecT.addRow([
            rankLabel,
            fmtDateDow(dk),
            fmtDateDow(retDk),
            nights,
            globalMin,
            cnt,
            `${bestCell.outSlot.slice(0,5)} (${bestCell.outTime})`,
            spreadShort,
            advice
        ]);
        r.getCell(1).alignment = { horizontal: 'center' };
        r.getCell(4).alignment = { horizontal: 'center' };
        r.getCell(5).numFmt   = '# ##0';
        r.getCell(5).alignment = { horizontal: 'center' };
        r.getCell(6).alignment = { horizontal: 'center' };
        r.getCell(7).alignment = { horizontal: 'center' };
        r.getCell(9).alignment = { wrapText: true };
        r.height = 32;
        const ratio = grp.globalMin / groups[0].globalMin;
        if (ratio <= 1.05) r.getCell(5).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFE2EFDA'} };
    });

    wsRecT.addConditionalFormatting({
        ref: `E2:E${groups.length + 1}`,
        rules: [CF_RULE]
    });

    // ── 5в. Рек - По дате ─────────────────────────────────────────────────────
    const groupsByDate = [...groups].sort((a, b) => {
        const [da, ma] = a.dk.split('.').map(Number);
        const [db, mb] = b.dk.split('.').map(Number);
        const diff = doy(da, ma) - doy(db, mb);
        return diff !== 0 ? diff : a.nights - b.nights;
    });

    wsRecTD.getColumn(1).width = 14;
    wsRecTD.getColumn(2).width = 14;
    wsRecTD.getColumn(3).width = 6;
    wsRecTD.getColumn(4).width = 18;
    wsRecTD.getColumn(5).width = 9;
    wsRecTD.getColumn(6).width = 18;
    wsRecTD.getColumn(7).width = 26;
    wsRecTD.getColumn(8).width = 64;

    const tdHeaders = ['Вылет', 'Возврат', 'Ноч.', 'Мин цена (2 пас.)', 'Кол-во', 'Лучший слот вылета', 'Разброс', '💡 Совет'];
    const tdRow = wsRecTD.addRow(tdHeaders);
    tdRow.eachCell(c => Object.assign(c, styleHeader));
    wsRecTD.views = [{ state: 'frozen', ySplit: 1 }];

    groupsByDate.forEach(grp => {
        const { dk, nights, retDk, globalMin, bestCell, spread, advice } = grp;
        const spreadPct   = Math.round((spread - 1) * 100);
        const spreadShort = spread < 1.20 ? `небольшой (+${spreadPct}%)`
                          : spread < 1.50 ? `умеренный (+${spreadPct}%)`
                          :                 `большой (+${spreadPct}%)`;
        const cnt = (ticketCount[dk] && ticketCount[dk][nights]) || 0;
        const r = wsRecTD.addRow([
            fmtDateDow(dk),
            fmtDateDow(retDk),
            nights,
            globalMin,
            cnt,
            `${bestCell.outSlot.slice(0,5)} (${bestCell.outTime})`,
            spreadShort,
            advice
        ]);
        r.getCell(3).alignment = { horizontal: 'center' };
        r.getCell(4).numFmt   = '# ##0';
        r.getCell(4).alignment = { horizontal: 'center' };
        r.getCell(5).alignment = { horizontal: 'center' };
        r.getCell(6).alignment = { horizontal: 'center' };
        r.getCell(8).alignment = { wrapText: true };
        r.height = 32;
    });

    wsRecTD.addConditionalFormatting({
        ref: `D2:D${groupsByDate.length + 1}`,
        rules: [CF_RULE]
    });

    // ── 5г. Рек - По окнам ─────────────────────────────────────────────────────
    wsRecW.getColumn(1).width = 16;
    sortedNights.forEach((_, i) => wsRecW.getColumn(i + 2).width = 14);
    wsRecW.getColumn(sortedNights.length + 2).width = 64;

    function dateInWindow(dk, wd) {
        const [dd, dm] = dk.split('.').map(Number);
        const [fd, fm] = wd['от'].split('.').map(Number);
        const [td, tm] = wd['до'].split('.').map(Number);
        return doy(dd, dm) >= doy(fd, fm) && doy(dd, dm) <= doy(td, tm);
    }

    const winCfRanges = [];

    windowDefs.forEach((wd, wi) => {
        const winDates = sortedDates.filter(dk => dateInWindow(dk, wd));
        if (winDates.length === 0) return;

        // Лучший вариант в окне
        let winBest = null;
        winDates.forEach(dk => {
            sortedNights.forEach(nights => {
                const a = analyzeGroup(dk, nights);
                if (!a) return;
                if (!winBest || a.globalMin < winBest.globalMin) {
                    const [d, m] = dk.split('.').map(Number);
                    const ret = addDaysTo(d, m, nights);
                    const retDk = `${String(ret.d).padStart(2,'0')}.${String(ret.m).padStart(2,'0')}`;
                    winBest = { dk, nights, retDk, ...a };
                }
            });
        });

        // Заголовок окна
        const wr = wsRecW.addRow([`Окно ${wi+1}:  вылет ${fmtDateDow(wd['от'])} — ${fmtDateDow(wd['до'])}`]);
        wsRecW.mergeCells(wr.number, 1, wr.number, sortedNights.length + 2);
        Object.assign(wr.getCell(1), {
            font: { bold:true, size:12, color:{argb:'FFFFFFFF'} },
            fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF4472C4'} }
        });
        wr.height = 20;

        // Заголовок таблицы
        const wh = wsRecW.addRow(['Дата вылета', ...sortedNights.map(n => `${n} ноч.`), '💡 Лучший вариант на этот день']);
        wh.eachCell(c => Object.assign(c, styleHeader));

        const dataStart = wsRecW.lastRow.number + 1;

        winDates.forEach(dk => {
            // Лучшие ночи для этой даты
            let bestNightsForDate = null, bestPriceForDate = Infinity;
            sortedNights.forEach(nights => {
                const a = analyzeGroup(dk, nights);
                if (a && a.globalMin < bestPriceForDate) {
                    bestPriceForDate = a.globalMin;
                    bestNightsForDate = nights;
                }
            });
            const bestA = bestNightsForDate !== null ? analyzeGroup(dk, bestNightsForDate) : null;

            const rowVals = [fmtDateDow(dk)];
            sortedNights.forEach(nights => {
                const a = analyzeGroup(dk, nights);
                rowVals.push(a ? a.globalMin : null);
            });
            rowVals.push(bestA
                ? `${bestNightsForDate}н — от ${fmtRub(bestA.globalMin)}. ${bestA.advice}`
                : '—');

            const r = wsRecW.addRow(rowVals);
            sortedNights.forEach((_, i) => {
                const c = r.getCell(i + 2);
                if (c.value !== null) { c.numFmt = '# ##0'; c.alignment = { horizontal: 'center' }; }
                else { c.value = '—'; Object.assign(c, styleMiss); }
            });
            r.getCell(sortedNights.length + 2).alignment = { wrapText: true };
            r.height = 32;
        });

        const dataEnd = wsRecW.lastRow.number;
        if (dataEnd >= dataStart) {
            winCfRanges.push(
                `${wsRecW.getCell(dataStart, 2).address}:${wsRecW.getCell(dataEnd, sortedNights.length + 1).address}`
            );
        }

        // Итоговый баннер окна
        if (winBest) {
            const sumText = `Лучший в окне: ${fmtDateDow(winBest.dk)} → ${fmtDateDow(winBest.retDk)} (${winBest.nights}н) — от ${fmtRub(winBest.globalMin)}. ${winBest.advice}`;
            const sr = wsRecW.addRow([sumText]);
            wsRecW.mergeCells(sr.number, 1, sr.number, sortedNights.length + 2);
            Object.assign(sr.getCell(1), {
                font:      { bold:true, italic:true, color:{argb:'FF1F6B35'} },
                fill:      { type:'pattern', pattern:'solid', fgColor:{argb:'FFE2EFDA'} },
                alignment: { wrapText: true }
            });
            sr.height = 36;
        }

        wsRecW.addRow([]);
    });

    if (winCfRanges.length > 0) {
        wsRecW.addConditionalFormatting({ ref: winCfRanges.join(' '), rules: [CF_RULE] });
    }

    // ── 5е. Итог ──────────────────────────────────────────────────────────────
    {
        const overallMin  = groups[0].globalMin;
        const GOOD_T      = overallMin * 1.20;
        const OK_T        = tripConfig['бюджет_макс'] || overallMin * 1.45;
        const pax         = tripConfig['маршрут']['пассажиры'] || 1;
        const routeNames  = tripConfig['маршрут']['названия'];
        const routeStr    = routeNames[0] + ' → ' + routeNames[1];
        const windowStr   = windowDefs.map(w => `${w['от']}–${w['до']}`).join('  и  ');

        const greenGroups = groups.filter(g => g.globalMin <= GOOD_T);
        const okGroups    = groups.filter(g => g.globalMin > GOOD_T && g.globalMin <= OK_T);
        const redGroups   = groups.filter(g => g.globalMin > OK_T);

        wsItog.getColumn(1).width = 90;

        const sTitle  = { font:{bold:true,size:14,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF1F3864'}}, alignment:{wrapText:false} };
        const sSub    = { font:{size:10,color:{argb:'FF44546A'}} };
        const sGreenH = { font:{bold:true,size:11,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF375623'}} };
        const sOkH    = { font:{bold:true,size:11,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF7D6608'}} };
        const sRedH   = { font:{bold:true,size:11,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF9C0006'}} };
        const sItem   = { font:{size:11}, alignment:{wrapText:false,indent:1} };
        const sConcl  = { font:{italic:true,size:11,color:{argb:'FF1F3864'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFDCE6F1'}}, alignment:{wrapText:true} };
        const sConcH  = { font:{bold:true,size:12,color:{argb:'FF1F3864'}} };

        function iRow(text, style, h) {
            const r = wsItog.addRow([text]);
            if (style) Object.assign(r.getCell(1), style);
            r.getCell(1).alignment = { ...(r.getCell(1).alignment || {}), wrapText: true };
            if (h) r.height = h;
            return r;
        }

        function grpLine(g) {
            const pct    = Math.round((g.globalMin / overallMin - 1) * 100);
            const pctStr = pct === 0 ? '  ← минимум' : `  +${pct}% к минимуму`;
            const pp     = pax > 1 ? `  (≈ ${fmtRub(Math.round(g.globalMin / pax))}/чел)` : '';
            const slot   = g.spread >= 1.20 ? `  ·  лучше ${g.bestCell.outSlot.slice(0,5)} (${g.bestCell.outTime})` : '  ·  время вылета почти не влияет';
            return `    •  ${fmtDateDow(g.dk)} → ${fmtDateDow(g.retDk)}  (${g.nights}н)  —  от ${fmtRub(g.globalMin)} на ${pax} пасс.${pp}${pctStr}${slot}`;
        }

        iRow(`Итог поиска: ${routeStr}`, sTitle, 28);
        iRow(`Окна вылета: ${windowStr}  ·  длительность: ${sortedNights.join(' или ')} ноч.  ·  пассажиров: ${pax}`, sSub);
        wsItog.addRow([]);

        iRow(`✅  БЕРЁМ  —  зелёная зона (не дороже +20% от минимума)`, sGreenH);
        greenGroups.forEach(g => iRow(grpLine(g), sItem, 20));
        wsItog.addRow([]);

        if (okGroups.length) {
            iRow(`⚠️  ЧУТЬ ДОРОГО, НО МОЖНО  —  до ${fmtRub(OK_T)}`, sOkH);
            okGroups.forEach(g => iRow(grpLine(g), sItem, 20));
            wsItog.addRow([]);
        }

        if (redGroups.length) {
            iRow(`🚫  НЕ БЕРЁМ  —  дороже ${fmtRub(OK_T)}`, sRedH);
            redGroups.forEach(g => iRow(grpLine(g), sItem, 20));
            wsItog.addRow([]);
        }

        iRow(`💡  ВЫВОД`, sConcH);

        const best = groups[0];
        iRow(
            `Самый выгодный: ${fmtDateDow(best.dk)} → ${fmtDateDow(best.retDk)} (${best.nights}н) — от ${fmtRub(best.globalMin)} на ${pax} пасс. ${best.advice}`,
            sConcl, 40
        );
        wsItog.addRow([]);

        windowDefs.forEach((wd, wi) => {
            const [fd, fm] = wd['от'].split('.').map(Number);
            const [td, tm] = wd['до'].split('.').map(Number);
            const winDates = sortedDates.filter(dk => {
                const [dd, dm] = dk.split('.').map(Number);
                return doy(dd, dm) >= doy(fd, fm) && doy(dd, dm) <= doy(td, tm);
            });
            let winBest = null;
            winDates.forEach(dk => {
                sortedNights.forEach(nights => {
                    const a = analyzeGroup(dk, nights);
                    if (!a) return;
                    if (!winBest || a.globalMin < winBest.globalMin) {
                        const [d, m] = dk.split('.').map(Number);
                        const ret = addDaysTo(d, m, nights);
                        const retDk = `${String(ret.d).padStart(2,'0')}.${String(ret.m).padStart(2,'0')}`;
                        winBest = { dk, nights, retDk, ...a };
                    }
                });
            });
            if (!winBest) return;
            const winLabel = `Окно ${wi+1} (${wd['от']}–${wd['до']}): лучший — ${fmtDateDow(winBest.dk)} → ${fmtDateDow(winBest.retDk)} (${winBest.nights}н), от ${fmtRub(winBest.globalMin)}. ${winBest.advice}`;
            iRow(winLabel, sConcl, 40);
        });
    }

    // ── 6. Пояснения ──────────────────────────────────────────────────────────
    wsHelp.getColumn(1).width = 24;
    wsHelp.getColumn(2).width = 88;
    [
        ['Вкладка', 'Что показывает'],
        ['По времени вылета', 'Строки = время вылета из Москвы (Ночь/Утро/День/Вечер). Столбцы = время обратного вылета из Минска. Ячейка — минимальная цена. Строка "Лучший ↓" внизу каждой группы = дешевле всего вылететь для этого времени возврата. Столбец "Лучший →" = дешевле всего вернуться для данного времени вылета.'],
        ['По датам', 'Минимальная цена по дате × кол-во ночей. Числа с цветовой шкалой.'],
        ['Все варианты', 'Все рейсы по возрастанию цены. Используй Ctrl+Shift+L (фильтр).'],
        ['Рек - Карточки', 'Детальные карточки на каждый (дата + ночи) вариант. Отсортированы от дешёвого к дорогому. Каждая карточка: диапазон цен, цвет-рейтинг по слотам вылета, итоговый совет.'],
        ['Рек - Таблица', 'Все варианты в одной строке: вылет, возврат, ночей, мин цена, лучший слот, разброс, совет. Удобно для быстрого сравнения.'],
        ['Рек - По окнам', 'Анализ по окнам вылета. Для каждого окна: сетка дата × ночи с минимальными ценами и цветовой шкалой, совет по каждой дате, итоговый лучший вариант в окне.'],
        ['Все сводки', 'Детализация из каждого отдельного поиска.'],
        ['', ''],
        ['Цены', 'Всё за двух пассажиров туда + обратно. На одного = ÷ 2.'],
        ['MOW', 'Москва (SVO / DME / VKO)'],
        ['MSQ', 'Минск — Национальный аэропорт'],
        ['Данные', new Date().toLocaleString('ru-RU')],
    ].forEach((row, i) => {
        const r = wsHelp.addRow(row);
        if (i === 0) r.eachCell(c => Object.assign(c, styleHeader));
        r.getCell(2).alignment = { wrapText: true };
    });

    // ── Сохранить и открыть ───────────────────────────────────────────────────
    const now = new Date();
    const ds = (now.getMonth()+1).toString().padStart(2,'0') + '-' + now.getDate().toString().padStart(2,'0');
    const ts = now.toTimeString().slice(0,5).replace(':', '-');
    const routeLabel = tripConfig['маршрут']['названия'].slice(0, -1).join('-');
    const outFile = path.join(OUTPUT_DIR, `as_СВОДКИ_${routeLabel}_${ds}_${ts}.xlsx`);
    await wb.xlsx.writeFile(outFile);
    console.log(`\n✓ ${outFile}`);
    execSync(`open "${outFile}"`);
}

build().catch(console.error);
