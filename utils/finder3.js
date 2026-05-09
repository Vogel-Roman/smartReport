const fs = require('fs');
const path = require('path');
// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("datasourse.json"))
    errFinish("Отсутсвует файл настроек: settings.json");
//  Считываем данные и удаляем BOM если он есть
let settings_data = fs.readFileSync("datasourse.json", { encoding: "utf-8" });
if (settings_data.charCodeAt(0) === 0xFEFF)
    settings_data = settings_data.slice(1);



/**
 * Очистка строки: нижний регистр, только буквы и цифры, нормализация пробелов.
 * Возвращает массив слов.
 */
function getWordsArray(str) {
    if (typeof str !== 'string') return [];
    let cleaned = str.toLowerCase();
    cleaned = cleaned.replace(/[^\p{L}\p{N}]+/gu, ' ');
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    return cleaned.length === 0 ? [] : cleaned.split(' ');
}

/**
 * Количество уникальных общих слов между двумя массивами.
 */
function countCommonWords(arr1, arr2) {
    const set2 = new Set(arr2);
    const common = new Set(arr1.filter(word => set2.has(word)));
    return common.size;
}

/**
 * Сопоставляет элементы исходного массива с элементами массива поиска.
 * @param {string[]} sources - исходные строки (эталоны)
 * @param {string[]} pool - массив для поиска (будет сокращаться)
 * @param {number} minCommon - минимальное количество общих слов для считанного совпадения (по умолчанию 2)
 * @returns {Array<[number, number]>} массив пар [индекс в sources, индекс в pool]
 */
function matchSourcesToPool(sources, pool, minCommon = 2) {
    // Предварительно вычисляем слова для всех источников (оптимизация)
    const sourceWordsList = sources.map(s => getWordsArray(s));
    // Доступные индексы в pool (не мутируем сам pool, только список индексов)
    let availableIndices = pool.map((_, idx) => idx);
    const matches = [];

    for (let srcIdx = 0; srcIdx < sources.length; srcIdx++) {
        const refWords = sourceWordsList[srcIdx];
        if (refWords.length === 0) continue; // пустой источник пропускаем

        let bestPoolIdx = null;
        let bestCommonCount = -1;

        // Перебираем все ещё доступные строки поиска
        for (const poolIdx of availableIndices) {
            const poolWords = getWordsArray(pool[poolIdx]);
            const common = countCommonWords(refWords, poolWords);
            if (common > bestCommonCount) {
                bestCommonCount = common;
                bestPoolIdx = poolIdx;
            }
        }

        // Если нашли хотя бы minCommon общих слов – считаем совпадение
        if (bestCommonCount >= minCommon && bestPoolIdx !== null) {
            matches.push([srcIdx, bestPoolIdx]);
            // Удаляем использованный индекс из доступных
            availableIndices = availableIndices.filter(idx => idx !== bestPoolIdx);
        }
    }

    return matches;
}

// Пример использования
// ========== Пример использования ==========
// const sources = [
//     'ЛДСП «Серый перламутровый» U763 ST9 16 мм, EGGER',
//     'ЛДСП «Дуб Давос трюфель» H1333 ST12 16 мм, EGGER',
//     'ЛДСП «Дуб Гладстоун песочный» H3309 ST28 16 мм, EGGER',
//     'ЛДСП «Индиго синий» U599 ST9 16 мм, EGGER',
//     'ЛДСП «Хромикс белый» U637 ST10 16 мм, EGGER'
// ];


//  Данные настроек внешнего файла
const datasourse = JSON.parse(settings_data);
const searchPool = datasourse.template21;
const sources = datasourse.template22;

console.log(sources.length);
console.log(searchPool.length);



const matches = matchSourcesToPool(sources, searchPool, 2);
console.log(matches);
console.log(`Сопоставлено: ${matches.length}`);
