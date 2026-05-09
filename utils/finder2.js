const fs = require('fs');
const path = require('path');
// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("datasourse.json"))
    errFinish("Отсутсвует файл настроек: settings.json");
//  Считываем данные и удаляем BOM если он есть
let settings_data = fs.readFileSync("datasourse.json", { encoding: "utf-8" });
if (settings_data.charCodeAt(0) === 0xFEFF)
    settings_data = settings_data.slice(1);

//  Данные настроек внешнего файла
const datasourse = JSON.parse(settings_data);

const searchPool = datasourse.template1;

/**
 * Модуль поиска похожих строк по общим ключевым словам
 * Использует итеративное увеличение минимального количества совпадений
 */

// Очищает строку: нижний регистр, только буквы и цифры, нормализует пробелы
function getWordsArray(str) {
    if (typeof str !== 'string') return [];
    let cleaned = str.toLowerCase();
    cleaned = cleaned.replace(/[^\p{L}\p{N}]+/gu, ' ');
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    return cleaned.length === 0 ? [] : cleaned.split(' ');
}

// Количество общих слов (уникальных) между двумя массивами
function countCommonWords(arr1, arr2) {
    const set2 = new Set(arr2);
    const common = new Set(arr1.filter(word => set2.has(word)));
    return common.size;
}

// Поиск индексов строк, максимально похожих на эталон
// startCount = минимальное количество общих слов на первом шаге
function findBestMatchingIndices(searchPool, referenceWords, startCount = 2) {
    let candidates = searchPool.map((_, idx) => idx);
    let lastValidCandidates = [];
    let required = startCount;

    while (candidates.length > 0) {
        lastValidCandidates = candidates;
        candidates = candidates.filter(idx => {
            const poolWords = getWordsArray(searchPool[idx]);
            return countCommonWords(referenceWords, poolWords) >= required;
        });
        required++;
    }
    return lastValidCandidates;
}

// Упрощённая версия: возвращает первую (или лучшую) найденную строку
function findBestMatchingString(searchPool, referenceString, startCount = 2) {
    const refWords = getWordsArray(referenceString);
    const indices = findBestMatchingIndices(searchPool, refWords, startCount);
    if (indices.length === 0) return null;
    // Можно выбрать один индекс (например, первый) или все
    return searchPool[indices[0]];
}

// Пример использования
const reference = 'ЛДСП «Миндаль бежевый» U211 ST19 16 мм, EGGER';

const best = findBestMatchingString(searchPool, reference, 2);
console.log('Лучшее совпадение:', best);