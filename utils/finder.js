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

// utils.js или прямо в коде

/**
 * Разбивает строку на значимые токены (слова + цифровые блоки)
 * Убирает знаки препинания, лишние пробелы, приводит к нижнему регистру.
 */
function tokenize(str) {
    // Убираем всё, кроме букв, цифр, пробелов и слэша (для артикулов типа U763/ST9)
    const cleaned = str.replace(/[^\w\s\/-]/g, ' ').toLowerCase();
    return cleaned.split(/\s+/).filter(t => t.length > 0);
}

/**
 * Автоматический поиск: все токены эталона должны присутствовать в целевой строке.
 * @param {string} target - искомая строка
 * @param {string[]} mainTokens - токены эталонной строки
 * @returns {boolean}
 */
function isSimilarAutomatic(target, mainTokens) {
    const targetTokens = tokenize(target);
    return mainTokens.every(token => targetTokens.includes(token));
}

/**
 * Обучение на одном примере соответствия: строим RegExp, который допускает
 * перестановку токенов и любые вставки между ними.
 * @param {string} main - эталонная строка
 * @param {string} example - похожая строка (из массива поиска)
 * @returns {RegExp} регулярное выражение для поиска других похожих строк
 */
function buildRegexFromExample(main, example) {
    const mainTokens = tokenize(main);
    const exTokens = tokenize(example);

    // Находим общие токены в том порядке, в котором они встречаются в example
    // (это даст нам порядок следования частей в "правильном" шаблоне)
    const commonTokens = [];
    for (const token of exTokens) {
        if (mainTokens.includes(token) && !commonTokens.includes(token)) {
            commonTokens.push(token);
        }
    }

    // Если общих токенов мало – возвращаем "либеральное" выражение
    if (commonTokens.length < 2) {
        return new RegExp(mainTokens.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('.*?'), 'i');
    }

    // Экранируем специальные символы в токенах
    const escapedTokens = commonTokens.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    // Строим шаблон: между токенами может быть что угодно (.*?), а до первого и после последнего – тоже что угодно
    const pattern = `^.*?${escapedTokens.join('.*?')}.*$`;
    return new RegExp(pattern, 'i');
}

/**
 * Главная функция поиска.
 * @param {string} mainString - эталонная строка
 * @param {string[]} searchArray - массив строк, среди которых ищем
 * @param {Array<[string, string]>} examples - опциональные пары (основная, похожая) для обучения
 * @returns {string[]} найденные похожие строки
 */
function findSimilarStrings(mainString, searchArray, examples = []) {
    // Если есть хотя бы один пример – используем обучение
    if (examples.length > 0) {
        // Берём первый пример для построения RegExp (можно усреднять несколько)
        const [mainEx, matchEx] = examples[0];
        // Для надёжности можно проверить, что mainEx совпадает с mainString
        const regex = buildRegexFromExample(mainEx, matchEx);
        return searchArray.filter(str => regex.test(str));
    } else {
        // Автоматический режим
        const mainTokens = tokenize(mainString);
        return searchArray.filter(str => isSimilarAutomatic(str, mainTokens));
    }
}

// ============ Пример использования с вашими данными ============
const main = 'ЛДСП «Серый перламутровый» U763 ST9 16 мм, EGGER';

//#region deepseek


// // 1. Автоматический режим (без обучения)
// console.log('=== Автоматический режим ===');
// const autoResult = findSimilarStrings(main, searchPool);
// console.log(autoResult);
// // Ожидаем все три похожие строки (первые две + четвёртую/пятую), кроме "ДСП 22 мм Белый"

// // 2. Режим с обучением (вручную укажем, что основной соответствует первому элементу поиска)
// console.log('\n=== Режим с обучением ===');
// const examples = [[main, searchPool[0]]]; // говорим: вот эталон, а вот его похожий вариант
// const trainedResult = findSimilarStrings(main, searchPool, examples);
// console.log(trainedResult);
// // Должен найти те же строки, но возможно точнее (учтёт порядок токенов из примера)

// // 3. Исключение найденных строк (пример, как можно уменьшать массив поиска)
// console.log('\n=== Исключение найденных ===');
// let remainingPool = [...searchPool];
// const found = findSimilarStrings(main, remainingPool);
// // Удаляем найденные
// remainingPool = remainingPool.filter(s => !found.includes(s));
// console.log('Осталось для следующих итераций:', remainingPool);

//#endregion

//#region Vogel

//  Берем из строки три самых динных слова


function getWordsArray(str) {
    if (typeof str !== 'string') return [undefined, undefined, undefined];

    // 1. Приводим к нижнему регистру
    let cleaned = str.toLowerCase();

    // 2. Заменяем все символы, кроме букв и цифр, на пробел
    //    \p{L} - любая буква, \p{N} - любая цифра (с флагом u)
    cleaned = cleaned.replace(/[^\p{L}\p{N}]+/gu, ' ');

    // 3. Нормализуем пробелы: заменяем последовательности пробелов на один пробел
    cleaned = cleaned.replace(/\s+/g, ' ');

    // 4. Удаляем пробелы в начале и конце
    cleaned = cleaned.trim();

    // 5. Разбиваем по пробелам
    let words = cleaned.length === 0 ? [] : cleaned.split(' ');

    // 6. Сортируем по убыванию длины
    const sorted = [...words].sort((a, b) => b.length - a.length);

    // 7. Формируем результат: первые три элемента или undefined
    const result = [];
    for (let i = 0; i < sorted.length; i++) {
        result.push(sorted[i] !== undefined ? sorted[i] : undefined);
    }
    return result;
}

function hasCommon(arr1, arr2) {
    // 1. Обрезаем undefined с конца обоих массивов
    const trimUndefined = (arr) => {
        const copy = [...arr];
        while (copy.length > 0 && copy[copy.length - 1] === undefined) {
            copy.pop();
        }
        return copy;
    };

    const a = trimUndefined(arr1);
    const b = trimUndefined(arr2);

    // 2. Считаем количество уникальных совпадений (без учёта порядка)
    const commonWords = [];

    for (let i = 0; i < 3; i++) {
        const wordA = a[i];
        // Если это слово уже учтено в commonWords – пропускаем (чтобы не считать дважды)
        if (commonWords.includes(wordA)) continue;

        // Проверяем, есть ли такое же слово в массиве b
        if (b.includes(wordA)) commonWords.push(wordA);
    }

    // 3. Если совпадений больше 2 – true
    return commonWords.length > 2;
}


const input = 'ЛДСП «Дуб Давос трюфель» H1333 ST12 16 мм, EGGER';
const inputArr = getWordsArray(input);

let ind = undefined;
searchPool.forEach((poolStr, index) => {
    let arr1 = inputArr;
    let arr2 = getWordsArray(poolStr);
    if (hasCommon(arr1, arr2)) {
        ind = index;
        return;
    };
});

if (ind) {
    console.log(searchPool[ind]);
}

//#endregion