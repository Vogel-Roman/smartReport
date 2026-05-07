/**
 * Умная сортировка массива объектов по нескольким полям
 * @param {Array} arr - массив объектов для сортировки
 * @param {Array} options - массив критериев сортировки [["поле", "направление"], ...]
 * @returns {Array} - отсортированный массив
 */
function smartSort(arr, options) {
    if (!arr || !Array.isArray(arr)) return [];
    if (!options || !Array.isArray(options) || options.length == 0) return arr;

    function getNestedValue(obj, path) {
        return path.split('.').reduce((item, key) => {
            return item && item[key] !== undefined ? item[key] : undefined;
        }, obj);
    };

    //   Сравнение значений разных типов
    function compareValues(a, b, field) {
        // Числа
        if (typeof a === 'number' && typeof b === 'number') {
            return a - b;
        };

        // Даты
        if (a instanceof Date && b instanceof Date) {
            return a.getTime() - b.getTime();
        };

        // Строки (регистронезависимое сравнение для строк)
        if (typeof a === 'string' && typeof b === 'string') {
            //return a.localeCompare(b, 'ru', { sensitivity: 'base' });
            return a.localeCompare(b, undefined, { sensitivity: 'base' });
        };

        // Разные типы или другие случаи
        return String(a).localeCompare(String(b));
    };

    return [...arr].sort((a, b) => {
        for (let [field, direction] of options) {

            // Получаем значения для сравнения
            let val_a = getNestedValue(a, field);
            let val_b = getNestedValue(b, field);

            // Обработка null/undefined
            val_a = val_a ?? '';
            val_b = val_b ?? '';

            // Сравнение
            let comparison = compareValues(val_a, val_b, field);

            if (comparison !== 0) {
                // Применяем направление сортировки
                if (direction === 'desc' || direction === 'des') {
                    return -comparison;
                } else {
                    return comparison
                };
            };
            // Если равны, переходим к следующему критерию
        };
        return 0;
    });
};

// /**
//  * Получение вложенного значения по пути (например, "user.name")
//  */
// function getNestedValue(obj, path) {
//     return path.split('.').reduce((current, key) => {
//         return current && current[key] !== undefined ? current[key] : undefined;
//     }, obj);
// };

/**
 * Сравнение значений разных типов
 */
// function compareValues(a, b, field) {
//     // Числа
//     if (typeof a === 'number' && typeof b === 'number') {
//         return a - b;
//     }

//     // Даты
//     if (a instanceof Date && b instanceof Date) {
//         return a.getTime() - b.getTime();
//     }

//     // Строки (регистронезависимое сравнение для строк)
//     if (typeof a === 'string' && typeof b === 'string') {
//         return a.localeCompare(b, 'ru', { sensitivity: 'base' });
//     }

//     // Разные типы или другие случаи
//     return String(a).localeCompare(String(b));
// }


const products = [
    { category: 'Мебель', price: 5000, name: 'Стол' },
    { category: 'Мебель', price: 3000, name: 'Стул' },
    { category: 'Фурнитура', price: 500, name: 'Петля' },
    { category: 'Фурнитура', price: 1000, name: 'Ручка' },
    { category: 'Мебель', price: 4000, name: 'Шкаф' }
];

// Сначала по category (возрастание), затем по price (убывание)
const sorted2 = smartSort(products, [
    ['category', 'asc'],
    ['price', 'desc']
]);

console.log(sorted2);
// Результат:
// Фурнитура: Ручка (1000), Петля (500)
// Мебель: Шкаф (4000), Стол (5000), Стул (3000) - по убыванию цены