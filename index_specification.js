/******************************************************************************/
/***       Скрипт на формирование отчетов из Проектов Базис в Excel         ***/
/***                         SmartWood Reports v1.0                         ***/
/******************************************************************************/

const fs = require('fs');
const fsPromises = require('fs').promises;
const path = require('path');
const firebird = require('node-firebird');
const { imageSize } = require('image-size');
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const QRCode = require("qrcode");
const { XMLParser } = require('fast-xml-parser');

//#region Инициализация

//  Массивы наименований материалов для поиска в БД
const BOARD_ARRAY = [];         //  Массив названий листовых материалов
const BUTT_ARRAY = [];          //  Массив названий кромочных материалов
const PROFILE_ARRAY = [];       //  Массив названий профильных материалов
const FURNITURE_ARRAY = [];     //  Массив названий фурнитуры
const ID_MAT_ARRAY = [];        //  Массив ID материалов
const TEMP_DIR = path.join(__dirname, 'temp');  // Путь к временной папке

// Допуск для сравнения с нулем
const EPS = 1e-10;

//  Расширения обрабатываемых файлов
const extensions = ['.fr3d', '.b3d'];

// Путь к конвертеру
const CONVERTER_EXE = path.resolve(__dirname, 'service', 'converter_batch.exe');

// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("settings.json"))
    errFinish("Отсутсвует файл настроек: settings.json");

//  Считываем данные и удаляем BOM если он есть
let settings_data = fs.readFileSync("settings.json", { encoding: "utf-8" });
if (settings_data.charCodeAt(0) === 0xFEFF)
    settings_data = settings_data.slice(1);

//  Проверка структуры файла settings.json
if (!checkSettingsData(settings_data))
    errFinish("Ошибки в строуктуре файла конфигурации: settings.json");

//  Данные настроек внешнего файла
const settings = JSON.parse(settings_data);

//  Директория сохранения отчетов
const SW_FOLDER = settings.reportFolderName;

//  Классы группировки данных
const defClass = "M0";
const classes = [...Object.keys(settings.estimate.classes), defClass];

//  Инициализация констант
const PROJECT_FILE = system.askFileName('bprj');
if (!PROJECT_FILE) errFinish("Файл проекта не выбран");
const PROJECT_NAME = system.getFileNameWithoutExtension(PROJECT_FILE);

//  Путь к папке для сохранения результата
let message = "Укажите папку для сохранения файлов спецификаций";
const FOLDER = system.askFolder(message, path.dirname(PROJECT_FILE));
if (!FOLDER) errFinish("Директория сохранения результата не выбрана");

//#endregion

//#region Служебные функции

//  Функция обработки ошибок
function errFinish(str) {
    console.log(str);
    Action.Finish();
};

//  Функция проверка структуры файла settings.json
function checkSettingsData(data) {
    //  Написать алгоритм проверки струкруты файла
    return true;
};

//  Функция огругления
function round(a, b) {
    b = b || 0;
    return Math.round(a * Math.pow(10, b)) / Math.pow(10, b);
};

//  Функция очистки значения от неточности записи
function clean(val) {
    if (Math.abs(val) < EPS) return 0;
    return Math.round(val * 10) / 10;
};

//  Функция добавляющая лидирующие нули
function addZero(value, length = 3) {
    // Преобразуем в строку и добавляем нули
    return String(value).padStart(length, '0');
};

//  Функция получения артикула и названия материала из имени
function getMaterialName(matname) {
    let mName = matname;
    let mArt = "";
    if (mName.indexOf("\r") > 0) {
        mArt = mName.split("\r")[1];
        mName = mName.split("\r")[0];
    };
    return [mName, mArt];
};

// Безопасное имя файла
function sanitizeFileName(name) {
    return name
        .replace(/[<>:"/\\|?*]/g, '')
        .replace(/\s+/g, ' ')
        .substring(0, 200);
};

//  Функция сортировки массива по нескольким полям
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

//  Преобразует дату в формат ДД.ММ.ГГГГ
function formatDate(date) {
    const d = typeof date === 'string' ? new Date(date) : date;
    if (!(d instanceof Date) || isNaN(d)) return 'Некорректная дата';
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}.${month}.${year}`;
};

// Функция поиска отверстий принадлежжащих панели
function findPanelHolesList(panel) {

    //  Функция вычисления конечной точки отверстия
    function getHoleEndPoint(hole, fastener, panel) {
        // 1. Вычисляем конец отверстия в локальной системе фурнитуры
        const k = 0.8; //   Коэффициент отступа от конечной точки
        const dir = hole.Direction;
        const depth = hole.Depth * k;

        // Нормализуем направление (на случай если вектор не единичный)
        const len = Math.hypot(dir.x, dir.y, dir.z);
        const normDir = {
            x: dir.x / len,
            y: dir.y / len,
            z: dir.z / len
        };

        // Точка конца: устье + направление * глубина
        const endLocal = {
            x: hole.Position.x + normDir.x * depth,
            y: hole.Position.y + normDir.y * depth,
            z: hole.Position.z + normDir.z * depth
        };

        // 2. Переводим в глобальные координаты
        let endGlobal = fastener.ToGlobal(endLocal);

        // 3. Переводим в локальную систему координат панели
        let endInPanel = panel.ToObject(endGlobal);

        return endInPanel;
    };

    function cleanPoint(point, decimal) {
        return {
            x: clean(point.x),
            y: clean(point.y),
            z: clean(point.z)
        };
    };

    function isPointInBounds(point, minPoint, maxPoint) {
        // Определяем реальные минимумы и максимумы на случай,
        // если minPoint и maxPoint переданы не в правильном порядке
        const minX = Math.min(minPoint.x, maxPoint.x);
        const maxX = Math.max(minPoint.x, maxPoint.x);
        const minY = Math.min(minPoint.y, maxPoint.y);
        const maxY = Math.max(minPoint.y, maxPoint.y);
        const minZ = Math.min(minPoint.z, maxPoint.z);
        const maxZ = Math.max(minPoint.z, maxPoint.z);

        // Проверяем, что точка находится в пределах по каждой оси
        return point.x >= minX && point.x <= maxX &&
            point.y >= minY && point.y <= maxY &&
            point.z >= minZ && point.z <= maxZ;
    };

    //  Массив отверстий
    const result = [];

    //  Массив фурнитуры, принадлежащей панели
    let fasteners = panel.FindConnectedFasteners();
    if (!fasteners) return result;

    fasteners.forEach(fastener => {
        if (!fastener) return;

        fastener.Holes.List.forEach(hole => {
            if (!hole) return;

            //  Фильрация отверстий по минимальному диаметру
            const minDiameter = settings.exclude.minHoleDiameter;

            if (hole.Radius * 2 < minDiameter) return;

            let posInPanel =
                cleanPoint(panel.ToObject(fastener.ToGlobal(hole.Position)), 1);
            let dirInPanel =
                panel.NToObject(fastener.NToGlobal(hole.Direction));
            dirInPanel = cleanPoint(dirInPanel, 1);

            // Анализируем dirInPanel.z
            const isPerpendicular = Math.abs(dirInPanel.z) > 0.99;
            const isParallel = Math.abs(dirInPanel.z) < 0.01;

            // Координаты конечной точки отверстия в ГСК
            let endInPanel =
                cleanPoint(getHoleEndPoint(hole, fastener, panel), 1);

            /**
             * Если обсолютное значение координат x или y свойства dirInPanel
             * равно 1, то это торцевое отверстие. А если абсолютное значение
             * координаты z равно 1, то это отвестие в пласть (по умолчанию)
             */
            let typeHole =
                (
                    Math.abs(dirInPanel.x) == 1 ||
                    Math.abs(dirInPanel.y) == 1
                ) ? 2 : 1;   //  Отверстие в пласть 1, торцевое 2

            if (isPointInBounds(endInPanel, panel.GMin, panel.GMax)) {
                result.push({
                    depth: round(hole.Depth, 2),
                    diameter: hole.Radius * 2,
                    depth: round(hole.Depth, 1),
                    type: typeHole, //  1 в пласть, 2 торцевой
                    drillMode: hole.DrillMode,  // 1 сквозное, 2 глухое
                    dirInPanel: dirInPanel,
                    positionInPanel: posInPanel
                });
            };
        });
    });
    return result;
};

//  Функция получения информации о кромках панели
function findPanelButtsList(panel, k) {
    //  count - это количество изделий в проекте
    const result = [];

    for (let i = 0; i < panel.Contour.Count; i++) {
        if (!panel.Contour[i].Data.Butt) continue;

        const elem = panel.Contour[i].Data.Butt;
        if (!elem.Thickness) continue;

        const overhung = elem.Overhung;
        const contour = panel.Contour[i];
        const length = 0.001 * k * (contour.ObjLength() + overhung * 2);

        const material = getMaterialName(elem.Material);

        //  Дополняем массив наименований материалов кромки
        if (!BUTT_ARRAY.includes(material[0])) BUTT_ARRAY.push(material[0]);

        result.push({
            material: elem.Material,        //  Имя материала кромки
            materialName: material[0],      //  Имя материала
            materialArticle: material[1],   //  Артикул материала кромки
            materialSyncExternal: "",       //  Код синхронизации (DB)
            materialID: undefined,          //  ID материала (DB)
            materialUnit: "",               //  Единица измерения (DB)
            materialPrice: 0,               //  Цена материала (DB)
            allowance: elem.Allowance,      //  Припуск на прифуговку
            clipPanel: elem.ClipPanel,      //  Св-во "подрезать панель"
            sign: elem.Sign,                //  Обозначение кромки
            overhung: overhung,             //  Величина свеса кромки
            thickness: elem.Thickness,      //  Толщина кромки
            width: elem.Width,              //  Ширина кромки
            length: length                  //  Длина кромки
        });
    };

    return result;
};

//  Функция получения информации о пазах панели
function findPanelCutsList(panel) {
    function testNameRegExp(str, patterns) {
        if (!str || typeof str !== 'string') return false;
        return patterns.some(pattern => {
            if (typeof pattern !== 'string') return false;

            // Если нет звездочки - только точное совпадение
            if (!pattern.includes('*')) return str === pattern;

            // Обработка звездочек
            // Случай: звездочка в конце (например "R*", "Фаска*")
            if (pattern.endsWith('*') && !pattern.slice(0, -1).includes('*')) {
                const prefix = pattern.slice(0, -1);
                return str.startsWith(prefix);
            };

            // Случай: звездочка в начале (например "*пласти*")
            if (pattern.startsWith('*') && !pattern.slice(1).includes('*')) {
                const suffix = pattern.slice(1);
                return str === suffix; // ТОЧНОЕ совпадение для "*пласти"
            };

            // Случай: звездочка в начале и в конце (например "*пласти*")
            if (pattern.startsWith('*') && pattern.endsWith('*')) {
                const middle = pattern.slice(1, -1);
                return str.includes(middle);
            };

            // Случай: звездочка посередине (например "R*лка")
            const regexPattern = pattern.replace(/\*/g, '.*');
            const regex = new RegExp(`^${regexPattern}$`);
            return regex.test(str);
        });
    };

    const result = [];

    for (let i = 0; i < panel.Cuts.Count; i++) {
        const cutNames = settings.exclude.cutNames;
        const cut = panel.Cuts[i];
        if (testNameRegExp(cut.Name, cutNames)) continue;
        if (!cut.Params) {
            //  Паз выемка
            const w = cut.Contour.Width;
            const h = cut.Contour.Height;
            let area = round(w * h * 0.000001, 2);

            result.push({
                materialSyncExternal: "",       //  Код синхронизации (DB)
                materialUnit: "",               //  Единица измерения (DB)
                name: cut.Name,                 //  Имя паза
                sign: cut.Sign,                 //  Обозначение паза
                cutType: 11,    //  Тип паза
                area: area,
                length: 0
            });
        } else {
            //  Исключаем пазы тпов: 8 и 10
            if (
                cut.Params.CutType == 8 ||
                cut.Params.CutType == 10
            ) continue;

            const length = round(cut.Trajectory.ObjLength(), 2);
            result.push({
                materialSyncExternal: "",       //  Код синхронизации (DB)
                materialUnit: "",               //  Единица измерения (DB)
                name: cut.Name,                 //  Имя паза
                sign: cut.Sign,                 //  Обозначение паза
                cutType: cut.Params.CutType,    //  Тип паза
                length: length                  //  Длина траектории паза
            });
        };
    };
    return result;
};

//  Функция получения информации о облицовки пласти
function findPanelPlasticList(panel) {
    const result = [];
    const letters = [
        'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
        'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'
    ];

    //  Цикл по массиву облицовок панели
    for (let i = 0; i < panel.Plastics.Count; i++) {
        const plastic = panel.Plastics[i];
        result.push({
            material: plastic.Material,
            tkn: plastic.Thickness,
            ltr: letters[i]
        });
    };

    return result;
};

//  Функция считывания с объекта дополнительных материалов.
function findAdditionalMaterials(obj, modelData) {
    //  Проверка существования дополнительных материалов.
    let attNode = obj.ParamSectionNode('MaterialAttendance', true);
    let objNode = attNode.FindNode('List');

    //  На случай, если такого узла нет, для старых версий файлов
    if (!objNode) return [];

    const result = [];
    for (let i = 0; i < objNode.Count; i++) {
        const elem = objNode.Nodes[i];
        let materialNode = elem.FindNode('Name');
        let countNode = elem.FindNode('Count');

        if (materialNode) {
            const material = getMaterialName(materialNode.Value);

            //  Дополняем массив наименований листовых материалов
            if (!FURNITURE_ARRAY.includes(material[0])) {
                FURNITURE_ARRAY.push(material[0]);
            };

            modelData.data.furnitureMaterials.push({
                material: materialNode.Value,   //  Название фурнитуры
                materialID: undefined,          //  ID Материала (DB)
                materialName: material[0],      //  Имя фурнитуры
                materialArticle: material[1],   //  Артикул фурнитуры
                materialSyncExternal: "",       //  Код синхронизации (DB)
                materialPrice: 0,               //  Цена из базы данных (DB)
                materialUnit: "",               //  Единица измерения (DB)
                prjCount: modelData.count,      //  Количество
                count: modelData.count          //  Количество
            });
        };
    };
    return result;
};

//  Функция получения буквы колонки по ее номеру
function getColumnLetter(col) {
    let letter = '';
    while (col > 0) {
        col--;
        letter = String.fromCharCode(65 + (col % 26)) + letter;
        col = Math.floor(col / 26);
    };
    return letter;
};

//  Транслитерация кириллицы в латиницу
function transliterate(text, separator = '_') {
    const map = {
        'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e',
        'ё': 'jo', 'ж': 'g', 'з': 'z', 'и': 'i', 'й': 'j', 'к': 'k',
        'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r',
        'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'c',
        'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '',
        'э': 'e', 'ю': 'yu', 'я': 'ya',
        'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E',
        'Ё': 'JO', 'Ж': 'G', 'З': 'Z', 'И': 'I', 'Й': 'J', 'К': 'K',
        'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R',
        'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'KH', 'Ц': 'C',
        'Ч': 'CH', 'Ш': 'SH', 'Щ': 'SCH', 'Ъ': '', 'Ы': 'Y', 'Ь': '',
        'Э': 'E', 'Ю': 'YU', 'Я': 'Ya'
    };

    let result = '';

    for (let char of text) {
        if (map[char] !== undefined) {
            result += map[char];
        } else if (char === ' ') {
            result += separator;
        } else if (/[a-zA-Z0-9]/.test(char)) {
            result += char;
        } else {
            result += char; // оставляем другие символы как есть
        }
    }

    // Убираем множественные разделители
    result = result.replace(new RegExp(`${separator}+`, 'g'), separator);
    // Убираем разделители в начале и конце
    result = result.replace(new RegExp(`^${separator}|${separator}$`, 'g'), '');

    return result;
};

//  Функция конвертации PDF файлов чертежей в PNG файлы
async function convertPDFtoPNG(fileNames, baseDIR) {

    //  Проверяем существование конвертера
    try {
        await fsPromises.access(CONVERTER_EXE);
    } catch {
        console.log(`Файл конвертера не найден.`);
        return;
    };

    Action.Hint = "Конвертируем изображения чертежей в PNG...";
    const sourseDIR = path.join(baseDIR, 'pdf');
    const outputDIR = path.join(baseDIR, 'png');

    // Создаём выходную папку, если её нет
    await fsPromises.mkdir(outputDIR, { recursive: true });

    // Формируем массив имён с расширением .pdf
    let pdfFiles = fileNames.map(name => `${name}.pdf`);

    // Проверяем, что все файлы существуют (опционально)
    const missing = [];
    for (const file of pdfFiles) {
        const fullPath = path.join(sourseDIR, file);
        try {
            await fsPromises.access(fullPath);
        } catch {
            missing.push(file);
        };
    };
    if (missing.length > 0) {
        console.warn(`Следующие файлы не найдены и будут пропущены:`, missing.join(', '));
        // Удаляем их из списка, чтобы не пытаться обрабатывать
        pdfFiles = pdfFiles.filter(f => !missing.includes(f));
    };
    if (pdfFiles.length === 0) {
        console.warn('Нет доступных PDF-файлов для конвертации.');
        return;
    };

    // Создаём временный JSON-файл
    const jsonFilePath = path.join(__dirname, 'service', '_filenames.json');
    await fsPromises.writeFile(jsonFilePath, JSON.stringify(pdfFiles, null, 2), 'utf8');

    // Запускаем конвертер
    const command = `"${CONVERTER_EXE}" "${jsonFilePath}" "${sourseDIR}" "${outputDIR}"`;
    try {
        const { stdout, stderr } = await execPromise(command);
        if (stdout) console.log(stdout);
        if (stderr) console.warn('ERROR', stderr);
    } catch (error) {
        console.error('Ошибка при выполнении конвертера:', error.message);
        if (error.stdout) console.log(error.stdout);
        if (error.stderr) console.error(error.stderr);
    } finally {
        try {
            await fsPromises.unlink(jsonFilePath); // Удаляем временный JSON-файл
            console.log(`Временный файл удалён: ${jsonFilePath}`);
        } catch (e) {
            console.warn('Не удалось удалить временный файл:', e.message);
        };
        Action.Hint = "Изображения сформированы";
        return true;
    };
};

//  Функция поиска сборочной единицы элемента модели
function findAssemblyUnitID(item) {
    let elem = item.Owner;
    while (elem) {
        elem = elem.Owner;
        if (elem && elem.IsAssemblyUnit && elem.Owner instanceof TModel3D) {
            return {
                UID: elem.UID,
                name: elem.Name,
                pos: elem.ArtPos,
                designation: elem.Designation
            };
        };
    };
    return undefined;
};

//  Функция получения пользовательского свойства Модели
function getUserProderty(obj, propName) {
    if (obj.UserProperty[propName]) return obj.UserProperty[propName]
    return;
};

//#endregion

//#region Функции рекурсивного обхода

//  Функция рекурсивного обхода модели
function forEachInList(list, func, data) {
    if (!func) return;
    for (let i = 0; i < list.Count; i++) {
        let elem = list.Objects[i];
        func(elem, data);
        if (
            elem.List &&
            (elem instanceof TFurnBlock || elem instanceof TDraftBlock)
        ) forEachInList(elem.AsList(), func, data);
    };
};

//  Call-back функция
function callbackFunc(item, data) {
    if (item instanceof TFurnPanel) {
        //  Панель
        panelProcessing(item, data);
    } else if (item instanceof TExtrusionBody) {
        //  Профиль
        profileProcessing(item, data);
    } else if (item instanceof TFurnAsm || item instanceof TFastener) {
        //  Фурнитура или сборка (покупное изделие)
        furnitureProcessing(item, data);
    } else if (item instanceof TFurnBlock || item instanceof TDraftBlock) {
        //  Поиск дополнительных материалов к блоку или полуфабрикату
        findAdditionalMaterials(item, data);
    };
};

//  Функция получения данных файлов Проекта
function readProjectFilesData(prj_file) {
    try {
        // Читаем файл проекта (XML-файл по структуре);
        const xmlData = fs.readFileSync(PROJECT_FILE, 'utf8');

        // Настройки парсера
        const options = {
            ignoreAttributes: false,  // Не игнорировать атрибуты
            attributeNamePrefix: "@_", // Префикс для атрибутов
            isArray: (name, jpath, isLeafNode, isAttribute) => {
                // Автоматически превращать в массив эти теги
                return ['File'].includes(name);
            }
        };
        const parser = new XMLParser(options);
        const prjData = parser.parse(xmlData);

        // Получаем доступ к данным списка файлов Проекта
        const listFiles = prjData.Document.DataProject.ListFiles.File;

        //  Массив результата обхода структуры проекта
        const result_array = [];

        //  Обходим циклом массив и собираем данные о файле проекта.
        listFiles.forEach(elem => {
            //  Путь до текущего файла проекта
            const filePath = path.join(settings.coreDIR, elem.Name);
            if (!fs.existsSync(filePath))
                errFinish(`Файл по указанному пути не существует: ${filePath}`);

            //  Добавляем данные из файла Проекта в массив
            result_array.push({
                type: elem.Type,
                name: path.basename(filePath),
                dirname: path.dirname(filePath),
                filepath: filePath,
                sign: elem.Sign,
                count: elem.Count,
                subname: elem.SubName,
                note: elem.Note,
                comment: elem.Comment,
                estimate: elem.HasUse_Estimate == "Y" ? 1 : 0,
                cutting: elem.HasUse_Cutting == "Y" ? 1 : 0,
                cnc: elem.HasUse_CNC == "Y" ? 1 : 0,
                data: {     //  Контейнер для информации из Модели
                    panelMaterials: [],
                    profileMaterials: [],
                    furnitureMaterials: []
                },
                estimate_data: {}
            });
        });

        //  Возвращаем результат
        return result_array;
    } catch (e) {
        errFinish(e.message);
    };
};

//#endregion

//#region Функции обработки объектов Модели

//  Функция обработки данных панели
function panelProcessing(panel, modelData) {
    const material = getMaterialName(panel.MaterialName);

    //  Сборочная единица элемента
    const AssemblyUnit = findAssemblyUnitID(panel);

    //  Поиск дополнительных материалов к панели
    findAdditionalMaterials(panel, modelData);

    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.panelMaterial;
    if (excludeMaterial.includes(material[0])) return;

    //  Дополняем массив наименований листовых материалов
    if (!BOARD_ARRAY.includes(material[0])) BOARD_ARRAY.push(material[0]);

    const dps = settings.delimPrjSign;
    const dpn = settings.delimPrjName;
    const ms = modelData.sign;

    //  Размеры панели
    const w = round(panel.ContourWidth, 1);
    const h = round(panel.ContourHeight, 1);

    //  Размеры панели
    let w_contour = panel.ContourWidth;
    let h_contour = panel.ContourHeight;

    if (panel.TextureOrientation == 2) {
        //  Изменено направление текстуры
        w_contour = panel.ContourHeight;
        h_contour = panel.ContourWidth;
    };

    //  Цикл на определение обозначения кромок L1L2 W1W2
    const ll = [];
    const ww = [];
    for (let i = 0; i < panel.Contours[0].Count; i++) {
        const contuor = panel.Contours[0][i];
        if (contuor.IsLine() && contuor.ObjLength() >= w_contour * 0.8) {
            if (contuor.Data) ll.push(contuor.Data.Butt ? contuor.Data.Butt.Thickness : 0);
        } else if (contuor.IsLine() && contuor.ObjLength() >= h_contour * 0.8) {
            if (contuor.Data) ww.push(contuor.Data.Butt ? contuor.Data.Butt.Thickness : 0);
        };
    };
    const llww = [...ll, ...ww];

    //  Площадь панели в метрах
    const panelArea = round(w * h * 0.000001, 2)

    //  Длина контура панели в метрах
    const contourLength = round(panel.Contour.ObjLength() * 0.001, 2);

    //  Позиция панели в проекте M2-0012
    const projectPos = panel.ArtPos ?
        modelData.sign + dps + addZero(panel.ArtPos) : "";

    //  Обозначение панели в проекте M2-0012
    const projectDes =
        panel.Designation ? ms + dps + addZero(panel.Designation) : "";

    //  Текст в QR-коде (Позиция – ArtPos)
    const barcodeData = panel.ArtPos ?
        PROJECT_NAME + dpn + projectPos : "";

    //  Текст в QR-коде (Обозначение – Designation)
    const barcodeDataDes = panel.Designation ?
        PROJECT_NAME + dpn + projectDes : "";

    //  Информация о кромках панели
    const buttInfoArray = findPanelButtsList(panel, modelData.count);

    //  Информация о пазах панели
    const cutInfoArray = findPanelCutsList(panel);

    //  Информация о присадке панелей
    const drillInfoArray = findPanelHolesList(panel);

    //  Информация о облицовки пласти
    const plasticInfoArray = findPanelPlasticList(panel);

    const materialName = panel.MaterialName.trim();

    modelData.data.panelMaterials.push({
        name: panel.Name,               //  Имя панели
        material: materialName,         //  Материал панели
        materialID: undefined,          //  ID Материала (DB)
        materialName: material[0],      //  Имя материала панели
        materialArticle: material[1],   //  Артикул материала панели
        materialSyncExternal: "",       //  Код синхронизации материала (DB)
        materialPrice: 0,               //  Цена из базы данных (DB)
        materialUnit: "",               //  Единица измерения (DB)
        materialWidth: 0,               //  Длина листа (DB) мм
        materialHeight: 0,              //  Ширина листа (DB) мм
        materialTkn: panel.Thickness,   //  Толщина материала
        prjCount: modelData.count,      //  Количество
        pos: panel.ArtPos,              //  Позиция в модели
        des: panel.Designation,         //  Обозначение в модели
        prjPos: projectPos,             //  Позиция в проекте
        prjDes: projectDes,             //  Обозначение в проекте
        barcode: barcodeData,           //  Код панели в проекте (Pos)
        barcode_des: barcodeDataDes,    //  Код панели в проекте (Designation)
        width: w,                       //  Длина панели
        height: h,                      //  Ширина панели
        area: panelArea,                //  Площадь панели
        contourLength: contourLength,   //  Длина контура панели
        buttInfo: buttInfoArray,        //  Массив кромок панели
        cutInfo: cutInfoArray,          //  Массив пазов панели
        drillInfo: drillInfoArray,      //  Массив отверстий панели
        assemblyUnit: AssemblyUnit,     //  Сборочная единица
        tkn: panel.ZThickness,          //  Толщина детали по Z
        llww: llww,                     //  Информация о кромке на детали
        isPlastic: false
    });

    //  Обработка массива облицовки панели
    if (!plasticInfoArray.length) return;

    for (let i = 0; i < plasticInfoArray.length; i++) {
        const pls = plasticInfoArray[i];

        const mat = getMaterialName(pls.material);
        //  Игнорируем исключенные листовые материалы
        if (excludeMaterial.includes(mat[0])) continue;

        //  Дополняем массив наименований листовых материалов
        if (!BOARD_ARRAY.includes(mat[0])) BOARD_ARRAY.push(mat[0]);

        const plsName = panel.Name + '_' + pls.ltr;
        const plsPnlDes = panel.Designation ? panel.Designation + pls.ltr : "";
        const plsPrjDes = projectDes ? projectDes + pls.ltr : "";
        const bcDataDes = barcodeDataDes ? barcodeDataDes + pls.ltr : "";

        //  Используется допущение, что размер облицовки равен размеру детали!!!
        modelData.data.panelMaterials.push({
            name: plsName,                  //  Имя панели
            material: pls.material,         //  Материал панели
            materialID: undefined,          //  ID Материала (DB)
            materialName: mat[0].trim(),    //  Имя материала панели
            materialArticle: mat[1],        //  Артикул материала панели
            materialSyncExternal: "",       //  Код синхронизации материала (DB)
            materialPrice: 0,               //  Цена из базы данных (DB)
            materialUnit: "",               //  Единица измерения (DB)
            materialWidth: 0,               //  Длина листа (DB) мм
            materialHeight: 0,              //  Ширина листа (DB) мм
            materialTkn: pls.tkn,           //  Толщина материала
            prjCount: modelData.count,      //  Количество
            pos: panel.ArtPos + pls.ltr,    //  Позиция в модели
            des: plsPnlDes,                 //  Обозначение в модели
            prjPos: projectPos + pls.ltr,   //  Позиция в проекте
            prjDes: plsPrjDes,              //  Обозначение в проекте
            barcode: barcodeData + pls.ltr, //  Код пан. в проекте (Pos)
            barcode_des: bcDataDes,         //  Код пан. в проекте (Designation)
            width: w,                       //  Длина панели
            height: h,                      //  Ширина панели
            area: panelArea,                //  Площадь панели
            contourLength: contourLength,   //  Длина контура панели
            buttInfo: [],                   //  Массив кромок панели
            cutInfo: [],                    //  Массив пазов панели
            drillInfo: [],                  //  Массив отверстий панели
            assemblyUnit: AssemblyUnit,     //  Сборочная единица
            tkn: panel.ZThickness,          //  Толщина детали по Z
            isPlastic: true,
        });
    };
};

//  Функция обработки данных профиля
function profileProcessing(profile, modelData) {

    //  Поиск дополнительных материалов к профилю
    findAdditionalMaterials(profile, modelData);

    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.profileMaterial;
    if (excludeMaterial.includes(profile.MaterialName)) return;

    //const material = getMaterialName(profile.MaterialName);

};

//  Функция обработки данных фурнитуры
function furnitureProcessing(fstn, modelData) {

    //  Игнорируем фурнитуру не используемую в смете
    if (!fstn.UseInEstimate) return;

    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.furnitureMaterial;

    //  Поиск составной фурнитуры
    function getFastenerElements(furn) {
        let data = furn.AdvParamData;
        if (data && (data = data.FindNode('Elements'))) {
            let result = [];
            for (let i = 0; i < data.Count; i++) result.push(data[i].Value);
            return result;
        };
        return [];
    };

    //  Поиск дополнительных материалов к фурнитуре
    findAdditionalMaterials(fstn, modelData);

    /**
     * Надо исключить попадание в массив фурнитуры, глухих и сквозных 
     * параметрических отверстий. Для этого нужно протестировать имя на
     * соответствие регулярным выраждениям. /^\d+x\d+/ и /^\d/;
     */
    const regexp1 = /^\d+(,\d+)?x\d+(,\d+)?$/;  //  Имя глухих отверстий
    const regexp2 = /^\d+(,\d+)?$/;             //  Имя сквозных отверстий

    let arr = getFastenerElements(fstn);
    if (arr.length) {
        //  Составная фурнитура
        //console.log('Составная фурниутра');
        for (let i = 0; i < arr.length; i++) {
            if (arr[i].match(regexp1) || arr[i].match(regexp2)) continue;
            const material = getMaterialName(arr[i]);

            //  Игнорируем исключеннeю фурнитуру
            if (excludeMaterial.includes(material[0])) continue;

            //  Дополняем массив наименований листовых материалов
            if (!FURNITURE_ARRAY.includes(material[0])) {
                FURNITURE_ARRAY.push(material[0]);
            };

            //  Добавляем данные в массив объектов фурнитуры
            modelData.data.furnitureMaterials.push({
                material: arr[i],               //  Название фурнитуры
                materialID: undefined,          //  ID Материала (DB)
                materialName: material[0],      //  Имя фурнитуры
                materialArticle: material[1],   //  Артикул фурнитуры
                materialSyncExternal: "",       //  Код синхронизации (DB)
                materialPrice: 0,               //  Цена из базы данных (DB)
                materialUnit: "",               //  Единица измерения (DB)
                prjCount: modelData.count,      //  Количество
                count: modelData.count          //  Количество
            });
        };
    } else {
        //  Обрабатываем как обычную фурнитуру
        if (fstn.Name.match(regexp1) || fstn.Name.match(regexp2)) return;
        const material = getMaterialName(fstn.Name);

        //  Игнорируем исключеннeю фурнитуру
        if (excludeMaterial.includes(material[0])) return;

        //  Дополняем массив наименований листовых материалов
        if (!FURNITURE_ARRAY.includes(material[0])) {
            FURNITURE_ARRAY.push(material[0]);
        };

        //  Добавляем данные в массив объектов фурнитуры
        modelData.data.furnitureMaterials.push({
            material: fstn.Name,            //  Название фурнитуры
            materialID: undefined,          //  ID Материала (DB)
            materialName: material[0],      //  Имя фурнитуры
            materialArticle: material[1],   //  Артикул фурнитуры
            materialSyncExternal: "",       //  Код синхронизации (DB)
            materialPrice: 0,               //  Цена из базы данных (DB)
            materialUnit: "",               //  Единица измерения (DB)
            prjCount: modelData.count,      //  Количество
            count: modelData.count          //  Количество
        });
    };

};

//#endregion

//#region Функции формирования данных

//  Функция выполнения запроса в БД
async function executeQuery(sql, options, params = []) {
    return new Promise((resolve, reject) => {
        firebird.attach(options, (err, db) => {
            if (err) {
                console.log(err.message);
                reject(err);
                return;
            };
            db.query(sql, params, (err, result) => {
                db.detach();
                if (err) {
                    reject(err);
                } else {
                    resolve(result);
                };
            });
        });
    });
};

//  Функция получения данных о материалах из Базы материалов
async function getDBMaterialInfo() {

    //  Формируем общий массив наименований для запроса в БД
    const matnames = [
        ...BOARD_ARRAY, ...BUTT_ARRAY,
        ...PROFILE_ARRAY, ...FURNITURE_ARRAY
    ];

    try {
        const options = settings.db_options || null;

        //  1 строка запроса на получение массива id материалов
        const sqlString1 = `
            SELECT
                m.ID_M,
                m.NAME_MAT
            FROM MATERIAL AS m
        `;
        const allMaterials = await executeQuery(sqlString1, options);
        const nameToIdsMap = new Map();
        for (const material of allMaterials) {
            const name = material.name_mat;
            const id = material.id_m;
            if (!nameToIdsMap.has(name)) nameToIdsMap.set(name, []);
            nameToIdsMap.get(name).push(id);
        };

        // 3. Проходим по массиву matnames и собираем ID
        const resultIds = [];

        for (const searchName of matnames) {
            const ids = nameToIdsMap.get(searchName);
            if (ids && ids.length > 0) {
                //console.log(ids[0], searchName);

                // Добавляем все ID для этого названия (если их несколько)
                resultIds.push(...ids);
            };
        };

        const placeholders = resultIds.map(() => '?').join(',');

        //  2 строка запроса на получение данных материлов
        const sqlString2 = `
        SELECT
            m.ID_M,
            m.NAME_MAT,
            m.PRICE,
            m.SYNC_EXTERNAL,
            ma.LENGTH,
            ma.WIDTH,
            ma.THICKNESS,
            me.NAME_MEAS
        FROM MATERIAL AS m
            LEFT JOIN MATERIAL_ADVANCE AS ma ON m.ID_M = ma.ID_M
            LEFT JOIN MEASURE AS me ON m.ID_MS = me.ID_MS
        WHERE
            m.ID_M IN (${placeholders})
        `;

        // //const placeholders = matnames.map(() => '?').join(',');

        // //  1 строка запроса на получение данных материлов
        // const sqlString = `
        // SELECT
        //     m.ID_M,
        //     m.NAME_MAT,
        //     m.PRICE,
        //     m.SYNC_EXTERNAL,
        //     ma.LENGTH,
        //     ma.WIDTH,
        //     ma.THICKNESS,
        //     me.NAME_MEAS
        // FROM MATERIAL AS m
        //     LEFT JOIN MATERIAL_ADVANCE AS ma ON m.ID_M = ma.ID_M
        //     LEFT JOIN MEASURE AS me ON m.ID_MS = me.ID_MS
        // WHERE
        //     m.ID_M IN (${placeholders})
        // `;

        //  Запрос в БД
        const result_db = await executeQuery(sqlString2, options, resultIds);

        const result = [
            [], //  Листовые материалы
            [], //  Кромочные материалы
            [], //  Погонные материалы
            []  //  Фурнитура
        ];

        //  Сортировка результата по группам
        result_db.forEach(elem => {
            //console.log(elem.id_m, elem.name_mat);

            if (BOARD_ARRAY.includes(elem.name_mat)) {
                result[0].push(elem);
                //console.log('-- панель');
            }
            if (BUTT_ARRAY.includes(elem.name_mat)) {
                result[1].push(elem);
                //console.log('-- кромка');
            }
            if (PROFILE_ARRAY.includes(elem.name_mat)) {
                result[2].push(elem);
                //console.log('-- профиль');
            }
            if (FURNITURE_ARRAY.includes(elem.name_mat)) {
                result[3].push(elem);
                //console.log('-- фурнитура');
            }
            ID_MAT_ARRAY.push(elem.id_m);
        });
        return result;
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

//  Функция дополнения данных из базы материалов
async function unionMaterialData(prj_arr, mat_arr) {
    const pnl_mat = new Map(mat_arr[0].map(item => [item.name_mat, item]));
    const butt_mat = new Map(mat_arr[1].map(item => [item.name_mat, item]));
    const prfl_mat = new Map(mat_arr[2].map(item => [item.name_mat, item]));
    const furn_mat = new Map(mat_arr[3].map(item => [item.name_mat, item]));

    //  Цикл по моделям проекта
    prj_arr.forEach(model => {

        //  Листовые материалы
        let panels = model.data.panelMaterials;
        panels.forEach(pnl => {
            const pnl_m = pnl_mat.get(pnl.materialName);
            if (pnl_m) {
                pnl.materialSyncExternal = pnl_m.sync_external;
                pnl.materialID = pnl_m.id_m;
                pnl.materialUnit = pnl_m.name_meas;
                pnl.materialPrice = pnl_m.price;
                pnl.materialWidth = pnl_m.length;
                pnl.materialHeight = pnl_m.width;
            } else {
                pnl.materialID = pnl.materialName;
            };

            //  Кромочные материалы
            pnl.buttInfo.forEach(butt => {
                const butt_m = butt_mat.get(butt.materialName);
                if (butt_m) {
                    butt.materialSyncExternal = butt_m.sync_external;
                    butt.materialID = butt_m.id_m;
                    butt.materialUnit = butt_m.name_meas;
                    butt.materialPrice = butt_m.price;
                } else {
                    butt.materialID = butt.materialName;
                };
            });
        });

        //  Погонные материалы
        let profiles = model.data.profileMaterials;
        profiles.forEach(prfl => {
            const prfl_m = prfl_mat.get(prfl.materialName);
            if (prfl_m) {
                prfl.materialSyncExternal = prfl_m.sync_external;
                prfl.materialID = prfl_m.id_m;
                prfl.materialUnit = prfl_m.name_meas;
                prfl.materialPrice = prfl_m.price;
                prfl.materialWidth = prfl_m.length;
                prfl.materialHeight = prfl_m.width;
            } else {
                prfl.materialID = prfl.materialName;
            };
        });

        //  Фурнитура
        let furnitures = model.data.furnitureMaterials;
        furnitures.forEach(furn => {
            const furn_m = furn_mat.get(furn.materialName);
            if (furn_m) {
                furn.materialSyncExternal = furn_m.sync_external;
                furn.materialID = furn_m.id_m;
                furn.materialUnit = furn_m.name_meas;
                furn.materialPrice = furn_m.price;
            } else {
                furn.materialID = furn.materialName;
            };
        });
    });
};

//  Функция получения данных о классах материалов
async function getDBCLassMaterialInfo(ids) {

    try {
        const options = settings.db_options || null;
        const classes = Object.keys(settings.estimate.classes);

        if (!classes) errFinish("Не заполнены классы материалов");

        //  Строка запроса
        const sqlString = `
            SELECT 
                LMC.ID_M,
                GCM.LABEL_CLASS || CM.CUR_NUM_REC AS CLASS_CODE,
                CM.NAME_CLASS,
                GCM.NAME_GROUP
            FROM 
                LINK_MATERIAL_CLASS LMC
                INNER JOIN CLASS_MATERIAL CM ON CM.ID_CM = LMC.ID_CM
                INNER JOIN GROUP_CLASS_MAT GCM ON GCM.ID_GRCM = CM.ID_GRCM
            WHERE 
                LMC.ID_M IN (${ids.map(() => '?').join(',')})
            ORDER BY 
                LMC.ID_M, GCM.LABEL_CLASS, CM.CUR_NUM_REC
        `;
        //  Запрос в БД
        const result_db = await executeQuery(sqlString, options, ids);
        const result = {};

        //  Обработка результата
        result_db.forEach(elem => {
            //  Игнорируем записи классов не входящих в список
            if (classes.indexOf(elem.class_code) < 0) return;

            //  Фильтруем только уникальные записи ID материалов
            if (!result[elem.id_m]) result[elem.id_m] = elem;
        });
        //console.log(JSON.stringify(result, null, 2));
        return result;
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

//  Функция компоновки данных для сметы
function setProjectEstimateData(db_data, prj) {

    //  Цикл по объектам Проекта
    prj.forEach(model => {
        //  Создаем объект сметы модели с ключами по классам материалов
        const estimate = {};
        for (const key of classes) {
            estimate[key] = {
                class_code: key,
                name_class: "",
                name_group: "",
                items: []
            };
        };

        estimate[defClass].name_class = "Материалы без класса";
        estimate[defClass].name_group = "Материалы без группы класса";

        //  Функция сортировки материалов по ключам сметы
        function addClassItems(obj) {
            for (const key in obj) {
                const item = obj[key];
                const class_code = db_data[key] ?
                    db_data[key].class_code :
                    defClass;

                //  Если не находим такого класса материалов в списке сметы
                if (!estimate[class_code]) {
                    estimate[defClass].items.push(item);
                    continue;
                };

                //  Если данные класса еще не заполнены
                if (!estimate[class_code].name_class) {
                    estimate[class_code].name_class = db_data[key].name_class;
                    estimate[class_code].name_group = db_data[key].name_group;
                };

                //  Добавляем элемент в массив
                estimate[class_code].items.push(item);
            };
        };

        //  Массивы элементов модели для суммирования данных
        const panels = model.data.panelMaterials;
        const butts = panels.flatMap(pnl => pnl.buttInfo || []);
        const profiles = model.data.profileMaterials;
        const furnitures = model.data.furnitureMaterials;

        //  Суммируем данные по листовым материалам
        const board_acc = panels.reduce((acc, item) => {
            //  Ключ материала (ID)
            const key = item.materialID;
            if (!acc[key]) {
                acc[key] = {
                    material: item.material,
                    materialID: key,
                    materialName: item.materialName,
                    materialArticle: item.materialArticle,
                    materialSyncExternal: item.materialSyncExternal,
                    materialPrice: item.materialPrice,
                    materialUnit: item.materialUnit,
                    materialWidth: item.materialWidth,
                    materialHeight: item.materialHeight,
                    materialTkn: item.materialTkn,
                    area: 0,
                    contourLength: 0,
                    drillInfo: [],
                    buttInfo: []
                };
            };

            //  Площадь деталей и длина контуров деталей
            acc[key].area += item.area * item.prjCount || 0;
            acc[key].contourLength += item.contourLength * item.prjCount || 0;

            // Обработка отверстий (присадки)
            if (item.drillInfo && Array.isArray(item.drillInfo)) {
                item.drillInfo.forEach(drill => {
                    // Ищем существующее отверстие с такими же параметрами
                    const existingDrill = acc[key].drillInfo.find(
                        d => d.diameter === drill.diameter &&
                            d.type === drill.type &&
                            d.depth === drill.depth &&
                            d.drillMode === drill.drillMode
                    );
                    if (existingDrill) {
                        // Увеличиваем счетчик
                        existingDrill.count += item.prjCount || 1;
                    } else {
                        // Добавляем новое отверстие
                        acc[key].drillInfo.push({
                            diameter: drill.diameter,
                            depth: drill.depth,
                            type: drill.type,
                            drillMode: drill.drillMode,
                            count: item.prjCount || 1
                        });
                    };
                });
            };

            //  Обработка кромки (данные с учетом свесов!!!)
            if (item.buttInfo && Array.isArray(item.drillInfo)) {
                item.buttInfo.forEach(butt => {
                    // Ищем существующее кромки с такими же параметрами
                    const existingButt = acc[key].buttInfo.find(
                        b => b.thickness === butt.thickness &&
                            b.width === butt.width
                    );
                    if (existingButt) {
                        // Увеличиваем длину
                        existingButt.length += butt.length;
                    } else {
                        // Добавляем новую кромку
                        acc[key].buttInfo.push({
                            thickness: butt.thickness,
                            width: butt.width,
                            length: butt.length
                        });
                    };
                });
            };
            return acc;
        }, {});

        //  Суммируем данные по кромочным материалам
        const butt_acc = butts.reduce((acc, item) => {
            //  Ключ материала (ID)
            const key = item.materialID;
            if (!acc[key]) {
                acc[key] = {
                    material: item.material,
                    materialID: key,
                    materialName: item.materialName,
                    materialArticle: item.materialArticle,
                    materialSyncExternal: item.materialSyncExternal,
                    materialPrice: item.materialPrice,
                    materialUnit: item.materialUnit,
                    materialSign: item.sign,
                    materialWidth: item.width,
                    materialTkn: item.thickness,
                    length: 0
                };
            };
            //  Длина кромки
            acc[key].length += item.length || 0;
            return acc;
        }, {});

        //  Суммируем данные по погонным материалам
        const prfl_acc = profiles.reduce((acc, item) => {
            //  Ключ материала (ID)
            const key = item.materialID;
            if (!acc[key]) {
                acc[key] = {
                    material: item.material,
                    materialID: key,
                    materialName: item.materialName,
                    materialArticle: item.materialArticle,
                    materialSyncExternal: item.materialSyncExternal,
                    materialPrice: item.materialPrice,
                    materialUnit: item.materialUnit,
                    materialSign: item.sign,
                    length: 0
                };
            };
            //  Длина кромки
            acc[key].length += item.length || 0;
            return acc;
        }, {});

        //  Суммируем данные по листовым материалам
        const furn_acc = furnitures.reduce((acc, item) => {
            //  Ключ материала (ID)
            const key = item.materialID;
            if (!acc[key]) {
                acc[key] = {
                    material: item.material,
                    materialID: key,
                    materialName: item.materialName,
                    materialArticle: item.materialArticle,
                    materialSyncExternal: item.materialSyncExternal,
                    materialPrice: item.materialPrice,
                    materialUnit: item.materialUnit,
                    count: 0
                };
            };
            //  Площадь деталей и длина контуров деталей
            acc[key].count += item.count || 0;
            return acc;
        }, {});

        addClassItems(board_acc);
        addClassItems(butt_acc);
        addClassItems(prfl_acc);
        addClassItems(furn_acc);

        //  Присваиваем данные
        model.estimate_data = estimate;
        //console.log(model.name);
    });
};

// Функция для агрегации материалов из массива prj_arr
function aggregateMaterials(prj_arr) {
    // Результирующий объект с группировкой по классам
    const result = {};

    // Объект для хранения Map по каждому классу
    // Ключ: className, Значение: Map для быстрого поиска материалов в этом классе
    const classMaps = {};

    // Проходим по всем элементам проекта
    for (const projectItem of prj_arr) {
        // Получаем объект estimate_data
        const estimateData = projectItem.estimate_data;

        // Проверяем, существует ли estimate_data и является ли он объектом
        if (!estimateData || typeof estimateData !== 'object') continue;

        // Получаем все ключи (классы материалов) динамически
        const classKeys = Object.keys(estimateData);

        // Проходим по каждому классу материалов
        for (const className of classKeys) {
            const classData = estimateData[className];

            // Проверяем наличие массива items
            if (!classData || !Array.isArray(classData.items)) continue;

            // Инициализируем структуру для класса, если её ещё нет
            if (!result[className]) {
                result[className] = { items: [] };
                classMaps[className] = new Map();
            }

            // Проходим по всем материалам в текущем классе
            for (const material of classData.items) {
                // Проверяем наличие обязательных полей
                if (!material.material) continue;

                // Определяем, какое количество использовать
                let quantity = 0;
                let quantityType = null; // 'area', 'length' или 'count'

                if (material.area !== undefined) {
                    quantity = material.area;
                    quantityType = 'area';
                } else if (material.length !== undefined) {
                    quantity = material.length;
                    quantityType = 'length';
                } else if (material.count !== undefined) {
                    quantity = material.count;
                    quantityType = 'count';
                } else {
                    continue;
                }

                // Создаем уникальный ключ для поиска в мапе конкретного класса
                // Для area используем material, для остальных - material
                const uniqueKey = material.material;

                // Проверяем, существует ли уже такой материал в этом классе
                if (classMaps[className].has(uniqueKey)) {
                    // Если существует - добавляем количество
                    const existingItem = classMaps[className].get(uniqueKey);
                    existingItem.quantity += quantity;

                    // Обновляем дублирующее поле (area/length/count)
                    if (quantityType === 'area') {
                        existingItem.area = (existingItem.area || 0) + quantity;
                    } else if (quantityType === 'length') {
                        existingItem.length = (existingItem.length || 0) + quantity;
                    } else if (quantityType === 'count') {
                        existingItem.count = (existingItem.count || 0) + quantity;
                    }
                } else {
                    // Вычисляем коэффициент k (если есть area)
                    const k = material.area !== undefined && material.materialWidth && material.materialHeight
                        ? 1 / (material.materialHeight * material.materialWidth * 0.000001)
                        : 1;

                    // Создаем новый объект материала
                    const newItem = {
                        material: material.material,
                        materialName: material.materialName || '',
                        materialArticle: material.materialArticle || '',
                        materialSyncExternal: material.materialSyncExternal || null,
                        materialUnit: material.materialUnit || '',
                        materialPrice: material.materialPrice || 0,
                        quantityType: quantityType,
                        quantity: quantity,
                        k: k
                    };

                    // Добавляем специфичные поля в зависимости от типа количества
                    if (quantityType === 'area') {
                        newItem.materialWidth = material.materialWidth || 0;
                        newItem.materialHeight = material.materialHeight || 0;
                        newItem.area = quantity;
                    } else if (quantityType === 'length') {
                        newItem.length = quantity;
                    } else if (quantityType === 'count') {
                        newItem.count = quantity;
                    }

                    // Добавляем в массив items для этого класса
                    result[className].items.push(newItem);
                    // Сохраняем в мапу для быстрого доступа
                    classMaps[className].set(uniqueKey, newItem);
                }
            }
        }
    };

    return result;
};

//  Функция для получения данных сборочных чертежей
function getAssemblyName(list, prj_item) {
    let result = [];
    for (let i = 0; i < list.Count; i++) {
        const item = list[i];
        if (
            (item instanceof TFurnBlock || item instanceof TDraftBlock) &&
            !item.JointData &&
            item.IsAssemblyUnit === true
        ) {
            const sbName = transliterate(item.Name);
            const sign = transliterate(prj_item.sign);
            const prjName = transliterate(PROJECT_NAME);

            const delimPrjName = settings.delimPrjName;
            const delimPrjSign = settings.delimPrjSign;
            const modelDrawName = `${prjName}${delimPrjName}${sign}_SB`;
            const drawName = `${modelDrawName}_na_${sbName}`;

            const UID = item.UID;
            let panelMaterialsAU = prj_item.data.panelMaterials.filter(item => {
                return item.assemblyUnit && item.assemblyUnit.UID === UID;
            });

            result.push({
                prjName: PROJECT_NAME,
                auArtPos: item.ArtPos,          //  Позиция СБЕ
                auDesignation: item.Designation,//  Обозначение СБЕ
                cloudLink: prj_item.cloudLink,  //  Ссылка на облако
                sign: prj_item.sign,            //  Обозначение в проекте изд.
                drawName: drawName,             //  Название СБ (tr)
                auName: item.Name,              //  Название СБ
                modelName: prj_item.modelName,  //  Название модели
                modelDrawName: modelDrawName,   //  Название схемы сборки (tr)
                items: {
                    panelMaterials: panelMaterialsAU,
                    profileMaterials: []
                },
            });
        };
    };
    return result;
};

//#endregion

//#region Формирование документов

//  Функуция создания файла Калькуляции проекта
async function createEsimateExcelFile(prj_arr) {
    const totalData = [];

    //#region Настройки стилей
    if (!settings.estimate.classes)
        errFinish("Ошибка файла настроек - classes");
    if (!settings.estimate.fillColor)
        errFinish("Ошибка файла настроек - fillColor");

    const fontFamily = settings.estimate.fontFamily ?
        settings.estimate.fontFamily : "Arial";

    const row_height = 14;
    const font_size = 9;
    const doc_font = { name: fontFamily, size: font_size + 7, bold: true };
    const tabl_font = { name: fontFamily, size: font_size + 2, bold: true };
    const h_font = { name: fontFamily, size: font_size, bold: true };
    const r_font = { name: fontFamily, size: font_size - 1, bold: false };

    const c_koef = settings.estimate.classes;
    const fill = settings.estimate.fillColor;

    //  Заливка ячеек
    const fill_green = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: fill.green }
    };
    const fill_yellow = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: fill.yellow }
    };
    const fill_orange = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: fill.orange }
    };
    const fill_red = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: fill.red }
    };
    const fill_blue = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: fill.blue }
    };

    //  Форматирование чисел
    const f_format = '#,##0.00';
    const sf_format = '#,##0.0';
    const n_format = '# ##0';

    const algn_left = { indent: 1, horizontal: 'left', vertical: 'middle' };
    const algn_right = { indent: 1, horizontal: 'right', vertical: 'middle' };
    const algn_center = { horizontal: 'center', vertical: 'middle' };

    const pageSetup = {
        orientation: 'portrait',    // 'portrait' | 'landscape'
        margins: {                  // Поля страницы в дюймах (1д = 2.54см)
            top: 0.5,               // Верхнее поле
            bottom: 0.5,            // Нижнее поле
            left: 0.39,             // Левое поле
            right: 0.39,            // Правое поле
            header: 0.3,            // Отступ для колонтитула сверху
            footer: 0.3             // Отступ для колонтитула снизу
        },
        // Масштабирование
        fitToPage: true,            // Вписать в страницу
        fitToWidth: 1,              // Вписать по ширине (1 страница)
        fitToHeight: 0,             // По высоте (0 = автоматически)
    };

    //#endregion

    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.author;
    workbook.created = new Date();

    //  Создаем новую вкладку документа (Сводная таблица)
    const total_worksheet = workbook.addWorksheet('Сводная таблица');

    //  Сортировка массива по обозначению изделия в проекте
    prj_arr = smartSort(prj_arr, [["sign", "asc"]]); //desc
    // Цикл создания вкладок
    prj_arr.forEach(model => {
        const estimate = model.estimate_data;   //  Объект данных Сметы
        const name = model.modelName;           //  Имя изделия
        const sign = model.sign;                //  Обозначение в Проекте
        const sheet_name = `${sign}_${name}`.substring(0, 31);
        const worksheet = workbook.addWorksheet(sheet_name);
        const prj_count = model.count;

        worksheet.pageSetup = pageSetup;

        // Колонки
        worksheet.columns = [
            { width: 0.67 },    //  Отступ
            { width: 5.42 },    //  Номер 
            { width: 18 },      //  Артикул
            { width: 60 },      //  Наименование 78
            { width: 8 },       //  Количество
            { width: 8 },       //  Ед.
            { width: 12 },      //  Цена
            { width: 12 }       //  Сумма
        ];

        //  Заголовки основной таблицы
        const headers = ['№', 'Артикул', 'Наименование',
            'Кол-во', 'Ед.', 'Цена', 'Сумма'];
        //  Заголовки вспомогательной таблицы
        const v_headers = ['k', 'Цена (DB)', 'Сумма (DB)', 'ВД'];
        const offset = 3;           //  Смещение вспомогательной таблицы
        const headerRowInd = 6;     //  Индекс верхней строки таблицы
        const scol = 2;             //  Начальная колонка
        const scol_second = scol + headers.length + offset;
        const headerRow = worksheet.getRow(headerRowInd);

        //#region Шапка документа
        worksheet.getRow(2).height = setRowHeght(24);
        worksheet.getRow(3).height = setRowHeght(5);
        styleCellRange(worksheet.getRow(3), scol, worksheet.columns.length, {
            border: {
                left: { style: 'none' }, right: { style: 'none' },
                top: { style: 'medium' }, bottom: { style: 'none' }
            }
        });

        //  Строка суммы (дубликат)
        const top_total_row = worksheet.getRow(headerRowInd - 2);
        top_total_row.height = setRowHeght(row_height);
        worksheet.getRow(headerRowInd - 1).height = setRowHeght(5);

        //  Название документа
        const docNameRow = worksheet.getRow(2);
        docNameRow.alignment = { horizontal: 'left', vertical: 'middle' };
        docNameRow.getCell(scol).font = { ...doc_font };
        docNameRow.getCell(scol).value =
            `Расчет изделия — ${model.modelName}`;
        //#endregion   

        //#region Таблица материалов и фурнитуры

        //  Название таблицы материалов и фурнитуры
        const mTableHeaderCell = top_total_row.getCell(scol);
        mTableHeaderCell.value = "Спецификация материалов и фурнитуры";
        mTableHeaderCell.font = { ...tabl_font };
        mTableHeaderCell.alignment = { horizontal: 'left', vertical: 'middle' };

        //  Шапка основной таблицы
        for (let i = 0; i < headers.length; i++) {
            headerRow.getCell(scol + i).value = headers[i];
        };
        setRowTableStyle(
            headerRow,
            scol,
            headers.length,
            row_height,
            'header',
            h_font
        );

        //  Шапка вспомогательной таблицы
        for (let i = 0; i < v_headers.length; i++) {
            headerRow.getCell(scol_second + i).value = v_headers[i];
        };
        setRowTableStyle(
            headerRow,
            scol_second,
            v_headers.length,
            row_height,
            'header',
            h_font
        );

        //  Индекс строки таблицы сразу после заголовка
        let ind = headerRowInd + 1;

        //  Литеры колонок вспомогательной таблицы
        const v_col = scol_second;
        const coef_ltr = getColumnLetter(v_col + 0);
        const price_ltr = getColumnLetter(v_col + 1);
        const sum_ltr = getColumnLetter(v_col + 2);
        const res_ltr = getColumnLetter(v_col + 3);

        //  Устанавливаем ширины колонок вспомогательной таблицы
        worksheet.getColumn(`${coef_ltr}`).width = 10;
        worksheet.getColumn(`${price_ltr}`).width = 12;
        worksheet.getColumn(`${sum_ltr}`).width = 12;
        worksheet.getColumn(`${res_ltr}`).width = 12;

        //  Тело таблицы
        let row_counter = 1;
        classes.forEach(key => {
            if (!estimate[key].items.length) return;
            //  Сортировка материалов в пределах класса
            let sort_options = [["materialName", "desc"]];
            if (estimate[key].items[0].area) {
                // Площадной материалв
                sort_options = [
                    ["materialTkn", "desc"], ["materialName", "asc"]
                ];
            } else if (estimate[key].items[0].length) {
                //  Кромка или погонный материал
                sort_options = [
                    ["materialName", "desc"]
                ];
            } else if (estimate[key].items[0].count) {
                //  Фурнитура
                sort_options = [
                    ["materialName", "desc"]
                ];
            };
            //  Сортировка
            const items = smartSort(estimate[key].items, sort_options);

            //  Коэффициент наценки
            const k = c_koef[key] ? c_koef[key] : 0;

            //  Цикл по массиву класса key
            for (let i = 0; i < items.length; i++) {
                //  Объект материала
                const item = items[i];

                //  Текущая строка
                const rn = ind + i; //  Индекс текущей строки
                const row = worksheet.getRow(rn);

                //  Стиль строки основной таблицы
                setRowTableStyle(
                    row,
                    scol,
                    headers.length,
                    row_height - 1,
                    'main',
                    r_font
                );
                //  Стиль строки расчетной таблицы
                setRowTableStyle(
                    row,
                    scol + headers.length + offset,
                    v_headers.length,
                    row_height - 1,
                    'main',
                    r_font
                );

                //  Выбор измеряемой величины
                let val = "area";
                if (item.length) val = "length";
                if (item.count) val = "count";

                //  Литеры колонок цены и количества
                const cpl = getColumnLetter(scol + 3);
                const cvl = getColumnLetter(scol + 5);

                //  Ячейка номера строки
                const counterCell = row.getCell(scol + 0);
                counterCell.value = row_counter++;
                counterCell.alignment = algn_right;

                //  Ячейка артикула материала
                const articleCell = row.getCell(scol + 1);
                articleCell.value = item.materialArticle;
                articleCell.alignment = algn_left;

                //  Ячейка названия материала
                const nameCell = row.getCell(scol + 2);
                nameCell.value = item.materialName;
                nameCell.alignment = algn_left;

                //  Ячейка количества
                const countCell = row.getCell(scol + 3);
                countCell.alignment = algn_right;
                countCell.value = item[val];
                countCell.numFmt = val == "count" ? n_format : f_format;

                //  Ячейка ед. изм.
                const unitCell = row.getCell(scol + 4);
                unitCell.value = item.materialUnit;
                unitCell.alignment = algn_center;

                //  Ячейка цены с наценкой
                const priceCoeffCell = row.getCell(scol + 5);
                priceCoeffCell.alignment = algn_right;
                priceCoeffCell.value = {
                    formula: `${price_ltr}${rn}*$${coef_ltr}$${ind}`
                };
                priceCoeffCell.numFmt = f_format;

                //  Ячейка суммы с наценкой
                const coefSumCell = row.getCell(scol + 6);
                coefSumCell.value = {
                    formula: `${cvl}${rn}*${cpl}${rn}`
                };
                coefSumCell.alignment = algn_right;
                coefSumCell.numFmt = f_format;

                //  Индикация материалов без цены
                if (!item.materialUnit) {
                    unitCell.fill = fill_red;
                    priceCoeffCell.fill = fill_red;
                    coefSumCell.fill = fill_red;
                };
                //------------------------------------------------------------//

                //  Коэффициент наценки класса
                const coefCell = worksheet.getCell(`${coef_ltr}${ind}`);
                coefCell.alignment = algn_center;
                coefCell.value = k;
                coefCell.numFmt = f_format;
                coefCell.fill = fill_yellow;

                //  Цена материала из базы
                const dbPriceCell = worksheet.getCell(`${price_ltr}${rn}`);
                dbPriceCell.alignment = algn_right;
                dbPriceCell.value = item.materialPrice;
                dbPriceCell.numFmt = f_format;

                //  Сумма материала из базы
                const dbSumCell = worksheet.getCell(`${sum_ltr}${rn}`)
                dbSumCell.alignment = algn_right;
                dbSumCell.value = { formula: `${price_ltr}${rn}*${cpl}${rn}` };
                dbSumCell.numFmt = f_format;

                //  Разница между суммой с наценкой и без
                const valCell = worksheet.getCell(`${res_ltr}${rn}`);
                valCell.alignment = algn_right;
                valCell.value = {
                    formula: `${cpl}${rn}*${cvl}${rn}-${sum_ltr}${rn}`
                };
                valCell.numFmt = f_format
            };

            //  Объединяем все ячейки в этой колонке для диапазона строк
            let last_ind = ind + items.length - 1;
            worksheet.mergeCells(`${coef_ltr}${ind}:${coef_ltr}${last_ind}`);
            ind += items.length;
        });

        const row = worksheet.getRow(ind);
        const end_res_col = scol + headers.length;
        const endMainTableCol = scol + headers.length - 1;
        //  Стиль строки основной таблицы
        setRowTableStyle(
            row,
            scol,
            headers.length,
            row_height,
            'empty',
            h_font
        );
        //  Стиль строки расчетной таблицы
        setRowTableStyle(
            row,
            scol_second,
            v_headers.length,
            row_height,
            'empty',
            h_font
        );

        //  Литера колонки Суммы основной таблицы
        const tsm_ltr = getColumnLetter(endMainTableCol);
        const tres_sum_ltr = getColumnLetter(endMainTableCol + offset + 3);
        const tres_vd_ltr = getColumnLetter(endMainTableCol + offset + 4);

        //  --- Основная таблица
        //  текст ИТОГО основной таблицы
        const total_row_text = row.getCell(endMainTableCol - 1);
        total_row_text.value = "Итого:";
        total_row_text.font.size = font_size;
        total_row_text.alignment = algn_right;
        total_row_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };
        //  Дубликат в шапке таблицы текста ИТОГО
        const d_total_text = top_total_row.getCell(endMainTableCol - 1);
        d_total_text.value = total_row_text.value;
        d_total_text.font = { ...tabl_font };
        d_total_text.alignment = algn_right;

        //  Дубликат суммы в шапке таблицы
        const d_total_val = top_total_row.getCell(endMainTableCol);
        d_total_val.value = {
            formula: `SUM(${tsm_ltr}${headerRowInd + 1}:${tsm_ltr}${ind - 1})`
        };
        d_total_val.font = { ...tabl_font };
        d_total_val.numFmt = n_format;
        d_total_val.alignment = algn_right;

        //  Ячейка итоговой суммы с наценкой
        const total_row_val = row.getCell(endMainTableCol);
        total_row_val.value = {
            formula: `SUM(${tsm_ltr}${headerRowInd + 1}:${tsm_ltr}${ind - 1})`
        };
        total_row_val.numFmt = n_format;
        total_row_val.alignment = algn_right;

        //  --- Вспомогательная таблица
        //  Текст ИТОГО вспомогательной таблицы
        const total_row_res_text = row.getCell(endMainTableCol + offset + 2);
        total_row_res_text.value = "Итого:";
        total_row_res_text.font.size = font_size;
        total_row_res_text.alignment = algn_right;
        total_row_res_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };
        //  Дубликат в шапке вспомогательной таблицы текста ИТОГО
        const d_totalres_text = top_total_row.getCell(endMainTableCol + offset + 2);
        d_totalres_text.value = total_row_res_text.value;
        d_totalres_text.font = { ...tabl_font };
        d_totalres_text.alignment = algn_right;

        //  Ячейка итоговой суммы из Базы
        const total_row_res_val = row.getCell(endMainTableCol + offset + 3);
        total_row_res_val.value = {
            formula: `SUM(${tres_sum_ltr}${headerRowInd + 1}:${tres_sum_ltr}${ind - 1})`
        };
        total_row_res_val.numFmt = n_format;
        total_row_res_val.alignment = algn_right;
        total_row_res_val.border = {
            left: { style: 'thin' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        //  Дубликат суммы из базы в шапке таблицы
        const d_totalres_val = top_total_row.getCell(endMainTableCol + offset + 3);
        d_totalres_val.value = {
            formula: `SUM(${tres_sum_ltr}${headerRowInd + 1}:${tres_sum_ltr}${ind - 1})`
        };
        d_totalres_val.font = { ...tabl_font };
        d_totalres_val.numFmt = n_format;
        d_totalres_val.alignment = algn_right;

        //  Ячейка итоговой суммы ВД
        const total_row_res_vd_val = row.getCell(endMainTableCol + offset + 4);
        total_row_res_vd_val.value = {
            formula: `SUM(${tres_vd_ltr}${headerRowInd + 1}:${tres_vd_ltr}${ind - 1})`
        };
        total_row_res_vd_val.numFmt = n_format;
        total_row_res_vd_val.alignment = algn_right;

        //  Дубликат итоговой суммы ВД
        const d_totalvd_val = top_total_row.getCell(endMainTableCol + offset + 4);
        d_totalvd_val.value = {
            formula: `SUM(${tres_vd_ltr}${headerRowInd + 1}:${tres_vd_ltr}${ind - 1})`
        };
        d_totalvd_val.font = { ...tabl_font };
        d_totalvd_val.numFmt = n_format;
        d_totalvd_val.alignment = algn_right;

        //#endregion

        const materialTotalCol = getColumnLetter(endMainTableCol);
        const materialTotalRow = ind;

        //#region Таблица Операций

        const operStartInd = worksheet.rowCount + 2;
        const topTotalOperRow = worksheet.getRow(operStartInd);

        //  Название таблицы материалов и фурнитуры
        const oTableNameCell = topTotalOperRow.getCell(scol);
        oTableNameCell.value = "Спецификация операций";
        oTableNameCell.font = { ...tabl_font };
        oTableNameCell.alignment = { horizontal: 'left', vertical: 'middle' };

        worksheet.getRow(operStartInd + 1).height = setRowHeght(5);
        const oTableHeaderRow = worksheet.getRow(operStartInd + 2);

        //  Шапка основной таблицы операций
        for (let i = 0; i < headers.length; i++) {
            oTableHeaderRow.getCell(scol + i).value = headers[i];
        };
        setRowTableStyle(
            oTableHeaderRow,    //  Строка
            scol,        //  Начало диапазона ячеек строки
            headers.length,     //  Конец диапазона ячеек строки
            row_height,         //  Высота строки
            'header',           //  Тип строки
            h_font              //  Стиль шрифта
        );

        //  Шапка вспомогательной таблицы операций
        for (let i = 0; i < v_headers.length; i++) {
            oTableHeaderRow.getCell(scol_second + i).value = v_headers[i];
        };
        setRowTableStyle(
            oTableHeaderRow,    //  Строка
            scol_second,        //  Начало диапазона ячеек строки
            v_headers.length,   //  Конец диапазона ячеек строки
            row_height,         //  Высота строки
            'header',           //  Тип строки
            h_font              //  Стиль шрифта
        );

        //  Индекс строки таблицы сразу после заголовка
        ind = operStartInd + 1;

        //  Объект операций из файла настроек
        const operations = settings.estimate.operations;
        ind += 1;

        classes.forEach(key => {
            if (!estimate[key].items.length) return;
            //  Исключенные классы материалов для раскроя
            //  Не считается раскрой и не считается кромка
            const exClasses = settings.estimate.excludeCuttingClasses;
            const bool = !exClasses.includes(key);
            const items = estimate[key].items;

            //  Игонорируем материалы, у которые есть свойство count
            if (
                !items.length ||
                items[0].count ||       //  Игнорируем фурнитуру
                items[0].length ||      //  Игнорируем кромку
                !items[0].materialUnit  //  Игонорируем материалы не из базы
            ) return;

            //  Игнорируем мастериалы у которых нет данных о операциях
            ind++;

            //  Цикл по массиву класса key
            for (let i = 0; i < items.length; i++) {
                //  Объект материала
                const item = items[i];
                if (!item.buttInfo && !item.drillInfo) {
                    ind--;
                    continue;
                }

                //  Текущая строка
                ind = ind + i;
                const rn = ind; //  Индекс текущей строкиind + i
                const row = worksheet.getRow(rn);
                //  Стиль строки основной таблицы
                setRowTableStyle(
                    row,
                    scol,
                    headers.length,
                    row_height - 1,
                    'main',
                    r_font
                );
                //  Стиль строки расчетной таблицы
                setRowTableStyle(
                    row,
                    scol_second,
                    v_headers.length,
                    row_height - 1,
                    'main',
                    r_font
                );

                //  Литеры колонок цены и количества
                const cpl = getColumnLetter(scol + 3);
                const cvl = getColumnLetter(scol + 5);

                //  Ячейка номера строки
                const counterCell = row.getCell(scol + 0);
                counterCell.value = "—";
                counterCell.alignment = algn_right;

                //  Ячейка артикула материала
                const articleCell = row.getCell(scol + 1);
                articleCell.value = item.materialArticle;
                articleCell.alignment = algn_left;

                //  Ячейка названия материала
                const nameCell = row.getCell(scol + 2);
                nameCell.value = item.materialName;
                nameCell.alignment = algn_left;

                //  Ячейка количества
                const countCell = row.getCell(scol + 3);
                countCell.alignment = algn_right;
                countCell.value = "—";

                //  Ячейка ед. изм.
                const unitCell = row.getCell(scol + 4);
                unitCell.value = "—";
                unitCell.alignment = algn_center;

                //  Ячейка цены
                const priceCoeffCell = row.getCell(scol + 5);
                priceCoeffCell.alignment = algn_right;
                priceCoeffCell.value = "—";

                //  Ячейка суммы
                const coefSumCell = row.getCell(scol + 6);
                coefSumCell.value = "—";
                coefSumCell.alignment = algn_right;

                //  Коэффициент наценки класса
                const coefCell = worksheet.getCell(`${coef_ltr}${ind}`);
                coefCell.alignment = algn_center;
                coefCell.value = "—";

                //  Цена материала из базы
                const dbPriceCell = worksheet.getCell(`${price_ltr}${ind}`);
                dbPriceCell.alignment = algn_right;
                dbPriceCell.value = "—";

                //  Сумма материала из базы
                const dbSumCell = worksheet.getCell(`${sum_ltr}${ind}`)
                dbSumCell.alignment = algn_right;
                dbSumCell.value = "—";

                //  Разница между суммой с наценкой и без
                const valCell = worksheet.getCell(`${res_ltr}${rn}`);
                valCell.alignment = algn_right;
                valCell.value = "—";

                //  Индикация строки материала
                counterCell.fill = fill_blue;
                articleCell.fill = fill_blue;
                nameCell.fill = fill_blue;
                countCell.fill = fill_blue;
                unitCell.fill = fill_blue;
                priceCoeffCell.fill = fill_blue;
                coefSumCell.fill = fill_blue;
                coefCell.fill = fill_blue;
                dbPriceCell.fill = fill_blue;
                dbSumCell.fill = fill_blue;
                valCell.fill = fill_blue;

                //  Операция раскроя
                if (bool) {

                    const cutting = settings.estimate.operations.cutting;
                    let cuttingName = "Раскрой плиты";
                    let ckf = 1;
                    for (let k = 0; k < cutting.length; k++) {
                        if (
                            item.materialTkn > cutting[k].tkn[0] &&
                            item.materialTkn <= cutting[k].tkn[1]
                        ) {
                            cuttingPrice = cutting[k].price;
                            cuttingName += ` ${item.materialTkn} мм`;
                            ckf = cutting[k].k;
                            break;
                        };
                    };
                    ind++;
                    const cRow = worksheet.getRow(ind);
                    //  Стиль строки основной таблицы
                    setRowTableStyle(
                        cRow,
                        scol,
                        headers.length,
                        row_height - 1,
                        'main',
                        r_font
                    );
                    //  Стиль строки расчетной таблицы
                    setRowTableStyle(
                        cRow,
                        scol_second,
                        v_headers.length,
                        row_height - 1,
                        'main',
                        r_font
                    );

                    //  Ячейка номера строки
                    const cCounterCell = cRow.getCell(scol + 0);
                    cCounterCell.alignment = algn_right;
                    cCounterCell.value = "";

                    //  Ячейка артикула материала
                    const cArticleCell = cRow.getCell(scol + 1);
                    cArticleCell.alignment = algn_left;
                    cArticleCell.value = "";

                    //  Ячейка названия материала
                    const cNameCell = cRow.getCell(scol + 2);
                    cNameCell.alignment = algn_left;
                    cNameCell.value = cuttingName;

                    //  Ячейка количества
                    const cCountCell = cRow.getCell(scol + 3);
                    cCountCell.alignment = algn_right;
                    cCountCell.value = round(item.contourLength, 0);

                    //  Ячейка ед. изм.
                    const cUnitCell = cRow.getCell(scol + 4);
                    cUnitCell.alignment = algn_center;
                    cUnitCell.value = "м";

                    //  Ячейка цены
                    const cPriceCoeffCell = cRow.getCell(scol + 5);
                    cPriceCoeffCell.alignment = algn_right;
                    cPriceCoeffCell.value = {
                        formula: `${price_ltr}${ind}*$${coef_ltr}$${ind}`
                    };
                    cPriceCoeffCell.numFmt = f_format;

                    //  Ячейка суммы
                    const cCoefSumCell = cRow.getCell(scol + 6);
                    cCoefSumCell.alignment = algn_right;
                    cCoefSumCell.value = {
                        formula: `${cvl}${ind}*${cpl}${ind}`
                    };
                    cCoefSumCell.numFmt = f_format;
                    //------------------------------------------------------------//

                    //  Коэффициент наценки класса
                    const coefCell = worksheet.getCell(`${coef_ltr}${ind}`);
                    coefCell.alignment = algn_center;
                    coefCell.value = ckf;
                    coefCell.numFmt = f_format;
                    coefCell.fill = fill_yellow;

                    //  Цена материала из базы
                    const dbPriceCell = worksheet.getCell(`${price_ltr}${ind}`);
                    dbPriceCell.alignment = algn_right;
                    dbPriceCell.value = cuttingPrice;
                    dbPriceCell.numFmt = f_format;

                    //  Сумма материала из базы
                    const dbSumCell = worksheet.getCell(`${sum_ltr}${ind}`)
                    dbSumCell.alignment = algn_right;
                    dbSumCell.value = { formula: `${price_ltr}${ind}*${cpl}${ind}` };
                    dbSumCell.numFmt = f_format;

                    //  Разница между суммой с наценкой и без
                    const valCell = worksheet.getCell(`${res_ltr}${ind}`);
                    valCell.alignment = algn_right;
                    valCell.value = {
                        formula: `${cpl}${ind}*${cvl}${ind}-${sum_ltr}${ind}`
                    };
                    valCell.numFmt = f_format
                };

                //  Операции облицовки кромкой
                if (item.buttInfo && item.buttInfo.length > 0) {
                    const buttArray = smartSort(item.buttInfo, [
                        ["width", "asc"], ["thickness", "asc"]
                    ]);
                    const butts = settings.estimate.operations.butt;
                    buttArray.forEach(butt => {
                        ind++;
                        const bRow = worksheet.getRow(ind);
                        //  Стиль строки основной таблицы
                        setRowTableStyle(
                            bRow,
                            scol,
                            headers.length,
                            row_height - 1,
                            'main',
                            r_font
                        );
                        //  Стиль строки расчетной таблицы
                        setRowTableStyle(
                            bRow,
                            scol_second,
                            v_headers.length,
                            row_height - 1,
                            'main',
                            r_font
                        );

                        //  Поиск цены кромки
                        let kf = 1;
                        let buttPrice = 0;
                        let buttName = "Наклейка кромки";
                        for (let k = 0; k < butts.length; k++) {
                            if (
                                butt.thickness > butts[k].tkn[0] &&
                                butt.thickness <= butts[k].tkn[1] &&
                                butt.width > butts[k].width[0] &&
                                butt.width <= butts[k].width[1]
                            ) {
                                buttPrice = butts[k].price;
                                buttName += ` ${butt.thickness}x${butt.width} мм`;
                                kf = butts[k].k;
                                break;
                            };
                        };

                        //  Ячейка номера строки
                        const bCounterCell = bRow.getCell(scol + 0);
                        bCounterCell.alignment = algn_right;
                        bCounterCell.value = "";

                        //  Ячейка артикула материала
                        const bArticleCell = bRow.getCell(scol + 1);
                        bArticleCell.alignment = algn_left;
                        bArticleCell.value = "";

                        //  Ячейка названия материала
                        const bNameCell = bRow.getCell(scol + 2);
                        bNameCell.alignment = algn_left;
                        bNameCell.value = buttName;

                        //  Ячейка количества
                        const bCountCell = bRow.getCell(scol + 3);
                        bCountCell.alignment = algn_right;
                        bCountCell.value = round(butt.length, 0);

                        //  Ячейка ед. изм.
                        const bUnitCell = bRow.getCell(scol + 4);
                        bUnitCell.alignment = algn_center;
                        bUnitCell.value = "м";

                        //  Ячейка цены
                        const bPriceCoeffCell = bRow.getCell(scol + 5);
                        bPriceCoeffCell.alignment = algn_right;
                        bPriceCoeffCell.value = {
                            formula: `${price_ltr}${ind}*$${coef_ltr}$${ind}`
                        };
                        bPriceCoeffCell.numFmt = f_format;

                        //  Ячейка суммы
                        const dCoefSumCell = bRow.getCell(scol + 6);
                        dCoefSumCell.alignment = algn_right;
                        dCoefSumCell.value = {
                            formula: `${cvl}${ind}*${cpl}${ind}`
                        };
                        dCoefSumCell.numFmt = f_format;
                        //------------------------------------------------------------//

                        //  Коэффициент наценки класса
                        const coefCell = worksheet.getCell(`${coef_ltr}${ind}`);
                        coefCell.alignment = algn_center;
                        coefCell.value = kf;
                        coefCell.numFmt = f_format;
                        coefCell.fill = fill_yellow;

                        //  Цена материала из базы
                        const dbPriceCell = worksheet.getCell(`${price_ltr}${ind}`);
                        dbPriceCell.alignment = algn_right;
                        dbPriceCell.value = buttPrice;
                        dbPriceCell.numFmt = f_format;

                        //  Сумма материала из базы
                        const dbSumCell = worksheet.getCell(`${sum_ltr}${ind}`)
                        dbSumCell.alignment = algn_right;
                        dbSumCell.value = { formula: `${price_ltr}${ind}*${cpl}${ind}` };
                        dbSumCell.numFmt = f_format;

                        //  Разница между суммой с наценкой и без
                        const valCell = worksheet.getCell(`${res_ltr}${ind}`);
                        valCell.alignment = algn_right;
                        valCell.value = {
                            formula: `${cpl}${ind}*${cvl}${ind}-${sum_ltr}${ind}`
                        };
                        valCell.numFmt = f_format
                    });
                };

                //  Операции присадки
                if (item.drillInfo && item.drillInfo.length > 0) {
                    const drillArray = smartSort(item.drillInfo, [
                        ["type", "asc"], ["diameter", "desc"]
                    ]);
                    const drill = settings.estimate.operations.drill;
                    drillArray.forEach(hole => {
                        ind++;
                        const dRow = worksheet.getRow(ind);
                        //  Стиль строки основной таблицы
                        setRowTableStyle(
                            dRow,
                            scol,
                            headers.length,
                            row_height - 1,
                            'main',
                            r_font
                        );
                        //  Стиль строки расчетной таблицы
                        setRowTableStyle(
                            dRow,
                            scol_second,
                            v_headers.length,
                            row_height - 1,
                            'main',
                            r_font
                        );

                        //  Поиск цены отверстия
                        let d = hole.diameter;
                        let drillPrice = 0;
                        let kf = 1;
                        for (let k = 0; k < drill.length; k++) {
                            if (
                                drill[k].type == hole.type &&
                                drill[k].drillMode == hole.drillMode &&
                                drill[k].diameters[d]
                            ) {
                                drillPrice = drill[k].diameters[d];
                                kf = drill[k].k;
                                break;
                            };
                        };

                        let drillName = "";
                        if (hole.type == 1 && hole.drillMode == 1) {
                            //  Отверстие сквозное в пласть
                            drillName =
                                `Присадка сквозного отверстия в пласть D${d} мм`;
                        } else if (hole.type == 1 && hole.drillMode == 2) {
                            //  Отверстие глухое в пласть
                            drillName =
                                `Присадка глухого отверстия в пласть D${d}x${hole.depth} мм`;
                            //  Отверстие в торец
                        } else {
                            drillName = `Присадка торцевого отверстия D${d}x${hole.depth} мм`;
                        };

                        //  Ячейка номера строки
                        const dCounterCell = dRow.getCell(scol + 0);
                        dCounterCell.alignment = algn_right;
                        dCounterCell.value = "";

                        //  Ячейка артикула материала
                        const dArticleCell = dRow.getCell(scol + 1);
                        dArticleCell.alignment = algn_left;
                        dArticleCell.value = "";

                        //  Ячейка названия материала
                        const dNameCell = dRow.getCell(scol + 2);
                        dNameCell.alignment = algn_left;
                        dNameCell.value = drillName;

                        //  Ячейка количества
                        const dCountCell = dRow.getCell(scol + 3);
                        dCountCell.alignment = algn_right;
                        dCountCell.value = hole.count;

                        //  Ячейка ед. изм.
                        const dUnitCell = dRow.getCell(scol + 4);
                        dUnitCell.alignment = algn_center;
                        dUnitCell.value = "шт.";

                        //  Ячейка цены
                        const dPriceCoeffCell = dRow.getCell(scol + 5);
                        dPriceCoeffCell.alignment = algn_right;
                        dPriceCoeffCell.value = {
                            formula: `${price_ltr}${ind}*$${coef_ltr}$${ind}`
                        };
                        dPriceCoeffCell.numFmt = f_format;

                        //  Ячейка суммы
                        const dCoefSumCell = dRow.getCell(scol + 6);
                        dCoefSumCell.alignment = algn_right;
                        dCoefSumCell.value = {
                            formula: `${cvl}${ind}*${cpl}${ind}`
                        };;
                        dCoefSumCell.numFmt = f_format;

                        //------------------------------------------------------------//

                        //  Коэффициент наценки класса
                        const coefCell = worksheet.getCell(`${coef_ltr}${ind}`);
                        coefCell.alignment = algn_center;
                        coefCell.value = kf;
                        coefCell.numFmt = f_format;
                        coefCell.fill = fill_yellow;

                        //  Цена материала из базы
                        const dbPriceCell = worksheet.getCell(`${price_ltr}${ind}`);
                        dbPriceCell.alignment = algn_right;
                        dbPriceCell.value = drillPrice;
                        dbPriceCell.numFmt = f_format;

                        //  Сумма материала из базы
                        const dbSumCell = worksheet.getCell(`${sum_ltr}${ind}`)
                        dbSumCell.alignment = algn_right;
                        dbSumCell.value = { formula: `${price_ltr}${ind}*${cpl}${ind}` };
                        dbSumCell.numFmt = f_format;

                        //  Разница между суммой с наценкой и без
                        const valCell = worksheet.getCell(`${res_ltr}${ind}`);
                        valCell.alignment = algn_right;
                        valCell.value = {
                            formula: `${cpl}${ind}*${cvl}${ind}-${sum_ltr}${ind}`
                        };
                        valCell.numFmt = f_format
                    });
                };
            };
        });
        ind++;
        const oRow = worksheet.getRow(ind);
        //  Стиль строки основной таблицы
        setRowTableStyle(
            oRow,
            scol,
            headers.length,
            row_height,
            'empty',
            h_font
        );
        //  Стиль строки расчетной таблицы
        setRowTableStyle(
            oRow,
            scol_second,
            v_headers.length,
            row_height,
            'empty',
            h_font
        );

        //  --- Основная таблица
        //  текст ИТОГО основной таблицы
        const o_total_row_text = oRow.getCell(endMainTableCol - 1);//scol + 5
        o_total_row_text.value = "Итого:";
        o_total_row_text.font.size = font_size;
        o_total_row_text.alignment = algn_right;
        o_total_row_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        //  Дубликат в шапке таблицы текста ИТОГО
        const o_d_total_text = topTotalOperRow.getCell(endMainTableCol - 1);
        o_d_total_text.value = o_total_row_text.value;
        o_d_total_text.font = { ...tabl_font };
        o_d_total_text.alignment = algn_right;

        //  Дубликат суммы в шапке таблицы
        const o_d_total_val = topTotalOperRow.getCell(endMainTableCol);
        o_d_total_val.value = {
            formula: `SUM(${tsm_ltr}${operStartInd + 1}:${tsm_ltr}${ind - 1})`
        };
        o_d_total_val.font = { ...tabl_font };
        o_d_total_val.numFmt = n_format;
        o_d_total_val.alignment = algn_right;

        //  Ячейка итоговой суммы с наценкой
        const o_total_row_val = oRow.getCell(endMainTableCol);
        o_total_row_val.value = {
            formula: `SUM(${tsm_ltr}${operStartInd + 1}:${tsm_ltr}${ind - 1})`
        };
        o_total_row_val.numFmt = n_format;
        o_total_row_val.alignment = algn_right;

        //  --- Вспомогательная таблица
        //  Текст ИТОГО вспомогательной таблицы
        const o_total_row_res_text = oRow.getCell(endMainTableCol + offset + 2);
        o_total_row_res_text.value = "Итого:";
        o_total_row_res_text.font.size = font_size;
        o_total_row_res_text.alignment = algn_right;
        o_total_row_res_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };
        //  Дубликат в шапке вспомогательной таблицы текста ИТОГО
        const o_d_totalres_text = topTotalOperRow.getCell(endMainTableCol + offset + 2);
        o_d_totalres_text.value = o_total_row_res_text.value;
        o_d_totalres_text.font = { ...tabl_font };
        o_d_totalres_text.alignment = algn_right;

        //  Ячейка итоговой суммы из Базы
        const o_total_row_res_val = oRow.getCell(endMainTableCol + offset + 3);
        o_total_row_res_val.value = {
            formula: `SUM(${tres_sum_ltr}${operStartInd + 1}:${tres_sum_ltr}${ind - 1})`
        };
        o_total_row_res_val.numFmt = n_format;
        o_total_row_res_val.alignment = algn_right;
        o_total_row_res_val.border = {
            left: { style: 'thin' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        //  Дубликат суммы из базы в шапке таблицы
        const o_d_totalres_val = topTotalOperRow.getCell(endMainTableCol + offset + 3);
        o_d_totalres_val.value = {
            formula: `SUM(${tres_sum_ltr}${operStartInd + 1}:${tres_sum_ltr}${ind - 1})`
        };
        o_d_totalres_val.font = { ...tabl_font };
        o_d_totalres_val.numFmt = n_format;
        o_d_totalres_val.alignment = algn_right;

        //  Ячейка итоговой суммы ВД
        const o_total_row_res_vd_val = oRow.getCell(endMainTableCol + offset + 4);
        o_total_row_res_vd_val.value = {
            formula: `SUM(${tres_vd_ltr}${operStartInd + 1}:${tres_vd_ltr}${ind - 1})`
        };
        o_total_row_res_vd_val.numFmt = n_format;
        o_total_row_res_vd_val.alignment = algn_right;

        //  Дубликат итоговой суммы ВД
        const o_d_totalvd_val = topTotalOperRow.getCell(endMainTableCol + offset + 4);
        o_d_totalvd_val.value = {
            formula: `SUM(${tres_vd_ltr}${operStartInd + 1}:${tres_vd_ltr}${ind - 1})`
        };
        o_d_totalvd_val.font = { ...tabl_font };
        o_d_totalvd_val.numFmt = n_format;
        o_d_totalvd_val.alignment = algn_right;

        //#endregion

        // Литера колонки для итоговой суммы в таблице операций
        const operationTotalCol = getColumnLetter(endMainTableCol);
        const operationTotalRow = ind;

        totalData.push({
            modelName: name,
            modelSign: sign,
            sheet_name: sheet_name,
            qnt: prj_count,
            materialTotalCell: `${materialTotalCol}${materialTotalRow}`,
            operationTotalCell: `${operationTotalCol}${operationTotalRow}`,
        });

        //  Получаем номер последней строки, где есть данные
        //  Устанавливаем область печати: колонки A-scm, все строки с данными
        worksheet.pageSetup.printArea = `A1:${tsm_ltr}${worksheet.rowCount}`;
    });

    //  Создание сводной таблицы по проекту
    function createTotalPage(td, total_worksheet) {
        //  Создание сводной таблицы проекта
        const headers = ['№', 'ID', 'Изделие', 'Кол-во', 'Цена', 'Стоимость'];
        const columns = [
            { width: 0.67 },    //  Отступ
            { width: 5.42 },    //  Номер 
            { width: 7 },       //  Обозначенеи изделия 
            { width: 60 },      //  Название издедия
            { width: 8 },       //  Количество в проекте
            { width: 15 },      //  Цена
            { width: 15 }       //  Сумма
        ];

        const rh = 14;
        total_worksheet.pageSetup = pageSetup;
        total_worksheet.columns = columns;

        const hri = 6;     //  Индекс верхней строки таблицы
        const scol = 2;    //  Начальная колонка
        const hr = total_worksheet.getRow(hri);

        //  Шапка документа
        total_worksheet.getRow(2).height = setRowHeght(24);
        total_worksheet.getRow(3).height = setRowHeght(5);
        styleCellRange(total_worksheet.getRow(3), scol, total_worksheet.columns.length, {
            border: {
                left: { style: 'none' }, right: { style: 'none' },
                top: { style: 'medium' }, bottom: { style: 'none' }
            }
        });

        //  Название документа
        const docNameRow = total_worksheet.getRow(2);
        docNameRow.alignment = { horizontal: 'left', vertical: 'middle' };
        docNameRow.getCell(scol).font = { ...doc_font };
        docNameRow.getCell(scol).value = `Сводная таблица проекта ${PROJECT_NAME}`;

        //  Линия под заголовком страницы
        const lineRow = total_worksheet.getRow(3);
        lineRow.height = setRowHeght(5);
        styleCellRange(lineRow, scol, columns.length, {
            border: {
                left: { style: 'none' }, right: { style: 'none' },
                top: { style: 'medium' }, bottom: { style: 'none' }
            }
        });

        //  Заголовок таблицы
        const headerTableRow = total_worksheet.getRow(4);
        headerTableRow.height = setRowHeght(rh);
        headerTableRow.alignment = { horizontal: 'left', vertical: 'middle' };
        headerTableRow.getCell(scol).font = { ...tabl_font };
        headerTableRow.getCell(scol).value = 'Список изделий';

        //  Строка под заголовком таблицы
        const underTHRow = total_worksheet.getRow(5);
        underTHRow.height = setRowHeght(5);

        //  Шапка таблицы
        for (let i = 0; i < headers.length; i++) {
            hr.getCell(scol + i).value = headers[i];
        };

        setRowTableStyle(
            hr,
            scol,
            headers.length,
            rh,
            'header',
            h_font
        );

        for (let i = 0; i < td.length; i++) {
            const model = td[i];
            const name = model.modelName;
            const sign = model.modelSign;
            const sheetName = model.sheet_name;
            const matTotalCell = model.materialTotalCell;
            const operTotalCell = model.operationTotalCell;
            const qnt = model.qnt;

            const rn = hri + 1 + i; //  Индекс текущей строки
            const row = total_worksheet.getRow(rn);
            //  Ячейка номера строки
            const numCell = row.getCell(scol);
            numCell.value = i + 1;
            numCell.alignment = algn_right;

            //  Ячейка обозначения изделия в проекте
            const signCell = row.getCell(scol + 1);
            signCell.value = sign;
            signCell.alignment = algn_right;

            //  Ячейка названия изделия
            const nameCell = row.getCell(scol + 2);
            nameCell.value = name;
            nameCell.alignment = algn_left;

            //  Ячейка количества
            const qntCell = row.getCell(scol + 3);
            qntCell.value = qnt;
            qntCell.alignment = algn_right;

            //  Ячейка цены
            const priceCell = row.getCell(scol + 4);
            priceCell.value = {
                formula: `('${sheetName}'!${matTotalCell}+'${sheetName}'!${operTotalCell})/${qnt}`
            };
            priceCell.alignment = algn_right;
            priceCell.numFmt = n_format;

            //  Ячейка суммы
            const sumCell = row.getCell(scol + 5);
            sumCell.value = 10000;
            sumCell.value = {
                formula: `'${sheetName}'!${matTotalCell}+'${sheetName}'!${operTotalCell}`
            };
            sumCell.alignment = algn_right;
            sumCell.numFmt = n_format;

            //  Стилизация строк таблицы
            setRowTableStyle(       //  Основная строка
                row,                //  Строка
                scol,               //  Начало диапазона ячеек строки
                headers.length,     //  Конец диапазона ячеек строки
                rh - 1,             //  Высота строки
                'main',             //  Тип строки
                r_font              //  Стиль шрифта
            );
        };
        //  Последняя строка таблицы страницы
        const lastRow = total_worksheet.getRow(hri + td.length);
        setRowTableStyle(       //  Последняя строка таблицы
            lastRow,            //  Строка
            scol,               //  Начало диапазона ячеек строки
            headers.length,     //  Конец диапазона ячеек строки
            rh - 1,             //  Высота строки
            'end',              //  Тип строки
            r_font              //  Стиль шрифта
        );

        const total_row = total_worksheet.getRow(hri + td.length + 1);
        const total_row_text = total_row.getCell(columns.length - 1);
        total_row_text.value = "Итого:";
        total_row_text.alignment = algn_right;
        setRowTableStyle(       //  Последняя строка таблицы
            total_row,          //  Строка
            scol,               //  Начало диапазона ячеек строки
            headers.length,     //  Конец диапазона ячеек строки
            rh - 1,             //  Высота строки
            'empty',            //  Тип строки
            h_font              //  Стиль шрифта
        );
        total_row_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        const totalSumCell = total_row.getCell(scol + headers.length - 1);
        // Литера колонки суммы

        const sumLtr = getColumnLetter(scol + headers.length - 1);
        const startSum = hri + 1;
        const endSum = startSum + td.length - 1;
        totalSumCell.value = {
            formula: `SUM(${sumLtr}${startSum}:${sumLtr}${endSum})`
        };
        totalSumCell.alignment = algn_right;
        totalSumCell.numFmt = n_format;
    };
    createTotalPage(totalData, total_worksheet);

    const fileName = `${sanitizeFileName(
        PROJECT_NAME + "_Калькуляция проекта"
    )}.xlsx`;

    //Сохранение документа. Проверяем/создаем директорию
    const RF = path.join(FOLDER, SW_FOLDER);
    if (!fs.existsSync(RF)) fs.mkdirSync(RF, { recursive: true });
    const filePath = path.join(RF, fileName);
    await workbook.xlsx.writeFile(filePath);
};

//  Функция создания спецификации для загрузки в 1С
async function createSpecificationForImport(materials_data) {
    //#region Проверка настроек
    if (!settings.estimate.classes)
        errFinish("Ошибка файла настроек - classes");
    if (!settings.estimate.fillColor)
        errFinish("Ошибка файла настроек - fillColor");
    //#endregion

    //#region Настройки стилей
    const fontFamily = settings.estimate.fontFamily || "Arial";
    const rh = 14;
    const font_size = 9;
    const h_font = { name: fontFamily, size: font_size, bold: true };
    const r_font = { name: fontFamily, size: font_size - 1, bold: false };
    const fill = settings.estimate.fillColor;

    // Заливка ячеек
    const fill_yellow = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: fill.yellow }
    };

    // Форматирование чисел
    const f_format = '#,##0.00';
    const n_format = '# ##0';

    const algnLeft = { indent: 1, horizontal: 'left', vertical: 'middle' };
    const algnRight = { indent: 1, horizontal: 'right', vertical: 'middle' };
    const algnCenter = { horizontal: 'center', vertical: 'middle' };
    //#endregion

    // Создаем файл спецификации
    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.author;
    workbook.created = new Date();

    // Создание и настройка страницы
    const worksheet = workbook.addWorksheet('Спецификация');
    worksheet.pageSetup = {
        orientation: 'portrait',
        margins: {
            top: 0.5,
            bottom: 0.5,
            left: 0.39,
            right: 0.39,
            header: 0.3,
            footer: 0.3
        },
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
    };

    // Колонки таблицы (только те, что нужны для 1С)
    worksheet.columns = [
        { width: 12 },      // Код (materialSyncExternal)
        { width: 20 },      // Артикул (materialArticle)
        { width: 75 },      // Наименование (materialName)
        { width: 12 },      // Количество (quantity)
        { width: 20 },      // Ед. (materialUnit)
    ];

    //  Заголовки колонок таблицы
    const startRowInd = 1;
    const startColInd = 1;
    const headers = ['Код', 'Артикул', 'Номенклатура', 'Количество', 'Единица измерения'];

    // Добавляем заголовки колонок (одна строка)
    const headerRow = worksheet.getRow(startRowInd);
    headerRow.alignment = algnCenter;
    headerRow.height = rh;

    for (let i = 0; i < headers.length; i++) {
        headerRow.getCell(startColInd + i).value = headers[i];
    };
    setRowTableStyle(
        headerRow,
        startColInd,
        headers.length,
        rh,
        'header',
        h_font
    );
    //  Индекс строки таблицы сразу после заголовка
    let ind = startRowInd + 1;

    // Проходим по классам и заполняем таблицу
    classes.forEach(key => {
        if (!materials_data[key].items.length) return;

        //  Сортировка материалов в пределах класса
        let sort_options = [["materialName", "desc"]];
        if (materials_data[key].items[0].area) {
            // Площадной материалв
            sort_options = [
                ["materialTkn", "desc"], ["materialName", "asc"]
            ];
        } else if (materials_data[key].items[0].length) {
            //  Кромка или погонный материал
            sort_options = [
                ["materialName", "desc"]
            ];
        } else if (materials_data[key].items[0].count) {
            //  Фурнитура
            sort_options = [
                ["materialName", "desc"]
            ];
        };

        // Сортируем материалы в классе
        const items = smartSort(materials_data[key].items, sort_options);

        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const row = worksheet.getRow(ind + i);

            //  Количество
            let unit = item.materialUnit;
            let price = item.materialPrice;
            let quantity = 0;
            if (item.area) {
                quantity = item.area;
            } else if (item.length) {
                quantity = item.length;
            } else if (item.count) {
                quantity = item.count;
            };

            //  Преобразование площадного материала  
            if (settings.estimate.useCoefficient) {
                let k = item.k ? item.k : 1;
                price = k * price;
                if (item.area) {
                    quantity = Math.round(quantity * k);
                    unit = 'шт';
                };
            } else {
                if (item.area) unit = 'м2';
            };

            // Ячейка кода номенклатуры
            const codeCell = row.getCell(startColInd);
            codeCell.value = item.materialSyncExternal;
            codeCell.alignment = algnLeft;
            if (!codeCell.value) codeCell.fill = fill_yellow;

            // Ячейка артикула
            const articleCell = row.getCell(startColInd + 1);
            articleCell.value = item.materialArticle;
            articleCell.alignment = algnLeft;

            // Ячейка названия номенклатуры
            const materialCell = row.getCell(startColInd + 2);
            materialCell.value = item.materialName;
            materialCell.alignment = algnLeft;

            // Ячейка количества
            const quantityCell = row.getCell(startColInd + 3);
            quantityCell.value = quantity;
            quantityCell.alignment = algnRight;
            quantityCell.numFmt = f_format;

            // Ячейка единицы измерения
            const unitCell = row.getCell(startColInd + 4);
            unitCell.value = unit;
            if (!unitCell.value) unitCell.fill = fill_yellow;
            unitCell.alignment = algnRight;

            //  Стилизация строк таблицы
            setRowTableStyle(
                row,
                startColInd,
                headers.length,
                rh - 1,
                'main',
                r_font
            );
        };
        ind += items.length;
    });

    //  Последняя строка таблицы страницы
    const lastRow = worksheet.getRow(ind - 1);
    setRowTableStyle(       //  Последняя строка таблицы
        lastRow,            //  Строка
        startColInd,        //  Начало диапазона ячеек строки
        headers.length,     //  Конец диапазона ячеек строки
        rh - 1,             //  Высота строки
        'end',              //  Тип строки
        r_font              //  Стиль шрифта
    );

    const fileName = `${sanitizeFileName(
        PROJECT_NAME + "_Спецификация для загрузки в 1С" + ` (${ind})`
    )}.xlsx`;

    // Проверяем/создаем директорию
    const RF = path.join(FOLDER, SW_FOLDER);
    if (!fs.existsSync(RF)) fs.mkdirSync(RF, { recursive: true });
    const filePath = path.join(RF, fileName);
    await workbook.xlsx.writeFile(filePath);
};

//  Функция создания спецификации материалов проекта
async function createSpecificationProjectFile(agr_arr, prj_arr, settings) {

    //#region Настройки стилей документа

    //  Шрифт документа
    const fontFamily = settings.estimate.fontFamily ?
        settings.estimate.fontFamily : "Arial";
    const font_size = 9;        //  Размер шрифта 
    const doc_font = {          //  Настройки шрифта страницы
        name: fontFamily,
        size: font_size + 7,
        bold: true
    };
    const tabl_font = {         //  Настройки шрифта названия таблицы
        name: fontFamily,
        size: font_size + 2,
        bold: true
    };
    const h_font = {            //  Настройик шрифта заглавной строки таблицы
        name: fontFamily,
        size: font_size,
        bold: true
    };
    const r_font = {            //  Настройки шрифта строки таблицы
        name: fontFamily,
        size: font_size - 1,
        bold: false
    };
    const h1 = 24;              //  Высота строки заголовка страницы
    const rh = 14;              //  Высота строки данных (таблицы)   
    const scol = 2;             //  Начальная колонка таблицы
    const headerRowInd = 6;     //  Индекс верхней строки таблицы

    //  Настройки цветов заливки ячеек
    const fill = settings.estimate.fillColor;
    const fill_green = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill.green } };
    const fill_yellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill.yellow } };
    const fill_orange = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill.orange } };
    const fill_red = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill.red } };
    const fill_blue = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill.blue } };

    //  Настройки выравнивания в ячейках
    const algnLeft = { indent: 1, horizontal: 'left', vertical: 'middle' };
    const algnRight = { indent: 1, horizontal: 'right', vertical: 'middle' };
    const algnCenter = { horizontal: 'center', vertical: 'middle' };

    //  Форматирование чисел
    const f_format = '#,##0.00';
    const sf_format = '#,##0.0';
    const n_format = '# ##0';

    //#endregion

    //  Название колонок таблицы
    const headers = ['№', 'Код', 'Артикул', 'Наименование',
        'Кол-во', 'Ед.', 'Цена', 'Сумма'];

    //  Настройка ширины колонок документа
    const columns = [
        { width: 0.67 },    //  Отступ
        { width: 5.42 },    //  Номер
        { width: 16 },      //  Код номенклатуры
        { width: 18 },      //  Артикул
        { width: 60 },      //  Наименование 78
        { width: 8 },       //  Количество
        { width: 7 },       //  Ед.
        { width: 11 },      //  Цена из Базы материалов
        { width: 12 }       //  Сумма
    ];

    //  Создаем новую книгу документа
    const workbook = new ExcelJS.Workbook();    //  Новая книга
    workbook.creator = settings.author;         //  Автор документа
    workbook.created = new Date();              //  Дата документа

    const pageSetup = {
        orientation: 'portrait',    // 'portrait' | 'landscape'
        margins: {                  // Поля страницы в дюймах (1д = 2.54см)
            top: 0.5,               // Верхнее поле
            bottom: 0.5,            // Нижнее поле
            left: 0.39,             // Левое поле
            right: 0.39,            // Правое поле
            header: 0.3,            // Отступ для колонтитула сверху
            footer: 0.3             // Отступ для колонтитула снизу
        },
        // Масштабирование
        fitToPage: true,            // Вписать в страницу
        fitToWidth: 1,              // Вписать по ширине (1 страница)
        fitToHeight: 0,             // По высоте (0 = автоматически)
    };

    //  Функция создания таблицы вкладки
    function createSheet(data) {
        const {
            classes,
            estimate,
            columns,
            headers,
            sheet_name,
            sheet_header,
            union_page = false,
        } = data;

        //  Создаем новую вкладку документа
        const worksheet = workbook.addWorksheet(sheet_name);
        worksheet.pageSetup = pageSetup;
        worksheet.columns = columns;

        //  Шапка страницы
        const docNameRow = worksheet.getRow(2);
        docNameRow.height = setRowHeght(h1);
        docNameRow.alignment = { horizontal: 'left', vertical: 'middle' };
        docNameRow.getCell(scol).font = { ...doc_font };
        docNameRow.getCell(scol).value = sheet_header;

        //  Линия под заголовком страницы
        const lineRow = worksheet.getRow(3);
        lineRow.height = setRowHeght(rh);
        styleCellRange(lineRow, scol, columns.length, {
            border: {
                left: { style: 'none' }, right: { style: 'none' },
                top: { style: 'medium' }, bottom: { style: 'none' }
            }
        });

        //  Заголовок таблицы
        const headerTableRow = worksheet.getRow(4);
        headerTableRow.height = setRowHeght(rh);
        headerTableRow.alignment = { horizontal: 'left', vertical: 'middle' };
        headerTableRow.getCell(scol).font = { ...tabl_font };
        headerTableRow.getCell(scol).value = 'Спецификация материалов и фурнитуры';

        const del_t_row = worksheet.getRow(5);
        del_t_row.height = setRowHeght(5);

        //  Шапка таблицы
        const headerRow = worksheet.getRow(headerRowInd);
        for (let i = 0; i < headers.length; i++) {
            headerRow.getCell(scol + i).value = headers[i];
        };
        setRowTableStyle(
            headerRow,
            scol,
            headers.length,
            rh,
            'header',
            h_font
        );
        //  Индекс строки таблицы сразу после заголовка
        let ind = headerRowInd + 1;

        //  Тело таблицы
        classes.forEach(key => {
            if (!estimate[key].items.length) return;

            //  Сортировка материалов в пределах класса
            let sort_options = [["materialName", "desc"]];
            if (estimate[key].items[0].area) {
                // Площадной материалв
                sort_options = [
                    ["materialTkn", "desc"], ["materialName", "asc"]
                ];
            } else if (estimate[key].items[0].length) {
                //  Кромка или погонный материал
                sort_options = [
                    ["materialName", "desc"]
                ];
            } else if (estimate[key].items[0].count) {
                //  Фурнитура
                sort_options = [
                    ["materialName", "desc"]
                ];
            };

            //  Сортировка
            const items = smartSort(estimate[key].items, sort_options);

            //  Цикл по массиву класса
            for (let i = 0; i < items.length; i++) {

                //  Объект материала
                const item = items[i];
                const rn = ind + i; //  Индекс текущей строки
                const row = worksheet.getRow(rn);

                let unit = item.materialUnit;
                let price = item.materialPrice;

                //  Количество
                let quantity = 0;
                if (item.area) {
                    quantity = item.area;
                } else if (item.length) {
                    quantity = item.length;
                } else if (item.count) {
                    quantity = item.count;
                };

                //  Преобразование площадного материала 
                const use_coeff = settings.estimate.useCoefficient;
                if (use_coeff && union_page) {
                    let k = item.k ? item.k : 1;
                    price = price / k;  // вычисляем цену листа
                    if (item.area) {
                        quantity = Math.round(quantity * k);
                        unit = 'шт';
                    };
                } else {
                    if (item.area) unit = 'м2';
                };

                //  Ячейка номера строки
                const numCell = row.getCell(scol);
                numCell.value = i + 1;
                numCell.alignment = algnRight;

                //  Ячейка кода номенклатуры
                const codeCell = row.getCell(scol + 1);
                codeCell.value = item.materialSyncExternal;
                codeCell.alignment = algnLeft;

                //  Ячейка артикула номенклатуры
                const articleCell = row.getCell(scol + 2);
                articleCell.value = item.materialArticle;
                articleCell.alignment = algnLeft;

                //  Ячейка названия номенклатуры
                const materialCell = row.getCell(scol + 3);
                materialCell.value = item.materialName;
                materialCell.alignment = algnLeft;

                //  Ячейка количества
                const quantityCell = row.getCell(scol + 4);
                quantityCell.value = quantity;
                quantityCell.alignment = algnRight;
                quantityCell.numFmt = f_format;

                //  Ячейка единицы измерения
                const unitCell = row.getCell(scol + 5);
                unitCell.value = unit;
                unitCell.alignment = algnRight;

                //  Ячейка цены
                const priceCell = row.getCell(scol + 6);
                priceCell.value = price;
                priceCell.alignment = algnRight;
                priceCell.numFmt = f_format;

                //  Литеры ячеек количества и цены
                const qnt_ltr = getColumnLetter(scol + 4);
                const price_ltr = getColumnLetter(scol + 6);

                //  Ячейка суммы
                const sumCell = row.getCell(scol + 7);
                sumCell.value = {
                    formula: `${price_ltr}${rn}*${qnt_ltr}${rn}`
                };
                sumCell.alignment = algnRight;
                sumCell.numFmt = f_format;

                //  Стилизация строк таблицы
                setRowTableStyle(       //  Основная строка
                    row,                //  Строка
                    scol,               //  Начало диапазона ячеек строки
                    headers.length,     //  Конец диапазона ячеек строки
                    rh - 1,             //  Высота строки
                    'main',             //  Тип строки
                    r_font              //  Стиль шрифта
                );
            };
            ind += items.length;
        });

        //  Последняя строка таблицы страницы
        const lastRow = worksheet.getRow(ind - 1);
        setRowTableStyle(       //  Последняя строка таблицы
            lastRow,            //  Строка
            scol,               //  Начало диапазона ячеек строки
            headers.length,     //  Конец диапазона ячеек строки
            rh - 1,             //  Высота строки
            'end',              //  Тип строки
            r_font              //  Стиль шрифта
        );
    };

    //  Создаем сводную страницу проекта
    createSheet({
        classes: classes,
        estimate: agr_arr,
        columns: columns,
        headers: headers,
        sheet_name: `Спецификация проекта`.substring(0, 31),
        sheet_header: `Сводная таблица проекта ${PROJECT_NAME}`,
        union_page: true,
    });


    prj_arr.forEach(model => {
        const estimate = model.estimate_data;   //  Объект данных Сметы
        const name = model.modelName;      //  Имя изделия
        const sign = model.sign;                //  Обозначение в Проекте
        const sheet_name = `${sign}_${name}`.substring(0, 31);
        const sheet_header = `${name}`;

        //  Создаем вкладки по изделиям проекта
        createSheet({
            classes: classes,
            estimate: estimate,
            columns: columns,
            headers: headers,
            sheet_name: sheet_name,
            sheet_header: sheet_header
        });
    });

    // Сохранение документа
    const fileName = `${sanitizeFileName(
        PROJECT_NAME + "_Спецификация проекта")}.xlsx`;

    // Проверяем/создаем директорию
    const RF = path.join(FOLDER, SW_FOLDER);
    if (!fs.existsSync(RF)) fs.mkdirSync(RF, { recursive: true });
    const filePath = path.join(RF, fileName);
    await workbook.xlsx.writeFile(filePath);
};

//  Функция создания PDF сборочных чертежей
async function createPDFAssemblyDrawings(assembly_array, settings) {

    const mm = 2.83465;
    const fontName = 'Arial';

    //  Группировка деталей по позиции и суммированием количество
    function groupPartsByPos(items, groupField = 'pos') {
        if (!items || !Array.isArray(items) || items.length === 0) return [];

        const groups = {};
        for (const item of items) {
            const key = item[groupField];
            if (!key) continue;

            // Проверяем: если в конце наименования есть _любая_буква - пропускаем
            if (item.name && /_[а-яА-Яa-zA-Z]$/.test(item.name)) {
                continue; // Пропускаем эту деталь
            }

            if (!groups[key]) {
                // Сохраняем первый встреченный объект, добавляем поле count
                groups[key] = {
                    ...item,
                    count: 0
                };
            };
            // Увеличиваем счетчик
            groups[key].count += 1;
        };

        // Преобразуем объект в массив и сортируем по pos
        return Object.values(groups).sort((a, b) => {
            const aVal = parseFloat(a[groupField]) || 0;
            const bVal = parseFloat(b[groupField]) || 0;
            return aVal - bVal;
        });
    };

    //  Функция рисует прямоугольник и текст внутри
    function drawCell(doc, options = {}) {
        const {
            x,
            y,
            width,
            height,
            text = '',
            align = 'center',
            valign = 'center',
            fontName: font = 'Arial',
            fontSize = 8,
            fontColor = '#000000',
            bold = false,
            italic = false,
            underline = false,
            strike = false,
            border = [0, 0, 0, 0],      // [верх, низ, лево, право]
            borderColor = '#000000',
            padding = 2,
            fillColor = null,           // Цвет заливки
            buttInfo = [],              // [верхняя_линия, нижняя_линия]
            buttOffsetTop = 0.5,        // Смещение верхней линии от верха ячейки (в пунктах)
            buttOffsetBottom = 0.5,     // Смещение нижней линии от низа ячейки (в пунктах)
            buttLineWidth = 0.6,        // Ширина линии (в процентах от ширины ячейки)
            buttDashLength = 2,         // Длина штриха для пунктира
            buttDashSpace = 2           // Длина пробела для пунктира
        } = options;

        doc.save();

        // Заливка фона
        if (fillColor) {
            doc.fillColor(fillColor)
                .rect(x, y, width, height)
                .fill();
        }

        // Рисуем рамку
        if (border && Array.isArray(border) && border.length === 4) {
            const [top, bottom, left, right] = border;

            if (top > 0) {
                doc.lineWidth(top)
                    .moveTo(x, y)
                    .lineTo(x + width, y)
                    .stroke(borderColor);
            };

            if (bottom > 0) {
                doc.lineWidth(bottom)
                    .moveTo(x, y + height)
                    .lineTo(x + width, y + height)
                    .stroke(borderColor);
            };

            if (left > 0) {
                doc.lineWidth(left)
                    .moveTo(x, y)
                    .lineTo(x, y + height)
                    .stroke(borderColor);
            };

            if (right > 0) {
                doc.lineWidth(right)
                    .moveTo(x + width, y)
                    .lineTo(x + width, y + height)
                    .stroke(borderColor);
            };
        };

        // Рисуем линии buttInfo (сверху и снизу текста)
        if (buttInfo && Array.isArray(buttInfo) && buttInfo.length > 0) {
            const lineWidth_ = width * (buttLineWidth || 0.7);
            const startX = x + (width - lineWidth_) / 2;

            // Верхняя линия (buttInfo[0])
            if (buttInfo[0] !== undefined && buttInfo[0] !== null) {
                const value = buttInfo[0];
                const yLine = y + padding + (buttOffsetTop || 0);
                drawButtLine(doc, startX, yLine, lineWidth_, value, borderColor, buttDashLength, buttDashSpace);
            }

            // Нижняя линия (buttInfo[1])
            if (buttInfo[1] !== undefined && buttInfo[1] !== null) {
                const value = buttInfo[1];
                const yLine = y + height - padding - (buttOffsetBottom || 0);
                drawButtLine(doc, startX, yLine, lineWidth_, value, borderColor, buttDashLength, buttDashSpace);
            }
        }

        // Формируем имя шрифта с учетом стилей
        let fontWithStyle = font;
        if (font === 'Arial' || font === 'Helvetica') {
            if (bold && italic) fontWithStyle = 'Arial-BoldItalic';
            else if (bold) fontWithStyle = 'Arial-Bold';
            else if (italic) fontWithStyle = 'Arial-Italic';
            else fontWithStyle = 'Arial';
        } else {
            if (bold && italic) fontWithStyle = font + '-BoldItalic';
            else if (bold) fontWithStyle = font + '-Bold';
            else if (italic) fontWithStyle = font + '-Italic';
            else fontWithStyle = font;
        }

        // Позиция текста
        let textX = x + padding;
        let textY = y + padding;

        // Корректируем textY если есть верхняя линия
        if (buttInfo && buttInfo.length > 0 && buttInfo[0] !== undefined && buttInfo[0] !== null) {
            const topLineOffset = 4 + (buttOffsetTop || 0);
            textY = y + padding + topLineOffset;
        }

        if (valign === 'center') {
            textY = y + height / 2;
            if (buttInfo && buttInfo.length > 0) {
                const hasTop = buttInfo[0] !== undefined && buttInfo[0] !== null;
                const hasBottom = buttInfo[1] !== undefined && buttInfo[1] !== null;
                const topOffset = buttOffsetTop || 0;
                const bottomOffset = buttOffsetBottom || 0;

                if (hasTop && hasBottom) {
                    textY = y + height / 2 + (topOffset - bottomOffset) / 2;
                } else if (hasTop) {
                    textY = y + height / 2 + topOffset + 3;
                } else if (hasBottom) {
                    textY = y + height / 2 - bottomOffset - 3;
                }
            }
        } else if (valign === 'bottom') {
            textY = y + height - padding;
            if (buttInfo && buttInfo.length > 0 && buttInfo[1] !== undefined && buttInfo[1] !== null) {
                const bottomLineOffset = 4 + (buttOffsetBottom || 0);
                textY = y + height - padding - bottomLineOffset;
            }
        }

        const textWidth = width - padding * 2;

        // Применяем шрифт
        doc.fontSize(fontSize)
            .font(fontWithStyle)
            .fillColor(fontColor);

        // Рисуем текст
        if (valign === 'center') {
            doc.text(text, textX, textY, {
                width: textWidth,
                align: align,
                baseline: 'middle',
                lineBreak: false
            });
        } else {
            doc.text(text, textX, textY, {
                width: textWidth,
                align: align,
                baseline: 'top',
                lineBreak: false
            });
        }

        // Подчеркивание и зачеркивание
        if (underline || strike) {
            const textWidth_ = doc.widthOfString(text, {
                fontSize: fontSize,
                font: fontWithStyle
            });

            const lineWidth_ = Math.min(textWidth_, textWidth);

            if (underline) {
                const lineY = valign === 'center' ?
                    textY + fontSize * 0.15 :
                    textY + fontSize * 0.8;
                doc.moveTo(textX, lineY)
                    .lineTo(textX + lineWidth_, lineY)
                    .stroke(fontColor);
            }

            if (strike) {
                const strikeY = valign === 'center' ?
                    textY :
                    textY + fontSize * 0.4;
                doc.moveTo(textX, strikeY)
                    .lineTo(textX + lineWidth_, strikeY)
                    .stroke(fontColor);
            }
        }

        doc.restore();
    };

    /**
     * Рисует линию в зависимости от значения
     * @param {PDFDocument} doc - экземпляр PDF документа
     * @param {number} x - координата X начала линии
     * @param {number} y - координата Y линии
     * @param {number} width - ширина линии
     * @param {number} value - значение для определения типа линии
     * @param {string} color - цвет линии
     * @param {number} dashLength - длина штриха для пунктира
     * @param {number} dashSpace - длина пробела для пунктира
     */
    function drawButtLine(doc, x, y, width, value, color = '#000000', dashLength = 2, dashSpace = 2) {
        if (value > 0 && value <= 0.5) {
            // Пунктирная линия
            doc.save();
            doc.lineWidth(0.5)
                .dash(dashLength, { space: dashSpace })
                .moveTo(x, y)
                .lineTo(x + width, y)
                .stroke(color);
            doc.restore();
        } else if (value > 0.5 && value <= 1.5) {
            // Тонкая сплошная (0.5)
            doc.save();
            doc.lineWidth(0.5)
                .moveTo(x, y)
                .lineTo(x + width, y)
                .stroke(color);
            doc.restore();
        } else if (value > 1.5) {
            // Толстая сплошная (1.0)
            doc.save();
            doc.lineWidth(1.0)
                .moveTo(x, y)
                .lineTo(x + width, y)
                .stroke(color);
            doc.restore();
        }
    };

    //  Функция вставки изображения
    function addImageFixedWidth(doc, options = {}) {
        const {
            imagePath,
            x,
            y,
            containerWidth,
            containerHeight,
            fit = 'contain',
            align = 'center',
            valign = 'center',
            debug = false,                      // Включить режим отладки
            debugColor = '#E8F4FD',           // Цвет фона (бледно-голубой)
            debugBorderColor = '#4A90D9',     // Цвет рамки
            debugBorderWidth = 0.5              // Толщина рамки
        } = options;

        // Проверяем существование файла
        if (!imagePath || !fs.existsSync(imagePath)) {
            console.warn(`⚠️ Изображение не найдено: ${imagePath}`);

            // В режиме отладки показываем пустой контейнер
            if (debug && containerWidth && containerHeight) {
                drawDebugContainer(doc, x, y, containerWidth, containerHeight, debugColor, debugBorderColor, debugBorderWidth, '❌ Изображение не найдено');
            }
            return null;
        };

        try {
            // Читаем файл в буфер
            const imageBuffer = fs.readFileSync(imagePath);

            // Получаем размеры изображения
            const dimensions = imageSize(imageBuffer);
            const imageWidth = dimensions.width;
            const imageHeight = dimensions.height;

            let finalWidth, finalHeight;
            let offsetX = 0, offsetY = 0;

            // Если задана высота контейнера - вписываем в контейнер
            if (containerHeight) {
                const scaleX = containerWidth / imageWidth;
                const scaleY = containerHeight / imageHeight;

                if (fit === 'contain') {
                    const scale = Math.min(scaleX, scaleY);
                    finalWidth = Math.floor(imageWidth * scale);
                    finalHeight = Math.floor(imageHeight * scale);
                } else if (fit === 'cover') {
                    const scale = Math.max(scaleX, scaleY);
                    finalWidth = Math.floor(imageWidth * scale);
                    finalHeight = Math.floor(imageHeight * scale);
                }

                // Вычисляем смещение для центрирования
                if (align === 'center') {
                    offsetX = (containerWidth - finalWidth) / 2;
                } else if (align === 'right') {
                    offsetX = containerWidth - finalWidth;
                }

                if (valign === 'center') {
                    offsetY = (containerHeight - finalHeight) / 2;
                } else if (valign === 'bottom') {
                    offsetY = containerHeight - finalHeight;
                }

            } else {
                // Только по ширине (старое поведение)
                const scale = containerWidth / imageWidth;
                finalWidth = containerWidth;
                finalHeight = Math.floor(imageHeight * scale);
            }

            // В режиме отладки рисуем контейнер
            if (debug && containerWidth && containerHeight) {
                drawDebugContainer(doc, x, y, containerWidth, containerHeight, debugColor, debugBorderColor, debugBorderWidth);

                // Добавляем информацию о размерах
                doc.save();
                doc.fontSize(6)
                    .font('Helvetica')
                    .fillColor('#666666')
                    .text(
                        `img: ${imageWidth}x${imageHeight} → ${finalWidth}x${finalHeight} (scale: ${(finalWidth / imageWidth).toFixed(2)})`,
                        x + 2,
                        y + containerHeight - 12,
                        { width: containerWidth - 4, align: 'left' }
                    );
                doc.restore();
            }

            // Добавляем изображение со смещением
            doc.image(imageBuffer, {
                x: x + offsetX,
                y: y + offsetY,
                width: finalWidth,
                height: finalHeight
            });

            return {
                width: finalWidth,
                height: finalHeight,
                offsetX: offsetX,
                offsetY: offsetY,
                scale: finalWidth / imageWidth,
                originalWidth: imageWidth,
                originalHeight: imageHeight
            };
        } catch (err) {
            console.error(`❌ Ошибка добавления изображения: ${err.message}`);

            // В режиме отладки показываем контейнер с ошибкой
            if (debug && containerWidth && containerHeight) {
                drawDebugContainer(doc, x, y, containerWidth, containerHeight, '#FFE8E8', '#D94A4A', 0.5, `❌ Ошибка: ${err.message}`);
            }
            return null;
        };
    };

    //  Функция отрисовки рамки вокруг изображения для отладки
    function drawDebugContainer(doc, x, y, width, height, fillColor, borderColor, borderWidth, message = '') {
        doc.save();
        // Заливка фона
        if (fillColor) {
            doc.fillColor(fillColor)
                .rect(x, y, width, height)
                .fill();
        };
        // Рамка контейнера
        if (borderColor) {
            doc.lineWidth(borderWidth || 0.5)
                .rect(x, y, width, height)
                .stroke(borderColor);
        };
        // Текст-заглушка
        if (message) {
            doc.fontSize(8)
                .font('Helvetica')
                .fillColor('#666666')
                .text(message, x + 5, y + height / 2 - 4, {
                    width: width - 10,
                    align: 'center',
                    baseline: 'middle'
                });
        };
        // Крестики по углам для визуального контроля размера
        const crossSize = 8;
        doc.lineWidth(0.3)
            .strokeColor('#999999');
        // Верхний левый угол
        doc.moveTo(x, y + crossSize).lineTo(x, y).lineTo(x + crossSize, y);
        // Верхний правый угол
        doc.moveTo(x + width - crossSize, y).lineTo(x + width, y).lineTo(x + width, y + crossSize);
        // Нижний левый угол
        doc.moveTo(x, y + height - crossSize).lineTo(x, y + height).lineTo(x + crossSize, y + height);
        // Нижний правый угол
        doc.moveTo(x + width - crossSize, y + height).lineTo(x + width, y + height).lineTo(x + width, y + height - crossSize);
        doc.restore();
    };

    //  Функция создания штампа сборочного чертежа
    function createDocStamp(doc, pageWidth, pageHeight, options = {}) {
        const {
            data,
            title = '',
            designation = '',
            sheet = '1',
            totalSheets = '1',
            margin = 5 * mm,
            fontSize = 9,
            font = 'Arial'
        } = options;

        const auSign = data.auDesignation;
        const auPos = data.auArtPos;

        // Твои размеры
        const ch = 7 * mm;
        const cw = 99 * mm;
        const squareSize = 21 * mm;
        const m_border = [1.5, 1.5, 1.5, 1.5];

        // Координаты правого нижнего угла штампа
        let x = pageWidth - margin - cw - squareSize;
        let y = pageHeight - margin - ch * 3;

        //  Название чертежа
        drawCell(doc, {
            x: x,
            y: y,
            width: cw,
            height: ch,
            text: `Сборочный чертеж на ${data.auName}`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: m_border,
        });

        //  Название модели
        drawCell(doc, {
            x: x,
            y: y + ch,
            width: cw,
            height: ch,
            text: `${data.modelName}`,
            fontName: font,
            fontSize: fontSize,
            border: m_border,
        });

        //  Дата печати отчетов
        drawCell(doc, {
            x: x,
            y: y + 2 * ch,
            width: cw - 24 * mm,
            height: ch,
            text: `Дата печати: ${formatDate(new Date())} г.`,
            fontName: font,
            fontSize: fontSize,
            border: m_border,
        });

        //  Лист / кол-во листов
        drawCell(doc, {
            x: x + cw - 24 * mm,
            y: y + 2 * ch,
            width: 24 * mm,
            height: ch,
            text: `Лист ${sheet} / ${totalSheets}`,
            fontName: font,
            fontSize: fontSize,
            border: m_border,
        });

        // Квадрат справа от штампа
        drawCell(doc, {
            x: pageWidth - margin - squareSize,
            y: pageHeight - margin - squareSize,
            width: squareSize,
            height: squareSize,
            text: `${auSign ? auSign : auPos}`,
            fontName: font,
            fontSize: 24,
            border: m_border,
        });

        //  Прямоугольник вверху слева
        drawCell(doc, {
            x: pageWidth - margin - 45 * mm,
            y: margin,
            width: 45 * mm,
            height: 7 * mm,
            text: `${data.prjName} / ${data.sign}`,
            fontName: font,
            fontSize: fontSize + 4,
            bold: true,
            border: m_border,
        });
    };

    //  Создание штампа листа схемы сборки
    function createDocUnionStamp(doc, pageWidth, pageHeight, options = {}) {
        const {
            data,
            title = '',
            designation = '',
            sheet = '1',
            totalSheets = '1',
            margin = 5 * mm,
            fontSize = 8,
            font = 'Arial'
        } = options;

        const ch = 7 * mm;
        const cw = 99 * mm;
        const squareSize = 21 * mm;
        const m_border = [1.5, 1.5, 1.5, 1.5];

        const width_header = pageWidth - 2 * margin - 90 * mm;
        const x_header = (pageWidth - width_header) / 2;
        const y_header = margin + 5 * mm;

        //  Добавляем название листа
        drawCell(doc, {
            x: x_header,
            y: y_header,
            width: width_header,
            height: ch,
            text: `Схема сборки`,
            fontName: fontName,
            fontSize: fontSize + 6,
            bold: true,
        });

        //  Добавляем название изделия
        drawCell(doc, {
            x: x_header,
            y: y_header + ch,
            width: width_header,
            height: ch,
            text: `${data.modelName}`,
            fontName: fontName,
            fontSize: fontSize + 6,
            bold: true,
        });

        // Координаты правого нижнего угла штампа
        let x = pageWidth - margin - cw - squareSize;
        let y = pageHeight - margin - ch * 3;

        //  Название чертежа
        drawCell(doc, {
            x: x,
            y: y,
            width: cw,
            height: ch,
            text: `Схема сборки на ${data.modelName}`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: m_border,
        });

        //  Название модели
        drawCell(doc, {
            x: x,
            y: y + ch,
            width: cw,
            height: ch,
            text: `${data.modelName}`,
            fontName: font,
            fontSize: fontSize,
            border: m_border,
        });

        //  Дата печати отчетов
        drawCell(doc, {
            x: x,
            y: y + 2 * ch,
            width: cw - 24 * mm,
            height: ch,
            text: `Дата печати: ${formatDate(new Date())} г.`,
            fontName: font,
            fontSize: fontSize,
            border: m_border,
        });

        //  Лист / кол-во листов
        drawCell(doc, {
            x: x + cw - 24 * mm,
            y: y + 2 * ch,
            width: 24 * mm,
            height: ch,
            text: `Лист ${sheet} / ${totalSheets}`,
            fontName: font,
            fontSize: fontSize,
            border: m_border,
        });

        // Квадрат справа от штампа
        drawCell(doc, {
            x: pageWidth - margin - squareSize,
            y: pageHeight - margin - squareSize,
            width: squareSize,
            height: squareSize,
            text: `${data.sign}`,
            fontName: font,
            fontSize: 24,
            border: m_border,
        });

        //  Название проекта / обозначение изделия в проекте
        drawCell(doc, {
            x: pageWidth - margin - 45 * mm,
            y: margin,
            width: 45 * mm,
            height: 7 * mm,
            text: `${data.prjName} / ${data.sign}`,
            fontName: font,
            fontSize: fontSize + 4,
            bold: true,
            border: m_border,
        });
    };

    //  Таблица деталей сборочной единицы
    function createAUPartsTable(doc, options) {
        const {
            data,
            title = '',
            designation = '',
            scale = '1:1',
            sheet = '1',
            totalSheets = '1',
            margin = 5 * mm,
            fontSize = 6,
            font = 'Arial'
        } = options;

        let x = 10 * mm;
        let y = 10 * mm;
        const ch = 5 * mm;
        const cw = 60 * mm;


        const b_r = [0.5, 0.5, 0.5, 1];
        const b_l = [0.5, 0.5, 1, 0.5];
        const b_cell = [0.5, 0.5, 0.5, 0.5];
        const b_t = [1, 0.5, 0.5, 0.5];
        const b_b = [0.5, 1, 0.5, 0.5];
        const b_tr = [1, 0.5, 0.5, 1];
        const b_tl = [1, 0.5, 1, 0.5];
        const b_br = [0.5, 1, 0.5, 1];
        const b_bl = [0.5, 1, 1, 0.5];

        //  Колонка номера
        const w1_col = 6 * mm;
        drawCell(doc, {
            x: x,
            y: y,
            width: w1_col,
            height: ch,
            text: `№`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: b_tl,
            align: 'center',
        });

        //  Колонка Pos
        const w2_col = 15 * mm;
        drawCell(doc, {
            x: x + w1_col,
            y: y,
            width: w2_col,
            height: ch,
            text: `Pos`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: b_t,
            align: 'center',
        });

        //  Колонка наименование
        const w3_col = 45 * mm;
        drawCell(doc, {
            x: x + w1_col + w2_col,
            y: y,
            width: w3_col,
            height: ch,
            text: `Наименование`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: b_t,
            align: 'center',
        });

        //  Количество
        const w4_col = 10 * mm;
        drawCell(doc, {
            x: x + w1_col + w2_col + w3_col,
            y: y,
            width: w4_col,
            height: ch,
            text: `Кол-во`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: b_t,
            align: 'center',
        });

        //  Размеры
        const w5_col = 20 * mm;
        drawCell(doc, {
            x: x + w1_col + w2_col + w3_col + w4_col,
            y: y,
            width: w5_col,
            height: ch,
            text: `Размеры, мм`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: b_t,
            align: 'center',
        });

        //  Материал
        const w6_col = 60 * mm;
        drawCell(doc, {
            x: x + w1_col + w2_col + w3_col + w4_col + w5_col,
            y: y,
            width: w6_col,
            height: ch,
            text: `Материал`,
            fontName: font,
            fontSize: fontSize,
            bold: true,
            border: b_tr,
            align: 'center',
        });

        y += ch;

        const panels = groupPartsByPos(data.items.panelMaterials);
        const array = [...panels];

        for (let i = 0; i < array.length; i++) {
            const item = array[i];
            let material = item.materialName;   //  Наименованеи материала
            drawCell(doc, { //  Колонка номера
                x: x,
                y: y + ch * i,
                width: w1_col,
                height: ch,
                text: `${i + 1}`,
                fontName: font,
                fontSize: fontSize,
                border: i == array.length - 1 ? b_bl : b_l,
                align: 'right',
            });

            drawCell(doc, { //  Колонка позиции
                x: x + w1_col,
                y: y + ch * i,
                width: w2_col,
                height: ch,
                text: `${item.prjPos}`,
                fontName: font,
                fontSize: fontSize,
                border: i == array.length - 1 ? b_b : b_cell,
                align: 'left',
            });

            drawCell(doc, { //  Колонка названия детали
                x: x + w1_col + w2_col,
                y: y + ch * i,
                width: w3_col,
                height: ch,
                text: `${item.name}${item.tkn ? ' (' + item.tkn + ')' : ''}`,
                fontName: font,
                fontSize: fontSize,
                border: i == array.length - 1 ? b_b : b_cell,
                align: 'left',
            });

            drawCell(doc, { //  Колонка количества
                x: x + w1_col + w2_col + w3_col,
                y: y + ch * i,
                width: w4_col,
                height: ch,
                text: `${item.count}`,
                fontName: font,
                fontSize: fontSize,
                border: i == array.length - 1 ? b_b : b_cell,
                align: 'right',
            });

            if (item.height && item.width) { // Для панелей
                drawCell(doc, { //  Колонка длины – width
                    x: x + w1_col + w2_col + w3_col + w4_col,
                    y: y + ch * i,
                    width: w5_col / 2,
                    height: ch,
                    text: `${item.width}`,
                    fontName: font,
                    fontSize: fontSize,
                    border: i == array.length - 1 ? b_b : b_cell,
                    align: 'center',
                    buttInfo: item.llww ? [item.llww[0], item.llww[1]] : [],
                });
                drawCell(doc, { //  Колонка длины – width
                    x: x + w1_col + w2_col + w3_col + w4_col + w5_col / 2,
                    y: y + ch * i,
                    width: w5_col / 2,
                    height: ch,
                    text: `${item.height}`,
                    fontName: font,
                    fontSize: fontSize,
                    border: i == array.length - 1 ? b_b : b_cell,
                    align: 'center',
                    buttInfo: item.llww ? [item.llww[2], item.llww[3]] : [],
                });
            };

            if (i > 0 && material === array[i - 1].materialName) {
                material = '';
            } else {
                material = array[i].materialName;
            };

            drawCell(doc, { //  Колонка материала
                x: x + w1_col + w2_col + w3_col + w4_col + w5_col,
                y: y + ch * i,
                width: w6_col,
                height: ch,
                text: `${material}`.substring(0, 48) + `${material.length ? '…' : ''}`,
                fontName: font,
                fontSize: fontSize,
                border: i == array.length - 1 ? b_br : b_r,
                align: 'left',
            });

        };
    };

    //  Функция создания листа
    function createDrawingSheet(doc, options = {}) {
        const {
            data,
            title = '',
            margin = 5 * mm,
            borderWidth = 1.5,
            orientation = 'landscape',
            isFirstPage = false,
            fontName = 'Arial',
            designation = '',
            sheet = '1',
            totalSheets = '1'
        } = options;

        if (!isFirstPage) {
            doc.addPage({
                size: 'A4',
                layout: orientation,
                margin: 0
            });
        }

        const pageWidth = doc.page.width;
        const pageHeight = doc.page.height;

        // 1. Основна рамка
        doc.save();
        doc.lineWidth(borderWidth)
            .rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2)
            .stroke('#000000');
        doc.restore();

        //  2. Штамп
        createDocStamp(doc, pageWidth, pageHeight, {
            data: data,
            title: title,
            designation: designation,
            sheet: sheet,
            totalSheets: totalSheets,
            margin: margin,
            font: fontName,
            fontSize: 8
        });

        //  3. Чертеж
        const drawWidth = (120 - 5) * mm;
        const drawHeight = (163 * mm);
        addImageFixedWidth(doc, {
            x: pageWidth - drawWidth - margin - 5 * mm,
            y: margin + (5 + 5) * mm,
            containerWidth: drawWidth,
            containerHeight: drawHeight,
            imagePath: path.join(
                settings.drawDIR,
                'png',
                data.drawName + '.png'
            )
        });

        //  4. Таблица деталей
        createAUPartsTable(doc, options);
    };

    //  Функция создания QR-кода
    async function addQRCode(doc, options = {}) {
        const {
            data,
            x,
            y,
            size = 20,
            color = '#000000',
            backgroundColor = '#FFFFFF',
            margin = 1,
            errorCorrectionLevel = 'M',
            debug = false,
            debugColor = '#E8F4FD',
            debugBorderColor = '#4A90D9'
        } = options;

        if (!data) {
            console.warn('⚠️ Нет данных для QR-кода');
            return null;
        }

        // Создаем временную папку, если её нет
        function ensureTempDir() {
            if (!fs.existsSync(TEMP_DIR)) {
                fs.mkdirSync(TEMP_DIR, { recursive: true });
                console.log('📁 Создана временная папка:', TEMP_DIR);
            }
        }

        /**
         * Генерирует уникальное имя файла
         */
        function generateTempFileName() {
            const timestamp = Date.now();
            const random = Math.floor(Math.random() * 10000);
            return `qr_${timestamp}_${random}.png`;
        }

        ensureTempDir();

        // Генерируем уникальное имя файла
        const fileName = generateTempFileName();
        const filePath = path.join(TEMP_DIR, fileName);

        try {
            // Генерируем QR-код и сохраняем в файл
            await QRCode.toFile(filePath, data, {
                errorCorrectionLevel: errorCorrectionLevel,
                margin: margin,
                color: {
                    dark: color,
                    light: backgroundColor
                },
                width: size * 4
            });

            // Проверяем, создался ли файл
            if (!fs.existsSync(filePath)) {
                throw new Error('Файл не создан');
            }

            // В режиме отладки рисуем контейнер
            if (debug) {
                doc.save();
                doc.fillColor(debugColor)
                    .rect(x, y, size, size)
                    .fill();
                doc.lineWidth(0.5)
                    .rect(x, y, size, size)
                    .stroke(debugBorderColor);
                doc.restore();
            }

            // Добавляем QR-код из файла
            doc.image(filePath, {
                x: x,
                y: y,
                width: size,
                height: size
            });

            // В режиме отладки добавляем информацию
            if (debug) {
                doc.save();
                doc.fontSize(5)
                    .font('Helvetica')
                    .fillColor('#666666')
                    .text(
                        `QR: ${data.length} симв.`,
                        x,
                        y + size + 1,
                        { width: size, align: 'center' }
                    );
                doc.restore();
            }

            // Удаляем временный файл
            try {
                fs.unlinkSync(filePath);
                if (debug) {
                    console.log(`🗑️ Удален временный файл: ${fileName}`);
                }
            } catch (unlinkErr) {
                // Если не удалось удалить - просто логируем
                console.warn(`⚠️ Не удалось удалить файл: ${fileName}`);
            }

            return {
                x: x,
                y: y,
                size: size,
                data: data
            };
        } catch (err) {
            console.error(`❌ Ошибка генерации QR-кода: ${err.message}`);

            // Если файл создался, но была ошибка - удаляем
            if (fs.existsSync(filePath)) {
                try {
                    fs.unlinkSync(filePath);
                } catch (e) { }
            }

            if (debug) {
                doc.save();
                doc.fillColor('#FFE8E8')
                    .rect(x, y, size, size)
                    .fill();
                doc.lineWidth(0.5)
                    .rect(x, y, size, size)
                    .stroke('#D94A4A');
                doc.fontSize(6)
                    .font('Helvetica')
                    .fillColor('#666666')
                    .text('❌ QR Error', x + 2, y + size / 2 - 3, {
                        width: size - 4,
                        align: 'center',
                        baseline: 'middle'
                    });
                doc.restore();
            }
            return null;
        }
    }

    //  Функция создания листа схемы сборки
    async function createUnionDrawingSheet(doc, options = {}) {
        const {
            data,
            title = 'Схема сборки',
            margin = 5 * mm,
            borderWidth = 1.5,
            orientation = 'landscape',
            isFirstPage = false,
            fontName = 'Arial',
            designation = '',
            sheet = '1',
            totalSheets = '1'
        } = options;

        if (!isFirstPage) {
            doc.addPage({ size: 'A4', layout: orientation, margin: 0 });
        };

        const pageWidth = doc.page.width;
        const pageHeight = doc.page.height;

        // 1. Основна рамка
        doc.save();
        doc.lineWidth(borderWidth)
            .rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2)
            .stroke('#000000');
        doc.restore();

        //  2. Штамп
        createDocUnionStamp(doc, pageWidth, pageHeight, {
            data: data,
            title: `Схема сборки на ${data.modelName} / ${designation}`,
            designation: designation,
            sheet: sheet,
            totalSheets: totalSheets,
            margin: margin,
            font: fontName,
            fontSize: 8
        });

        //  3. Чертеж
        const drawWidth = pageWidth - 4 * margin;
        const drawHeight = pageHeight - 60 * mm
        addImageFixedWidth(doc, {
            x: 2 * margin,
            y: margin + (24) * mm,
            containerWidth: drawWidth,
            containerHeight: drawHeight,
            imagePath: path.join(
                settings.drawDIR,
                'png',
                data.drawName + '.png'
            )
        });

        if (data.cloudLink) {
            await addQRCode(doc, {
                data: data.cloudLink,
                x: margin + 4 * mm,
                y: margin + 4 * mm,
                size: 20 * mm,
                debug: false
            });

        };
    };

    //  Функция сохранения документа
    function savePDF(doc, filePath) {
        return new Promise((resolve, reject) => {
            const writeStream = fs.createWriteStream(filePath);
            doc.pipe(writeStream);
            writeStream.on('finish', () => { resolve() });
            writeStream.on('error', (err) => { reject(err) });
            doc.end();
        });
    };

    //  Цикла создания сборочных чертежей по изделиям
    for (const array of assembly_array) {
        if (!array || array.length === 0) continue;

        const doc = new PDFDocument({
            size: 'A4',
            layout: 'landscape',
            margin: 0
        });

        // Регистрируем шрифты с поддержкой стилей
        try {
            doc.registerFont('Arial', 'C:/Windows/Fonts/arial.ttf');
            doc.registerFont('Arial-Bold', 'C:/Windows/Fonts/arialbd.ttf');
            doc.registerFont('Arial-Italic', 'C:/Windows/Fonts/ariali.ttf');
            doc.registerFont('Arial-BoldItalic', 'C:/Windows/Fonts/arialbi.ttf');
        } catch (err) {
            console.error('❌ Ошибка регистрации шрифтов:', err.message);
        };

        const prju = array[0];
        const docName = prju.prjName + settings.delimPrjName + prju.sign + '_Сборочные чертежи';
        const filePath = path.join(FOLDER, SW_FOLDER, `${docName}.pdf`);

        const pdfDir = path.join(FOLDER, SW_FOLDER);
        if (!fs.existsSync(pdfDir)) fs.mkdirSync(pdfDir, { recursive: true });

        //  Проверяем существование общего сборочного чертежа
        const sbImagePath = path.join(settings.drawDIR, 'png', array[0].modelDrawName + '.png');

        let unionDraw = 0;
        let totalPages = array.length;

        if (fs.existsSync(sbImagePath)) {
            unionDraw = 1;
            totalPages++;
            const sbData = array[0];
            // Запускаем создание общего сборочного чертежа
            await createUnionDrawingSheet(doc, {
                data: {
                    prjName: sbData.prjName,
                    drawName: sbData.modelDrawName,
                    modelName: sbData.modelName,
                    sign: sbData.sign,
                    cloudLink: sbData.cloudLink,
                },
                title: 'Схема сборки',
                margin: 5 * mm,
                orientation: 'landscape',
                isFirstPage: true,
                fontName: fontName,
                designation: sbData.sign,
                sheet: String(1),
                totalSheets: String(totalPages)
            });
        };

        for (let i = 0; i < array.length; i++) {
            const aUnit = array[i];
            createDrawingSheet(doc, {
                data: aUnit,
                title: aUnit.drawName,
                margin: 5 * mm,
                orientation: 'landscape',
                isFirstPage: unionDraw === 0,
                fontName: fontName,
                designation: aUnit.sign || '',
                sheet: String(i + unionDraw + 1),
                totalSheets: String(totalPages)
            });
        };
        await savePDF(doc, filePath);
    };
    console.log('✅ Все PDF документы созданы!');
};

//=== Стлизация таблиц =======================================================//

//  Функция стилизации строк таблицы
function setRowTableStyle(row, start, tclmns, rh, r_type = 'main', font_style) {
    const a = start;            //  Начальная колонка таблицы
    const b = a + tclmns - 1;   //  Конечная колонка таблицы
    const sb = 'medium';
    const st = 'thin';
    row.height = setRowHeght(rh);
    let bc = {    //  Граница основной ячейки
        left: { style: st }, right: { style: st },
        top: { style: st }, bottom: { style: st }
    };
    let blc = {   //  Граница левой ячейки
        left: { style: sb }, right: { style: st },
        top: { style: st }, bottom: { style: st }
    };
    let brc = {  //  Граница правой ячейки
        left: { style: st }, right: { style: sb },
        top: { style: st }, bottom: { style: st }
    };
    switch (r_type) {
        case 'empty':
            bc = {
                left: { style: 'none' }, right: { style: 'none' },
                top: { style: sb }, bottom: { style: sb }
            };
            blc = {
                left: { style: sb }, right: { style: 'none' },
                top: { style: sb }, bottom: { style: sb }
            };
            brc = {
                left: { style: 'none' }, right: { style: sb },
                top: { style: sb }, bottom: { style: sb }
            };
            styleCellRange(row, a, b, { border: bc, font: font_style });
            break;
        case 'header':
            bc = {
                left: { style: st }, right: { style: st },
                top: { style: sb }, bottom: { style: sb }
            };
            blc = {
                left: { style: sb }, right: { style: st },
                top: { style: sb }, bottom: { style: sb }
            };
            brc = {
                left: { style: st }, right: { style: sb },
                top: { style: sb }, bottom: { style: sb }
            };
            styleCellRange(row, a, b, {
                border: bc,
                font: font_style,
                alignment: { vertical: 'middle', horizontal: 'center' }
            });
            break;
        case 'end':
            bc = {
                left: { style: st }, right: { style: st },
                top: { style: st }, bottom: { style: sb }
            };
            blc = {
                left: { style: sb }, right: { style: st },
                top: { style: st }, bottom: { style: sb }
            };
            brc = {
                left: { style: st }, right: { style: sb },
                top: { style: st }, bottom: { style: sb }
            };
            styleCellRange(row, a, b, { border: bc, font: font_style });
            break;
        case 'main':
            styleCellRange(row, a, b, { border: bc, font: font_style });
            break;
        default:
            break;
    }
    styleCellRange(row, a, a, { border: blc });
    styleCellRange(row, b, b, { border: brc });
};

// Функция для стилизации диапазона ячеек
function styleCellRange(row, a, b, styles) {
    for (let col = a; col <= b; col++) {
        const cell = row.getCell(col);
        Object.assign(cell, styles);
    };
};

//  Установка высоты строки (пересчет)
function setRowHeght(num) {
    return Math.round(num / 0.75 * 10) / 10;
};
//============================================================================//

//#endregion

async function main() {

    let ind = 0;    //  Индекс файла Проекта
    let count = 0;  //  Счетчик обработанных файлов
    let prj_array = readProjectFilesData(PROJECT_FILE);
    let assembly_array = [];

    //  Функция рекурсивного обхода массива файлов Проекта
    async function proccessNextFile() {
        //  Путь до текущего файла проекта
        const filepath = prj_array[ind].filepath;

        // Вывод прогресса обработки файлов Проекта
        Action.Hint =
            `обработка файла ${ind + 1} из ${prj_array.length} – ${filepath}`;

        //  Проверяем расширение текущего файла
        if (extensions.indexOf(path.extname(filepath)) >= 0) {

            //  Загружаем текущую Модель Проекта
            if (Action.LoadModel(filepath)) {

                //  Рекурсивный обход текущей модели
                forEachInList(Model, callbackFunc, prj_array[ind]);

                //  Поиск дополнительных материалов к Модели
                findAdditionalMaterials(Model, prj_array[ind]);

                //  Добавляем наименование заказа в данные
                prj_array[ind].modelName = Article.Name;
                count++;

                // Проверяем есть ли свойство CloudLink у модели
                prj_array[ind].cloudLink = getUserProderty(Model, 'CloudLink');

                //  Обход блоков сборочных единиц 1 уровня
                let arr = getAssemblyName(Model.AsList(), prj_array[ind]);
                assembly_array.push(arr);
            };
        };

        ind++;
        if (ind < prj_array.length) {
            // Обработка следующего файла
            Action.AsyncExec(proccessNextFile);
        } else {

            Action.Hint = `получаем данные из Базы материалов...`;
            //  1. Получаем данные из Базы материалов
            let dbNames = await getDBMaterialInfo();

            //  2. Дополняем данные с модели данными из Базы материалов
            await unionMaterialData(prj_array, dbNames);

            //  3. Получаем данные по классам из Базы материалов
            let mat_classes = {};
            if (ID_MAT_ARRAY.length) {
                mat_classes = await getDBCLassMaterialInfo(ID_MAT_ARRAY);
            };

            //  4. Формируем обобщенные данные по моделям проекта
            setProjectEstimateData(mat_classes, prj_array);
            //  Сводные данные по проекту в целом
            const agrMat = aggregateMaterials(prj_array);

            //  5. Формируем файлы проекта
            try {
                Action.Hint = `Создание сборочных чертежей...`;

                //  Формируем картинки сборочных чертежей
                const drawNames = assembly_array.flatMap(item => item.map(el => el.drawName));

                //  Получаем имена сборочных чертежей модели.
                const modelDrawNames = assembly_array.map(item => item[0]?.modelDrawName);
                drawNames.push(...modelDrawNames);

                const pictureBool = await convertPDFtoPNG(drawNames, settings.drawDIR);
                if (!pictureBool) console.log('Картинки сборочных чертежей не созданы.');

                await createPDFAssemblyDrawings(assembly_array, settings);
            } catch (e) {
                errFinish('Ошибка создания сборочных чертежей: ' + e.message);
            };

            //Action.Finish();    //  Временная заглушка

            try {
                Action.Hint = `Сохраняем калькуляцию проекта...`;
                await createEsimateExcelFile(prj_array);
            } catch (e) {
                errFinish('Ошибка создания файла калькуляции проекта: ' + e.message);
            };

            try {
                Action.Hint = `Сохраняем спецификацию для загрузки в 1С...`;
                await createSpecificationForImport(agrMat);
            } catch (e) {
                errFinish('Ошибка файла загрузки для 1С: ' + e.message);
            };

            try {
                Action.Hint = `Сохраняем спецификацию проекта...`;
                await createSpecificationProjectFile(agrMat, prj_array, settings);
            } catch (e) {
                errFinish('Ошибка создания файла спецификации: ' + e.message);
            };

            //  Завершение обработки (выход из скрипта)
            //alert(`Обработано ${count} файлов из ${prj_array.length}`);
            Action.Finish();
        };
    };

    //  Запуск функции рекурсивного обхода файлов Проекта
    if (prj_array.length > 0) proccessNextFile();

};

main();
Action.Continue();
/******************************************************************************/