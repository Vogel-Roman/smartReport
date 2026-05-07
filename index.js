/******************************************************************************/
/***       Скрипт на формирование отчетов из Проектов Базис в Excel         ***/
/***                         SmartWood Reports v1.0                         ***/
/******************************************************************************/

//#region Инициализация

const fs = require('fs');
const path = require('path');
const firebird = require('node-firebird');
const ExcelJS = require('exceljs');
const { XMLParser } = require('fast-xml-parser');

//  Массивы наименований материалов для поиска в БД
const BOARD_ARRAY = [];         //  Массив названий листовых материалов
const BUTT_ARRAY = [];          //  Массив названий кромочных материалов
const PROFILE_ARRAY = [];       //  Массив названий профильных материалов
const FURNITURE_ARRAY = [];     //  Массив названий фурнитуры

const ID_MAT_ARRAY = [];        //  Массив ID материалов

// Допуск для сравнения с нулем
const EPS = 1e-10;

//  Расширения обрабатываемых файлов
const extensions = ['.fr3d', '.b3d'];

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

            if (isPointInBounds(endInPanel, panel.GMin, panel.GMax)) {
                result.push({
                    depth: round(hole.Depth, 2),
                    diameter: hole.Radius * 2,
                    depth: round(hole.Depth, 1),
                    drillMode: hole.DrillMode,
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

//#endregion

//#region Функции обработки объектов Модели

//  Функция обработки данных панели
function panelProcessing(panel, modelData) {
    const material = getMaterialName(panel.MaterialName);

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

    modelData.data.panelMaterials.push({
        name: panel.Name,               //  Имя панели
        material: panel.MaterialName,   //  Материал панели
        materialID: undefined,          //  ID Материала (DB)
        materialName: material[0],      //  Имя материала панели
        materialArticle: material[1],   //  Артикул материала панели
        materialSyncExternal: "",       //  Код синхронизации материала (DB)
        materialPrice: 0,               //  Цена из базы данных (DB)
        materialUnit: "",               //  Единица измерения (DB)
        materialWidth: 0,               //  Длина листа (DB) мм
        materilaHeight: 0,              //  Ширина листа (DB) мм
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
            materialName: mat[0],           //  Имя материала панели
            materialArticle: mat[1],        //  Артикул материала панели
            materialSyncExternal: "",       //  Код синхронизации материала (DB)
            materialPrice: 0,               //  Цена из базы данных (DB)
            materialUnit: "",               //  Единица измерения (DB)
            materialWidth: 0,               //  Длина листа (DB) мм
            materilaHeight: 0,              //  Ширина листа (DB) мм
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

//#region Функции обработки файлов Проекта

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
                pnl.materilaHeight = pnl_m.width;
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
                prfl.materilaHeight = prfl_m.width;
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
                    materilaHeight: item.materilaHeight,
                    materialTkn: item.materialTkn,
                    area: 0,
                    contourLength: 0,
                    drillInfo: []
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
                            drillMode: drill.drillMode,
                            count: item.prjCount || 1
                        });
                    }
                });
            }
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

//#endregion

//#region Формирование документов

// Функция для стилизации диапазона ячеек
function styleCellRange(row, startCol, endCol, styles) {
    for (let col = startCol; col <= endCol; col++) {
        const cell = row.getCell(col);
        Object.assign(cell, styles);
    };
};

function getColumnLetter(colNumber) {
    let letter = '';
    while (colNumber > 0) {
        colNumber--;
        letter = String.fromCharCode(65 + (colNumber % 26)) + letter;
        colNumber = Math.floor(colNumber / 26);
    }
    return letter;
};

//  Функция задания параметров страницы
function setWorksheetSettings(worksheet) {
    worksheet.pageSetup = {
        // Ориентация страницы
        orientation: 'portrait',  // 'portrait' | 'landscape'

        // Поля страницы (в ДЮЙМАХ! 1 дюйм = 2.54 см)
        margins: {
            top: 0.5,      // Верхнее поле
            bottom: 0.5,   // Нижнее поле
            left: 0.39,     // Левое поле
            right: 0.39,    // Правое поле
            header: 0.3,   // Отступ для колонтитула сверху
            footer: 0.3    // Отступ для колонтитула снизу
        },

        // Масштабирование
        fitToPage: true,    // Вписать в страницу
        fitToWidth: 1,      // Вписать по ширине (1 страница)
        fitToHeight: 0,     // По высоте (0 = автоматически)

        // Альтернатива — масштаб в процентах
        // scale: 85,

        // Центрирование на странице
        // horizontalCentered: true,
        // verticalCentered: false,

        // Сетка и заголовки
        // showGridLines: false,        // Печатать сетку
        // showRowColHeaders: false,    // Печатать заголовки строк/столбцов

        // Повторять строки/колонки на каждой странице
        // printTitlesRow: '6:6',      // Повторять 6-ю строку (заголовок)
        // printTitlesColumn: 'A:B',

        // Область печати
        // printArea: 'A1:G50',

        // Колонтитулы
        // headerFooter: {
        //     oddHeader: '&C&BСпецификация деталей',
        //     oddFooter: '&LСтраница &P из &N&RДата: &D',
        //     evenHeader: '&C&BСпецификация деталей',
        //     evenFooter: '&LСтраница &P из &N&RДата: &D'
        // },

        // Порядок страниц
        // pageOrder: 'downThenOver',   // 'overThenDown'

        // Номер первой страницы
        // firstPageNumber: 1,

        // Качество печати
        // blackAndWhite: false,
        // draft: false,

        // Количество копий
        // copies: 1
    };
};

//  Функуция создания файла Сметы
async function createEsimateExcelFile(prj_arr) {

    //#region Настройки стилей

    const row_height = 14;
    const font_size = 9;
    const doc_font = { name: 'Arial', size: font_size + 7, bold: true };
    const h_font = { name: 'Arial', size: font_size, bold: true };
    const r_font = { name: 'Arial', size: font_size - 1, bold: false };

    if (!settings.estimate.classes)
        errFinish("Ошибка файла настроек - classes");
    if (!settings.estimate.fillColor)
        errFinish("Ошибка файла настроек - fillColor");
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

    //  Форматирование чисел
    const f_format = '#,##0.00';
    const n_format = '# ##0';

    const algn_left = { indent: 1, horizontal: 'left', vertical: 'middle' };
    const algn_right = { indent: 1, horizontal: 'right', vertical: 'middle' };
    const algn_center = { horizontal: 'center', vertical: 'middle' };

    //#endregion

    //#region Стили строк таблицы

    //  Установка высоты строки (пересчет)
    function setRowHeght(num) {
        return Math.round(num / 0.75 * 10) / 10;
    };

    //  Стиль заглавной строки таблицы
    function setHeaderRowTableStyle(row, a, b) {
        const s_b = 'medium';
        const s_t = 'thin';
        row.height = setRowHeght(row_height);
        const algn = { vertical: 'middle', horizontal: 'center' }
        const brd_cell = {        //  Граница ячейки
            left: { style: s_t }, right: { style: s_t },
            top: { style: s_b }, bottom: { style: s_b }
        };
        const brd_left_cell = {   //  Граница левой ячейки
            left: { style: s_b }, right: { style: s_t },
            top: { style: s_b }, bottom: { style: s_b }
        };
        const brd_right_cell = {  //  Граница правой ячейки
            left: { style: s_t }, right: { style: s_b },
            top: { style: s_b }, bottom: { style: s_b }
        };
        let style = { border: brd_cell, font: h_font, alignment: algn, };
        //  Применяем стили к диапазону ячеек a…b
        styleCellRange(row, a, b, style);
        styleCellRange(row, a, a, { border: brd_left_cell });
        styleCellRange(row, b, b, { border: brd_right_cell });
    };

    //  Стиль строки таблицы
    function setRowTableStyle(row, a, b) {
        const s_b = 'medium';
        const s_t = 'thin';
        row.height = setRowHeght(row_height - 1);
        const brd_cell = {        //  Граница ячейки
            left: { style: s_t }, right: { style: s_t },
            top: { style: s_t }, bottom: { style: s_t }
        };
        const brd_left_cell = {   //  Граница левой ячейки
            left: { style: s_b }, right: { style: s_t },
            top: { style: s_t }, bottom: { style: s_t }
        };
        const brd_right_cell = {  //  Граница правой ячейки
            left: { style: s_t }, right: { style: s_b },
            top: { style: s_t }, bottom: { style: s_t }
        };
        let style = { border: brd_cell, font: r_font };
        //  Применяем стили к диапазону ячеек a…b
        styleCellRange(row, a, b, style);
        styleCellRange(row, a, a, { border: brd_left_cell });
        styleCellRange(row, b, b, { border: brd_right_cell });
    };

    //  Стиль последней строки таблицы
    function setEndRowTableStyle(row, a, b) {
        const s_b = 'medium';
        row.height = setRowHeght(row_height);
        const brd_cell = {        //  Граница ячейки
            left: { style: 'none' }, right: { style: 'none' },
            top: { style: s_b }, bottom: { style: s_b }
        };
        const brd_left_cell = {   //  Граница левой ячейки
            left: { style: s_b }, right: { style: 'none' },
            top: { style: s_b }, bottom: { style: s_b }
        };
        const brd_right_cell = {  //  Граница правой ячейки
            left: { style: 'none' }, right: { style: s_b },
            top: { style: s_b }, bottom: { style: s_b }
        };
        let style = { border: brd_cell, font: h_font };
        //  Применяем стили к диапазону ячеек a…b
        styleCellRange(row, a, b, style);
        styleCellRange(row, a, a, { border: brd_left_cell });
        styleCellRange(row, b, b, { border: brd_right_cell });
    };

    //#endregion

    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.author;
    workbook.created = new Date();

    // Цикл создания вкладок
    prj_arr.forEach(model => {
        const estimate = model.estimate_data;   //  Объект данных Сметы
        const name = model.name;                //  Имя изделия
        const sign = model.sign;                //  Обозначение в Проекте
        const sheet_name = `${sign}_${name}`.substring(0, 31);
        const worksheet = workbook.addWorksheet(sheet_name);
        setWorksheetSettings(worksheet);

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

        //  Количество колонок документа
        const col_count = worksheet.columns.length;

        // Таблица (header row)
        const headerRowInd = 6;     //  Индекс верхней строки таблицы
        const scol = 2;             //  Начальная колонка
        const headerRow = worksheet.getRow(headerRowInd);

        //#region Шапка документа

        //  Закрепляем строки до headerRowInd
        worksheet.views = [{ state: 'frozen', ySplit: headerRowInd }];

        worksheet.getRow(2).height = setRowHeght(24);
        worksheet.getRow(3).height = setRowHeght(5);
        styleCellRange(worksheet.getRow(3), scol, col_count, {
            border: {
                left: { style: 'none' }, right: { style: 'none' },
                top: { style: 'medium' }, bottom: { style: 'none' }
            }
        });

        //  Строка суммы (дуликат)
        const top_total_row = worksheet.getRow(4);

        worksheet.getRow(headerRowInd - 1).height = setRowHeght(5);

        const docNameRow = worksheet.getRow(2);
        docNameRow.alignment = { horizontal: 'left', vertical: 'middle' };
        docNameRow.getCell(scol).font = doc_font;
        docNameRow.getCell(scol).value =
            `Расчет изделия ${model.modelName} проекта ${PROJECT_NAME}`;

        //#endregion

        //#region Шапка основной таблицы

        // Устанавливаем значения заголовков вручную (начиная с B)
        const headers = [
            '№', 'Артикул', 'Наименование',
            'Кол-во', 'Ед.', 'Цена', 'Сумма'
        ];

        for (let col = scol; col <= headers.length + scol - 1; col++) {
            const cell = headerRow.getCell(col);
            cell.value = headers[col - scol];  // Заголовок
        };
        setHeaderRowTableStyle(headerRow, scol, col_count);

        const v_headers = ['k', 'Цена (DB)', 'Сумма (DB)', 'ВД'];
        for (let col = 0; col < v_headers.length; col++) {
            const cell = headerRow.getCell(scol + 8 + col);
            cell.value = v_headers[col];  // Заголовок
        };
        let end_col = col_count + scol + v_headers.length - 1;
        setHeaderRowTableStyle(headerRow, scol + 8, end_col);

        //#endregion

        //  Индекс строки таблицы сразу после заголовка
        let ind = headerRowInd + 1;

        //  Колонка для коэффициента
        const coef_ltr = getColumnLetter(scol + 8);
        const price_ltr = getColumnLetter(scol + 9);
        const sum_ltr = getColumnLetter(scol + 10);
        const res_ltr = getColumnLetter(scol + 11);

        //  Тело таблицы
        let row_counter = 1;
        classes.forEach(key => {

            const startRowInd = ind;
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
            //const items = estimate[key].items;

            const k = c_koef[key] ? c_koef[key] : 0;

            //if (!items.length) return;
            for (let i = 0; i < items.length; i++) {
                const item = items[i];

                //  Текущая строка
                const rn = ind + i;
                const row = worksheet.getRow(rn);

                //  Стиль строки основной таблицы
                setRowTableStyle(row, scol, col_count);

                //  Стиль строки расчетной таблицы
                setRowTableStyle(row, scol + 8, scol + 11);

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
                    formula: `${price_ltr}${rn}*$${coef_ltr}$${startRowInd}`
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
                //  Устанавливаем ширины колонок
                worksheet.getColumn(`${coef_ltr}`).width = 10;
                worksheet.getColumn(`${price_ltr}`).width = 12;
                worksheet.getColumn(`${sum_ltr}`).width = 12;
                worksheet.getColumn(`${res_ltr}`).width = 12;

                //  Коэффициент наценки класса
                const coefCell = worksheet.getCell(`${coef_ltr}${startRowInd}`);
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
                //------------------------------------------------------------//


            };

            ind += items.length;
            const endRowInd = ind - 1;

            //  Объединяем все ячейки в этой колонке для диапазона строк
            const cellAdress = `${coef_ltr}${startRowInd}`;
            worksheet.mergeCells(
                `${coef_ltr}${startRowInd}:${coef_ltr}${endRowInd}`
            );
        });

        //#region Подвал таблицы

        const row = worksheet.getRow(ind);
        setEndRowTableStyle(row, scol, col_count);

        const end_res_col = scol + 8;
        setEndRowTableStyle(row, end_res_col, end_col);

        const tsm_ltr = getColumnLetter(scol + 6);
        const tres_sum_ltr = getColumnLetter(end_res_col + 2);
        const tres_vd_ltr = getColumnLetter(end_res_col + 3);

        //  текст ИТОГО основной таблицы
        const total_row_res_text = row.getCell(end_res_col + 1);
        total_row_res_text.value = "Итого:";
        total_row_res_text.font.size = font_size;
        total_row_res_text.alignment = algn_right;
        total_row_res_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        //  Текст ИТОГО вспомогательной таблицы
        const total_row_text = row.getCell(scol + 5);
        total_row_text.value = "Итого:";
        total_row_text.font.size = font_size;
        total_row_text.alignment = algn_right;
        total_row_text.border = {
            left: { style: 'none' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        //  Ячейка суммы
        const total_row_val = row.getCell(scol + 6);
        total_row_val.value = {
            formula: `SUM(${tsm_ltr}${headerRowInd + 1}:${tsm_ltr}${ind - 1})`
        };
        total_row_val.numFmt = n_format;
        total_row_val.alignment = algn_right;

        //  Ячейка итоговой суммы из Базы
        const total_row_res_val = row.getCell(end_res_col + 2);
        total_row_res_val.value = {
            formula: `SUM(${tres_sum_ltr}${headerRowInd + 1}:${tres_sum_ltr}${ind - 1})`
        };
        total_row_res_val.numFmt = n_format;
        total_row_res_val.alignment = algn_right;
        total_row_res_val.border = {
            left: { style: 'thin' }, right: { style: 'thin' },
            top: { style: 'medium' }, bottom: { style: 'medium' }
        };

        //  Ячейка итоговой суммы ВД
        const total_row_res_vd_val = row.getCell(end_res_col + 3);
        total_row_res_vd_val.value = {
            formula: `SUM(${tres_vd_ltr}${headerRowInd + 1}:${tres_vd_ltr}${ind - 1})`
        };
        total_row_res_vd_val.numFmt = n_format;
        total_row_res_vd_val.alignment = algn_right;
        //#endregion

        //  Дубликаты расчетов в шапке документа
        const d_total_text = top_total_row.getCell(scol + 5);
        d_total_text.value = "Итого:";
        d_total_text.font = h_font;
        d_total_text.font.size = font_size;
        d_total_text.alignment = algn_right;

        const d_total_val = top_total_row.getCell(scol + 6);
        d_total_val.value = {
            formula: `SUM(${tsm_ltr}${headerRowInd + 1}:${tsm_ltr}${ind - 1})`
        };
        d_total_val.font = h_font;
        d_total_val.font.size = font_size;
        d_total_val.numFmt = n_format;
        d_total_val.alignment = algn_right;

        const d_totalres_text = top_total_row.getCell(end_res_col + 1);
        d_totalres_text.value = "Итого:";
        d_totalres_text.font = h_font;
        d_totalres_text.font.size = font_size;
        d_totalres_text.alignment = algn_right;

        const d_totalres_val = top_total_row.getCell(end_res_col + 2);
        d_totalres_val.value = {
            formula: `SUM(${tres_sum_ltr}${headerRowInd + 1}:${tres_sum_ltr}${ind - 1})`
        };
        d_totalres_val.font = h_font;
        d_totalres_val.font.size = font_size;
        d_totalres_val.numFmt = n_format;
        d_totalres_val.alignment = algn_right;

        const d_totalvd_val = top_total_row.getCell(end_res_col + 3);
        d_totalvd_val.value = {
            formula: `SUM(${tres_vd_ltr}${headerRowInd + 1}:${tres_vd_ltr}${ind - 1})`
        };
        d_totalvd_val.font = h_font;
        d_totalvd_val.font.size = font_size;
        d_totalvd_val.numFmt = n_format;
        d_totalvd_val.alignment = algn_right;

        //  Получаем номер последней строки, где есть данные
        //  Устанавливаем область печати: колонки A-scm, все строки с данными
        worksheet.pageSetup.printArea = `A1:${tsm_ltr}${worksheet.rowCount}`;
    });

    // Сохранение документа
    const fileName = `${sanitizeFileName(
        PROJECT_NAME + "_Калькуляция проекта"
    )}.xlsx`;
    const filePath = path.join(FOLDER, fileName);
    await workbook.xlsx.writeFile(filePath);
};

//#endregion

/****************************** ОСНОВНАЯ ФУНКЦИЯ ******************************/
async function main() {

    let ind = 0;    //  Индекс файла Проекта
    let count = 0;  //  Счетчик обработанных файлов

    let prj_array = readProjectFilesData(PROJECT_FILE);

    //  Функция рекурсивного обхода массива файлов Проекта
    async function processNextFile() {
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
                //console.log(prj_array[ind].name);
                count++;
            };
        };

        ind++;
        if (ind < prj_array.length) {
            // Обработка следующего файла
            Action.AsyncExec(processNextFile);
        } else {
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

            //  5. Формируем файлы спецификаций деталей

            //  6. Формируем файл Сметы проекта
            //await createEsimateExcelFile(prj_array);
            //  Завершение обработки (выход из скрипта)
            //console.log(JSON.stringify(prj_array, null, 2));
            console.log(`Обработано ${count} файлов из ${prj_array.length}`);
            Action.Finish();
        };
    };

    //  Запуск функции рекурсивного обхода файлов Проекта
    if (prj_array.length > 0) processNextFile();

    //console.log(JSON.stringify(prj_array, null, 2));
    //Action.Finish();
};

main();
Action.Continue();
/******************************************************************************/