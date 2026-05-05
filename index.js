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

// Функция поиска отверстий принадлежжащих панели
function findPanelHolesList(panel) {

    //  Функция вычисления конечной точки отверстия
    function getHoleEndPoint(hole, fastener, panel) {
        // 1. Вычисляем конец отверстия в локальной системе фурнитуры
        const dir = hole.Direction;
        const depth = hole.Depth;

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
    };

};

//#endregion

//#region Функции обработки объектов Модели

//  Функция обработки данных панели
function panelProcessing(panel, modelData) {
    const material = getMaterialName(panel.MaterialName);

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
    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.profileMaterial;
    if (excludeMaterial.includes(profile.MaterialName)) return;

    //const material = getMaterialName(profile.MaterialName);

};

//  Функция обработки данных фурнитуры
function furnitureProcessing(fastener, modelData) {
    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.furnitureMaterial;
    if (excludeMaterial.includes(fastener.MaterialName)) return;

    //  !!!! Тут нужна проверка на составную фурнитуру !!!!!!
    //const material = getMaterialName(fastener.MaterialName);
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

        //  Надо доработать получения классов материалов
        //  Строка запроса
        const sqlString = `
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
            INNER JOIN MATERIAL_ADVANCE AS ma ON m.ID_M = ma.ID_M
            LEFT JOIN MEASURE AS me ON m.ID_MS = me.ID_MS
        WHERE
            m.NAME_MAT IN (${matnames.map(() => '?').join(',')})
        `;
        //  Запрос в БД
        const result_db = await executeQuery(sqlString, options, matnames);

        const result = [
            [], //  Листовые материалы
            [], //  Кромочные материалы
            [], //  Погонные материалы
            []  //  Фурнитура
        ];

        //  Сортировка результата по группам
        result_db.forEach(elem => {
            if (BOARD_ARRAY.includes(elem.name_mat)) result[0].push(elem);
            if (BUTT_ARRAY.includes(elem.name_mat)) result[1].push(elem);
            if (PROFILE_ARRAY.includes(elem.name_mat)) result[2].push(elem);
            if (FURNITURE_ARRAY.includes(elem.name_mat)) result[3].push(elem);

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
                    contourLength: 0
                };
            };
            //  Площадь деталей и длина контуров деталей
            acc[key].area += item.area * item.prjCount || 0;
            acc[key].contourLength += item.contourLength * item.prjCount || 0;
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

        addClassItems(board_acc);
        addClassItems(butt_acc);
        addClassItems(prfl_acc);
        //addClassItems(furn_acc);

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


//  Функуция создания файла Сметы
async function createEsimateExcelFile(prj_arr) {

    const row_height = 14;
    const font_size = 9;

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

    //  Стиль заглавной строки таблицы
    function headerRowTableStyle(row, start, end) {
        let rowStyle = {
            border: {
                top: { style: 'medium' },
                left: { style: 'thin' },
                bottom: { style: 'medium' },
                right: { style: 'thin' }
            },
            alignment: { vertical: 'middle', horizontal: 'center' },
            font: {
                name: 'Arial',
                size: font_size,
                bold: true
            }
        };
        let height = row_height;
        row.height = Math.round(height / 0.75 * 10) / 10;

        styleCellRange(row, start, end, rowStyle);

        //  Первая ячейка
        styleCellRange(row, start, start, {
            border: {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'thin' }
            }
        });

        //  Последняя ячейка
        styleCellRange(row, end, end, {
            border: {
                top: { style: 'medium' },
                left: { style: 'thin' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            }
        });
    };

    //Стиль строки таблицы
    function rowTableStyle(row, start, end) {
        let rowStyle = {
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            },
            alignment: { vertical: 'middle' },
            font: {
                name: 'Arial',
                size: font_size - 1,
                bold: false
            }
        };
        let height = row_height - 0.5;
        row.height = Math.round(height / 0.75 * 10) / 10;

        styleCellRange(row, start, end, rowStyle);

        //  Первая ячейка
        styleCellRange(row, 2, 2, {
            border: {
                top: { style: 'thin' },
                left: { style: 'medium' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        });

        //  Последняя ячейка
        styleCellRange(row, end, end, {
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'medium' }
            }
        });
    };

    //  Стиль последней строки таблицы
    function endRowTableStyle(row, start, end) {
        let rowStyle = {
            border: {
                top: { style: 'medium' },
                left: { style: 'none' },
                bottom: { style: 'medium' },
                right: { style: 'none' }
            },
            font: {
                name: 'Arial',
                size: font_size - 1,
                bold: true
            }
        };
        let height = row_height;
        row.height = Math.round(height / 0.75 * 10) / 10;

        styleCellRange(row, start, end, rowStyle);

        //  Первая ячейка
        styleCellRange(row, 2, 2, {
            border: {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'none' }
            }
        });

        //  Последняя ячейка
        styleCellRange(row, end, end, {
            border: {
                top: { style: 'medium' },
                left: { style: 'none' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            }
        });
    };

    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.author;
    workbook.created = new Date();

    const c_koef = settings.estimate.classes;

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
            { width: 0.67 },    //  A Отступ
            { width: 5.42 },    //  B Номер 
            { width: 18 },      //  Артикул
            { width: 60 },      //  C Наименование 78
            { width: 8 },       //  D Количество
            { width: 8 },       //  E Ед. изм.
            { width: 12 },      //  F Цена
            { width: 12 }       //  G Сумма
        ];

        //  Количество колонок документа
        const col_count = worksheet.columns.length;

        // Таблица (header row)
        const headerRowInd = 6;     //  Индекс верхней строки таблицы
        const scol = 2;         //  Начальная колонка
        const headerRow = worksheet.getRow(headerRowInd);


        // Устанавливаем значения заголовков вручную (начиная с B)
        const headers = [
            '№', 'Артикул', 'Наименование',
            'Кол-во', 'Ед.', 'Цена', 'Сумма'
        ];
        for (let col = scol; col <= headers.length + scol - 1; col++) {
            const cell = headerRow.getCell(col);
            cell.value = headers[col - scol];  // Заголовок
        };
        headerRowTableStyle(headerRow, scol, col_count);

        const v_headers = ['k', 'Цена', 'Сумма', 'Доход'];
        for (let col = 0; col < v_headers.length; col++) {
            const cell = headerRow.getCell(scol + 8 + col);
            cell.value = v_headers[col];  // Заголовок
            console.log(v_headers[col]);

        };
        headerRowTableStyle(headerRow, scol + 8, col_count + scol + v_headers.length - 1);

        let ind = headerRowInd + 1;

        //  Колонка для коэффициента
        const coeffColLetter = getColumnLetter(scol + 8);
        const dbPriceLetter = getColumnLetter(scol + 9);
        const dbSumLetter = getColumnLetter(scol + 10);
        const vdLetter = getColumnLetter(scol + 11);

        let row_counter = 1;
        classes.forEach(key => {

            const startRowInd = ind;
            const items = estimate[key].items;
            const k = c_koef[key] ? c_koef[key] : 0;

            if (!items.length) return;
            for (let i = 0; i < items.length; i++) {
                const item = items[i];

                //  Текущая строка
                const rn = ind + i;
                const row = worksheet.getRow(rn);
                rowTableStyle(row, scol, col_count);

                //  Выбор измеряемой величины
                let val = "area";
                if (item.length) val = "length";
                if (item.count) val = "count";

                //  Колонка цены т количества
                const cpl = getColumnLetter(scol + 3);
                const cvl = getColumnLetter(scol + 5);

                worksheet.getCell(`${coeffColLetter}${startRowInd}`).value = k;

                //  Заполняем ячейки строки
                row.getCell(scol + 0).value = row_counter++;
                row.getCell(scol + 1).value = item.materialArticle;
                row.getCell(scol + 2).value = item.materialName;
                row.getCell(scol + 3).value = item[val];
                row.getCell(scol + 4).value = item.materialUnit;
                // row.getCell(scol + 5).value = item.materialPrice;

                //------------------------------------------------------------//
                //  Цена материала из базы
                const dbPriceCell = worksheet.getCell(`${dbPriceLetter}${rn}`);
                dbPriceCell.value = item.materialPrice;
                dbPriceCell.font = {
                    name: 'Arial',
                    size: font_size - 1,
                    bold: false
                };
                dbPriceCell.numFmt = '#,##0.00';
                dbPriceCell.alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };
                dbPriceCell.border = {
                    top: { style: 'thin' },
                    left: { style: 'medium' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };

                const dbSumCell = worksheet.getCell(`${dbSumLetter}${rn}`)

                //  Сумма материала из базы
                dbSumCell.value = {
                    formula: `${dbPriceLetter}${rn}*${cpl}${rn}`
                };
                dbSumCell.font = {
                    name: 'Arial',
                    size: font_size - 1,
                    bold: false
                };
                dbSumCell.numFmt = '#,##0.00';
                dbSumCell.alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };
                dbSumCell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };

                //  Разница между суммой с наценкой и без
                const valCell = worksheet.getCell(`${vdLetter}${rn}`)
                valCell.value = {
                    formula: `${cpl}${rn}*${cvl}${rn}-${dbSumLetter}${rn}`
                };
                valCell.font = {
                    name: 'Arial',
                    size: font_size - 1,
                    bold: false
                };
                valCell.numFmt = '#,##0.00';
                valCell.alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };
                valCell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'medium' }
                };

                //------------------------------------------------------------//

                //  Цена
                row.getCell(scol + 5).value = {
                    formula: `${dbPriceLetter}${rn}*${coeffColLetter}${startRowInd}`
                };

                //  Сумма
                row.getCell(scol + 6).value = {
                    formula: `${cvl}${rn}*${cpl}${rn}`
                };

                //  Форматирование ячеек

                //  Ячейка номера строки
                row.getCell(scol + 0).alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };

                //  Ячейка названия материалы
                row.getCell(scol + 1).alignment = {
                    indent: 1,
                    horizontal: 'left',
                    vertical: 'middle'
                };

                //  Ячейка названия материалы
                row.getCell(scol + 2).alignment = {
                    indent: 1,
                    horizontal: 'left',
                    vertical: 'middle'
                };

                //  Ячейка количества
                row.getCell(scol + 3).alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };

                //  Ячейка ед. изм.
                row.getCell(scol + 4).alignment = {
                    horizontal: 'center',
                    vertical: 'middle'
                };

                //  Ячейка цены
                row.getCell(scol + 5).alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };

                //  Ячейка суммы
                row.getCell(scol + 6).alignment = {
                    indent: 1,
                    horizontal: 'right',
                    vertical: 'middle'
                };

                //  Форматирование числа
                if (val == "count") {
                    row.getCell(scol + 3).numFmt = '# ##0';
                } else {
                    row.getCell(scol + 3).numFmt = '#,##0.00';
                };
                row.getCell(scol + 5).numFmt = '#,##0.00';
                row.getCell(scol + 6).numFmt = '#,##0.00';
            };

            ind += items.length;
            const endRowInd = ind - 1;

            //  Объединяем все ячейки в этой колонке для диапазона строк
            const cellAdress = `${coeffColLetter}${startRowInd}`;
            worksheet.mergeCells(
                `${coeffColLetter}${startRowInd}:${coeffColLetter}${endRowInd}`
            );
            const mergedCell = worksheet.getCell(cellAdress);
            mergedCell.alignment = {
                vertical: 'middle',
                horizontal: 'center'
            };
            mergedCell.numFmt = '#,##0.00';
            mergedCell.font = {
                name: 'Arial',
                size: font_size - 1,
                bold: false
            };
        });

        const row = worksheet.getRow(ind);
        endRowTableStyle(row, scol, col_count);
        const csm = getColumnLetter(scol + 6);

        //  Ячейка Итого
        row.getCell(scol + 5).value = "Итого:";
        row.getCell(scol + 5).alignment = {
            horizontal: 'right',
            vertical: 'middle'
        };
        row.getCell(scol + 5).border = {
            top: { style: 'medium' },
            left: { style: 'none' },
            bottom: { style: 'medium' },
            right: { style: 'thin' }
        };

        //  Ячейка суммы
        row.getCell(scol + 6).value = {
            formula: `SUM(${csm}${headerRowInd + 1}:${csm}${ind - 1})`
        };
        row.getCell(scol + 6).numFmt = '# ##0';
        row.getCell(scol + 6).alignment = {
            indent: 1,
            horizontal: 'right',
            vertical: 'middle'
        };

        //  Устанавливаем диапазон печати
        // Получаем номер последней строки, где есть данные
        // Устанавливаем область печати: колонки A-scm, все строки с данными
        worksheet.pageSetup.printArea = `A1:${csm}${worksheet.rowCount}`;

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
            await createEsimateExcelFile(prj_array);
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