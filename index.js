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
function findPanelButtsList(panel) {
    const result = [];

    for (let i = 0; i < panel.Contour.Count; i++) {
        if (!panel.Contour[i].Data.Butt) continue;

        const elem = panel.Contour[i].Data.Butt;
        if (!elem.Thickness) continue;

        const overhung = elem.Overhung;
        const length = round(panel.Contour[i].ObjLength(), 2) + overhung * 2;

        const material = getMaterialName(elem.Material);

        //  Дополняем массив наименований материалов кромки
        if (!BUTT_ARRAY.includes(material[0])) BUTT_ARRAY.push(material[0]);

        result.push({
            material: elem.Material,        //  Имя материала кромки
            materialName: material[0],      //  Имя материала
            materialArticle: material[1],   //  Артикул материала кромки
            materialSyncExternal: "",       //  Код синхронизации (DB)
            materialUnit: "",               //  Единица измерения (DB)
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
    const buttInfoArray = findPanelButtsList(panel);

    //  Информация о пазах панели
    const cutInfoArray = findPanelCutsList(panel);

    //  Информация о присадке панелей
    const drillInfoArray = findPanelHolesList(panel);

    //  Информация о облицовки пласти
    const plasticInfoArray = findPanelPlasticList(panel);

    modelData.data.panelMaterials.push({
        name: panel.Name,               //  Имя панели
        material: panel.MaterialName,   //  Материал панели
        materialName: material[0],      //  Имя материала панели
        materialArticle: material[1],   //  Артикул материала панели
        materialSyncExternal: "",       //  Код синхронизации материала (DB)
        materialPrice: 0,               //  Цена из базы данных (DB)
        materialUnit: "",               //  Единица измерения (DB)
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
            materialName: mat[0],           //  Имя материала панели
            materialArticle: mat[1],        //  Артикул материала панели
            materialSyncExternal: "",       //  Код синхронизации материала (DB)
            materialPrice: 0,               //  Цена из базы данных (DB)
            materialUnit: "",               //  Единица измерения (DB)
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
                esimate: {}
            });
        });

        //  Возвращаем результат
        return result_array;
    } catch (e) {
        errFinish(e.message);
    };
};

//#endregion


// ---------- ОСНОВНАЯ ФУНКЦИЯ ---------- 
function main() {

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
            //alert(`Обработано ${count} файлов из ${prj_array.length}`);


            console.log(`Обработано ${count} файлов из ${prj_array.length}`);
            //console.log(JSON.stringify(prj_array, null, 2));
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