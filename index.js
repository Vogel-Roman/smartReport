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

//  Функция добавляющая лидирующие нули
function addZero(value, length = 3) {
    // Преобразуем в строку и добавляем нули
    return String(value).padStart(length, '0');
};

//  Функция получения артикула и названия материала из имени
function getMaterialName(material) {
    let mName = material;
    let mArt = "";
    if (material.indexOf("\r") > 0) {
        mArt = mName.split("\r")[1];
        mName = mName.split("\r")[0];
    };
    return [mName, mArt];
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
    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.panelMaterial;
    if (excludeMaterial.includes(panel.MaterialName)) return;

    const material = getMaterialName(panel.MaterialName);

    //  Размеры панели
    const w = round(panel.ContourWidth, 1);
    const h = round(panel.ContourHeight, 1);

    //  Площадь панели в метрах
    const panelArea = round(w * h * 0.000001, 2)

    //  Длина контура панели в метрах
    const contourLength = round(panel.Contour.ObjLength() * 0.001, 2);

    //  Позиция панели в проекте M2-0012
    const projectPos =
        modelData.sign + settings.delimPrjSign + addZero(obj.ArtPos);

    //  Обозначение панели в проекте M2-0012
    const projectDes =
        modelData.sign + settings.delimPrjSign + addZero(obj.Designation);

    //  Текст в QR-коде (Позиция – ArtPos)
    const barcodeData = PROJECT_NAME + settings.delimPrjName + projectPos;

    //  Текст в QR-коде (Обозначение – Designation)
    const barcodeDataDes = PROJECT_NAME + settings.delimPrjName + projectDes;

    //  Информация о кромках панели
    const buttInfoArray = [];

    //  Информация о пазах панели
    const cutInfoArray = [];

    //  Информация о присадке панелей
    const drillInfoArray = [];

    //  Информация о облицовки пласти
    const plasticInfoArray = [];

    modelData.data.panelMaterials.push({
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
        name: panel.Name,               //  Имя панели
        width: w,                       //  Длина панели
        height: h,                      //  Ширина панели
        area: panelArea,                //  Площадь панели
        contourLength: contourLength,   //  Длина контура панели
        buttInfo: buttInfoArray,        //  Массив кромок панели
        cutInfo: cutInfoArray,          //  Массив пазов панели
        drillInfo: drillInfoArray,      //  Массив отверстий панели
        plasticInfo: plasticInfoArray   //  Массив облицовки пласти панели
    });
};

//  Функция обработки данных профиля
function profileProcessing(profile, modelData) {
    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.profileMaterial;
    if (excludeMaterial.includes(profile.MaterialName)) return;

    const material = getMaterialName(profile.MaterialName);

};

//  Функция обработки данных фурнитуры
function furnitureProcessing(fastener, modelData) {
    //  Игнорируем исключенные материалы
    const excludeMaterial = settings.exclude.furnitureMaterial;
    if (excludeMaterial.includes(fastener.MaterialName)) return;

    //  !!!! Тут нужна проверка на составную фурнитуру !!!!!!
    const material = getMaterialName(fastener.MaterialName);
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
                console.log(prj_array[ind].name);

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