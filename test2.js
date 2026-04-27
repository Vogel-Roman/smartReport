const fs = require('fs');
const path = require('path');
const { XMLParser } = require('fast-xml-parser');
const ExcelJS = require('exceljs');
const firebird = require('node-firebird');

//  Расширения обрабатываемых файлов
const extensions = ['.fr3d', '.b3d'];

// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("settings.json")) errFinish("Отсутсвует файл настроек: settings.json");

//  Считываем данные
let data = fs.readFileSync("settings.json", { encoding: "utf-8" });

// Удаляем BOM если он есть
if (data.charCodeAt(0) === 0xFEFF) data = data.slice(1);
const settings = JSON.parse(data);

//  Инициализация констант
const PROJECT_FILE = system.askFileName('bprj');
const PROFECT_NAME = system.getFileNameWithoutExtension(PROJECT_FILE);

let message = "Укажите папку для сохранения файлов спецификаций";
const FOLDER = system.askFolder(message, path.dirname(PROJECT_FILE));

//#region Функции рекурсивного обхода

//  Функция рекурсивного обхода модели
function forEachInList(list, func, obj_data) {
    if (!func) return;
    for (let i = 0; i < list.Count; i++) {
        let elem = list.Objects[i];
        func(elem, obj_data);
        if (
            elem.List &&
            (elem instanceof TFurnBlock || elem instanceof TDraftBlock)
        ) forEachInList(elem.AsList(), func, obj_data);
    };
};

//  Call-back функция получения информации о типах объектов
function selectObjectProcess(item, data) {
    if (item instanceof TFurnPanel) processingPanel(item, data);
};

//  Функция обработки данных панели
function processingPanel(obj, options) {
    //  Игнорируем исключенные материалы
    if (settings.excludeMaterials.includes(obj.MaterialName)) return;

    if (!options.data.panelMaterials) options.data.panelMaterials = [];

    let mName = obj.MaterialName;
    let mArt = "";

    if (obj.MaterialName.indexOf("\r") > 0) {
        mArt = mName.split("\r")[1];
        mName = mName.split("\r")[0];
    };

    let pnlWidth = round(obj.ContourWidth, 1);
    let pnlHeight = round(obj.ContourHeight, 1);
    let prjPos = options.sign + settings.delimPrjSign + addLZ(obj.ArtPos, 3);
    options.data.panelMaterials.push({
        material: obj.MaterialName,
        materialName: mName,
        materialArticle: mArt,
        materialSyncExternal: "",    //  Код синхронизации материала
        materialPrice: 0,
        materialUnit: "",
        materialTkn: obj.Thickness,
        barcode_data: PROFECT_NAME + settings.delimPrjName + prjPos,
        pos: obj.ArtPos,
        prjPos: prjPos,
        prjCount: options.count,
        dest: obj.Designation,
        name: obj.Name,
        width: pnlWidth,
        height: pnlHeight,
        area: round(pnlWidth * pnlHeight * 0.000001, 2),
        contourLength: round(obj.Contour.ObjLength() * 0.001, 2)
    });
};

//#endregion

//#region Служебные функции

//  Функция обработки ошибок
function errFinish(str) {
    console.log(str);
    Action.Finish();
};

//  Функция добавляющая лидирующие нули
function addLZ(value, length = 2) {
    // Преобразуем в строку и добавляем нули
    return String(value).padStart(length, '0');
};

// Безопасное имя файла
function sanitizeFileName(name) {
    return name
        .replace(/[<>:"/\\|?*]/g, '')
        .replace(/\s+/g, ' ')
        .substring(0, 200);
};

//  Функция огругления
function round(a, b) {
    b = b || 0;
    return Math.round(a * Math.pow(10, b)) / Math.pow(10, b);
};

//  Функция сортировки массива по свойству
function sortByProperty(array, property, ascending = true) {
    return [...array].sort((a, b) => {
        const valueA = a[property];
        const valueB = b[property];

        if (valueA < valueB) return ascending ? -1 : 1;
        if (valueA > valueB) return ascending ? 1 : -1;
        return 0;
    });
};


//#endregion

//#region Функции обработки файлов проекта

//  Функция заполнения общих данных текущей модели
function getModelData(data) {
    if (!data) return;
    data.model = {
        orderCode: Article.OrderName,
        orderName: Article.Name,
        enterprise: settings.enterprise || "SmartWood",
        author: settings.author || "Vogel",
    };
    //console.log(JSON.stringify(data.model, null, 2));
};

async function getEstimateData(data) {
    //console.log(JSON.stringify(data, null, 2));
    data.estimate = {
        panelMaterials: [],
        profileMaterials: [],
        furnitureMaterials: []
    };

    //  Функция суммирования значений данных Материала
    function sumByPanelMaterial(items) {
        const grouped = items.reduce((acc, item) => {
            const key = item.materialName;

            if (!acc[key]) {
                acc[key] = {
                    materialName: key,
                    materialArticle: item.materialArticle,
                    materialTkn: item.materialTkn,
                    materialSyncExternal: "",
                    materialClassArray: [],
                    materialPrice: 0,
                    area: 0,
                    contourLength: 0
                };
            };

            //  Площадь деталей
            acc[key].area += item.area || 0;

            //  Длина контуров деталей
            acc[key].contourLength += item.contourLength || 0;

            return acc;
        }, {});

        return Object.values(grouped);
    };

    if (data.data.panelMaterials) {
        let grouped = sumByPanelMaterial(data.data.panelMaterials);
        data.estimate.panelMaterials = [...grouped];
    };
};

//  Функция получения данных, описывающих файлы Проекта
function getProjectFilesData() {
    //  Проверка выбора файла Директории
    if (!PROJECT_FILE) errFinish("Файл проекта не выбран");

    try {

        // Читаем XML файл
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
        const result = [];
        for (let i = 0; i < listFiles.length; i++) {
            const elem = listFiles[i];
            const filePath = path.join(settings.coreDIR, elem.Name);

            if (!fs.existsSync(filePath))
                errFinish(`Файла по указанному пути не существует: ${filePath}`);

            //  Добавляем данные из файла Проекта в массив
            result.push({
                type: elem.Type,
                name: filePath,
                sign: elem.Sign,
                count: elem.Count,
                subname: elem.SubName,
                note: elem.Note,
                comment: elem.Comment,
                estimate: elem.HasUse_Estimate == "Y" ? 1 : 0,
                cutting: elem.HasUse_Cutting == "Y" ? 1 : 0,
                cnc: elem.HasUse_CNC == "Y" ? 1 : 0,
                data: {}    //  Контейнер для информации из Модели
            });
        };

        //  Возвращаем результат
        return result;
    } catch (e) {
        errFinish(e.message);
    };
};

//#endregion

//#region Функции формирования файлов excel

// Функция для стилизации диапазона ячеек
function styleCellRange(row, startCol, endCol, styles) {
    for (let col = startCol; col <= endCol; col++) {
        const cell = row.getCell(col);
        Object.assign(cell, styles);
    };
};

async function createMaterialExcel(data, outputDir) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.creator;
    workbook.created = new Date();

    const borderColor = 'FFD9D9D9';

    //  Стиль строки заголовка
    let headerRowStyle = {
        fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        },
        border: {
            top: { style: 'thin', color: { argb: borderColor } },
            left: { style: 'thin', color: { argb: borderColor } },
            bottom: { style: 'thin', color: { argb: borderColor } },
            right: { style: 'thin', color: { argb: borderColor } }
        },
        alignment: { vertical: 'middle', horizontal: 'center' },
        font: {
            name: 'Calibri',
            size: 11,
            color: { argb: 'FFFFFFFF' },
        }
    };

    //  Стиль строки
    let rowStyle = {
        border: {
            top: { style: 'thin', color: { argb: borderColor } },
            left: { style: 'thin', color: { argb: borderColor } },
            bottom: { style: 'thin', color: { argb: borderColor } },
            right: { style: 'thin', color: { argb: borderColor } }
        },
        alignment: { indent: 1, horizontal: 'right', vertical: 'middle' },
        font: { name: 'Calibri', size: 11 }
    };

    const row_h = 18;

    let material = data.material;
    let array = data.array;

    // Лист с именем материала (ограничение Excel — 31 символ)
    const sheetName = material.substring(0, 31);
    const worksheet = workbook.addWorksheet(sheetName);

    // ---------- КОЛОНКИ ----------
    worksheet.columns = [
        { header: 'barcode', key: 'barcode', width: 20 },
        { header: 'Поз.', key: 'prjPos', width: 12 },
        { header: 'Наименование', key: 'name', width: 25 },
        { header: 'Длина', key: 'width', width: 10 },
        { header: 'Ширина', key: 'height', width: 10 },
        { header: 'Кол-во', key: 'count', width: 10 },
        { header: 'Площадь', key: 'area', width: 10 },
        { header: 'Толщина', key: 'tkn', width: 10 },
        { header: 'Код синхр.', key: 'sync_ext', width: 20 },
        { header: 'Артикул', key: 'material_art', width: 20 },
        { header: 'Материал', key: 'material', width: 80 }
    ];

    // ---------- СТИЛЬ ШАПКИ ----------
    const headerRow = worksheet.getRow(1);
    headerRow.height = row_h;
    styleCellRange(headerRow, 1, worksheet.columns.length, headerRowStyle);

    // ---------- ДАННЫЕ ----------
    array.forEach((item, index) => {
        worksheet.addRow({
            barcode: item.barcode_data,
            prjPos: item.prjPos,
            name: item.name,
            tkn: item.materialTkn,
            width: item.width,
            height: item.height,
            count: item.prjCount,
            area: item.area,
            sync_ext: item.materialSyncExternal,
            material_art: item.materialArticle,
            material: item.materialName
        });

        const currentRow = worksheet.getRow(index + 2);
        currentRow.height = row_h;
        styleCellRange(currentRow, 1, worksheet.columns.length, rowStyle);

        // Чередование фона
        if (index % 2 === 0) {
            styleCellRange(currentRow, 1, worksheet.columns.length, {
                fill: {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFF2F2F2' }
                }
            });
        };

        // Выравнивание
        currentRow.getCell('name').alignment = { indent: 1, horizontal: 'left' };
        currentRow.getCell('material').alignment = { indent: 1, horizontal: 'left' };
    });

    // ---------- ЗАКРЕПЛЕНИЕ СТРОКИ ----------
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];

    // ---------- СОХРАНЕНИЕ ----------
    const fileName = `${sanitizeFileName(PROFECT_NAME)}_${sanitizeFileName(material)}.xlsx`;
    const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(filePath);
};

async function generateSpecificationFiles(array) {
    //const folder = "C:\\Users\\Roman\\Desktop\\test_project";

    //  Функция группировки объектов по материалам
    function groupByMaterial(items) {
        // Группируем
        const grouped = items.reduce((acc, item, index) => {
            const material = item.materialName;

            if (!acc[material]) {
                // Создаем новую группу
                acc[material] = {
                    id: Object.keys(acc).length + 1, // или можно index использовать
                    material: material,
                    array: []
                };
            };

            acc[material].array.push(item);
            return acc;
        }, {});

        // Превращаем объект в массив
        return Object.values(grouped);
    };

    //  Объединяем элементы из всех файлов в общие массивы
    let arr_panelMaterials = [];

    for (let i = 0; i < array.length; i++) {

        //  Проверяем наличие такого свойства у элемента массива array
        if (array[i].data.panelMaterials) {
            arr_panelMaterials.push(...array[i].data.panelMaterials);
        };

        //… аналогично для другних свойств

    };

    let result = groupByMaterial(arr_panelMaterials);
    for (let i = 0; i < result.length; i++) {
        //  Сортируем массив деталей по возрастанию позиции в проекте
        result[i].array = sortByProperty(result[i].array);

        //  Создание спецификации для загрузки
        await createMaterialExcel(result[i], FOLDER);

        //  Создание печатной спецификации.
    };
    Action.Finish();
};

//#endregion

//#region Функции формирвоания файла сметы

async function createEstimateExcelFile(arr_data, outputDir) {

    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.creator;
    workbook.created = new Date();

    const borderColor = 'FFD9D9D9';
    const row_h = 18;

    //  Стиль строки заголовка
    let headerRowStyle = {
        fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        },
        border: {
            top: { style: 'thin', color: { argb: borderColor } },
            left: { style: 'thin', color: { argb: borderColor } },
            bottom: { style: 'thin', color: { argb: borderColor } },
            right: { style: 'thin', color: { argb: borderColor } }
        },
        alignment: { vertical: 'middle', horizontal: 'center' },
        font: {
            name: 'Calibri',
            size: 11,
            color: { argb: 'FFFFFFFF' },
        }
    };

    //  Стиль строки
    let rowStyle = {
        border: {
            top: { style: 'thin', color: { argb: borderColor } },
            left: { style: 'thin', color: { argb: borderColor } },
            bottom: { style: 'thin', color: { argb: borderColor } },
            right: { style: 'thin', color: { argb: borderColor } }
        },
        alignment: { indent: 1, horizontal: 'right', vertical: 'middle' },
        font: { name: 'Calibri', size: 11 }
    };

    //  Создание вкладок по изделиям
    for (let i = 0; i < arr_data.length; i++) {
        const obj = arr_data[i];
        const sheetName = `${obj.sign}_${obj.model.orderName}`.substring(0, 31);
        const worksheet = workbook.addWorksheet(sheetName);

        // ---------- КОЛОНКИ ----------
        worksheet.columns = [
            { header: '№', key: 'ind', width: 10 },
            { header: 'Номенклатура', key: 'material', width: 80 },
            { header: 'Кол-во', key: 'count', width: 10 },
            { header: 'Ед. изм.', key: 'unit', width: 10 },
            { header: 'Цена', key: 'price', width: 20 },
            { header: 'Сумма', key: 'sum', width: 20 },
        ];

        // ---------- СТИЛЬ ШАПКИ ----------
        const headerRow = worksheet.getRow(1);
        headerRow.height = row_h;
        styleCellRange(headerRow, 1, worksheet.columns.length, headerRowStyle);

        // ---------- ДАННЫЕ ----------
        obj.estimate.panelMaterials.forEach((item, index) => {

            const rn = index + 2;           // +2 из-за заголовка
            worksheet.addRow({
                ind: index + 1,             //  A
                material: item.materialName,//  B
                count: item.area,           //  C
                unit: item.materialUnit,    //  D
                price: item.materialPrice,  //  E
                sum: 0                      //  F
            });

            //  Формула суммы
            worksheet.getCell(`F${rn}`).value = {
                formula: `C${rn}*E${rn}`,
                result: (item.price || 0) * (item.count || 0)
            };

            //  Устанавливаем форматирвоание числа 0,00
            worksheet.getCell(`C${rn}`).numFmt = '#,##0.00';
            worksheet.getCell(`E${rn}`).numFmt = '#,##0.00';
            worksheet.getCell(`F${rn}`).numFmt = '#,##0.00';

            const currentRow = worksheet.getRow(index + 2);
            currentRow.height = row_h;
            styleCellRange(currentRow, 1, worksheet.columns.length, rowStyle);

            // Чередование фона
            if (index % 2 === 0) {
                styleCellRange(currentRow, 1, worksheet.columns.length, {
                    fill: {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF2F2F2' }
                    }
                });
            };

            // Выравнивание
            currentRow.getCell('material').alignment = { indent: 1, horizontal: 'left' };
        });

        // ---------- ЗАКРЕПЛЕНИЕ СТРОКИ ----------
        worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    };

    // ---------- СОХРАНЕНИЕ ----------
    const fileName = `${sanitizeFileName(PROFECT_NAME + "_Калькуляция проекта")}.xlsx`;
    const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(filePath);
};

async function generateEstimateFile(arr_data) {

    function getListMaterialNames(array) {
        let result = [];
        array.map((item, index) => {
            result.push(item.materialName);
        });
        return result;
    };

    //  Функция группировки объектов по материалам
    function groupByPanelMaterial(items) {

        // Группируем
        const grouped = items.reduce((acc, item, index) => {

            const key = item.materialName;

            if (!acc[key]) {
                // Создаем новую группу
                acc[key] = {
                    id: Object.keys(acc).length + 1,
                    materialName: item.materialName,
                    materialArticle: item.materialArticle,
                    materialTkn: item.materialTkn,
                    materialSyncExternal: item.materialSyncExternal,
                    materialClassArray: [],
                    materialPrice: 0,
                    materialUnit: "",
                    area: 0,
                    contourLength: 0
                };
            };

            //  Площадь деталей
            acc[key].area += item.area || 0;

            //  Длина контуров деталей
            acc[key].contourLength += item.contourLength || 0;

            return acc;
        }, {});

        // Превращаем объект в массив
        return Object.values(grouped);
    };

    //  Объединяем элементы из всех файлов в общие массивы
    let arr_panelMaterials = [];

    for (let i = 0; i < arr_data.length; i++) {
        const model_data = arr_data[i];

        //  Заполняем массив листовых материалов
        arr_panelMaterials.push(...model_data.estimate.panelMaterials);

        //  Заполняем массив кромочных материалов

        //  Заполняем массив погонных материалов

        //  Заполняем массив фурнитуры
    };

    let grouped_panelMaterials = groupByPanelMaterial(arr_panelMaterials);

    //  Формируем список наименований материлов для обращения в Базу материалов
    let listMaterials = [...getListMaterialNames(grouped_panelMaterials)];

    //  Получение данных из Базы материалов
    const dbMaterialData = await getDBMaterialInfo(listMaterials);

    // Быстрый доступ через Map
    const dbDataMap = new Map(
        dbMaterialData.map(item => [item.name_mat.toLowerCase(), item])
    );

    // Дополняем grouped_panelMaterials
    for (const panelMaterial of grouped_panelMaterials) {
        const dbItem = dbDataMap.get(panelMaterial.materialName.toLowerCase());
        if (dbItem) {
            panelMaterial.materialPrice = dbItem.price;
            panelMaterial.syncExternal = dbItem.sync_external;
            panelMaterial.materialTkn = dbItem.thickness;
            panelMaterial.materialUnit = dbItem.name_meas;
            panelMaterial.dimensions = {
                length: dbItem.length,
                width: dbItem.width,
                tkn: dbItem.thickness
            };
        };
    };

    // Дополняем данные в спецификациях деталей и сметы
    for (let i = 0; i < arr_data.length; i++) {
        const obj = arr_data[i];
        for (const panelMaterial of obj.data.panelMaterials) {
            const dbItem = dbDataMap.get(panelMaterial.materialName.toLowerCase());
            if (dbItem) {
                panelMaterial.materialPrice = dbItem.price;
                panelMaterial.materialSyncExternal = dbItem.sync_external;
                panelMaterial.materialTkn = dbItem.thickness;
                panelMaterial.materialUnit = dbItem.name_meas;
            };
        };

        for (const panelMaterial of obj.estimate.panelMaterials) {
            const dbItem = dbDataMap.get(panelMaterial.materialName.toLowerCase());
            if (dbItem) {
                panelMaterial.materialPrice = dbItem.price;
                panelMaterial.materialSyncExternal = dbItem.sync_external;
                panelMaterial.materialTkn = dbItem.thickness;
                panelMaterial.materialUnit = dbItem.name_meas;
            };
        };
    };

    //console.log('done');
    await createEstimateExcelFile(arr_data, FOLDER);


};

//#endregion

//#region Запрос в Базу данных

//  Функция получения данных о материалах из Базы материалов
async function getDBMaterialInfo(matnames) {
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

    try {
        const options = settings.db_options || null;

        //  Строка запроса
        const sqlString = `
        SELECT
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
        WHERE m.NAME_MAT IN (${matnames.map(() => '?').join(',')})
        `;
        //  Запрос в БД
        const result = await executeQuery(sqlString, options, matnames);

        //console.log(`Получено записей: ${result.length}`);
        //console.log(JSON.stringify(result, null, 2));
        return result;
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

//#endregion

//  Основная функция
function main() {

    let ind = 0;
    let count = 0;

    //  Получаем массив данных о файлах Проекта
    let prj_arr = getProjectFilesData();

    //  Функция обработки файлов Проекта
    async function ProcessNextFile() {
        let fileName = prj_arr[ind].name;

        // вывод прогресса обработки файлов
        Action.Hint =
            `обработка файла ${ind + 1} из ${prj_arr.length} - ${fileName}`;
        if (extensions.indexOf(path.extname(fileName)) >= 0) {

            if (Action.LoadModel(fileName)) {
                /********   Блок обработки текущей модели   **********/

                //  Рекурсивный обход текущей модели
                forEachInList(Model, selectObjectProcess, prj_arr[ind]);

                //  Заполняем общую информацию о модели
                getModelData(prj_arr[ind]);

                //  Формируем данные для сметы
                await getEstimateData(prj_arr[ind]);


                //console.log(JSON.stringify(prj_arr[ind].estimate, null, 2));

                /*****************************************************/
                count++;
            };
        };

        ind++;
        if (ind < prj_arr.length) {
            // Обработка следующего файла
            Action.AsyncExec(ProcessNextFile);
        } else {
            //alert(`Обработано ${count} файлов из ${prj_arr.length}`);

            //  Завершение обработки массива файлов.

            //  Запуск функции создания Сметы
            await generateEstimateFile(prj_arr);

            //  Запуск функции генерации файлов EXCEL
            await generateSpecificationFiles(prj_arr);

            console.log(`Обработано ${count} файлов из ${prj_arr.length}`);
            Action.Finish();
        };
    };

    if (prj_arr.length > 0) ProcessNextFile();
};

main();
Action.Continue();
