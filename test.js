

const fs = require('fs');
const path = require('path');
const { XMLParser } = require('fast-xml-parser');
const ExcelJS = require('exceljs');
const { count } = require('console');

//  Расширения обрабатываемых файлов
const extensions = ['.fr3d', '.b3d'];

// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("settings.json")) errFinish("Отсутсвует файл настроек: settings.json");

//  Считываем данные
let data = fs.readFileSync("settings.json", { encoding: "utf-8" });

// Удаляем BOM если он есть
if (data.charCodeAt(0) === 0xFEFF) data = data.slice(1);
const settings = JSON.parse(data);

const PROJECT_FILE = system.askFileName('bprj');
const PROFECT_NAME = system.getFileNameWithoutExtension(PROJECT_FILE);

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

//  Функция получения информации о объектах модели
function selectObjectProcess(item, data) {
    if (item instanceof TFurnPanel) processingPanel(item, data);
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
}

//  Функция получения массива данных, описывающих файлы проекта
function getProjectFilesData(settings) {
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

        // Получаем доступ к данным
        const listFiles = prjData.Document.DataProject.ListFiles.File;
        const result = [];
        for (let i = 0; i < listFiles.length; i++) {
            const elem = listFiles[i];
            const filePath = path.join(settings.coreDIR, elem.Name);

            if (!fs.existsSync(filePath)) errFinish(`Файла по указанному пути не существует: ${filePath}`);

            //  Добавляем данные в массив результатов
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
                data: {}
            });
        };

        //  Возвращаем результат
        return result;
    } catch (e) {
        errFinish(e);
    };
};
//#endregion

//#region Основные функции 

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
        materialTkn: obj.Thickness,
        barcode_data: PROFECT_NAME + settings.delimPrjName + prjPos,
        pos: obj.ArtPos,
        prjPos: prjPos,
        prjCount: options.count,
        dest: obj.Designation,
        name: obj.Name,
        width: pnlWidth,
        height: pnlHeight,
        area: round(pnlWidth * pnlHeight * 0.000001, 2)
    });
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
        { header: 'Код синхр.', key: 'code', width: 20 },
        //{ header: 'Артикул', key: 'material_art', width: 12 },
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
            sync_ext: "",
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

async function generateAllExcelFiles(array) {
    let message = "Укажите папку для сохранения файлов спецификаций";
    const folder = system.askFolder(message, path.dirname(PROJECT_FILE));

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
        await createMaterialExcel(result[i], folder);

        //  Создание печатной спецификации.
    };
    Action.Finish();
};

//#endregion

//  Основная функция
function main() {

    let ind = 0;
    let count = 0;

    //  Получаем массив данных для обработки
    let array = getProjectFilesData(settings);

    //Функция обработки файлов
    async function ProcessNextFile() {
        let fileName = array[ind].name;

        // вывод прогресса обработки файлов
        Action.Hint =
            `обработка файла ${ind + 1} из ${array.length} - ${fileName}`;
        if (extensions.indexOf(path.extname(fileName)) >= 0) {

            if (Action.LoadModel(fileName)) {
                /******   Блок обработки текущей модели   ********/
                //  Сбор и заполнение данных о текущей модели
                forEachInList(Model, selectObjectProcess, array[ind]);
                /*************************************************/
                count++;
            };
        };

        ind++;
        if (ind < array.length) {
            // Обработка следующего файла
            Action.AsyncExec(ProcessNextFile);
        } else {
            //alert(`Обработано ${count} файлов из ${array.length}`);
            //  Завершение обработки массива файлов.
            await generateAllExcelFiles(array);

            console.log(`Обработано ${count} файлов из ${array.length}`);
            Action.Finish();
        };
    };

    if (array.length > 0) ProcessNextFile();
};

main();
Action.Continue();
