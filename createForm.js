const fs = require('fs');
const path = require('path');
const { XMLParser } = require('fast-xml-parser');
const ExcelJS = require('exceljs');
const firebird = require('node-firebird');

// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("settings.json")) errFinish("Отсутсвует файл настроек: settings.json");

//  Считываем данные
let data = fs.readFileSync("settings.json", { encoding: "utf-8" });

// Удаляем BOM если он есть
if (data.charCodeAt(0) === 0xFEFF) data = data.slice(1);
const settings = JSON.parse(data);

// Функция для стилизации диапазона ячеек
function styleCellRange(row, startCol, endCol, styles) {
    for (let col = startCol; col <= endCol; col++) {
        const cell = row.getCell(col);
        Object.assign(cell, styles);
    };
};


async function createForm() {

    const workbook = new ExcelJS.Workbook();
    workbook.creator = settings.creator;
    workbook.created = new Date();

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

    function headerRowTableStyle(row, start, end) {
        let headerRowStyle = {
            border: {
                top: { style: 'medium' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            },
            alignment: { vertical: 'middle', horizontal: 'center' },
            font: {
                name: 'Arial',
                size: 9,
                bold: true
            }
        };
        let height = 13.1;
        row.height = Math.round(height / 0.75 * 10) / 10;

        styleCellRange(row, start, end, headerRowStyle);

        //  Первая ячейка
        styleCellRange(row, 2, 2, {
            border: {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        });

        //  Последняя ячейка
        styleCellRange(row, end, end, {
            border: {
                top: { style: 'medium' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'medium' }
            }
        });
    };

    let arr_data = ["M1", "M2", "M3"];

    for (let i = 0; i < arr_data.length; i++) {
        const obj = arr_data[i];

        const sheetName = `${obj}`.substring(0, 31);
        const worksheet = workbook.addWorksheet(sheetName);
        setWorksheetSettings(worksheet);

        worksheet.columns = [
            { width: 0.67 },    //  A Отступ
            { width: 5.42 },    //  B Номер 
            { width: 78 },      //  C Наименование
            { width: 8 },       //  D Количество
            { width: 8 },       //  E Ед. изм.
            { width: 12 },      //  F Цена
            { width: 12 }       //  G Сумма
        ];

        // ---------- СТИЛЬ ШАПКИ ----------
        const headerRowInd = 6;
        const headerRow = worksheet.getRow(headerRowInd);
        const start_col = 2;

        //  Стиль строки заголовка
        headerRowTableStyle(headerRow, start_col, worksheet.columns.length);

        // Устанавливаем значения заголовков вручную (начиная с B)
        const headers = ['№', 'Наименование', 'Кол-во', 'Ед.', 'Цена', 'Сумма'];
        for (let col = start_col; col <= headers.length + start_col - 1; col++) {
            const cell = headerRow.getCell(col);
            cell.value = headers[col - start_col];  // Заголовок
        };

    };

    // ---------- СОХРАНЕНИЕ ----------
    const fileName = `${"_Калькуляция проекта"}.xlsx`;
    //const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(fileName);

};

createForm();