const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ==========================================
// ТВОИ ДАННЫЕ
// ==========================================
const materials = [
    {
        materialName: "ДСП 16мм Белый",
        items: [
            { pos: "1", width: 356, height: 230, name: "Панель горизонт" },
            { pos: "2", width: 356, height: 500, name: "Боковая панель левая" },
            { pos: "3", width: 356, height: 500, name: "Боковая панель правая" },
            { pos: "4", width: 500, height: 350, name: "Полка" }
        ]
    },
    {
        materialName: "МДФ 10мм Серый",
        items: [
            { pos: "1", width: 200, height: 300, name: "Фасад верхний" },
            { pos: "2", width: 200, height: 500, name: "Фасад нижний" }
        ]
    },
    {
        materialName: "ЛДСП 22мм Дуб",
        items: [
            { pos: "1", width: 600, height: 400, name: "Столешница" },
            { pos: "2", width: 150, height: 700, name: "Ножка стола" },
            { pos: "3", width: 150, height: 700, name: "Ножка стола 2" }
        ]
    }
];

// ==========================================
// НАСТРОЙКИ
// ==========================================
const dir_name = 'excel_output'; // Папка для сохранения файлов

// ==========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ==========================================

// Безопасное имя файла
function sanitizeFileName(name) {
    return name
        .replace(/[<>:"/\\|?*]/g, '')
        .replace(/\s+/g, '_')
        .substring(0, 200);
}

// Создание папки, если её нет
function ensureDir(dir) {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
}

// ==========================================
// ОСНОВНАЯ ФУНКЦИЯ СОЗДАНИЯ ФАЙЛА ДЛЯ ОДНОГО МАТЕРИАЛА
// ==========================================
async function createMaterialExcel(materialData, outputDir) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Мебельный конструктор';
    workbook.created = new Date();

    // Лист с именем материала (ограничение Excel — 31 символ)
    const sheetName = materialData.materialName.substring(0, 31);
    const worksheet = workbook.addWorksheet(sheetName);

    // ---------- КОЛОНКИ ----------
    worksheet.columns = [
        { header: '№ позиции', key: 'pos', width: 12 },
        { header: 'Наименование', key: 'name', width: 35 },
        { header: 'Ширина, мм', key: 'width', width: 15 },
        { header: 'Высота, мм', key: 'height', width: 15 },
        { header: 'Площадь, м²', key: 'area', width: 15 }
    ];

    // ---------- СТИЛЬ ШАПКИ ----------
    const headerRow = worksheet.getRow(1);
    headerRow.height = 25;
    headerRow.font = {
        name: 'Arial',
        size: 12,
        bold: true,
        color: { argb: 'FFFFFFFF' }
    };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4472C4' }
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    headerRow.eachCell((cell) => {
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // ---------- ДАННЫЕ ----------
    materialData.items.forEach((item, index) => {
        const area = parseFloat(((item.width * item.height) / 1_000_000).toFixed(3));

        const row = worksheet.addRow({
            pos: item.pos,
            name: item.name,
            width: item.width,
            height: item.height,
            area: area
        });

        const currentRow = worksheet.getRow(index + 2);
        currentRow.height = 20;
        currentRow.font = { name: 'Arial', size: 11 };

        // Чередование фона
        if (index % 2 === 0) {
            currentRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF2F2F2' }
            };
        }

        // Выравнивание
        currentRow.getCell('pos').alignment = { horizontal: 'center' };
        currentRow.getCell('width').alignment = { horizontal: 'center' };
        currentRow.getCell('height').alignment = { horizontal: 'center' };
        currentRow.getCell('area').alignment = { horizontal: 'center' };
        currentRow.getCell('name').alignment = { horizontal: 'left' };

        // Границы
        currentRow.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
                left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
                bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
                right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
            };
        });
    });

    // ---------- ИТОГОВАЯ СТРОКА ----------
    const totalRow = worksheet.addRow({
        pos: '',
        name: 'ИТОГО:',
        width: '',
        height: '',
        area: 0
    });

    const totalRowNumber = materialData.items.length + 2;
    const lastRow = worksheet.getRow(totalRowNumber);

    const totalArea = materialData.items.reduce((sum, item) => {
        return sum + (item.width * item.height) / 1_000_000;
    }, 0);

    lastRow.getCell('area').value = parseFloat(totalArea.toFixed(3));
    lastRow.font = { name: 'Arial', size: 11, bold: true };
    lastRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E2F3' }
    };

    lastRow.eachCell((cell) => {
        cell.border = {
            top: { style: 'medium', color: { argb: 'FF4472C4' } },
            left: { style: 'thin' },
            bottom: { style: 'medium' },
            right: { style: 'thin' }
        };
    });

    // ---------- АВТОФИЛЬТР + ЗАКРЕПЛЕНИЕ СТРОКИ ----------
    worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: materialData.items.length + 1, column: 5 }
    };

    worksheet.views = [{ state: 'frozen', ySplit: 1 }];

    // ---------- СОХРАНЕНИЕ ----------
    const fileName = `${sanitizeFileName(materialData.materialName)}.xlsx`;
    const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(filePath);

    console.log(`✅ ${fileName}`);
    console.log(`   Позиций: ${materialData.items.length} | Общая площадь: ${totalArea.toFixed(3)} м²\n`);
}

// ==========================================
// ГЛАВНАЯ ФУНКЦИЯ — ГЕНЕРАЦИЯ ВСЕХ ФАЙЛОВ
// ==========================================
async function generateAllExcelFiles(materialsData) {
    console.log('🚀 Запуск генерации Excel-файлов...\n');

    const directory = path.dirname(Action.ModelFilename);
    const output_dir = path.join(directory, dir_name);
    ensureDir(output_dir);

    for (const material of materialsData) {
        try {
            await createMaterialExcel(material, output_dir);
        } catch (error) {
            console.error(`❌ Ошибка для "${material.materialName}":`, error.message);
        }
    }

    console.log(`✨ Готово! Файлы сохранены в папку "${output_dir}"`);
    Action.Finish();
}

// ==========================================
// ЗАПУСК
// ==========================================
generateAllExcelFiles(materials);
Action.Continue();