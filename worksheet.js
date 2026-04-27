worksheet.pageSetup = {
    // Ориентация страницы
    orientation: 'landscape',  // 'portrait' (портретная) | 'landscape' (альбомная)

    // Размер бумаги (9 = A4)
    paperSize: 9,

    // Поля страницы (в ДЮЙМАХ! 1 дюйм = 2.54 см)
    margins: {
        top: 0.5,      // Верхнее поле
        bottom: 0.5,   // Нижнее поле
        left: 0.3,     // Левое поле
        right: 0.3,    // Правое поле
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
    horizontalCentered: true,
    verticalCentered: false,

    // Сетка и заголовки
    showGridLines: false,        // Печатать сетку
    showRowColHeaders: false,    // Печатать заголовки строк/столбцов

    // Повторять строки/колонки на каждой странице
    printTitlesRow: '6:6',      // Повторять 6-ю строку (заголовок)
    // printTitlesColumn: 'A:B',

    // Область печати
    // printArea: 'A1:G50',

    // Колонтитулы
    headerFooter: {
        oddHeader: '&C&BСпецификация деталей',
        oddFooter: '&LСтраница &P из &N&RДата: &D',
        evenHeader: '&C&BСпецификация деталей',
        evenFooter: '&LСтраница &P из &N&RДата: &D'
    },

    // Порядок страниц
    pageOrder: 'downThenOver',   // 'overThenDown'

    // Номер первой страницы
    firstPageNumber: 1,

    // Качество печати
    blackAndWhite: false,
    draft: false,

    // Количество копий
    copies: 1
};