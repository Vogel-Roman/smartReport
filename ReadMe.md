1. Необходимо прочитать файл настроек settings.json и извлечь из него путь до корневой папки BazisMebelschik

2. Вызвать диалоговое окно выбора файла ".bpprj". Это XML файл по структуре. Рапарсить его и получить список файлов проекта.

3. Формируем массив данных для обработки файлов основным скриптом. Для проверки вывести цикл.

Разделы файла data
panelMaterials:

profileMaterials:

furnitureMaterials:


materials:["Material1"],
panels:[
    {
        
        material:"Material1",
        materialName: "Имя материала панелей",
        materialArt: "Артикул материала",
        materialTkn: 16,
        array:[]
    },
    {…}
]

{
    name:"Имя файла полный путь",
    model:{},
    estimate:{},
    panelMaterials:[],
}