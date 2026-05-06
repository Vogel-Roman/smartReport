
const fs = require('fs');
const path = require('path');
const firebird = require('node-firebird');

// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("settings.json"))
    errFinish("Отсутсвует файл настроек: settings.json");

//  Считываем данные и удаляем BOM если он есть
let settings_data = fs.readFileSync("settings.json", { encoding: "utf-8" });
if (settings_data.charCodeAt(0) === 0xFEFF)
    settings_data = settings_data.slice(1);

//  Данные настроек внешнего файла
const settings = JSON.parse(settings_data);

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

    // //  Формируем общий массив наименований для запроса в БД
    // const matnames = [
    //     ...BOARD_ARRAY, ...BUTT_ARRAY,
    //     ...PROFILE_ARRAY, ...FURNITURE_ARRAY
    // ]; // "ЛДСП «Альпина белая» W1100 ST9 16 мм, EGGER", "Конфирмат 7х50 мм, Zn"


    const matnames = [
        "LIBRA CC2 Заглушка для навесов D12, пластик, белая",
        //"Шуруп 3,5х20 мм, Zn",
        "LIBRA H2 Скрытый навес универсальный",
        "LIBRA WP2 Планка для навесов с вертикальной регулировкой, сталь",
        // "Шуруп МЦП 45мм",
        // "Конфирмат 7х50 мм, Zn",
        //"Шкант 8х30 мм",
        "Сушка Vibo в мод.600 мм",
        "BI-MATERIALE Демпфер D5 мм, белый",
        "Заглушка на чашку петли",
        // "Шуруп 3,5х16 мм, Zn",
        "Петля Blum Clip top Blumotion 110°, накладная",
        "Заглушка на плечо петли (прямая)",
        "Монтажная планка h=0 (прям. с эксц.) Expando D-5мм",
        "Полкодержатель K-LINE с фиксацией, никель",
        // "Стяжка Rastex 15/15 D",
        //"Дюбель DU 232 Twister",
        "Монтажная планка h=0 (прям. с эксц.) под саморез",
        // "Чашка петли",
        "Петля междверн. Clip top без пружины",
        "Петля Blum Clip top 120° без пружины, накладная",
        "Комплект силовых маханизмов AVENTOS HF top 25 (кр. на саморезы 4х35)",
        //"Шуруп 3,5х35 мм, Zn",
        "Комплект рычагов AVENTOS HF top F35 (KH 600-910 мм)",
        "Комплект заглушек AVENTOS HF top, Белый"
    ];

    //console.log(JSON.stringify(matnames, null, 2));
    //console.log(JSON.stringify(FURNITURE_ARRAY, null, 2));
    //console.log(JSON.stringify(BOARD_ARRAY, null, 2));

    const placeholders = matnames.map(() => '?').join(',');


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
            m.NAME_MAT IN (${placeholders})
        `;


        //  Запрос в БД
        const result_db = await executeQuery(sqlString, options, matnames);

        const result = [
            [], //  Листовые материалы
            [], //  Кромочные материалы
            [], //  Погонные материалы
            []  //  Фурнитура
        ];

        // //  Сортировка результата по группам
        // result_db.forEach(elem => {
        //     if (BOARD_ARRAY.includes(elem.name_mat)) result[0].push(elem);
        //     if (BUTT_ARRAY.includes(elem.name_mat)) result[1].push(elem);
        //     if (PROFILE_ARRAY.includes(elem.name_mat)) result[2].push(elem);
        //     if (FURNITURE_ARRAY.includes(elem.name_mat)) result[3].push(elem);

        //     ID_MAT_ARRAY.push(elem.id_m);
        // });
        return result;
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

async function main() {
    let arr = await getDBMaterialInfo();
    Action.Finish();
};

main();
Action.Continue();