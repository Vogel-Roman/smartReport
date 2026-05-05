const fs = require('fs');
const firebird = require('node-firebird');


//  Считываем данные и удаляем BOM если он есть
let settings_data = fs.readFileSync("settings.json", { encoding: "utf-8" });
if (settings_data.charCodeAt(0) === 0xFEFF)
    settings_data = settings_data.slice(1);
//  Данные настроек внешнего файла
const settings = JSON.parse(settings_data);

const options = settings.db_options || null;

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

let arr = [15299, 2162, 18703, 18095, 17688, 18125, 16527];




//  Функция получения данных о классах материалов
async function getDBCLassMaterialInfo(ids) {

    try {
        const options = settings.db_options || null;
        const classes = settings.estimate.classes;

        const sqlStr1 = `
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
        const result_db = await executeQuery(sqlStr1, options, ids);

        let result = {};

        //  Сортировка результата по группам
        result_db.forEach(elem => {
            if (classes.indexOf(elem.class_code) < 0) return;

            if (!result[elem.id_m]) result[elem.id_m] = elem;

            //console.log(elem.class_code);
            //console.log(classes.indexOf(elem.class_code));


        });
        console.log(JSON.stringify(result, null, 2));
        //return result;
        Action.Finish();
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

getDBCLassMaterialInfo(arr);
Action.Continue();