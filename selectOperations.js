/******************************************************************************/
/***       Получение списка операций по сопутсвию группам материалов        ***/
/******************************************************************************/

const fs = require('fs');
const path = require('path');
const firebird = require('node-firebird');

//let Action = {};

// Проверяем существование файла и читаем файл настроек settings.json 
if (!fs.existsSync("settings.json")) errFinish("Нет файла settings.json");

//  Считываем данные
let data = fs.readFileSync("settings.json", { encoding: "utf-8" });

// Удаляем BOM если он есть
if (data.charCodeAt(0) === 0xFEFF) data = data.slice(1);
const settings = JSON.parse(data);

//#region Служебные функции

//  Функция обработки ошибок
function errFinish(str) {
    console.log(str);
    Action.Finish();
};

//#endregion

//#region Функция запроса в Базу данных Firebird

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
async function getDBMaterialInfo(matnames) {

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

        console.log(`Получено записей: ${result.length}`);
        console.log(JSON.stringify(result, null, 2));
        return result;
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

async function getAttendOperation() {

    try {
        const options = settings.db_options || null;

        //  Строка запроса
        const sqlString = `
        SELECT
            aogm.ID_GRM,
            aogm.ID_O,
            aogm.ATT_OUT_NORM,
            aogm.COUNT_ATT
        FROM ATTEND_OPER_GROUP_MAT AS aogm
        `;

        const sqlString2 = `
        WITH RECURSIVE group_operations AS (
            -- Прямые назначения операций на группы
            SELECT DISTINCT
                a.ID_O,
                a.ID_GRM
            FROM ATTEND_OPER_GROUP_MAT a
            
            UNION ALL
            
            -- Распространение на дочерние группы (наследование)
            SELECT 
                go.ID_O,
                gm.ID_GRM
            FROM group_operations go
            INNER JOIN GROUP_MATERIAL gm ON gm.ENTRY = go.ID_GRM
        )
        SELECT 
            ID_O,
            LIST(DISTINCT ID_GRM) AS group_ids
        FROM group_operations
        GROUP BY ID_O
        `;

        const sqlString3 = `
            WITH RECURSIVE group_operations AS (
                SELECT DISTINCT
                    a.ID_O,
                    a.ID_GRM
                FROM ATTEND_OPER_GROUP_MAT a
                
                UNION ALL
                
                SELECT 
                    go.ID_O,
                    gm.ID_GRM
                FROM group_operations go
                INNER JOIN GROUP_MATERIAL gm ON gm.ENTRY = go.ID_GRM
            ),
            ordered_groups AS (
                SELECT 
                    ID_O,
                    ID_GRM
                FROM group_operations
                ORDER BY ID_O, ID_GRM
            )
            SELECT 
                ID_O,
                LIST(ID_GRM) AS group_ids
            FROM ordered_groups
            GROUP BY ID_O
            `;

        const sqlString4 = `
        WITH RECURSIVE group_operations AS (
            -- Базовый уровень: операции, назначенные напрямую на группы
            SELECT 
                a.ID_GRM,
                a.ID_O
            FROM ATTEND_OPER_GROUP_MAT a
            
            UNION ALL
            
            -- Распространяем операции на дочерние группы
            SELECT 
                gm.ID_GRM,
                go.ID_O
            FROM group_operations go
            INNER JOIN GROUP_MATERIAL gm ON gm.ENTRY = go.ID_GRM
        )
        SELECT 
            ID_GRM,
            LIST(DISTINCT ID_O) AS operation_ids
        FROM group_operations
        GROUP BY ID_GRM
        ORDER BY ID_GRM
        `;
        //  Запрос в БД
        const res = await executeQuery(sqlString4, options, []);

        let result = {};

        res.forEach(gr => {
            result[gr.id_grm] = gr.operation_ids.split(',')
                .map(id => Number(id));
        });

        //console.log(`Получено записей: ${res.length}`);
        console.log(JSON.stringify(res, null, 2));
        return result;
    } catch (e) {
        console.error("Ошибка:", e.message);
        Action.Finish();
    };
};

//#endregion

async function main() {
    let orepationIdsByGroup = await getAttendOperation();

    //console.log(JSON.stringify(orepationIdsByGroup, null, 2));

    Action.Finish();
};

main();


Action.Continue();