
//#region 
//  Функция получения артикула и названия материала из имени
function getMaterialName(matname) {
    let mName = matname;
    let mArt = "";
    if (mName.indexOf("\r") > 0) {
        mArt = mName.split("\r")[1];
        mName = mName.split("\r")[0];
    };
    return [mName, mArt];
};

//  Функция огругления
function round(a, b) {
    b = b || 0;
    return Math.round(a * Math.pow(10, b)) / Math.pow(10, b);
};

function isStringExcluded(str, patterns) {
    if (!str || typeof str !== 'string') return false;

    return patterns.some(pattern => {
        if (typeof pattern !== 'string') return false;

        // Если нет звездочки - только точное совпадение
        if (!pattern.includes('*')) {
            return str === pattern;
        }

        // Обработка звездочек
        // Случай: звездочка в конце (например "R*", "Фаска*")
        if (pattern.endsWith('*') && !pattern.slice(0, -1).includes('*')) {
            const prefix = pattern.slice(0, -1);
            return str.startsWith(prefix);
        }

        // Случай: звездочка в начале (например "*пласти*")
        if (pattern.startsWith('*') && !pattern.slice(1).includes('*')) {
            const suffix = pattern.slice(1);
            return str === suffix; // ТОЧНОЕ совпадение для "*пласти"
        }

        // Случай: звездочка в начале и в конце (например "*пласти*")
        if (pattern.startsWith('*') && pattern.endsWith('*')) {
            const middle = pattern.slice(1, -1);
            return str.includes(middle);
        }

        // Случай: звездочка посередине (например "R*лка")
        const regexPattern = pattern.replace(/\*/g, '.*');
        const regex = new RegExp(`^${regexPattern}$`);
        return regex.test(str);
    });
}

function testNameRegExp(str, patterns) {
    if (!str || typeof str !== 'string') return false;
    return patterns.some(pattern => {
        if (typeof pattern !== 'string') return false;

        // Если нет звездочки - только точное совпадение
        if (!pattern.includes('*')) return str === pattern;

        // Обработка звездочек
        // Случай: звездочка в конце (например "R*", "Фаска*")
        if (pattern.endsWith('*') && !pattern.slice(0, -1).includes('*')) {
            const prefix = pattern.slice(0, -1);
            return str.startsWith(prefix);
        };

        // Случай: звездочка в начале (например "*пласти*")
        if (pattern.startsWith('*') && !pattern.slice(1).includes('*')) {
            const suffix = pattern.slice(1);
            return str === suffix; // ТОЧНОЕ совпадение для "*пласти"
        };

        // Случай: звездочка в начале и в конце (например "*пласти*")
        if (pattern.startsWith('*') && pattern.endsWith('*')) {
            const middle = pattern.slice(1, -1);
            return str.includes(middle);
        };

        // Случай: звездочка посередине (например "R*лка")
        const regexPattern = pattern.replace(/\*/g, '.*');
        const regex = new RegExp(`^${regexPattern}$`);
        return regex.test(str);
    });
};

const cutNames = [
    "**R*",
    "ромка",
    "Фаска*",
    "**пласти*",
    "Евро"
]

//#endregion

function main() {

    let panel = Model.Selected;
    if (!panel) Action.Finish();

    const result = [];

    for (let i = 0; i < panel.Cuts.Count; i++) {

        const cut = panel.Cuts[i];
        if (testNameRegExp(cut.Name, cutNames)) continue;

        if (!cut.Params) {
            //  Паз выемка

            const w = cut.Contour.Width;
            const h = cut.Contour.Height;

            let area = w * h * 0.000001;

            result.push({
                materialSyncExternal: "",       //  Код синхронизации (DB)
                materialUnit: "",               //  Единица измерения (DB)
                name: cut.Name,                 //  Имя паза
                sign: cut.Sign,                 //  Обозначение паза
                cutType: 11,    //  Тип паза
                area: area,
                length: 0
            });
            console.log(cut.Name);
        } else {

            if (
                cut.Params.CutType == 8 ||
                cut.Params.CutType == 10
            ) continue;

            const length = round(cut.Trajectory.ObjLength(), 2);
            result.push({
                materialSyncExternal: "",       //  Код синхронизации (DB)
                materialUnit: "",               //  Единица измерения (DB)
                name: cut.Name,                 //  Имя паза
                sign: cut.Sign,                 //  Обозначение паза
                cutType: cut.Params.CutType,    //  Тип паза
                length: length                  //  Длина траектории паза
            });
        };
    };
    console.log('*********');


    //return result;
    // panel.Butts.forEach(elem => {
    //     //if (!elem) return;
    //     

    // });
    console.log(JSON.stringify(result, null, 2));

    Action.Finish();
};

main();
Action.Continue();