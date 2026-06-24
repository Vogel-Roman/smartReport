
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

    const letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'];

    const result = [];

    for (let i = 0; i < panel.Plastics.Count; i++) {

        let plastic = panel.Plastics[i];
        console.log('---');

        result.push({
            material: plastic.Material,
            tkn: plastic.Thickness,
            ltr: letters[i]
        })


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