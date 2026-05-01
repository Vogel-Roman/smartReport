
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

//#endregion

function main() {

    let panel = Model.Selected;
    if (!panel) Action.Finish();

    const result = [];

    for (let i = 0; i < panel.Contour.Count; i++) {
        if (!panel.Contour[i].Data.Butt) continue;
        const elem = panel.Contour[i].Data.Butt;

        const overhung = elem.Overhung;
        const length = round(panel.Contour[i].ObjLength(), 2) + overhung * 2;

        const material = getMaterialName(elem.Material);
        result.push({
            material: elem.Material,        //  Имя материала кромки
            materialName: material[0],      //  Имя материала
            materialArticle: material[1],   //  Артикул материала кромки
            materialSyncExternal: "",       //  Код синхронизации (DB)
            materialUnit: "",               //  Единица измерения (DB)
            allowance: elem.Allowance,      //  Припуск на прифуговку
            clipPanel: elem.ClipPanel,      //  Св-во "подрезать панель"
            sign: elem.Sign,                //  Обозначение кромки
            overhung: elem.Overhung,        //  Величина свеса кромки
            thickness: elem.Thickness,      //  Толщина кромки
            width: elem.Width,              //  Ширина кромки
            length: length
        });
        console.log('---');

    };

    //return result;
    // panel.Butts.forEach(elem => {
    //     //if (!elem) return;
    //     

    // });
    Action.Finish();
};

main();
Action.Continue();