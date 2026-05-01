
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

    for (let i = 0; i < panel.Cuts.Count; i++) {

        const cut = panel.Cuts[i];
        const length = round(cut.Trajectory.ObjLength(), 2);

        result.push({
            materialSyncExternal: "",       //  Код синхронизации (DB)
            materialUnit: "",               //  Единица измерения (DB)
            name: cut.Name,                 //  Имя паза
            sign: cut.Sign,                 //  Обозначение паза
            //cutType: cut.CutType,           //  Тип паза
            //width: cut.Width,               //  Ширина паза
            length: length                  //  Длина траектории паза
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