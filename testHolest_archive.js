const EPS = 1e-10; // Допуск для сравнения с нулем

//  Функция огругления
function round(a, b) {
    b = b || 0;
    return Math.round(a * Math.pow(10, b)) / Math.pow(10, b);
};


function main() {
    function getHoleEndPoint(hole, fastener, panel) {
        // 1. Вычисляем конец отверстия в локальной системе фурнитуры
        const dir = hole.Direction;
        const depth = hole.Depth;

        // Нормализуем направление (на случай если вектор не единичный)
        const len = Math.hypot(dir.x, dir.y, dir.z);
        const normDir = {
            x: dir.x / len,
            y: dir.y / len,
            z: dir.z / len
        };

        // Точка конца: устье + направление * глубина
        const endLocal = {
            x: hole.Position.x + normDir.x * depth,
            y: hole.Position.y + normDir.y * depth,
            z: hole.Position.z + normDir.z * depth
        };

        // 2. Переводим в глобальные координаты
        let endGlobal = fastener.ToGlobal(endLocal);

        // 3. Переводим в локальную систему координат панели
        let endInPanel = panel.ToObject(endGlobal);

        return endInPanel;
    };

    function cleanPoint(point, decimal) {
        const factor = Math.pow(10, decimal);
        const clean = (val) => {
            if (Math.abs(val) < EPS) return 0;
            return Math.round(val * factor) / factor;
        };

        return {
            x: clean(point.x),
            y: clean(point.y),
            z: clean(point.z)
        };
    }

    let panel = Model.Selected;
    if (!panel) Action.Finish();

    let fasteners = panel.FindConnectedFasteners();
    if (!fasteners) Action.Finish();

    const panelGab = {
        w: panel.ContourWidth,
        h: panel.ContourHeight,
        t: panel.Thickness
    };

    const result = [];

    fasteners.forEach(fastener => {
        if (!fastener) return;

        console.log(fastener.Name);
        console.log('---');


        fastener.Holes.List.forEach(hole => {
            if (!hole) return;

            let posInPanel = cleanPoint(panel.ToObject(fastener.ToGlobal(hole.Position)), 1);
            let dirInPanel = panel.NToObject(fastener.NToGlobal(hole.Direction));
            dirInPanel = cleanPoint(dirInPanel, 1);

            // Анализируем dirInPanel.z
            const isPerpendicular = Math.abs(dirInPanel.z) > 0.99;
            const isParallel = Math.abs(dirInPanel.z) < 0.01;

            let endInPanel = cleanPoint(getHoleEndPoint(hole, fastener, panel), 1);

            if (
                (endInPanel.x >= 0 && endInPanel.x <= panelGab.w) &&
                (endInPanel.y >= 0 && endInPanel.y <= panelGab.h) &&
                (endInPanel.z >= 0 && endInPanel.z <= panelGab.t)
            ) {
                result.push({
                    depth: round(hole.Depth, 2),
                    diameter: hole.Radius * 2,
                    depth: round(hole.Depth, 1),
                    drillMode: hole.DrillMode,
                    dirInPanel: dirInPanel,
                    positionInPanel: posInPanel
                });
                //console.log(`\n${isPerpendicular ? "ВЕРТИКАЛЬНОЕ" : (isParallel ? "ГОРИЗОНТАЛЬНОЕ" : "НАКЛОННОЕ")}`);
                console.log(`Отверстие D${hole.Radius * 2}x${hole.Depth} мм принадлежит панели`);
            };
        });
    });
    console.log(JSON.stringify(result, null, 2));
    Action.Finish();
};

main();
Action.Continue();