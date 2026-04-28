
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

function main() {

    let panel = Model.Selected;
    if (!panel) Action.Finish();

    let fasteners = panel.FindConnectedFasteners();
    if (!fasteners) Action.Finish();

    const panelGab = {
        w: panel.ContourWidth,
        h: panel.ContourHeight,
        t: panel.Thickness
    };

    fasteners.forEach(fastener => {
        if (!fastener) return;

        console.log(fastener.Name);

        fastener.Holes.List.forEach(hole => {
            if (!hole) return;

            // ВСЕГО ДВЕ СТРОКИ для получения ориентации в системе панели!
            const posInPanel = panel.ToObject(fastener.ToGlobal(hole.Position));
            const dirInPanel = panel.NToObject(fastener.NToGlobal(hole.Direction));

            // Анализируем dirInPanel.z
            const isPerpendicular = Math.abs(dirInPanel.z) > 0.99;
            const isParallel = Math.abs(dirInPanel.z) < 0.01;

            let endInPanel = getHoleEndPoint(hole, fastener, panel);

            if (
                (endInPanel.x >= 0 && endInPanel.x <= panelGab.w) &&
                (endInPanel.y >= 0 && endInPanel.y <= panelGab.h) &&
                (endInPanel.z >= 0 && endInPanel.z <= panelGab.t)
            ) {
                console.log(`\n${isPerpendicular ? "ВЕРТИКАЛЬНОЕ" : (isParallel ? "ГОРИЗОНТАЛЬНОЕ" : "НАКЛОННОЕ")}`);
                console.log(`Отверстие D${hole.Radius * 2}x${hole.Depth} мм принадлежит панели`);
            };
        });
        console.log('-------------');

    });
    Action.Finish();
};

main();
Action.Continue();