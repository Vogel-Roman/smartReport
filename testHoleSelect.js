
const EPS = 1e-12;


//#region 

// // Матричные операции
// function transpose4x4(m) {
//     return [
//         m[0], m[4], m[8], m[12],
//         m[1], m[5], m[9], m[13],
//         m[2], m[6], m[10], m[14],
//         m[3], m[7], m[11], m[15]
//     ];
// };

// function multiply4x4(a, b) {
//     const r = new Array(16);
//     for (let i = 0; i < 4; i++) {
//         for (let j = 0; j < 4; j++) {
//             r[i * 4 + j] = a[i * 4] * b[j] + a[i * 4 + 1] * b[4 + j] + a[i * 4 + 2] * b[8 + j] + a[i * 4 + 3] * b[12 + j];
//         };
//     };
//     return r;
// };

// //  Очистка матрицы от приближенных нулей
// function clean(m, eps = EPS) {
//     return m.map(v => Math.abs(v) < eps ? 0 : v);
// };

// // Извлечение матрицы 3x3 из 4x4
// function extractRotation3x3(mat4x4) {
//     return [
//         mat4x4[0], mat4x4[1], mat4x4[2],
//         mat4x4[4], mat4x4[5], mat4x4[6],
//         mat4x4[8], mat4x4[9], mat4x4[10]
//     ];
// };

// // Функция для преобразования вектора
// function transformVector(v, R) {
//     return {
//         x: R[0] * v.x + R[1] * v.y + R[2] * v.z,
//         y: R[3] * v.x + R[4] * v.y + R[5] * v.z,
//         z: R[6] * v.x + R[7] * v.y + R[8] * v.z
//     };
// };

// // Нормализация вектора
// function normalize(v) {
//     const len = Math.hypot(v.x, v.y, v.z);
//     if (len < 1e-12) return { x: 0, y: 0, z: 0 };
//     return { x: v.x / len, y: v.y / len, z: v.z / len };
// };

// // Определение ориентации отверстия относительно плиты
// function getOrientationType(directionInSheetSpace) {
//     const dir = normalize(directionInSheetSpace);
//     const absZ = Math.abs(dir.z);
//     const absXY = Math.sqrt(dir.x * dir.x + dir.y * dir.y);

//     if (absZ > 0.99) {
//         return "PERPENDICULAR";  // Перпендикулярно плоскости
//     } else if (absXY > 0.99) {
//         return "PARALLEL";       // Параллельно плоскости
//     } else {
//         return "ANGLED";         // Под углом
//     };
// };

//#endregion

// // Эта функция анализирует одно отверстие
// function analyzeHole(hole, relativeMatrix, sheetDims) {
//     // Извлекаем матрицу поворота 3x3
//     const R_rel = extractRotation3x3(relativeMatrix);

//     // 1. Преобразуем позицию отверстия в систему плиты
//     const posInSheet = transformVector(hole.position, R_rel);

//     // 2. Преобразуем направление отверстия в систему плиты
//     const dirInSheet = transformVector(hole.direction, R_rel);

//     // 3. Определяем ориентацию
//     const orientation = getOrientationType(dirInSheet);

//     // 4. Проверяем находится ли внутри материала
//     const isInside = checkHoleInsideMaterial(posInSheet, hole, sheetDims);

//     // 5. Дополнительная информация для перпендикулярных отверстий
//     let depthInfo = null;
//     if (orientation === "PERPENDICULAR") {
//         const dirNorm = normalize(dirInSheet);
//         const availableDepth = dirNorm.z > 0
//             ? sheetDims.thickness - posInSheet.z
//             : posInSheet.z;
//         depthInfo = {
//             requiredDepth: hole.depth,
//             availableDepth: availableDepth,
//             isDepthSufficient: availableDepth >= hole.depth - 0.5
//         };
//     }

//     // 6. Для параллельных отверстий - проверка расстояния до краев
//     let edgeInfo = null;
//     if (orientation === "PARALLEL") {
//         const radius = hole.diameter / 2;
//         const dirNorm = normalize(dirInSheet);

//         let distToEdge = null;
//         if (Math.abs(dirNorm.x) > 0.9) {
//             distToEdge = Math.min(posInSheet.y, sheetDims.height - posInSheet.y,
//                 posInSheet.z, sheetDims.thickness - posInSheet.z);
//         } else if (Math.abs(dirNorm.y) > 0.9) {
//             distToEdge = Math.min(posInSheet.x, sheetDims.width - posInSheet.x,
//                 posInSheet.z, sheetDims.thickness - posInSheet.z);
//         }

//         edgeInfo = {
//             distanceToNearestEdge: distToEdge,
//             isSafeFromEdge: distToEdge >= radius + 5 // 5mm минимальный отступ
//         };
//     }

//     return {
//         position: posInSheet,
//         direction: dirInSheet,
//         orientation: orientation,
//         isInsideMaterial: isInside,
//         depthInfo: depthInfo,
//         edgeInfo: edgeInfo,
//         originalHole: hole
//     };
// }

// // Проверка нахождения отверстия в материале
// function checkHoleInsideMaterial(pos, dir, hole, sheet) {
//     const eps = 0.01; // небольшой допуск

//     // Переводим позицию в локальные координаты плиты
//     // Учитываем, что плита обычно: x: 0..width, y: 0..height, z: 0..thickness

//     const sheetLocal = {
//         x: pos.x,
//         y: pos.y,
//         z: pos.z
//     };

//     // Проверка, что середина отверстия внутри плиты
//     const isCenterInside = (
//         sheetLocal.x >= -eps && sheetLocal.x <= sheet.width + eps &&
//         sheetLocal.y >= -eps && sheetLocal.y <= sheet.height + eps &&
//         sheetLocal.z >= -eps && sheetLocal.z <= sheet.thickness + eps
//     );

//     // Для перпендикулярных отверстий проверяем глубину
//     if (Math.abs(dir.z) > 0.99) {
//         const depthNeeded = hole.depth || 10;
//         const availableDepth = (dir.z > 0)
//             ? sheet.thickness - sheetLocal.z
//             : sheetLocal.z;
//         return isCenterInside && availableDepth >= depthNeeded - eps;
//     }

//     // Для параллельных отверстий проверяем отступ от краев
//     if (Math.abs(dir.x) > 0.99 || Math.abs(dir.y) > 0.99) {
//         const radius = hole.diameter / 2;
//         const clearance = 5; // минимальное расстояние до края

//         let isSafe = true;
//         if (Math.abs(dir.x) > 0.99) {
//             // Отверстие вдоль X, нужно проверить расстояние до краев по Y и Z
//             isSafe = isSafe && sheetLocal.y >= radius + clearance;
//             isSafe = isSafe && sheetLocal.y <= sheet.height - radius - clearance;
//             isSafe = isSafe && sheetLocal.z >= radius + clearance;
//             isSafe = isSafe && sheetLocal.z <= sheet.thickness - radius - clearance;
//         } else if (Math.abs(dir.y) > 0.99) {
//             // Отверстие вдоль Y
//             isSafe = isSafe && sheetLocal.x >= radius + clearance;
//             isSafe = isSafe && sheetLocal.x <= sheet.width - radius - clearance;
//             isSafe = isSafe && sheetLocal.z >= radius + clearance;
//             isSafe = isSafe && sheetLocal.z <= sheet.thickness - radius - clearance;
//         }
//         return isCenterInside && isSafe;
//     }

//     return isCenterInside;
// }


function main() {

    let pnl = Model.Selected;
    if (!pnl) Action.Finish();

    // let arr_fstn = pnl.FindConnectedFasteners();
    // if (!arr_fstn) Action.Finish();

    // Получаем размеры плиты
    const sheetDims = {
        width: pnl.ContourWidth,      // замените на правильное свойство
        height: pnl.ContourHeight,    // замените на правильное свойство
        thickness: pnl.Thickness  // замените на правильное свойство
    };

    console.log("=== Размеры плиты ===");
    console.log(`Ширина: ${sheetDims.width}, Высота: ${sheetDims.height}, Толщина: ${sheetDims.thickness}`);

    // Вычисляем относительную матрицу поворота
    //const panelInv = transpose4x4(pnl.RotMatrix);

    // Супер-краткая версия

    let pnl = Model.Selected;
    if (!pnl) Action.Finish();

    let fasteners = pnl.FindConnectedFasteners();
    if (!fasteners) Action.Finish();

    console.log('SSS');

    // fasteners.forEach(fastener => {
    //     fastener.Holes.List.forEach(hole => {
    //         // ВСЕГО ДВЕ СТРОКИ для получения ориентации в системе панели!
    //         const posInPanel = pnl.ToObject(fastener.ToGlobal(hole.Position));
    //         const dirInPanel = pnl.NToObject(fastener.NToGlobal(hole.Direction));

    //         // Анализируем dirInPanel.z
    //         const isPerpendicular = Math.abs(dirInPanel.z) > 0.99;
    //         const isParallel = Math.abs(dirInPanel.z) < 0.01;

    //         console.log(`Отверстие: ${isPerpendicular ? "ВЕРТИКАЛЬНОЕ" : (isParallel ? "ГОРИЗОНТАЛЬНОЕ" : "НАКЛОННОЕ")}`);
    //     });
    // });

    Action.Finish();


    // arr_fstn.forEach((fstn, index) => {

    //     let fstn_RotMatrix = clean(fstn.RotMatrix);

    //     // console.log(JSON.stringify(fstn_RotMatrix));
    //     const relativeMatrix = clean(multiply4x4(panelInv, fstn_RotMatrix));

    //     for (let i = 0; i < fstn.Holes.List.length; i++) {
    //         const hData = fstn.Holes.List[i];
    //         const hole = {
    //             depth: hData.Depth,
    //             diameter: hData.Diameter,
    //             drillMode: hData.DrillMode,
    //             optional: hData.Optional,
    //             radius: hData.Radius,
    //             contour: hData.Contour,
    //             direction: hData.Direction,
    //             position: hData.Position,
    //         };

    //         let res = pnl.ToObject(fstn.Position);

    //         console.log(JSON.stringify(clean(fstn.RotMatrix), null, 2));

    //         let hole_point = {
    //             x: res.x + hole.position.x,
    //             y: res.y + hole.position.y,
    //             z: res.z + hole.position.z,
    //         }

    //         console.log(JSON.stringify(hole_point, null, 2));

    //         //console.log(JSON.stringify(hole, null, 2));

    //         // Анализируем отверстие
    //         //const analysis = analyzeHole(hole, relativeMatrix, sheetDims);

    //     };
    // });

    //Action.Finish();
}

main();
Action.Continue();