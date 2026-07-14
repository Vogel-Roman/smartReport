let panel = Model.Selected;

//  Функция огругления
function round(a, b) {
    b = b || 0;
    return Math.round(a * Math.pow(10, b)) / Math.pow(10, b);
};

//alert(panel.Name);

//  Размеры панели
let w_contour = panel.ContourWidth;
let h_contour = panel.ContourHeight;

if (panel.TextureOrientation == 2) {
    //  Изменено направление текстуры
    w_contour = panel.ContourHeight;
    h_contour = panel.ContourWidth;
};

const ll = [];
const ww = [];

for (let i = 0; i < panel.Contours[0].Count; i++) {
    const contuor = panel.Contours[0][i];
    if (contuor.IsLine() && contuor.ObjLength() >= w_contour * 0.8) {
        if (contuor.Data) ll.push(contuor.Data.Butt ? contuor.Data.Butt.Thickness : 0);
    } else if (contuor.IsLine() && contuor.ObjLength() >= h_contour * 0.8) {
        if (contuor.Data) ww.push(contuor.Data.Butt ? contuor.Data.Butt.Thickness : 0);
    };
};

console.log(JSON.stringify([...ll, ...ww]));

Action.Finish();