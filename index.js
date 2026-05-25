const pptxgen = require("pptxgenjs");

const W = 3.58;
const H = 2.17;
const BLUE = "2563EB";
const BLACK = "0A0A0A";
const GRAY  = "888888";
const LGRAY = "DDDDDD";

let pres = new pptxgen();
pres.defineLayout({ name: 'CARD', width: W, height: H });
pres.layout = 'CARD';

// ━━━━━━━━━━━━━━━━━━━━
// 表面
// ━━━━━━━━━━━━━━━━━━━━
let f = pres.addSlide();

// 白背景
f.addShape(pres.ShapeType.rect, { x:0,y:0,w:W,h:H, fill:{color:'FFFFFF'}, line:{color:'FFFFFF'} });

// ── Logo: Leica/Zeiss スタイル（3重同心円） ──
// 3重リング = カメラレンズの断面。シンプルで即わかる。
const lx = 0.23, ly = 0.20, lr = 0.155;

// 外リング
f.addShape(pres.ShapeType.ellipse, {
  x: lx - lr, y: ly - lr, w: lr*2, h: lr*2,
  fill:{type:'none'}, line:{color: BLACK, width: 2.0}
});
// 中リング
f.addShape(pres.ShapeType.ellipse, {
  x: lx - lr*0.62, y: ly - lr*0.62, w: lr*1.24, h: lr*1.24,
  fill:{type:'none'}, line:{color: BLACK, width: 1.2}
});
// 内円（塗り）
f.addShape(pres.ShapeType.ellipse, {
  x: lx - lr*0.28, y: ly - lr*0.28, w: lr*0.56, h: lr*0.56,
  fill:{color: BLUE}, line:{color: BLUE}
});

// ── Wordmark ──
f.addText([
  { text: "Gym",  options: { color: BLACK, bold: true, fontSize: 14, fontFace: "Arial Black" } },
  { text: "Lens", options: { color: BLUE,  bold: true, fontSize: 14, fontFace: "Arial Black" } },
], { x: 0.46, y: 0.115, w: 1.8, h: 0.28, align:"left", margin:0 });

// tagline
f.addText("AI Fitness Technology", {
  x: 0.46, y: 0.40, w: 2.0, h: 0.15,
  fontSize: 5.5, fontFace: "Arial", color: GRAY,
  align:"left", margin:0, charSpacing: 1.5
});

// ── 区切り線 ──
f.addShape(pres.ShapeType.line, {
  x: 0.20, y: 0.72, w: W-0.38, h: 0,
  line: { color: LGRAY, width: 0.6 }
});

// ── 名前 ──
f.addText("兼村 大暉", {
  x: 0.20, y: 0.82, w: 2.0, h: 0.30,
  fontSize: 15, fontFace: "Arial", color: BLACK,
  bold: true, align:"left", margin:0
});
f.addText("TAIKI KANEMURA", {
  x: 0.20, y: 1.12, w: 2.2, h: 0.16,
  fontSize: 6.5, fontFace: "Arial", color: BLUE,
  bold: true, align:"left", margin:0, charSpacing: 1.8
});

// 肩書き
f.addText("Founder & CEO", {
  x: 0.20, y: 1.30, w: 2.0, h: 0.15,
  fontSize: 6.5, fontFace: "Arial", color: GRAY,
  align:"left", margin:0
});

// ── 区切り線2 ──
f.addShape(pres.ShapeType.line, {
  x: 0.20, y: 1.52, w: W-0.38, h: 0,
  line: { color: LGRAY, width: 0.4 }
});

// ── 連絡先 ──
const contacts = [
  { label:"TEL",  val:"080-9804-6325" },
  { label:"Mail", val:"taikikanemura@gmail.com" },
  { label:"Web",  val:"gymlens.com" },
];
contacts.forEach((c, i) => {
  const y = 1.59 + i * 0.17;
  f.addText(c.label, {
    x:0.20, y, w:0.26, h:0.15,
    fontSize:5.5, fontFace:"Arial", color:BLUE,
    bold:true, align:"left", margin:0
  });
  f.addText(c.val, {
    x:0.50, y, w:2.15, h:0.15,
    fontSize:6.5, fontFace:"Arial", color:"333333",
    align:"left", margin:0
  });
});

// ── QRコード（右下・シンプル）──
const qx=2.92, qy=1.42, qw=0.52;

f.addShape(pres.ShapeType.rect, {
  x:qx, y:qy, w:qw, h:qw,
  fill:{color:'FFFFFF'}, line:{color:LGRAY, width:0.5}
});

// ファインダー描画
function finder(slide, fx, fy, sz) {
  const c = sz/7;
  slide.addShape(pres.ShapeType.rect,{x:fx,y:fy,w:sz,h:sz,fill:{color:BLACK},line:{color:BLACK}});
  slide.addShape(pres.ShapeType.rect,{x:fx+c,y:fy+c,w:sz-c*2,h:sz-c*2,fill:{color:'FFFFFF'},line:{color:'FFFFFF'}});
  slide.addShape(pres.ShapeType.rect,{x:fx+c*2,y:fy+c*2,w:sz-c*4,h:sz-c*4,fill:{color:BLACK},line:{color:BLACK}});
}
const c = qw/9, p = c*0.7, fs = c*2.2;
finder(f, qx+p,        qy+p,        fs);
finder(f, qx+qw-p-fs,  qy+p,        fs);
finder(f, qx+p,        qy+qw-p-fs,  fs);

// データドット
[[4,1],[5,2],[4,2],[6,2],[5,4],[4,4],[6,4],[1,4],[2,4],
 [1,5],[3,5],[5,5],[2,6],[4,6],[6,6],[1,7],[3,7],[5,7]]
.forEach(([col,row])=>{
  f.addShape(pres.ShapeType.rect,{
    x: qx+p+col*c*0.42, y: qy+p+row*c*0.42,
    w:c*0.36, h:c*0.36,
    fill:{color:BLACK}, line:{color:BLACK}
  });
});

// ━━━━━━━━━━━━━━━━━━━━
// 裏面（純黒・ミニマル）
// ━━━━━━━━━━━━━━━━━━━━
let b = pres.addSlide();

// 黒背景
b.addShape(pres.ShapeType.rect, { x:0,y:0,w:W,h:H, fill:{color:BLACK}, line:{color:BLACK} });

// 大きなリングロゴ（装飾・薄く右に）
const bx=W*0.75, by=H*0.5, br=0.75;
b.addShape(pres.ShapeType.ellipse, {
  x:bx-br, y:by-br, w:br*2, h:br*2,
  fill:{type:'none'}, line:{color:'1E1E1E', width: 14}
});
b.addShape(pres.ShapeType.ellipse, {
  x:bx-br*0.62, y:by-br*0.62, w:br*1.24, h:br*1.24,
  fill:{type:'none'}, line:{color:'1A1A1A', width:8}
});
b.addShape(pres.ShapeType.ellipse, {
  x:bx-br*0.28, y:by-br*0.28, w:br*0.56, h:br*0.56,
  fill:{color:'1C2A42'}, line:{color:'1C2A42'}
});

// 左: ロゴ
const blx=0.50, bly=H/2, blr=0.20;
b.addShape(pres.ShapeType.ellipse, {
  x:blx-blr, y:bly-blr, w:blr*2, h:blr*2,
  fill:{type:'none'}, line:{color:'FFFFFF', width:2.0}
});
b.addShape(pres.ShapeType.ellipse, {
  x:blx-blr*0.62, y:bly-blr*0.62, w:blr*1.24, h:blr*1.24,
  fill:{type:'none'}, line:{color:'FFFFFF', width:1.2}
});
b.addShape(pres.ShapeType.ellipse, {
  x:blx-blr*0.28, y:bly-blr*0.28, w:blr*0.56, h:blr*0.56,
  fill:{color:BLUE}, line:{color:BLUE}
});

// ワードマーク
b.addText([
  { text:"Gym",  options:{color:"FFFFFF", bold:true, fontSize:20, fontFace:"Arial Black"} },
  { text:"Lens", options:{color:BLUE,     bold:true, fontSize:20, fontFace:"Arial Black"} },
],{ x:0.82, y:H/2-0.24, w:2.2, h:0.40, align:"left", margin:0 });

b.addText("See Every Rep.", {
  x:0.82, y:H/2+0.18, w:2.2, h:0.18,
  fontSize:7.5, fontFace:"Arial", color:"555555",
  align:"left", margin:0, charSpacing:2.5
});

pres.writeFile({ fileName:"GymLens_Card.pptx" });
console.log("Done!");
