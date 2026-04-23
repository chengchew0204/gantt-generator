// Comprehensive palette of shapes mirroring Excel's Insert > Shapes menu.
// Each preset maps 1:1 to an ECMA-376 ST_ShapeType via the `prstGeom`
// string so the OOXML DrawingML codec round-trips losslessly.
//
// Runtime rendering: every preset exposes a `pathFn(w, h)` that returns an
// SVG path `d` string. Lines render stroke-only (fill ignored); every
// other preset renders as a filled path.

const TWO_PI = Math.PI * 2;

function polygon(points) {
  return 'M' + points.map(([x, y]) => `${x},${y}`).join(' L') + ' Z';
}

function regularPolygonPath(w, h, sides, rotationDeg = -90) {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const offset = (rotationDeg * Math.PI) / 180;
  const pts = [];
  for (let i = 0; i < sides; i++) {
    const angle = offset + (i * TWO_PI) / sides;
    pts.push([cx + rx * Math.cos(angle), cy + ry * Math.sin(angle)]);
  }
  return polygon(pts);
}

function starPath(w, h, points, innerRatio, rotationDeg = -90) {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const offset = (rotationDeg * Math.PI) / 180;
  const pts = [];
  const total = points * 2;
  for (let i = 0; i < total; i++) {
    const radiusX = i % 2 === 0 ? rx : rx * innerRatio;
    const radiusY = i % 2 === 0 ? ry : ry * innerRatio;
    const angle = offset + (i * Math.PI) / points;
    pts.push([cx + radiusX * Math.cos(angle), cy + radiusY * Math.sin(angle)]);
  }
  return polygon(pts);
}

function roundRectPath(w, h, r) {
  const rr = Math.max(0, Math.min(r, Math.min(w, h) / 2));
  return (
    `M${rr},0 L${w - rr},0 Q${w},0 ${w},${rr} ` +
    `L${w},${h - rr} Q${w},${h} ${w - rr},${h} ` +
    `L${rr},${h} Q0,${h} 0,${h - rr} ` +
    `L0,${rr} Q0,0 ${rr},0 Z`
  );
}

function ellipsePath(w, h) {
  return (
    `M0,${h / 2} A${w / 2},${h / 2} 0 1,0 ${w},${h / 2} ` +
    `A${w / 2},${h / 2} 0 1,0 0,${h / 2} Z`
  );
}

function linePath(w, h) {
  return `M0,0 L${w},${h}`;
}

function elbowConnectorPath(w, h) {
  return `M0,0 L${w},0 L${w},${h}`;
}

function curvedLinePath(w, h) {
  return `M0,${h} Q${w / 2},0 ${w},${h}`;
}

function straightArrowPath(w, h) {
  return `M0,${h} L${w},0`;
}

function rightArrowPath(w, h) {
  const bodyTop = h * 0.25;
  const bodyBottom = h * 0.75;
  const headStart = w * 0.7;
  return polygon([
    [0, bodyTop],
    [headStart, bodyTop],
    [headStart, 0],
    [w, h / 2],
    [headStart, h],
    [headStart, bodyBottom],
    [0, bodyBottom],
  ]);
}

function leftArrowPath(w, h) {
  const bodyTop = h * 0.25;
  const bodyBottom = h * 0.75;
  const headEnd = w * 0.3;
  return polygon([
    [w, bodyTop],
    [headEnd, bodyTop],
    [headEnd, 0],
    [0, h / 2],
    [headEnd, h],
    [headEnd, bodyBottom],
    [w, bodyBottom],
  ]);
}

function upArrowPath(w, h) {
  const bodyL = w * 0.25;
  const bodyR = w * 0.75;
  const headEnd = h * 0.3;
  return polygon([
    [bodyL, h],
    [bodyL, headEnd],
    [0, headEnd],
    [w / 2, 0],
    [w, headEnd],
    [bodyR, headEnd],
    [bodyR, h],
  ]);
}

function downArrowPath(w, h) {
  const bodyL = w * 0.25;
  const bodyR = w * 0.75;
  const headStart = h * 0.7;
  return polygon([
    [bodyL, 0],
    [bodyR, 0],
    [bodyR, headStart],
    [w, headStart],
    [w / 2, h],
    [0, headStart],
    [bodyL, headStart],
  ]);
}

function leftRightArrowPath(w, h) {
  const bodyTop = h * 0.25;
  const bodyBottom = h * 0.75;
  const headLeft = w * 0.2;
  const headRight = w * 0.8;
  return polygon([
    [0, h / 2],
    [headLeft, 0],
    [headLeft, bodyTop],
    [headRight, bodyTop],
    [headRight, 0],
    [w, h / 2],
    [headRight, h],
    [headRight, bodyBottom],
    [headLeft, bodyBottom],
    [headLeft, h],
  ]);
}

function upDownArrowPath(w, h) {
  const bodyL = w * 0.25;
  const bodyR = w * 0.75;
  const headTop = h * 0.2;
  const headBottom = h * 0.8;
  return polygon([
    [w / 2, 0],
    [w, headTop],
    [bodyR, headTop],
    [bodyR, headBottom],
    [w, headBottom],
    [w / 2, h],
    [0, headBottom],
    [bodyL, headBottom],
    [bodyL, headTop],
    [0, headTop],
  ]);
}

function chevronPath(w, h) {
  return polygon([
    [0, 0],
    [w * 0.75, 0],
    [w, h / 2],
    [w * 0.75, h],
    [0, h],
    [w * 0.25, h / 2],
  ]);
}

function homePlatePath(w, h) {
  return polygon([
    [0, 0],
    [w * 0.75, 0],
    [w, h / 2],
    [w * 0.75, h],
    [0, h],
  ]);
}

function parallelogramPath(w, h) {
  const sh = w * 0.25;
  return polygon([[sh, 0], [w, 0], [w - sh, h], [0, h]]);
}

function trapezoidPath(w, h) {
  return polygon([[w * 0.25, 0], [w * 0.75, 0], [w, h], [0, h]]);
}

function diamondPath(w, h) {
  return polygon([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
}

function trianglePath(w, h) {
  return polygon([[w / 2, 0], [w, h], [0, h]]);
}

function rtTrianglePath(w, h) {
  return polygon([[0, 0], [0, h], [w, h]]);
}

function plusPath(w, h) {
  const a = w / 3;
  const b = h / 3;
  return polygon([
    [a, 0], [w - a, 0], [w - a, b], [w, b], [w, h - b], [w - a, h - b],
    [w - a, h], [a, h], [a, h - b], [0, h - b], [0, b], [a, b],
  ]);
}

function piePath(w, h) {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const endAngle = Math.PI * 1.75;
  const endX = cx + rx * Math.cos(endAngle);
  const endY = cy + ry * Math.sin(endAngle);
  return `M${cx},${cy} L${cx + rx},${cy} A${rx},${ry} 0 1,1 ${endX},${endY} Z`;
}

function chordPath(w, h) {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const startAngle = Math.PI * 0.25;
  const endAngle = Math.PI * 1.75;
  const sx = cx + rx * Math.cos(startAngle);
  const sy = cy + ry * Math.sin(startAngle);
  const ex = cx + rx * Math.cos(endAngle);
  const ey = cy + ry * Math.sin(endAngle);
  return `M${sx},${sy} A${rx},${ry} 0 1,1 ${ex},${ey} Z`;
}

function teardropPath(w, h) {
  const r = Math.min(w, h) / 2;
  return (
    `M${w},${h / 2} ` +
    `A${r},${r} 0 0,1 ${w / 2},${h} ` +
    `A${w / 2},${h / 2} 0 0,1 0,${h / 2} ` +
    `A${w / 2},${h / 2} 0 0,1 ${w / 2},0 ` +
    `L${w},${h / 2} Z`
  );
}

function canPath(w, h) {
  const ry = Math.min(h * 0.12, w * 0.2);
  return (
    `M0,${ry} A${w / 2},${ry} 0 0,1 ${w},${ry} ` +
    `L${w},${h - ry} A${w / 2},${ry} 0 0,1 0,${h - ry} Z ` +
    `M0,${ry} A${w / 2},${ry} 0 0,0 ${w},${ry}`
  );
}

function cubePath(w, h) {
  const d = Math.min(w, h) * 0.25;
  return (
    `M0,${d} L${d},0 L${w},0 L${w - d},${d} L${w - d},${h} L0,${h} Z ` +
    `M${w - d},${d} L${w},0 ` +
    `M0,${d} L${w - d},${d}`
  );
}

function heartPath(w, h) {
  const cx = w / 2;
  return (
    `M${cx},${h * 0.3} ` +
    `C${cx},${h * 0.05} ${-w * 0.05},${-h * 0.05} 0,${h * 0.35} ` +
    `C0,${h * 0.55} ${cx},${h * 0.85} ${cx},${h} ` +
    `C${cx},${h * 0.85} ${w},${h * 0.55} ${w},${h * 0.35} ` +
    `C${w * 1.05},${-h * 0.05} ${cx},${h * 0.05} ${cx},${h * 0.3} Z`
  );
}

function snipRectPath(w, h) {
  const s = Math.min(w, h) * 0.15;
  return polygon([
    [0, 0], [w - s, 0], [w, s], [w, h], [0, h],
  ]);
}

function flowChartTerminatorPath(w, h) {
  return roundRectPath(w, h, h / 2);
}

function flowChartDataPath(w, h) {
  return parallelogramPath(w, h);
}

function flowChartPredefinedProcessPath(w, h) {
  const bar = w * 0.08;
  return (
    `M0,0 L${w},0 L${w},${h} L0,${h} Z ` +
    `M${bar},0 L${bar},${h} ` +
    `M${w - bar},0 L${w - bar},${h}`
  );
}

function flowChartDocumentPath(w, h) {
  const wave = h * 0.12;
  return (
    `M0,0 L${w},0 L${w},${h - wave} ` +
    `Q${w * 0.75},${h - wave * 2.5} ${w * 0.5},${h - wave} ` +
    `Q${w * 0.25},${h + wave * 0.5} 0,${h - wave} Z`
  );
}

function flowChartManualInputPath(w, h) {
  return polygon([[0, h * 0.25], [w, 0], [w, h], [0, h]]);
}

function flowChartDelayPath(w, h) {
  return (
    `M0,0 L${w * 0.5},0 ` +
    `A${w * 0.5},${h / 2} 0 0,1 ${w * 0.5},${h} ` +
    `L0,${h} Z`
  );
}

function smileyFacePath(w, h) {
  const rx = w / 2;
  const ry = h / 2;
  const eyeY = h * 0.35;
  const mouthY = h * 0.6;
  const eyeR = Math.min(w, h) * 0.05;
  return (
    `M0,${ry} A${rx},${ry} 0 1,0 ${w},${ry} A${rx},${ry} 0 1,0 0,${ry} Z ` +
    `M${w * 0.3 - eyeR},${eyeY} a${eyeR},${eyeR} 0 1,0 ${eyeR * 2},0 a${eyeR},${eyeR} 0 1,0 ${-eyeR * 2},0 ` +
    `M${w * 0.7 - eyeR},${eyeY} a${eyeR},${eyeR} 0 1,0 ${eyeR * 2},0 a${eyeR},${eyeR} 0 1,0 ${-eyeR * 2},0 ` +
    `M${w * 0.3},${mouthY} Q${w * 0.5},${h * 0.85} ${w * 0.7},${mouthY}`
  );
}

const PRESETS = [
  // Lines - rendered stroke-only.
  { id: 'line', label: 'Line', group: 'Lines', prstGeom: 'line', kind: 'line', pathFn: linePath },
  { id: 'straightArrow', label: 'Arrow', group: 'Lines', prstGeom: 'straightConnector1', kind: 'line', pathFn: straightArrowPath, arrowEnd: 'triangle' },
  { id: 'leftRightArrowLine', label: 'Double Arrow', group: 'Lines', prstGeom: 'straightConnector1', kind: 'line', pathFn: linePath, arrowStart: 'triangle', arrowEnd: 'triangle' },
  { id: 'elbowConnector', label: 'Elbow Connector', group: 'Lines', prstGeom: 'bentConnector3', kind: 'line', pathFn: elbowConnectorPath },
  { id: 'curvedLine', label: 'Curved Line', group: 'Lines', prstGeom: 'curvedConnector3', kind: 'line', pathFn: curvedLinePath },

  // Rectangles
  { id: 'rect', label: 'Rectangle', group: 'Rectangles', prstGeom: 'rect', kind: 'shape', pathFn: (w, h) => polygon([[0, 0], [w, 0], [w, h], [0, h]]) },
  { id: 'roundRect', label: 'Rounded Rectangle', group: 'Rectangles', prstGeom: 'roundRect', kind: 'shape', pathFn: (w, h) => roundRectPath(w, h, Math.min(w, h) * 0.1) },
  { id: 'snipRect', label: 'Snip Corner Rectangle', group: 'Rectangles', prstGeom: 'snip1Rect', kind: 'shape', pathFn: snipRectPath },
  { id: 'textBox', label: 'Text Box', group: 'Rectangles', prstGeom: 'rect', kind: 'shape', pathFn: (w, h) => polygon([[0, 0], [w, 0], [w, h], [0, h]]), isTextBox: true },
  { id: 'verticalTextBox', label: 'Vertical Text Box', group: 'Rectangles', prstGeom: 'rect', kind: 'shape', pathFn: (w, h) => polygon([[0, 0], [w, 0], [w, h], [0, h]]), isTextBox: true, vertical: true },

  // Basic Shapes
  { id: 'ellipse', label: 'Oval', group: 'Basic Shapes', prstGeom: 'ellipse', kind: 'shape', pathFn: ellipsePath },
  { id: 'triangle', label: 'Isosceles Triangle', group: 'Basic Shapes', prstGeom: 'triangle', kind: 'shape', pathFn: trianglePath },
  { id: 'rtTriangle', label: 'Right Triangle', group: 'Basic Shapes', prstGeom: 'rtTriangle', kind: 'shape', pathFn: rtTrianglePath },
  { id: 'parallelogram', label: 'Parallelogram', group: 'Basic Shapes', prstGeom: 'parallelogram', kind: 'shape', pathFn: parallelogramPath },
  { id: 'trapezoid', label: 'Trapezoid', group: 'Basic Shapes', prstGeom: 'trapezoid', kind: 'shape', pathFn: trapezoidPath },
  { id: 'diamond', label: 'Diamond', group: 'Basic Shapes', prstGeom: 'diamond', kind: 'shape', pathFn: diamondPath },
  { id: 'pentagon', label: 'Pentagon', group: 'Basic Shapes', prstGeom: 'pentagon', kind: 'shape', pathFn: (w, h) => regularPolygonPath(w, h, 5) },
  { id: 'hexagon', label: 'Hexagon', group: 'Basic Shapes', prstGeom: 'hexagon', kind: 'shape', pathFn: (w, h) => regularPolygonPath(w, h, 6, 0) },
  { id: 'heptagon', label: 'Heptagon', group: 'Basic Shapes', prstGeom: 'heptagon', kind: 'shape', pathFn: (w, h) => regularPolygonPath(w, h, 7) },
  { id: 'octagon', label: 'Octagon', group: 'Basic Shapes', prstGeom: 'octagon', kind: 'shape', pathFn: (w, h) => regularPolygonPath(w, h, 8, -67.5) },
  { id: 'star5', label: '5-Point Star', group: 'Basic Shapes', prstGeom: 'star5', kind: 'shape', pathFn: (w, h) => starPath(w, h, 5, 0.4) },
  { id: 'plus', label: 'Plus', group: 'Basic Shapes', prstGeom: 'plus', kind: 'shape', pathFn: plusPath },
  { id: 'pie', label: 'Pie', group: 'Basic Shapes', prstGeom: 'pie', kind: 'shape', pathFn: piePath },
  { id: 'chord', label: 'Chord', group: 'Basic Shapes', prstGeom: 'chord', kind: 'shape', pathFn: chordPath },
  { id: 'teardrop', label: 'Teardrop', group: 'Basic Shapes', prstGeom: 'teardrop', kind: 'shape', pathFn: teardropPath },
  { id: 'can', label: 'Cylinder', group: 'Basic Shapes', prstGeom: 'can', kind: 'shape', pathFn: canPath, compound: true },
  { id: 'cube', label: 'Cube', group: 'Basic Shapes', prstGeom: 'cube', kind: 'shape', pathFn: cubePath, compound: true },
  { id: 'heart', label: 'Heart', group: 'Basic Shapes', prstGeom: 'heart', kind: 'shape', pathFn: heartPath },
  { id: 'smileyFace', label: 'Smiley Face', group: 'Basic Shapes', prstGeom: 'smileyFace', kind: 'shape', pathFn: smileyFacePath, compound: true },

  // Block Arrows
  { id: 'rightArrow', label: 'Right Arrow', group: 'Block Arrows', prstGeom: 'rightArrow', kind: 'shape', pathFn: rightArrowPath },
  { id: 'leftArrow', label: 'Left Arrow', group: 'Block Arrows', prstGeom: 'leftArrow', kind: 'shape', pathFn: leftArrowPath },
  { id: 'upArrow', label: 'Up Arrow', group: 'Block Arrows', prstGeom: 'upArrow', kind: 'shape', pathFn: upArrowPath },
  { id: 'downArrow', label: 'Down Arrow', group: 'Block Arrows', prstGeom: 'downArrow', kind: 'shape', pathFn: downArrowPath },
  { id: 'leftRightArrow', label: 'Left-Right Arrow', group: 'Block Arrows', prstGeom: 'leftRightArrow', kind: 'shape', pathFn: leftRightArrowPath },
  { id: 'upDownArrow', label: 'Up-Down Arrow', group: 'Block Arrows', prstGeom: 'upDownArrow', kind: 'shape', pathFn: upDownArrowPath },
  { id: 'chevron', label: 'Chevron', group: 'Block Arrows', prstGeom: 'chevron', kind: 'shape', pathFn: chevronPath },
  { id: 'homePlate', label: 'Pentagon Arrow', group: 'Block Arrows', prstGeom: 'homePlate', kind: 'shape', pathFn: homePlatePath },

  // Flowchart
  { id: 'flowChartProcess', label: 'Process', group: 'Flowchart', prstGeom: 'flowChartProcess', kind: 'shape', pathFn: (w, h) => polygon([[0, 0], [w, 0], [w, h], [0, h]]) },
  { id: 'flowChartAlternateProcess', label: 'Alternate Process', group: 'Flowchart', prstGeom: 'flowChartAlternateProcess', kind: 'shape', pathFn: (w, h) => roundRectPath(w, h, Math.min(w, h) * 0.15) },
  { id: 'flowChartDecision', label: 'Decision', group: 'Flowchart', prstGeom: 'flowChartDecision', kind: 'shape', pathFn: diamondPath },
  { id: 'flowChartTerminator', label: 'Terminator', group: 'Flowchart', prstGeom: 'flowChartTerminator', kind: 'shape', pathFn: flowChartTerminatorPath },
  { id: 'flowChartData', label: 'Data', group: 'Flowchart', prstGeom: 'flowChartInputOutput', kind: 'shape', pathFn: flowChartDataPath },
  { id: 'flowChartPredefinedProcess', label: 'Predefined Process', group: 'Flowchart', prstGeom: 'flowChartPredefinedProcess', kind: 'shape', pathFn: flowChartPredefinedProcessPath, compound: true },
  { id: 'flowChartDocument', label: 'Document', group: 'Flowchart', prstGeom: 'flowChartDocument', kind: 'shape', pathFn: flowChartDocumentPath },
  { id: 'flowChartManualInput', label: 'Manual Input', group: 'Flowchart', prstGeom: 'flowChartManualInput', kind: 'shape', pathFn: flowChartManualInputPath },
  { id: 'flowChartConnector', label: 'Connector', group: 'Flowchart', prstGeom: 'flowChartConnector', kind: 'shape', pathFn: ellipsePath },
  { id: 'flowChartDelay', label: 'Delay', group: 'Flowchart', prstGeom: 'flowChartDelay', kind: 'shape', pathFn: flowChartDelayPath },
];

const PRESET_BY_ID = new Map(PRESETS.map((p) => [p.id, p]));
// First-wins mapping so prstGeom="rect" resolves to the plain rectangle
// preset rather than the later text-box variant. The shape extractor
// promotes a `rect` to `textBox`/`verticalTextBox` based on fill+outline
// and `<a:bodyPr vert="...">` heuristics.
const PRESET_BY_PRSTGEOM = new Map();
for (const p of PRESETS) {
  if (!PRESET_BY_PRSTGEOM.has(p.prstGeom)) PRESET_BY_PRSTGEOM.set(p.prstGeom, p);
}

export const SHAPE_GROUPS = ['Lines', 'Rectangles', 'Basic Shapes', 'Block Arrows', 'Flowchart'];

export function listShapePresets() {
  return PRESETS;
}

export function getShapePreset(id) {
  return PRESET_BY_ID.get(id) || null;
}

export function getPresetByPrstGeom(prstGeom) {
  return PRESET_BY_PRSTGEOM.get(prstGeom) || null;
}

export function groupShapePresets() {
  const out = {};
  for (const g of SHAPE_GROUPS) out[g] = [];
  for (const p of PRESETS) {
    if (out[p.group]) out[p.group].push(p);
  }
  return out;
}

export function renderShapePath(id, w, h) {
  const preset = getShapePreset(id);
  if (!preset) return '';
  return preset.pathFn(Math.max(1, w), Math.max(1, h));
}

export function isLineShape(id) {
  const preset = getShapePreset(id);
  return !!preset && preset.kind === 'line';
}
