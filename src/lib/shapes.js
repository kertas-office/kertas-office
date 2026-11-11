// Shapes for documents. Each is an SVG path in a 100 x 100 box (or a custom
// aspect), drawn with a fill and outline and inserted as a floating picture.
export const SHAPES = [
  { id: 'rect', label: 'Rectangle', d: 'M4 14h92v72H4z' },
  { id: 'rounded', label: 'Rounded rectangle', d: 'M18 14h64a14 14 0 0 1 14 14v44a14 14 0 0 1-14 14H18A14 14 0 0 1 4 72V28a14 14 0 0 1 14-14z' },
  { id: 'ellipse', label: 'Oval', d: 'M50 14a46 36 0 1 1 0 72a46 36 0 1 1 0-72z' },
  { id: 'circle', label: 'Circle', d: 'M50 6a44 44 0 1 1 0 88a44 44 0 1 1 0-88z' },
  { id: 'triangle', label: 'Triangle', d: 'M50 8L95 90H5z' },
  { id: 'right-triangle', label: 'Right triangle', d: 'M8 8v84h84z' },
  { id: 'diamond', label: 'Diamond', d: 'M50 4l46 46-46 46L4 50z' },
  { id: 'pentagon', label: 'Pentagon', d: 'M50 5l45 33-17 54H22L5 38z' },
  { id: 'hexagon', label: 'Hexagon', d: 'M27 8h46l23 42-23 42H27L4 50z' },
  { id: 'octagon', label: 'Octagon', d: 'M31 4h38l27 27v38L69 96H31L4 69V31z' },
  { id: 'star5', label: 'Star', d: 'M50 4l13.5 29.5 32 3.5-24 21.5 7 31.5L50 74 21.5 90l7-31.5-24-21.5 32-3.5z' },
  { id: 'star6', label: 'Six-point star', d: 'M50 4l11.5 20H85L73 46l12 21H61.5L50 88 38.5 67H15l12-21-12-22h23.5z' },
  { id: 'heart', label: 'Heart', d: 'M50 88C20 66 4 50 4 31A21 21 0 0 1 50 22a21 21 0 0 1 46 9c0 19-16 35-46 57z' },
  { id: 'arrow-right', label: 'Right arrow', d: 'M4 34h56V14l36 36-36 36V66H4z' },
  { id: 'arrow-left', label: 'Left arrow', d: 'M96 34H40V14L4 50l36 36V66h56z' },
  { id: 'arrow-up', label: 'Up arrow', d: 'M34 96V40H14L50 4l36 36H66v56z' },
  { id: 'arrow-down', label: 'Down arrow', d: 'M34 4v56H14l36 36 36-36H66V4z' },
  { id: 'arrow-both', label: 'Double arrow', d: 'M4 50l26-28v16h40V22l26 28-26 28V62H30v16z' },
  { id: 'chevron', label: 'Chevron', d: 'M4 14h62l30 36-30 36H4l30-36z' },
  { id: 'callout', label: 'Speech bubble', d: 'M12 8h76a8 8 0 0 1 8 8v48a8 8 0 0 1-8 8H42L22 92l4-20H12a8 8 0 0 1-8-8V16a8 8 0 0 1 8-8z' },
  { id: 'cloud', label: 'Cloud', d: 'M27 80a19 19 0 0 1-3-38 24 24 0 0 1 45-10 17 17 0 0 1 24 16 16 16 0 0 1-3 32z' },
  { id: 'plus', label: 'Plus', d: 'M36 6h28v30h30v28H64v30H36V64H6V36h30z' },
  { id: 'banner', label: 'Ribbon banner', d: 'M4 24h92l-10 26 10 26H4l10-26z' },
  { id: 'line', label: 'Line', d: 'M4 50h92', line: true },
  { id: 'arrow-line', label: 'Arrow line', d: 'M4 50h84M74 36l16 14-16 14', line: true },
];

export const SHAPE_COLORS = [
  { fill: '#4F46E5', stroke: '#3730A3' }, { fill: '#0EA5E9', stroke: '#0369A1' }, { fill: '#10B981', stroke: '#047857' },
  { fill: '#F59E0B', stroke: '#B45309' }, { fill: '#EF4444', stroke: '#B91C1C' }, { fill: '#FFFFFF', stroke: '#334155' },
];

/** SVG markup for a shape at a given size, with fill, outline and optional centred text */
export function shapeSvg(shape, { width = 200, height = 200, fill = '#4F46E5', stroke = '#3730A3', strokeWidth = 2, text = '', textColor = '#FFFFFF', font = 'Inter, Arial, sans-serif' } = {}) {
  const sw = strokeWidth * (100 / Math.min(width, height));
  const body = shape.line
    ? `<path d="${shape.d}" fill="none" stroke="${stroke}" stroke-width="${Math.max(sw, 3)}" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke"/>`
    : `<path d="${shape.d}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}" stroke-linejoin="round" vector-effect="non-scaling-stroke"/>`;
  const label = text ? `<text x="${width / 2}" y="${height / 2}" fill="${textColor}" font-family="${font}" font-size="${Math.min(height * 0.18, (width * 1.6) / Math.max(4, text.length))}" font-weight="600" text-anchor="middle" dominant-baseline="central">${String(text).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]))}</text>` : '';
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><g transform="scale(${width / 100} ${height / 100})">${body}</g>${label}</svg>`;
}

/** Small preview for the gallery */
export const shapeIcon = (shape) => shapeSvg(shape, { width: 100, height: 100, fill: 'currentColor', stroke: 'currentColor', strokeWidth: 1 }).replace('<svg ', '<svg style="color:var(--accent);fill-opacity:.85" ');

/** Renders a shape to a PNG data URL (2x for sharp printing) in the browser */
export async function shapePng(shape, opts) {
  const { width, height } = opts;
  const svg = shapeSvg(shape, opts);
  const img = new Image();
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await img.decode();
  const c = document.createElement('canvas');
  c.width = width * 2; c.height = height * 2;
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
  return c.toDataURL('image/png');
}
