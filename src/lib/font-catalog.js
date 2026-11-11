// The fonts Kertas ships. All are free (SIL Open Font License or Apache) from
// Google Fonts, packaged as @expo-google-fonts/* TTFs. Used for:
//   - the font lists in Writer and Sheets (shown in their own typeface)
//   - drawing those fonts on screen (@font-face, see font-assets.js)
//   - embedding the same fonts when Sheets exports a PDF
// Microsoft's standard fonts are listed by their own names and drawn with
// metric-compatible stand-ins (same widths), so documents lay out as in Word
// and Excel even on computers that do not have them. Installed copies win.
export const FONT_CATALOG = [
  // Word and Excel standards, with metric-compatible stand-ins
  { family: 'Calibri', pkg: 'carlito', file: 'Carlito', category: 'sans', standard: true },
  { family: 'Cambria', pkg: 'caladea', file: 'Caladea', category: 'serif', standard: true },
  { family: 'Arial', pkg: 'arimo', file: 'Arimo', category: 'sans', standard: true },
  { family: 'Times New Roman', pkg: 'tinos', file: 'Tinos', category: 'serif', standard: true },
  { family: 'Courier New', pkg: 'cousine', file: 'Cousine', category: 'mono', standard: true },
  // Sans serif
  { family: 'Inter', pkg: 'inter', file: 'Inter', category: 'sans' },
  { family: 'Roboto', pkg: 'roboto', file: 'Roboto', category: 'sans' },
  { family: 'Open Sans', pkg: 'open-sans', file: 'OpenSans', category: 'sans' },
  { family: 'Lato', pkg: 'lato', file: 'Lato', category: 'sans' },
  { family: 'Montserrat', pkg: 'montserrat', file: 'Montserrat', category: 'sans' },
  { family: 'Poppins', pkg: 'poppins', file: 'Poppins', category: 'sans' },
  { family: 'Source Sans 3', pkg: 'source-sans-3', file: 'SourceSans3', category: 'sans' },
  { family: 'Nunito', pkg: 'nunito', file: 'Nunito', category: 'sans' },
  { family: 'Work Sans', pkg: 'work-sans', file: 'WorkSans', category: 'sans' },
  { family: 'Raleway', pkg: 'raleway', file: 'Raleway', category: 'sans' },
  { family: 'Noto Sans', pkg: 'noto-sans', file: 'NotoSans', category: 'sans' },
  // Serif
  { family: 'Merriweather', pkg: 'merriweather', file: 'Merriweather', category: 'serif' },
  { family: 'Lora', pkg: 'lora', file: 'Lora', category: 'serif' },
  { family: 'Playfair Display', pkg: 'playfair-display', file: 'PlayfairDisplay', category: 'serif' },
  { family: 'EB Garamond', pkg: 'eb-garamond', file: 'EBGaramond', category: 'serif' },
  { family: 'Libre Baskerville', pkg: 'libre-baskerville', file: 'LibreBaskerville', category: 'serif' },
  { family: 'Source Serif 4', pkg: 'source-serif-4', file: 'SourceSerif4', category: 'serif' },
  // Monospace
  { family: 'JetBrains Mono', pkg: 'jetbrains-mono', file: 'JetBrainsMono', category: 'mono' },
  { family: 'Roboto Mono', pkg: 'roboto-mono', file: 'RobotoMono', category: 'mono' },
  // Handwriting
  { family: 'Caveat', pkg: 'caveat', file: 'Caveat', category: 'hand' },
  { family: 'Dancing Script', pkg: 'dancing-script', file: 'DancingScript', category: 'hand' },
  // Malayalam and Arabic
  { family: 'Manjari', pkg: 'manjari', file: 'Manjari', category: 'sans', script: 'Malayalam' },
  { family: 'Noto Sans Malayalam', pkg: 'noto-sans-malayalam', file: 'NotoSansMalayalam', category: 'sans', script: 'Malayalam' },
  { family: 'Cairo', pkg: 'cairo', file: 'Cairo', category: 'sans', script: 'Arabic' },
  { family: 'Amiri', pkg: 'amiri', file: 'Amiri', category: 'serif', script: 'Arabic' },
  { family: 'Noto Sans Arabic', pkg: 'noto-sans-arabic', file: 'NotoSansArabic', category: 'sans', script: 'Arabic' },
];

export const FACE_DIRS = { regular: '400Regular', italic: '400Regular_Italic', bold: '700Bold', boldItalic: '700Bold_Italic' };
const FALLBACK = { italic: ['regular'], bold: ['regular'], boldItalic: ['bold', 'italic', 'regular'] };
const CSS_FALLBACK = { sans: 'sans-serif', serif: 'serif', mono: 'monospace', hand: 'cursive' };

/** Package-relative path of a face file, e.g. '@expo-google-fonts/roboto/700Bold/Roboto_700Bold.ttf' */
export const facePath = (f, face) => `@expo-google-fonts/${f.pkg}/${FACE_DIRS[face]}/${f.file}_${FACE_DIRS[face]}.ttf`;

/** The catalog entry for a family name (case-insensitive), or null */
export function findFont(family) {
  const k = String(family || '').replace(/["']/g, '').split(',')[0].trim().toLowerCase();
  return FONT_CATALOG.find((f) => f.family.toLowerCase() === k) || null;
}

/** Which face file to use for a style, falling back when a family lacks it. has(path) tells what exists. */
export function pickFace(f, { bold = false, italic = false } = {}, has = () => true) {
  const want = bold && italic ? 'boldItalic' : bold ? 'bold' : italic ? 'italic' : 'regular';
  for (const face of [want, ...(FALLBACK[want] || [])]) if (has(facePath(f, face))) return face;
  return 'regular';
}

export const cssStack = (f) => `'${f.family}', ${CSS_FALLBACK[f.category] || 'sans-serif'}`;
