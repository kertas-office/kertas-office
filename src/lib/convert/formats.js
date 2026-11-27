// What the Convert app reads and what it can turn each kind of file into.

export const GROUPS = {
  image: {
    label: 'Picture',
    inputs: ['jpg', 'jpeg', 'png', 'webp', 'avif', 'gif', 'bmp', 'ico', 'svg', 'tif', 'tiff', 'heic', 'heif'],
    outputs: ['jpg', 'png', 'webp', 'avif', 'gif', 'bmp', 'tiff', 'ico', 'pdf'],
    default: 'jpg',
  },
  pdf: {
    label: 'PDF',
    inputs: ['pdf'],
    outputs: ['docx', 'xlsx', 'png', 'jpg', 'txt', 'compressed', 'pages'],
    default: 'docx',
  },
  doc: {
    label: 'Document',
    inputs: ['docx', 'docm', 'dotx', 'doc', 'rtf', 'odt', 'md', 'markdown', 'html', 'htm', 'txt'],
    outputs: ['pdf', 'docx', 'odt', 'rtf', 'html', 'md', 'txt'],
    default: 'pdf',
  },
  video: {
    label: 'Video',
    inputs: ['mp4', 'mov', 'mkv', 'webm', 'avi', 'wmv', 'flv', 'mpeg', 'mpg', 'm4v', '3gp', 'ts', 'mts', 'm2ts', 'ogv'],
    outputs: ['mp4', 'webm', 'mov', 'mkv', 'avi', 'gif', 'mp3', 'm4a', 'wav'],
    default: 'mp4',
    media: true,
  },
  audio: {
    label: 'Audio',
    inputs: ['mp3', 'm4a', 'aac', 'wav', 'flac', 'ogg', 'oga', 'opus', 'wma', 'aiff', 'aif', 'amr'],
    outputs: ['mp3', 'm4a', 'wav', 'flac', 'ogg', 'opus'],
    default: 'mp3',
    media: true,
  },
  sheet: {
    label: 'Spreadsheet',
    inputs: ['xlsx', 'xlsm', 'xls', 'xlsb', 'ods', 'csv', 'tsv', 'json'],
    outputs: ['xlsx', 'ods', 'csv', 'tsv', 'json', 'html', 'pdf'],
    default: 'xlsx',
  },
};

export const LABELS = {
  jpg: 'JPG', png: 'PNG', webp: 'WEBP', avif: 'AVIF', gif: 'GIF', bmp: 'BMP', tiff: 'TIFF', ico: 'ICO (icon)', pdf: 'PDF',
  docx: 'Word (.docx)', xlsx: 'Excel (.xlsx)', odt: 'ODT (OpenDocument)', rtf: 'RTF (rich text)', ods: 'ODS (OpenDocument)',
  html: 'Web page (.html)', md: 'Markdown (.md)', txt: 'Plain text (.txt)', csv: 'CSV', tsv: 'TSV (tab separated)', json: 'JSON',
  compressed: 'Smaller PDF (compress)', pages: 'One PDF per page (split)',
  mp4: 'MP4 video', webm: 'WebM video', mov: 'MOV video', mkv: 'MKV video', avi: 'AVI video', mp3: 'MP3 audio', m4a: 'M4A audio (AAC)',
  wav: 'WAV audio', flac: 'FLAC audio (lossless)', ogg: 'OGG audio', opus: 'Opus audio',
};

/** Outputs only LibreOffice can write */
export const needsLibreOffice = (to) => to === 'odt' || to === 'rtf';

const SAME = { jpeg: 'jpg', tif: 'tiff', htm: 'html', markdown: 'md', heif: 'heic', docm: 'docx', dotx: 'docx', xlsm: 'xlsx', mpg: 'mpeg', oga: 'ogg', aif: 'aiff' };
/** One name per format: jpeg is jpg, tif is tiff... */
export const canonical = (ext) => SAME[ext] || ext;

export const extOf = (name) => String(name).toLowerCase().split('.').pop();
export const baseOf = (name) => String(name).replace(/\.[^.\\/]+$/, '');

export function groupOf(name) {
  const ext = extOf(name);
  for (const [id, g] of Object.entries(GROUPS)) if (g.inputs.includes(ext)) return id;
  return null;
}

/** The outputs offered for a file: its group's, less its own format */
export function outputsFor(name) {
  const g = GROUPS[groupOf(name)];
  if (!g) return [];
  const ext = canonical(extOf(name));
  // only Word and the older word-processor formats go to .odt and .rtf (through LibreOffice)
  const plain = ['md', 'html', 'txt'].includes(ext);
  return g.outputs.filter((o) => o !== ext && !(plain && needsLibreOffice(o)));
}

// What people most often turn these into, ahead of their group's usual choice
const FAVOURITE = { xlsx: 'pdf', mp4: 'mp3', png: 'jpg', webp: 'jpg', heic: 'jpg', csv: 'xlsx', html: 'pdf', md: 'pdf', txt: 'pdf' };

/** The output a file starts with: its favourite, the group's default, or the first other one */
export function defaultOutput(name) {
  const outs = outputsFor(name);
  const fav = FAVOURITE[canonical(extOf(name))];
  if (fav && outs.includes(fav)) return fav;
  const g = GROUPS[groupOf(name)];
  return outs.includes(g?.default) ? g.default : outs[0] || null;
}
