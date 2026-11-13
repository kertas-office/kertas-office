// Icon helper backed by Lucide (ISC). Only the icons listed here are bundled.
import {
  createElement, Save, SaveAll, FolderOpen, FilePlus2, FileText, Sheet, FileType2, Printer,
  FileDown, X, Sun, Moon, House, Plus, Search, Undo2, Redo2, ZoomIn, ZoomOut, Maximize2,
  MousePointer2, Type, Highlighter, Square, PenLine, Signature, Eraser, Files, RotateCcw,
  RotateCw, Trash2, ArrowUp, ArrowDown, ArrowLeft, ArrowRight, Eye, PenTool, ListChecks, Sparkles, Merge, Clock,
  ChevronDown, Coffee, Heading1, Heading2, Heading3, Heading, SquareCheck, FoldVertical, UnfoldVertical, Command, Layers, FileSpreadsheet, ScanText, Stamp, StickyNote, Lock, Scissors,
  Image as ImageIcon, LayoutGrid, PanelLeft, Check,
  Underline, Strikethrough, Circle, Slash, MoveUpRight, CalendarDays, TextCursorInput, ListRestart,
  Layers2, FilePlus, Droplet, PanelTop, Hash, Bookmark, LockOpen, SquareSlash, Expand, FileImage,
  ClipboardPaste, Copy, ClipboardType, Superscript, Subscript, CaseUpper, CaseSensitive, RemoveFormatting,
  Pilcrow, PaintBucket, SquareDashed, PanelBottom, Replace, TextSelect, SeparatorHorizontal, SquareSplitVertical,
  Table, Minus, Link, Omega, Frame, RectangleVertical, RectangleHorizontal, File, Columns2, Columns3,
  ListOrdered, TableOfContents, RefreshCw, Footprints, BookOpenText, Captions, Sigma, MessagesSquare, List,
  BookOpen, Focus, Ruler, Percent, Reply, LayoutTemplate, History, ArchiveRestore,
  ChartColumnBig, ChartBarBig, ChartLine, ChartArea, ChartPie, ChartScatter, Donut, Funnel, FunnelX,
  ArrowDownAZ, ArrowDownZA, ArrowUpDown, Snowflake, SquareFunction, Table2, ShieldCheck, Palette, CopyPlus,
  BetweenHorizontalStart, BetweenVerticalStart, TableProperties, Paintbrush, Grid3x3, MessageSquarePlus,
  EyeOff, Calculator, WandSparkles, Rows3, Eraser as EraserIcon, ListFilter, Group, ScanEye, Pin,
  PilcrowLeft, PilcrowRight, IndentIncrease, IndentDecrease, SpellCheck, Shapes, Radical, TextQuote, Mail, Mails,
  BookMarked, Library, FileOutput, ScanText as ScanTextIcon, FileSignature, TrendingUp, ListTree, Ungroup, Triangle, Star,
  ArrowBigRight, MessageCircle, BringToFront, SendToBack, Ellipsis, Move, LayersPlus, WrapText, AlignLeft, AlignCenter, AlignRight, Bold, Italic, Baseline, PaintbrushVertical, AlignJustify, AArrowUp, AArrowDown, Grid2x2, TableCellsMerge, DollarSign, DecimalsArrowLeft, DecimalsArrowRight, AlignVerticalJustifyStart, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, ArrowDownUp, Hexagon, Diamond, TextCursor, BadgeCheck, KeyRound, ScanLine, Languages, Type as TypeIcon,
  PinOff, CircleDot, SquareChevronDown, FileArchive, LetterText, Settings2, SlidersHorizontal, PanelsTopLeft,
  Crop, FlipHorizontal2, FlipVertical2, ArrowRightToLine,
} from 'lucide';

const ICONS = {
  'pin-off': PinOff,
  'shield-check': ShieldCheck,
  'field-text': TextCursorInput, 'field-para': LetterText, 'field-check': SquareCheck, 'field-radio': CircleDot, 'field-drop': SquareChevronDown, 'field-list': ListChecks, 'field-props': Settings2,
  compress: FileArchive, slicer: SlidersHorizontal, 'split-panes': PanelsTopLeft,
  save: Save, 'save-as': SaveAll, open: FolderOpen, new: FilePlus2, doc: FileText, sheet: Sheet,
  bold: Bold, italic: Italic, highlighter: Highlighter, 'font-color': Baseline, painter: PaintbrushVertical, 'list-ol': ListOrdered, 'list-ul': List,
  'indent-less': IndentDecrease, 'indent-more': IndentIncrease, justify: AlignJustify, 'line-spacing': ArrowDownUp, grow: AArrowUp, shrink: AArrowDown,
  'change-case': CaseSensitive, fill: PaintBucket, 'borders-all': Grid2x2, merge: TableCellsMerge, currency: DollarSign, 'num-format': Hash,
  'dec-less': DecimalsArrowLeft, 'dec-more': DecimalsArrowRight, 'v-top': AlignVerticalJustifyStart, 'v-mid': AlignVerticalJustifyCenter, 'v-bottom': AlignVerticalJustifyEnd, rotate: RotateCw,
  wrap: WrapText, 'align-left': AlignLeft, 'align-center': AlignCenter, 'align-right': AlignRight,
  front: BringToFront, back: SendToBack, more: Ellipsis, move: Move, 'layer-plus': LayersPlus, duplicate: CopyPlus,
  pdf: FileType2, print: Printer, export: FileDown, close: X, sun: Sun, moon: Moon, home: House,
  plus: Plus, search: Search, undo: Undo2, redo: Redo2, 'zoom-in': ZoomIn, 'zoom-out': ZoomOut,
  fit: Maximize2, select: MousePointer2, text: Type, highlight: Highlighter, rect: Square,
  ink: PenLine, sign: Signature, erase: Eraser, merge: Merge, files: Files, 'rotate-l': RotateCcw,
  'rotate-r': RotateCw, trash: Trash2, up: ArrowUp, down: ArrowDown, left: ArrowLeft, right: ArrowRight, 'up-down': ArrowUpDown, view: Eye, edit: PenTool,
  track: ListChecks, sparkles: Sparkles, clock: Clock, chevron: ChevronDown, command: Command,
  layers: Layers, xlsx: FileSpreadsheet, ocr: ScanText, stamp: Stamp, note: StickyNote, lock: Lock,
  split: Scissors, image: ImageIcon, grid: LayoutGrid, sidebar: PanelLeft, check: Check,
  underline: Underline, strike: Strikethrough, ellipse: Circle, line: Slash, arrow: MoveUpRight,
  calendar: CalendarDays, form: TextCursorInput, 'text-cursor': TextSelect, 'eraser-form': ListRestart, flatten: Layers2,
  'file-plus': FilePlus, droplet: Droplet, header: PanelTop, hash: Hash, bookmark: Bookmark,
  unlock: LockOpen, redact: SquareSlash, 'fit-page': Expand, 'file-image': FileImage,
  paste: ClipboardPaste, cut: Scissors, copy: Copy, 'paste-plain': ClipboardType, superscript: Superscript,
  subscript: Subscript, caps: CaseUpper, smallcaps: CaseSensitive, 'clear-format': RemoveFormatting,
  paragraph: Pilcrow, paint: PaintBucket, border: SquareDashed, 'border-bottom': PanelBottom, 'border-top': PanelTop,
  pilcrow: Pilcrow, replace: Replace, 'select-all': TextSelect, 'page-break': SeparatorHorizontal,
  section: SquareSplitVertical, table: Table, minus: Minus, link: Link, footer: PanelBottom, omega: Omega,
  margins: Frame, orientation: RectangleVertical, portrait: RectangleVertical, landscape: RectangleHorizontal,
  paper: File, columns: Columns2, col1: Square, col3: Columns3, 'line-numbers': ListOrdered, toc: TableOfContents,
  refresh: RefreshCw, footnote: Footprints, endnote: BookOpenText, caption: Captions, 'word-count': Sigma,
  messages: MessagesSquare, list: List, book: BookOpen, focus: Focus, ruler: Ruler, percent: Percent, reply: Reply,
  template: LayoutTemplate, history: History, restore: ArchiveRestore,
  'chart-column': ChartColumnBig, 'chart-bar': ChartBarBig, 'chart-line': ChartLine, 'chart-area': ChartArea,
  'chart-pie': ChartPie, 'chart-donut': Donut, 'chart-scatter': ChartScatter, filter: Funnel, 'filter-x': FunnelX,
  'sort-az': ArrowDownAZ, 'sort-za': ArrowDownZA, sort: ArrowUpDown, freeze: Snowflake, function: SquareFunction,
  pivot: Table2, validation: ShieldCheck, palette: Palette, dedupe: CopyPlus, 'insert-row': BetweenHorizontalStart,
  'insert-col': BetweenVerticalStart, 'table-style': TableProperties, brush: Paintbrush, gridlines: Grid3x3,
  'note-add': MessageSquarePlus, 'eye-off': EyeOff, calculator: Calculator, magic: WandSparkles, rows: Rows3,
  clear: EraserIcon, 'list-filter': ListFilter, group: Group, 'scan-eye': ScanEye, pin: Pin, sigma: Sigma,
  ltr: PilcrowRight, rtl: PilcrowLeft, 'indent-more': IndentIncrease, 'indent-less': IndentDecrease, spell: SpellCheck,
  shapes: Shapes, equation: Radical, quote: TextQuote, mail: Mail, mails: Mails, citation: BookMarked, bibliography: Library,
  'file-out': FileOutput, 'scan-text': ScanTextIcon, 'file-sign': FileSignature, sparkline: TrendingUp, outline: ListTree,
  tabs: ArrowRightToLine, crop: Crop, flip: FlipHorizontal2, 'flip-v': FlipVertical2,
  ungroup: Ungroup, fold: FoldVertical, coffee: Coffee, 'heading-1': Heading1, 'heading-2': Heading2, 'heading-3': Heading3, heading: Heading, 'check-square': SquareCheck, unfold: UnfoldVertical, triangle: Triangle, star: Star, 'arrow-shape': ArrowBigRight, callout: MessageCircle, hexagon: Hexagon,
  diamond: Diamond, 'text-box': TextCursor, verified: BadgeCheck, key: KeyRound, 'scan-line': ScanLine, languages: Languages, font: TypeIcon,
};

// Cell alignment: text lines placed in a box, top/middle/bottom by left/centre/right
const cellAlignIcon = (v, h) => {
  const y = { top: 7.5, middle: 11, bottom: 14.5 }[v];
  const line = (len, dy) => { const x = h === 'left' ? 7 : h === 'right' ? 17 - len : 12 - len / 2; return ['path', { d: `M${x} ${y + dy}h${len}` }]; };
  return [['rect', { width: '18', height: '18', x: '3', y: '3', rx: '2' }], line(10, 0), line(6, 3)];
};
for (const v of ['top', 'middle', 'bottom']) for (const h of ['left', 'center', 'right']) ICONS[`cell-${v}-${h}`] = cellAlignIcon(v, h);

const cache = new Map();

/** Returns an SVG string for the named icon */
export function icon(name, size = 18) {
  const key = `${name}:${size}`;
  if (cache.has(key)) return cache.get(key);
  const def = ICONS[name];
  if (!def) return '';
  const el = createElement(def);
  el.setAttribute('width', size);
  el.setAttribute('height', size);
  el.setAttribute('stroke-width', '1.75');
  el.setAttribute('aria-hidden', 'true');
  const html = el.outerHTML;
  cache.set(key, html);
  return html;
}
