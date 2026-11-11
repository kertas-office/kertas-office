// Excel's margin presets in points. t, b, l, r: margins; h, f: where the header and
// footer text sit. Shared by the sheet printer and the Excel page setup reader.
const inch = (n) => n * 72;
export const MARGINS = {
  normal: { t: inch(0.75), b: inch(0.75), l: inch(0.7), r: inch(0.7), h: inch(0.3), f: inch(0.3) },
  narrow: { t: inch(0.75), b: inch(0.75), l: inch(0.25), r: inch(0.25), h: inch(0.3), f: inch(0.3) },
  wide: { t: inch(1), b: inch(1), l: inch(1), r: inch(1), h: inch(0.5), f: inch(0.5) },
};
