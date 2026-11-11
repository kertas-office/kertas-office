// Node-side font loader for tests and scripts (the app loads the same files by URL)
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
export const FONT_FILES = {
  base: '@expo-google-fonts/noto-sans/400Regular/NotoSans_400Regular.ttf',
  baseBold: '@expo-google-fonts/noto-sans/700Bold/NotoSans_700Bold.ttf',
  arabic: '@expo-google-fonts/noto-sans-arabic/400Regular/NotoSansArabic_400Regular.ttf',
  malayalam: '@expo-google-fonts/noto-sans-malayalam/400Regular/NotoSansMalayalam_400Regular.ttf',
  devanagari: '@expo-google-fonts/noto-sans-devanagari/400Regular/NotoSansDevanagari_400Regular.ttf',
  bengali: '@expo-google-fonts/noto-sans-bengali/400Regular/NotoSansBengali_400Regular.ttf',
  tamil: '@expo-google-fonts/noto-sans-tamil/400Regular/NotoSansTamil_400Regular.ttf',
  thai: '@expo-google-fonts/noto-sans-thai/400Regular/NotoSansThai_400Regular.ttf',
  sc: '@expo-google-fonts/noto-sans-sc/400Regular/NotoSansSC_400Regular.ttf',
  kr: '@expo-google-fonts/noto-sans-kr/400Regular/NotoSansKR_400Regular.ttf',
};
export const nodeFontLoader = async (key) => {
  const path = key.startsWith('face:') ? key.slice(5) : FONT_FILES[key];
  try { return new Uint8Array(await readFile(require.resolve(path))); } catch { return null; }
};
