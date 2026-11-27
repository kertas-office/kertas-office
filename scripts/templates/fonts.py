"""Fonts for the Slides templates: open-licence (SIL OFL) families from Google Fonts, cut to
static Regular and Bold faces with the Latin letters, and packed as EOT for embedding in a .pptx
(<p:embeddedFontLst>), so a deck looks the same on any computer, in Kertas and in PowerPoint.
The cut faces are kept in scripts/templates/fonts/ with their licence; this file only rebuilds them."""
import io, os, struct, urllib.request
from fontTools.ttLib import TTFont
from fontTools.varLib import instancer
from fontTools import subset

HERE = os.path.join(os.path.dirname(__file__), 'fonts')
SRC = 'https://raw.githubusercontent.com/google/fonts/main/ofl/'
# family: (file in google/fonts, regular weight, bold weight or None for a single-face family)
FAMILIES = {
    'Plus Jakarta Sans': ('plusjakartasans/PlusJakartaSans%5Bwght%5D.ttf', 400, 800),
    'Manrope': ('manrope/Manrope%5Bwght%5D.ttf', 400, 800),
    'Inter': ('inter/Inter%5Bopsz,wght%5D.ttf', 400, 700),
    'Space Grotesk': ('spacegrotesk/SpaceGrotesk%5Bwght%5D.ttf', 400, 700),
    'Outfit': ('outfit/Outfit%5Bwght%5D.ttf', 400, 700),
    'Fraunces': ('fraunces/Fraunces%5BSOFT,WONK,opsz,wght%5D.ttf', 400, 700),
    'DM Serif Display': ('dmserifdisplay/DMSerifDisplay-Regular.ttf', None, None),
    'Instrument Serif': ('instrumentserif/InstrumentSerif-Regular.ttf', None, None),
}
LATIN = list(range(0x20, 0x7F)) + list(range(0xA0, 0x180)) + [0x2013, 0x2014, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2026, 0x20AC, 0x2122, 0x2190, 0x2192, 0x2191, 0x2193, 0x2713, 0x00B7, 0x2212, 0x2009, 0x202F, 0x00D7]


def face_path(family, bold):
    return os.path.join(HERE, f"{family.replace(' ', '')}-{'Bold' if bold else 'Regular'}.ttf")


def build():
    os.makedirs(HERE, exist_ok=True)
    for family, (file, reg, bold) in FAMILIES.items():
        raw = urllib.request.urlopen(SRC + file).read()
        for is_bold, wght in ((False, reg), (True, bold)):
            if is_bold and bold is None:
                continue
            f = TTFont(io.BytesIO(raw))
            if 'fvar' in f:
                axes = {a.axisTag: a.defaultValue for a in f['fvar'].axes}
                axes['wght'] = wght
                if 'opsz' in axes:
                    axes['opsz'] = 32 if family == 'Inter' else axes['opsz']
                f = instancer.instantiateVariableFont(f, axes)
            opts = subset.Options()
            opts.layout_features = ['kern', 'liga', 'calt', 'ccmp', 'locl', 'mark', 'mkmk', 'tnum', 'lnum']
            opts.hinting = False
            opts.name_IDs = [0, 1, 2, 3, 4, 5, 6, 13, 14]
            opts.notdef_outline = True
            sub = subset.Subsetter(opts)
            sub.populate(unicodes=LATIN)
            sub.subset(f)
            style = 'Bold' if is_bold else 'Regular'
            name = f['name']
            for rec in list(name.names):
                if rec.nameID in (16, 17, 21, 22, 25):
                    name.removeNames(nameID=rec.nameID)
            name.setName(family, 1, 3, 1, 0x409)
            name.setName(style, 2, 3, 1, 0x409)
            name.setName(f'{family} {style}', 4, 3, 1, 0x409)
            name.setName(f"{family.replace(' ', '')}-{style}", 6, 3, 1, 0x409)
            os2 = f['OS/2']
            os2.usWeightClass = 700 if is_bold else 400
            os2.fsSelection = (os2.fsSelection & ~0b1100001) | (0b100000 if is_bold else 0b1000000)
            os2.fsType = 0
            f['head'].macStyle = 1 if is_bold else 0
            f.save(face_path(family, is_bold))
            print(family, style, os.path.getsize(face_path(family, is_bold)))
        licence = urllib.request.urlopen(SRC + file.split('/')[0] + '/OFL.txt').read()
        open(os.path.join(HERE, f"{family.replace(' ', '')}-OFL.txt"), 'wb').write(licence)


def eot(path):
    """An uncompressed EOT (version 2.1) holding the TrueType file"""
    data = open(path, 'rb').read()
    f = TTFont(path)
    os2, name = f['OS/2'], f['name']
    p = os2.panose
    panose = bytes([p.bFamilyType, p.bSerifStyle, p.bWeight, p.bProportion, p.bContrast, p.bStrokeVariation, p.bArmStyle, p.bLetterForm, p.bMidline, p.bXHeight])

    def s(text):
        b = (text or '').encode('utf-16-le')
        return struct.pack('<H', len(b)) + b
    body = panose + struct.pack('<BBIHH', 1, 1 if os2.fsSelection & 1 else 0, os2.usWeightClass, os2.fsType, 0x504C)
    body += struct.pack('<IIIIII', os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3, os2.ulUnicodeRange4, os2.ulCodePageRange1, os2.ulCodePageRange2)
    body += struct.pack('<I', f['head'].checkSumAdjustment) + b'\0' * 16 + b'\0\0'
    body += s(name.getDebugName(1)) + b'\0\0' + s(name.getDebugName(2)) + b'\0\0' + s(name.getDebugName(5)) + b'\0\0' + s(name.getDebugName(4))
    body += b'\0\0' + struct.pack('<H', 0)  # Padding5, RootStringSize
    size = 16 + len(body) + len(data)
    return struct.pack('<IIII', size, len(data), 0x00020001, 0) + body + data


if __name__ == '__main__':
    build()
