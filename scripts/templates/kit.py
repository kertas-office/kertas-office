"""The design kit for the Slides templates: themes, slide layouts and the finishing step that
embeds the fonts and sets the deck's theme fonts. make-slides.py writes the decks with it."""
import io, os, random, re, zipfile
from PIL import Image, ImageDraw, ImageFilter
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

import fonts as fontkit

W, H = Inches(13.333), Inches(7.5)
IN = Inches
L, C, R = PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT
TOP, MID, BOT = MSO_ANCHOR.TOP, MSO_ANCHOR.MIDDLE, MSO_ANCHOR.BOTTOM


def rgb(h):
    return RGBColor.from_string(h.lstrip('#'))


def hex2(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ------------------------------------------------------------------ pictures
def mesh(colors, size=(1600, 900), seed=1, base=None, grain=6, blobs=None, scrim=0.72):
    """A soft mesh gradient: big blurred blobs of the colours over a base colour, with a little grain"""
    rnd = random.Random(seed)
    w, h = size
    img = Image.new('RGB', (w // 4, h // 4), hex2(base or colors[0]))
    d = ImageDraw.Draw(img)
    pts = blobs or [(rnd.uniform(-.1, 1.1), rnd.uniform(-.1, 1.1), rnd.uniform(.35, .7)) for _ in colors[1:]]
    for col, (x, y, r) in zip(colors[1:], pts):
        cx, cy, rr = x * w / 4, y * h / 4, r * w / 4
        d.ellipse([cx - rr, cy - rr * .8, cx + rr, cy + rr * .8], fill=hex2(col))
    img = img.filter(ImageFilter.GaussianBlur(w / 4 * .12)).resize(size, Image.BICUBIC)
    if scrim:
        # the base colour comes back over the left side, where the words sit, so they stay readable
        mask = Image.linear_gradient('L').rotate(-90).resize(size)  # 255 on the left, 0 on the right
        mask = mask.point(lambda v: int(v * scrim))
        img = Image.composite(Image.new('RGB', size, hex2(base or colors[0])), img, mask)
    if grain:
        noise = Image.effect_noise(size, grain * 4).convert('RGB')
        img = Image.blend(img, noise, .035)
    out = io.BytesIO()
    img.save(out, 'JPEG', quality=86, optimize=True)
    out.seek(0)
    return out


def photo(colors, seed=1, size=(1200, 900), shade=False):
    """A stand-in picture: an abstract, photo-like gradient with soft light (to be replaced)"""
    rnd = random.Random(seed)
    w, h = size
    img = Image.new('RGB', (w // 3, h // 3), hex2(colors[0]))
    d = ImageDraw.Draw(img)
    for i in range(7):
        col = hex2(colors[i % len(colors)])
        x, y = rnd.uniform(0, w / 3), rnd.uniform(0, h / 3)
        r = rnd.uniform(.15, .45) * w / 3
        d.ellipse([x - r, y - r, x + r, y + r], fill=col)
    # a horizon of light, like a landscape or a lit room
    d.rectangle([0, h / 3 * rnd.uniform(.55, .7), w / 3, h / 3], fill=hex2(colors[-1]))
    img = img.filter(ImageFilter.GaussianBlur(w / 3 * .06)).resize(size, Image.BICUBIC)
    if shade:
        # darker towards the bottom, for words written over the picture
        mask = Image.linear_gradient('L').resize(size).point(lambda v: int(max(0, v - 90) * 1.25))
        img = Image.composite(Image.new('RGB', size, (10, 10, 12)), img, mask)
    noise = Image.effect_noise(size, 30).convert('RGB')
    img = Image.blend(img, noise, .04)
    out = io.BytesIO()
    img.save(out, 'JPEG', quality=84, optimize=True)
    out.seek(0)
    return out


# ------------------------------------------------------------------ deck
class Deck:
    """t: dict(head, body, bg, surface, ink, muted, line, accent, accent2, accent3, dark, on_dark,
    on_dark_muted, mesh=[...], photo=[...], title)"""

    def __init__(self, t):
        self.t = t
        self.p = Presentation()
        self.p.slide_width, self.p.slide_height = W, H
        self.blank = self.p.slide_layouts[6]
        self.p.core_properties.author = ''
        self.p.core_properties.title = t.get('title', '')
        self.n = 0
        self.fonts = {t['head'], t['body']}

    # -- slides
    def slide(self, bg=None, picture=None):
        s = self.p.slides.add_slide(self.blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb(bg or self.t['bg'])
        if picture is not None:
            pic = s.shapes.add_picture(picture, 0, 0, W, H)
            pic.name = 'Background'
        self.n += 1
        return s

    # -- shapes
    def box(self, s, x, y, w, h, color=None, radius=None, line=None, shape=None, gradient=None, alpha=None):
        shape = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
        r = s.shapes.add_shape(shape, x, y, w, h)
        if gradient:
            r.fill.gradient()
            r.fill.gradient_angle = gradient[2] if len(gradient) > 2 else 45
            stops = r.fill.gradient_stops
            stops[0].color.rgb = rgb(gradient[0]); stops[0].position = 0
            stops[1].color.rgb = rgb(gradient[1]); stops[1].position = 1
        elif color is None:
            r.fill.background()
        else:
            r.fill.solid()
            r.fill.fore_color.rgb = rgb(color)
            if alpha is not None:
                sf = r.fill._xPr.find(qn('a:solidFill'))
                clr = sf[0]
                a = clr.makeelement(qn('a:alpha'), {'val': str(int(alpha * 100000))})
                clr.append(a)
        if line:
            r.line.color.rgb = rgb(line)
            r.line.width = Pt(1)
        else:
            r.line.fill.background()
        # no theme shadow or glow: flat, like current design tools
        sp = r._element.spPr
        sp.append(sp.makeelement(qn('a:effectLst'), {}))
        style = r._element.find(qn('p:style'))
        if style is not None:
            r._element.remove(style)
        return r

    def card(self, s, x, y, w, h, color=None, radius_in=0.22, line=None, gradient=None):
        r = self.box(s, x, y, w, h, color if color is not None else self.t['surface'], radius=True, line=line, gradient=gradient)
        r.adjustments[0] = min(.5, IN(radius_in) / min(w, h))
        return r

    def pill(self, s, x, y, text, fg=None, bg=None, size=12, pad=0.18, line=None, bold=True, alpha=None):
        w = IN(pad * 2 + len(text) * size * 0.0082 + 0.02)
        h = IN(size * 0.028 + 0.14)
        b = self.box(s, x, y, w, h, bg if bg is not None else self.t['surface'], radius=True, line=line, alpha=alpha)
        b.adjustments[0] = .5
        self.text(s, x, y, w, h, text, size, fg or self.t['ink'], bold=bold, align=C, anchor=MID, font=self.t['body'], tracking=40, inset=0)
        return w

    def picture(self, s, stream, x, y, w, h, radius_in=0.22, crop=True):
        pic = s.shapes.add_picture(stream, x, y, w, h)
        if crop:
            iw, ih = pic.image.size
            target = w / h
            src = iw / ih
            if src > target:
                cut = (1 - target / src) / 2
                pic.crop_left = pic.crop_right = cut
            else:
                cut = (1 - src / target) / 2
                pic.crop_top = pic.crop_bottom = cut
        if radius_in:
            geom = pic._element.spPr.find(qn('a:prstGeom'))
            geom.set('prst', 'roundRect')
            av = geom.find(qn('a:avLst'))
            gd = av.makeelement(qn('a:gd'), {'name': 'adj', 'fmla': f'val {int(min(50000, IN(radius_in) / min(w, h) * 100000))}'})
            av.append(gd)
        return pic

    # -- text
    def text(self, s, x, y, w, h, lines, size=18, color=None, bold=False, align=L, anchor=TOP,
             bullets=False, spacing=1.12, after=0, italic=False, font=None, tracking=None, inset=0.04, autofit=False):
        tb = s.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = IN(inset)
        tf.margin_top = tf.margin_bottom = IN(0.02)
        if isinstance(lines, str):
            lines = [lines]
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            para.line_spacing = spacing
            para.space_after = Pt(after)
            parts = line if isinstance(line, list) else [(line, {})]
            if bullets:
                pPr = para._p.get_or_add_pPr()
                pPr.set('marL', str(IN(0.32)))
                pPr.set('indent', str(-IN(0.32)))
                bu = pPr.makeelement(qn('a:buClr'), {})
                bu.append(bu.makeelement(qn('a:srgbClr'), {'val': self.t['accent'].lstrip('#')}))
                pPr.append(bu)
                bf = pPr.makeelement(qn('a:buFont'), {'typeface': 'Arial'})
                pPr.append(bf)
                pPr.append(pPr.makeelement(qn('a:buChar'), {'char': '•'}))
            for part, st in parts:
                run = para.add_run()
                run.text = part
                f = run.font
                f.name = st.get('font', font or self.t['body'])
                f.size = Pt(st.get('size', size))
                f.bold = st.get('bold', bold)
                f.italic = st.get('italic', italic)
                f.color.rgb = rgb(st.get('color', color or self.t['ink']))
                if tracking is not None or 'tracking' in st:
                    run._r.get_or_add_rPr().set('spc', str(st.get('tracking', tracking)))
        return tb

    def head(self, s, x, y, w, h, text, size=40, color=None, align=L, anchor=TOP, spacing=1.0, tracking=None):
        t = self.t
        bold = not t.get('head_light')
        return self.text(s, x, y, w, h, text, size, color or t['ink'], bold, align, anchor, font=t['head'], spacing=spacing,
                         tracking=tracking if tracking is not None else t.get('head_tracking', -40 if bold else -10))

    def kicker(self, s, x, y, text, color=None, size=12):
        return self.text(s, x, y, IN(8), IN(0.4), text.upper(), size, color or self.t['accent'], True, tracking=180)

    def footer(self, s, label, dark=False):
        col = self.t['on_dark_muted'] if dark else self.t['muted']
        self.text(s, IN(0.7), IN(6.9), IN(8), IN(0.3), label, 10, col, tracking=40)
        self.text(s, IN(11.6), IN(6.9), IN(1.03), IN(0.3), f'{self.n:02d}', 10, col, align=R, tracking=40)

    def title_block(self, s, kicker, title, sub=None, dark=False, w=11.9, size=40):
        t = self.t
        y = 0.62
        if kicker:
            self.kicker(s, IN(0.7), IN(y), kicker, t['accent'] if not dark else t.get('on_dark_accent', t['accent']))
            y += 0.42
        self.head(s, IN(0.7), IN(y), IN(w), IN(1.0), title, size, t['on_dark'] if dark else t['ink'])
        if sub:
            self.text(s, IN(0.72), IN(y + 0.95), IN(w), IN(0.5), sub, 17, t['on_dark_muted'] if dark else t['muted'])

    # -- data
    def table(self, s, x, y, w, rows, widths, size=14, row_h=0.56, head_bg=None, head_fg=None):
        t = self.t
        shape = s.shapes.add_table(len(rows), len(rows[0]), x, y, w, IN(row_h * len(rows)))
        tbl = shape.table
        tblPr = tbl._tbl.tblPr
        for k in ('bandRow', 'firstRow'):
            tblPr.set(k, '0')
        style = tblPr.find(qn('a:tableStyleId'))
        if style is not None:
            style.text = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'  # no style, no grid
        for i, cw in enumerate(widths):
            tbl.columns[i].width = IN(cw)
        for r, row in enumerate(rows):
            tbl.rows[r].height = IN(row_h)
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb((head_bg or t['ink']) if r == 0 else (t['bg'] if r % 2 else t['surface']))
                cell.margin_left = cell.margin_right = IN(0.16)
                cell.vertical_anchor = MID
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                for side in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                    ln = tcPr.makeelement(qn(side), {'w': '0'})
                    ln.append(ln.makeelement(qn('a:noFill'), {}))
                    tcPr.append(ln)
                p = cell.text_frame.paragraphs[0]
                p.text = ''
                run = p.add_run()
                run.text = val
                run.font.name = t['body']
                run.font.size = Pt(size)
                run.font.bold = r == 0
                run.font.color.rgb = rgb((head_fg or t['bg']) if r == 0 else t['ink'])
        return tbl

    def chart(self, s, kind, x, y, w, h, cats, series, colors, legend=None, labels=False, dark=False, number_format=None, gap=60):
        t = self.t
        data = CategoryChartData(number_format=number_format) if number_format else CategoryChartData()
        data.categories = cats
        for name, values in series:
            data.add_series(name, values)
        ch = s.shapes.add_chart(kind, x, y, w, h, data).chart
        ink = t['on_dark_muted'] if dark else t['muted']
        ch.has_legend = legend if legend is not None else len(series) > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(12)
            ch.legend.font.color.rgb = rgb(ink)
            ch.legend.font.name = t['body']
        ch.font.name = t['body']
        ch.font.size = Pt(12)
        ch.font.color.rgb = rgb(ink)
        line_kinds = (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS)
        for i, ser in enumerate(ch.series):
            if kind in line_kinds:
                ser.format.line.color.rgb = rgb(colors[i])
                ser.format.line.width = Pt(3.5)
                ser.smooth = True
                ser.marker.style = 8  # circle
                ser.marker.size = 8
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = rgb(colors[i])
                ser.marker.format.line.color.rgb = rgb(t['bg'] if not dark else t['dark'])
            elif kind in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE):
                for j, pt in enumerate(ser.points):
                    pt.format.fill.solid()
                    pt.format.fill.fore_color.rgb = rgb(colors[j % len(colors)])
                    pt.format.line.color.rgb = rgb(t['dark'] if dark else t['bg'])
                    pt.format.line.width = Pt(2)
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = rgb(colors[i])
                ser.format.line.fill.background()
            if labels:
                dl = ser.data_labels
                dl.show_value = True
                dl.font.size = Pt(12)
                dl.font.bold = True
                dl.font.name = t['body']
                dl.font.color.rgb = rgb(t['on_dark'] if dark else t['ink'])
                if number_format:
                    dl.number_format = number_format
                    dl.number_format_is_linked = False
                if kind not in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE):
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END if kind not in line_kinds else XL_LABEL_POSITION.ABOVE
        if kind not in (XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.PIE):
            va, ca = ch.value_axis, ch.category_axis
            va.has_major_gridlines = True
            va.major_gridlines.format.line.color.rgb = rgb(t['line'] if not dark else t.get('dark_line', '333333'))
            va.major_gridlines.format.line.width = Pt(0.75)
            va.format.line.fill.background()
            va.tick_labels.font.size = Pt(11)
            ca.format.line.color.rgb = rgb(t['line'] if not dark else t.get('dark_line', '333333'))
            ca.tick_labels.font.size = Pt(12)
            from pptx.enum.chart import XL_TICK_MARK
            va.major_tick_mark = XL_TICK_MARK.NONE
            ca.major_tick_mark = XL_TICK_MARK.NONE
            if labels:
                va.visible = False
                va.has_major_gridlines = False
            try:
                ch.plots[0].gap_width = gap
                ch.plots[0].overlap = -10 if len(series) > 1 else 0
            except Exception:
                pass
        else:
            ch.plots[0].has_data_labels = labels
            hole = ch.plots[0]._element.find(qn('c:holeSize'))
            if hole is not None:
                hole.set('val', '68')
        return ch

    # -- finishing
    def save(self, path):
        buf = io.BytesIO()
        self.p.save(buf)
        data = finish(buf.getvalue(), self.t['head'], self.t['body'], sorted(self.fonts))
        with open(path, 'wb') as f:
            f.write(data)
        return len(data)


def finish(pptx, head, body, families):
    """Embeds the fonts, makes them the deck's theme fonts, and drops python-pptx's defaults"""
    zin = zipfile.ZipFile(io.BytesIO(pptx))
    files = {n: zin.read(n) for n in zin.namelist()}
    rels = files['ppt/_rels/presentation.xml.rels'].decode()
    pres = files['ppt/presentation.xml'].decode()
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    nxt = max(ids) + 1
    entries = []
    k = 1
    for fam in families:
        parts = {}
        for bold in (False, True):
            path = fontkit.face_path(fam, bold)
            if not os.path.exists(path):
                continue
            name = f'fonts/font{k}.fntdata'
            files[f'ppt/{name}'] = fontkit.eot(path)
            rid = f'rId{nxt}'
            nxt += 1
            k += 1
            rels = rels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="{name}"/></Relationships>')
            parts['bold' if bold else 'regular'] = rid
        serif = fam in ('Fraunces', 'DM Serif Display', 'Instrument Serif')
        inner = ''.join(f'<p:{st} r:id="{rid}"/>' for st, rid in parts.items())
        entries.append(f'<p:embeddedFont><p:font typeface="{fam}" pitchFamily="{18 if serif else 34}" charset="0"/>{inner}</p:embeddedFont>')
    lst = f'<p:embeddedFontLst>{"".join(entries)}</p:embeddedFontLst>'
    pres = re.sub(r'(<p:notesSz[^>]*/>)', r'\1' + lst, pres, count=1)
    pres = pres.replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ', 1)
    files['ppt/presentation.xml'] = pres.encode()
    files['ppt/_rels/presentation.xml.rels'] = rels.encode()
    ct = files['[Content_Types].xml'].decode()
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
    files['[Content_Types].xml'] = ct.encode()
    for n in list(files):
        if re.match(r'ppt/theme/theme\d+\.xml$', n):
            th = files[n].decode()
            th = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*"', r'\g<1>' + head + '"', th)
            th = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*"', r'\g<1>' + body + '"', th)
            files[n] = th.encode()
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        order = ['[Content_Types].xml'] + [n for n in files if n != '[Content_Types].xml']
        for n in order:
            z.writestr(n, files[n])
    return out.getvalue()
