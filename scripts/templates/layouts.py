"""Slide layouts for the templates. Each takes the Deck and the slide's words; colours and fonts
come from the deck's theme, so one layout serves every template."""
from kit import IN, W, H, L, C, R, TOP, MID, BOT, MSO_SHAPE, XL_CHART_TYPE, mesh, photo, qn


def _ellipse(pic):
    pic._element.spPr.find(qn('a:prstGeom')).set('prst', 'ellipse')


# ------------------------------------------------------------------ covers
def cover_mesh(d, kicker, title, sub, meta, light_text=True, size=64, seed=3):
    t = d.t
    s = d.slide(t['dark'], mesh(t['mesh'], seed=seed, base=t['mesh'][0]))
    fg = t['on_dark'] if light_text else t['ink']
    fg2 = t['on_dark_muted'] if light_text else t['muted']
    d.pill(s, IN(0.8), IN(0.8), kicker.upper(), fg, bg='FFFFFF' if light_text else t['ink'], size=11, alpha=0.16 if light_text else 0.08)
    d.head(s, IN(0.8), IN(2.15), IN(10.5), IN(2.9), title, size, fg, anchor=BOT, spacing=0.95)
    d.text(s, IN(0.82), IN(5.2), IN(8.5), IN(0.9), sub, 20, fg2, spacing=1.2)
    d.text(s, IN(0.82), IN(6.75), IN(11.7), IN(0.4), meta, 12, fg2, tracking=40)
    return s


def cover_split(d, kicker, title, sub, meta, seed=4, size=56):
    t = d.t
    s = d.slide()
    d.picture(s, photo(t['photo'], seed=seed), IN(7.0), IN(0.45), IN(5.88), IN(6.6), radius_in=0.32)
    d.kicker(s, IN(0.8), IN(0.9), kicker)
    d.head(s, IN(0.8), IN(1.5), IN(5.9), IN(3.6), title, size, spacing=0.95, anchor=MID)
    d.text(s, IN(0.82), IN(5.1), IN(5.6), IN(1.0), sub, 18, t['muted'], spacing=1.25)
    d.text(s, IN(0.82), IN(6.65), IN(5.6), IN(0.4), meta, 12, t['muted'], tracking=40)
    return s


def cover_type(d, kicker, title, sub, meta, size=88):
    t = d.t
    s = d.slide()
    d.kicker(s, IN(0.8), IN(0.85), kicker)
    d.head(s, IN(0.75), IN(1.4), IN(11.8), IN(3.9), title, size, spacing=0.9, anchor=BOT)
    d.box(s, IN(0.82), IN(5.55), IN(1.1), IN(0.09), t['accent'])
    d.text(s, IN(0.82), IN(5.8), IN(8), IN(0.6), sub, 20, t['muted'])
    d.text(s, IN(8.8), IN(6.75), IN(3.8), IN(0.4), meta, 12, t['muted'], align=R, tracking=40)
    return s


# ------------------------------------------------------------------ structure
def section(d, num, title, sub=None, use_mesh=False, seed=5):
    t = d.t
    s = d.slide(t['dark'], mesh(t['mesh'], seed=seed, base=t['mesh'][0]) if use_mesh else None)
    d.text(s, IN(0.8), IN(1.0), IN(4), IN(1.6), num, 96, t.get('on_dark_accent', t['accent']), True, font=t['head'], tracking=-40)
    d.head(s, IN(0.8), IN(3.6), IN(11), IN(1.6), title, 54, t['on_dark'], anchor=BOT)
    if sub:
        d.text(s, IN(0.82), IN(5.35), IN(9), IN(0.8), sub, 20, t['on_dark_muted'])
    d.footer(s, t['short'], dark=True)
    return s


def agenda(d, title, items, kicker='Agenda'):
    t = d.t
    s = d.slide()
    d.kicker(s, IN(0.8), IN(0.75), kicker)
    d.head(s, IN(0.8), IN(1.2), IN(4.6), IN(3.5), title, 46, spacing=0.95)
    col_x = [IN(6.0), IN(9.45)]
    per = (len(items) + 1) // 2
    for i, (head, body) in enumerate(items):
        x = col_x[i // per]
        y = IN(1.0 + (i % per) * 1.55)
        d.box(s, x, y, IN(3.05), IN(0.02), t['line'])
        d.text(s, x, y + IN(0.18), IN(0.8), IN(0.5), f'{i + 1:02d}', 15, t['accent'], True, tracking=40)
        d.text(s, x, y + IN(0.55), IN(3.05), IN(0.45), head, 20, t['ink'], True, font=t['head'])
        d.text(s, x, y + IN(0.95), IN(3.05), IN(0.5), body, 13, t['muted'])
    d.footer(s, t['short'])
    return s


def statement(d, before, highlight, after='', kicker=None, dark=False):
    t = d.t
    s = d.slide(t['dark'] if dark else None)
    if kicker:
        d.kicker(s, IN(0.8), IN(0.85), kicker, t.get('on_dark_accent', t['accent']) if dark else None)
    fg = t['on_dark'] if dark else t['ink']
    hl = t.get('on_dark_accent', t['accent']) if dark else t['accent']
    bold = not t.get('head_light')
    d.text(s, IN(0.8), IN(1.4), IN(11.6), IN(4.8), [[(before, {}), (highlight, {'color': hl}), (after, {})]], 50, fg, bold,
           anchor=MID, font=t['head'], spacing=1.02, tracking=t.get('head_tracking', -40 if bold else -10))
    d.footer(s, t['short'], dark=dark)
    return s


# ------------------------------------------------------------------ content
def bento(d, kicker, title, big, cells, wide):
    """big: (number, label, body) on the accent colour; cells: two (head, body); wide: (head, body) on dark"""
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    top = IN(2.25)
    d.card(s, IN(0.7), top, IN(4.3), IN(4.45), t['accent'])
    d.text(s, IN(1.05), top + IN(0.4), IN(3.6), IN(1.4), big[0], 72, t.get('on_accent', 'FFFFFF'), True, font=t['head'], tracking=-50)
    d.text(s, IN(1.05), top + IN(1.75), IN(3.6), IN(0.5), big[1], 18, t.get('on_accent', 'FFFFFF'), True)
    d.text(s, IN(1.05), top + IN(2.9), IN(3.6), IN(1.3), big[2], 14, t.get('on_accent_muted', 'F1F1F1'), spacing=1.25, anchor=BOT)
    for i, (h, b) in enumerate(cells):
        x = IN(5.2 + i * 3.82)
        d.card(s, x, top, IN(3.62), IN(2.12))
        d.box(s, x + IN(0.32), top + IN(0.32), IN(0.42), IN(0.42), t['accent2'] if i == 0 else t['accent3'], shape=MSO_SHAPE.OVAL)
        d.text(s, x + IN(0.32), top + IN(0.9), IN(3.0), IN(0.45), h, 18, t['ink'], True, font=t['head'])
        d.text(s, x + IN(0.32), top + IN(1.32), IN(3.0), IN(0.75), b, 13, t['muted'], spacing=1.2)
    d.card(s, IN(5.2), top + IN(2.33), IN(7.44), IN(2.12), t['dark'])
    d.text(s, IN(5.55), top + IN(2.68), IN(6.7), IN(0.5), wide[0], 20, t['on_dark'], True, font=t['head'])
    d.text(s, IN(5.55), top + IN(3.2), IN(6.7), IN(1.0), wide[1], 14, t['on_dark_muted'], spacing=1.25)
    d.footer(s, t['short'])
    return s


def stats(d, kicker, title, items, note=None, dark=False):
    t = d.t
    s = d.slide(t['dark'] if dark else None)
    d.title_block(s, kicker, title, dark=dark, size=36)
    n = len(items)
    colw = 11.9 / n
    for i, (num, label, body) in enumerate(items):
        x = IN(0.7 + i * colw)
        d.box(s, x, IN(2.75), IN(colw - 0.35), IN(0.03), t.get('on_dark_accent', t['accent']) if dark else t['accent'])
        d.text(s, x, IN(2.95), IN(colw - 0.3), IN(1.3), num, 64 if n < 4 else 56, t['on_dark'] if dark else t['ink'], True, font=t['head'], tracking=-50)
        d.text(s, x, IN(4.2), IN(colw - 0.35), IN(0.5), label, 17, t['on_dark'] if dark else t['ink'], True)
        d.text(s, x, IN(4.65), IN(colw - 0.45), IN(1.2), body, 13, t['on_dark_muted'] if dark else t['muted'], spacing=1.25)
    if note:
        d.text(s, IN(0.7), IN(6.3), IN(11.9), IN(0.4), note, 11, t['on_dark_muted'] if dark else t['muted'], italic=True)
    d.footer(s, t['short'], dark=dark)
    return s


def chart_insight(d, kicker, title, kind, cats, series, colors, insight, labels=True, number_format=None, legend=None):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    d.chart(s, kind, IN(0.6), IN(2.15), IN(8.1), IN(4.55), cats, series, colors, labels=labels, number_format=number_format, legend=legend)
    d.card(s, IN(9.05), IN(2.25), IN(3.6), IN(4.4), t['surface'])
    d.text(s, IN(9.4), IN(2.6), IN(2.95), IN(1.2), insight[0], 40, t['accent'], True, font=t['head'], tracking=-40)
    d.text(s, IN(9.4), IN(3.55), IN(2.95), IN(0.6), insight[1], 16, t['ink'], True)
    d.text(s, IN(9.4), IN(4.15), IN(2.95), IN(2.3), insight[2], 13, t['muted'], spacing=1.3)
    d.footer(s, t['short'])
    return s


def donut(d, kicker, title, parts, colors, centre, note):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    d.chart(s, XL_CHART_TYPE.DOUGHNUT, IN(0.7), IN(2.0), IN(5.2), IN(4.7), [p[0] for p in parts], [('Share', [p[1] for p in parts])], colors, legend=False, number_format='0%')
    d.text(s, IN(0.7), IN(3.85), IN(5.2), IN(0.8), centre[0], 40, t['ink'], True, align=C, font=t['head'], tracking=-40)
    d.text(s, IN(0.7), IN(4.55), IN(5.2), IN(0.4), centre[1], 13, t['muted'], align=C)
    for i, (label, share, body) in enumerate(parts):
        y = IN(2.3 + i * 1.08)
        d.box(s, IN(6.6), y + IN(0.08), IN(0.22), IN(0.22), colors[i], shape=MSO_SHAPE.OVAL)
        d.text(s, IN(7.0), y - IN(0.04), IN(3.5), IN(0.45), label, 18, t['ink'], True)
        d.text(s, IN(7.0), y + IN(0.38), IN(4.6), IN(0.5), body, 13, t['muted'])
        d.text(s, IN(10.9), y - IN(0.06), IN(1.7), IN(0.5), f'{round(share * 100)}%', 24, t['ink'], True, align=R, font=t['head'])
    if note:
        d.text(s, IN(6.6), IN(6.3), IN(6), IN(0.4), note, 11, t['muted'], italic=True)
    d.footer(s, t['short'])
    return s


def timeline(d, kicker, title, steps, now=1):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    n = len(steps)
    x0, x1 = 0.9, 12.4
    gap = (x1 - x0) / (n - 1)
    y = IN(3.55)
    d.box(s, IN(x0), y, IN(x1 - x0), IN(0.05), t['line'])
    if now >= 0:
        d.box(s, IN(x0), y, IN(min(x1 - x0, gap * now + gap * 0.45)), IN(0.05), t['accent'])
    for i, (when, what, body) in enumerate(steps):
        x = x0 + i * gap
        done = i <= now
        d.box(s, IN(x - 0.17), y - IN(0.145), IN(0.34), IN(0.34), t['accent'] if done else t['bg'], shape=MSO_SHAPE.OVAL, line=None if done else t['muted'])
        tx = IN(max(0.6, min(x - 1.2, 13.333 - 0.6 - 2.4)))
        al = L if i == 0 else R if i == n - 1 else C
        d.text(s, tx, IN(2.65), IN(2.4), IN(0.45), when, 13, t['accent'] if done else t['muted'], True, align=al, tracking=60)
        d.text(s, tx, IN(4.0), IN(2.4), IN(0.5), what, 18, t['ink'], True, align=al, font=t['head'])
        d.text(s, tx, IN(4.45), IN(2.4), IN(1.4), body, 12.5, t['muted'], align=al, spacing=1.25)
    if 0 <= now < n - 1:
        x = x0 + now * gap + gap * 0.45
        d.pill(s, IN(x - 0.45), IN(3.05), 'TODAY', t.get('on_accent', 'FFFFFF'), t['accent'], size=10)
    d.footer(s, t['short'])
    return s


def compare(d, kicker, title, left, right):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    for i, (head, items) in enumerate((left, right)):
        x = IN(0.7 + i * 6.05)
        dark = i == 1
        d.card(s, x, IN(2.2), IN(5.85), IN(4.45), t['dark'] if dark else t['surface'])
        d.pill(s, x + IN(0.4), IN(2.55), head.upper(), t['on_dark'] if dark else t['ink'], t['accent'] if dark else t['bg'], size=11)
        d.text(s, x + IN(0.4), IN(3.25), IN(5.0), IN(3.2), items, 17, t['on_dark'] if dark else t['ink'], bullets=True, after=12, spacing=1.2)
    d.footer(s, t['short'])
    return s


def quote(d, text, who, role, use_mesh=True, seed=7):
    t = d.t
    s = d.slide(t['dark'], mesh(t['mesh'], seed=seed, base=t['mesh'][0]) if use_mesh else None)
    d.text(s, IN(0.8), IN(0.7), IN(2), IN(1.6), '“', 140, t.get('on_dark_accent', t['accent']), True, font=t['head'])
    d.head(s, IN(0.8), IN(2.0), IN(11.2), IN(3.0), text, 40, t['on_dark'], spacing=1.1)
    d.text(s, IN(0.82), IN(5.5), IN(8), IN(0.45), who, 18, t['on_dark'], True)
    d.text(s, IN(0.82), IN(5.92), IN(8), IN(0.45), role, 14, t['on_dark_muted'])
    return s


def team(d, kicker, title, people, seed=20):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    n = len(people)
    w = 11.9 / n
    for i, (name, role, body) in enumerate(people):
        x = 0.7 + i * w
        pic = d.picture(s, photo(t['photo'], seed=seed + i, size=(700, 700)), IN(x), IN(2.3), IN(min(2.1, w - 0.4)), IN(min(2.1, w - 0.4)), radius_in=0)
        _ellipse(pic)
        d.text(s, IN(x), IN(4.6), IN(w - 0.3), IN(0.45), name, 19, t['ink'], True, font=t['head'])
        d.text(s, IN(x), IN(5.02), IN(w - 0.3), IN(0.4), role, 14, t['accent'], True)
        d.text(s, IN(x), IN(5.45), IN(w - 0.35), IN(1.0), body, 12.5, t['muted'], spacing=1.25)
    d.footer(s, t['short'])
    return s


def image_text(d, kicker, title, body, points=None, seed=9, left=False):
    t = d.t
    s = d.slide()
    px = IN(0.45) if left else IN(7.0)
    tx = 7.0 if left else 0.8
    d.picture(s, photo(t['photo'], seed=seed), px, IN(0.45), IN(5.88), IN(6.6), radius_in=0.32)
    d.kicker(s, IN(tx), IN(0.95), kicker)
    d.head(s, IN(tx), IN(1.4), IN(5.6), IN(1.3), title, 38, spacing=0.98)
    d.text(s, IN(tx + 0.02), IN(2.85), IN(5.4), IN(1.3), body, 16, t['muted'], spacing=1.3)
    if points:
        d.text(s, IN(tx), IN(4.3), IN(5.5), IN(2.2), points, 16, t['ink'], bullets=True, after=8)
    if left:
        d.text(s, IN(11.6), IN(6.9), IN(1.03), IN(0.3), f'{d.n:02d}', 10, t['muted'], align=R, tracking=40)
    else:
        d.footer(s, t['short'])
    return s


def gallery(d, kicker, title, captions, seed=30):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    n = len(captions)
    w = (11.95 - 0.3 * (n - 1)) / n
    for i, cap in enumerate(captions):
        x = 0.7 + i * (w + 0.3)
        h = 3.4
        d.picture(s, photo(t['photo'], seed=seed + i, size=(900, 1100)), IN(x), IN(2.25), IN(w), IN(h), radius_in=0.22)
        d.text(s, IN(x), IN(2.25 + h + 0.15), IN(w), IN(0.4), cap[0], 15, t['ink'], True)
        d.text(s, IN(x), IN(2.25 + h + 0.5), IN(w), IN(0.4), cap[1], 12, t['muted'])
    d.footer(s, t['short'])
    return s


def full_photo(d, title, caption, seed=40):
    t = d.t
    s = d.slide(t['dark'])
    d.picture(s, photo(t['photo'], seed=seed, size=(1600, 900), shade=True), 0, 0, W, H, radius_in=0)
    d.head(s, IN(0.8), IN(4.95), IN(11.5), IN(1.2), title, 46, 'FFFFFF')
    d.text(s, IN(0.82), IN(6.1), IN(10), IN(0.6), caption, 16, 'E5E5E5')
    return s


def plans(d, kicker, title, items, pick=1):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    n = len(items)
    w = (11.95 - 0.3 * (n - 1)) / n
    for i, (name, price, per, feats) in enumerate(items):
        x = IN(0.7 + i * (w + 0.3))
        hot = i == pick
        d.card(s, x, IN(2.2), IN(w), IN(4.5), t['dark'] if hot else t['surface'])
        fg = t['on_dark'] if hot else t['ink']
        fg2 = t['on_dark_muted'] if hot else t['muted']
        d.text(s, x + IN(0.38), IN(2.55), IN(w - 0.76), IN(0.4), name, 16, fg, True)
        if hot:
            d.pill(s, x + IN(w - 1.75), IN(2.52), 'POPULAR', t.get('on_accent', 'FFFFFF'), t['accent'], size=10)
        d.text(s, x + IN(0.38), IN(3.0), IN(w - 0.76), IN(1.0), price, 48, fg, True, font=t['head'], tracking=-40)
        d.text(s, x + IN(0.38), IN(3.95), IN(w - 0.76), IN(0.4), per, 12, fg2)
        d.box(s, x + IN(0.38), IN(4.45), IN(w - 0.76), IN(0.015), t['dark_line'] if hot else t['line'])
        d.text(s, x + IN(0.38), IN(4.65), IN(w - 0.76), IN(1.9), feats, 14, fg, bullets=True, after=6)
    d.footer(s, t['short'])
    return s


def table_slide(d, kicker, title, rows, widths, note=None):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    d.table(s, IN(0.7), IN(2.25), IN(sum(widths)), rows, widths, 14, row_h=0.62)
    if note:
        d.text(s, IN(0.7), IN(6.4), IN(11.9), IN(0.4), note, 11, t['muted'], italic=True)
    d.footer(s, t['short'])
    return s


def process(d, kicker, title, steps):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    n = len(steps)
    w = (11.95 - 0.35 * (n - 1)) / n
    for i, (head, body) in enumerate(steps):
        x = 0.7 + i * (w + 0.35)
        d.card(s, IN(x), IN(2.3), IN(w), IN(3.4), t['surface'])
        d.text(s, IN(x + 0.32), IN(2.55), IN(1.5), IN(1.0), f'{i + 1:02d}', 44, t['accent'], True, font=t['head'], tracking=-40)
        d.text(s, IN(x + 0.32), IN(3.7), IN(w - 0.6), IN(0.5), head, 19, t['ink'], True, font=t['head'])
        d.text(s, IN(x + 0.32), IN(4.2), IN(w - 0.6), IN(1.9), body, 13, t['muted'], spacing=1.3)
        if i < n - 1:
            d.text(s, IN(x + w - 0.06), IN(3.95), IN(0.47), IN(0.5), '→', 22, t['muted'], align=C)
    d.footer(s, t['short'])
    return s


def checklist(d, kicker, title, items, side):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36, w=7.5)
    for i, item in enumerate(items):
        y = IN(2.35 + i * 0.92)
        d.box(s, IN(0.75), y + IN(0.06), IN(0.44), IN(0.44), t['accent'], shape=MSO_SHAPE.OVAL)
        d.text(s, IN(0.75), y + IN(0.06), IN(0.44), IN(0.44), '✓', 15, t.get('on_accent', 'FFFFFF'), True, align=C, anchor=MID)
        d.text(s, IN(1.45), y, IN(6.2), IN(0.6), item, 18, t['ink'], anchor=MID)
    d.card(s, IN(8.35), IN(0.7), IN(4.3), IN(6.0), t['dark'])
    d.kicker(s, IN(8.75), IN(1.15), side[0], t.get('on_dark_accent', t['accent']), size=11)
    d.head(s, IN(8.75), IN(1.6), IN(3.6), IN(2.4), side[1], 30, t['on_dark'], spacing=1.0)
    d.text(s, IN(8.75), IN(4.3), IN(3.6), IN(2.1), side[2], 14, t['on_dark_muted'], spacing=1.3, anchor=BOT)
    d.footer(s, t['short'])
    return s


def three_cards(d, kicker, title, cards, icons=True):
    t = d.t
    s = d.slide()
    d.title_block(s, kicker, title, size=36)
    cols = [t['accent'], t['accent2'], t['accent3']]
    for i, (head, body) in enumerate(cards):
        x = IN(0.7 + i * 4.05)
        d.card(s, x, IN(2.3), IN(3.85), IN(3.3), t['surface'])
        if icons:
            d.card(s, x + IN(0.38), IN(2.68), IN(0.62), IN(0.62), cols[i % 3], radius_in=0.16)
            d.text(s, x + IN(0.38), IN(2.68), IN(0.62), IN(0.62), str(i + 1), 18, t.get('on_accent', 'FFFFFF'), True, align=C, anchor=MID, font=t['head'])
        d.text(s, x + IN(0.38), IN(3.6), IN(3.1), IN(0.55), head, 21, t['ink'], True, font=t['head'])
        d.text(s, x + IN(0.38), IN(4.15), IN(3.1), IN(1.9), body, 14, t['muted'], spacing=1.3)
    d.footer(s, t['short'])
    return s


def closing(d, title, sub, contact, seed=11, use_mesh=True):
    t = d.t
    s = d.slide(t['dark'], mesh(t['mesh'], seed=seed, base=t['mesh'][0]) if use_mesh else None)
    d.head(s, IN(0.8), IN(1.6), IN(11.5), IN(2.6), title, 72, t['on_dark'], anchor=BOT, spacing=0.92)
    d.text(s, IN(0.82), IN(4.45), IN(9), IN(0.8), sub, 20, t['on_dark_muted'], spacing=1.25)
    for i, c in enumerate(contact):
        d.text(s, IN(0.82 + i * 4.0), IN(6.4), IN(3.8), IN(0.4), c, 14, t['on_dark'], True)
    return s
