"""Builds the Slides templates in src/templates/ (python-pptx). Run: python3 scripts/templates/make-slides.py
Each deck is 16:9, uses Calibri (Carlito where Calibri is missing), and keeps every text in a
text box or table so it can be replaced in place. Previews are taken from the Slides editor
itself (scripts/templates/slide-previews.mjs)."""
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu
import os

OUT = os.path.join(os.path.dirname(__file__), '..', '..', 'src', 'templates')
W, H = Inches(13.333), Inches(7.5)
FONT = 'Calibri'


def rgb(h):
    return RGBColor.from_string(h.lstrip('#'))


class Deck:
    def __init__(self, theme):
        self.t = theme
        self.p = Presentation()
        self.p.slide_width, self.p.slide_height = W, H
        self.blank = self.p.slide_layouts[6]
        self.p.core_properties.author = ''
        self.p.core_properties.title = theme.get('title', '')

    def slide(self, bg=None):
        s = self.p.slides.add_slide(self.blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb(bg or self.t['bg'])
        return s

    def rect(self, s, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, line=None):
        r = s.shapes.add_shape(shape, x, y, w, h)
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(color)
        if line:
            r.line.color.rgb = rgb(line)
            r.line.width = Pt(1)
        else:
            r.line.fill.background()
        r.shadow.inherit = False
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            r.adjustments[0] = 0.08
        return r

    def text(self, s, x, y, w, h, lines, size=18, color=None, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, bullets=False, spacing=1.1, after=6, italic=False):
        tb = s.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        if isinstance(lines, str):
            lines = [lines]
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            para.line_spacing = spacing
            para.space_after = Pt(after)
            run = para.add_run()
            run.text = (u'•  ' if bullets else '') + line
            f = run.font
            f.name = FONT
            f.size = Pt(size)
            f.bold = bold
            f.italic = italic
            f.color.rgb = rgb(color or self.t['ink'])
        return tb

    def title(self, s, text, sub=None, color=None):
        self.rect(s, Inches(0.6), Inches(0.62), Inches(0.12), Inches(0.62), self.t['accent'])
        self.text(s, Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.9), text, 34, color or self.t['ink'], True, anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            self.text(s, Inches(0.92), Inches(1.32), Inches(11.5), Inches(0.5), sub, 16, self.t['muted'])

    def footer(self, s, n, label):
        self.text(s, Inches(0.6), Inches(6.95), Inches(8), Inches(0.35), label, 11, self.t['muted'])
        self.text(s, Inches(11.7), Inches(6.95), Inches(1.05), Inches(0.35), str(n), 11, self.t['muted'], align=PP_ALIGN.RIGHT)

    def table(self, s, x, y, w, rows, widths=None, size=14, head=None, row_h=0.5):
        shape = s.shapes.add_table(len(rows), len(rows[0]), x, y, w, Inches(row_h * len(rows)))
        tbl = shape.table
        if widths:
            for i, cw in enumerate(widths):
                tbl.columns[i].width = Inches(cw)
        for r, row in enumerate(rows):
            tbl.rows[r].height = Inches(row_h)
            for c, val in enumerate(row):
                cell = tbl.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(head or self.t['accent']) if r == 0 else rgb(self.t['card'] if r % 2 else self.t['bg'])
                cell.margin_left = cell.margin_right = Inches(0.12)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.paragraphs[0].text = ''
                run = tf.paragraphs[0].add_run()
                run.text = val
                run.font.name = FONT
                run.font.size = Pt(size)
                run.font.bold = r == 0
                run.font.color.rgb = rgb('FFFFFF' if r == 0 else self.t['ink'])
        return tbl

    def chart(self, s, kind, x, y, w, h, cats, series, colors, legend=True, labels=False):
        data = CategoryChartData()
        data.categories = cats
        for name, values in series:
            data.add_series(name, values)
        ch = s.shapes.add_chart(kind, x, y, w, h, data).chart
        ch.has_legend = legend and len(series) > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(12)
            ch.legend.font.color.rgb = rgb(self.t['muted'])
        ch.font.name = FONT
        ch.font.size = Pt(12)
        ch.font.color.rgb = rgb(self.t['muted'])
        for i, ser in enumerate(ch.series):
            fmt = ser.format
            if kind in (XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.LINE):
                fmt.line.color.rgb = rgb(colors[i])
                fmt.line.width = Pt(3)
                ser.smooth = False
            else:
                fmt.fill.solid()
                fmt.fill.fore_color.rgb = rgb(colors[i])
            if labels:
                ser.data_labels.show_value = True
                ser.data_labels.font.size = Pt(12)
                ser.data_labels.font.color.rgb = rgb(self.t['ink'])
        try:
            ch.value_axis.has_major_gridlines = True
            ch.value_axis.major_gridlines.format.line.color.rgb = rgb(self.t['line'])
            ch.value_axis.format.line.fill.background()
            ch.category_axis.format.line.color.rgb = rgb(self.t['line'])
            if kind in (XL_CHART_TYPE.COLUMN_CLUSTERED,):
                ch.plots[0].gap_width = 70
        except Exception:
            pass
        return ch

    def card(self, s, x, y, w, h, head, body, color=None, size=15, head_size=20, fill=None):
        self.rect(s, x, y, w, h, fill or self.t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
        self.rect(s, x + Inches(0.3), y + Inches(0.32), Inches(0.5), Inches(0.08), color or self.t['accent'])
        self.text(s, x + Inches(0.28), y + Inches(0.5), w - Inches(0.56), Inches(0.6), head, head_size, self.t['ink'], True)
        self.text(s, x + Inches(0.28), y + Inches(1.1), w - Inches(0.56), h - Inches(1.25), body, size, self.t['muted'])

    def save(self, name):
        self.p.save(os.path.join(OUT, f'{name}.pptx'))


# ------------------------------------------------------------------ Pitch deck
def pitch():
    t = dict(bg='FFFFFF', ink='0F172A', muted='475569', accent='4F46E5', accent2='06B6D4', card='F1F5F9', line='E2E8F0', dark='1E1B4B', title='Pitch deck')
    d = Deck(t)
    s = d.slide(t['dark'])
    d.rect(s, Inches(8.6), Inches(-1.2), Inches(6.2), Inches(6.2), '312E81', MSO_SHAPE.OVAL)
    d.rect(s, Inches(10.4), Inches(3.9), Inches(4.2), Inches(4.2), '4338CA', MSO_SHAPE.OVAL)
    d.text(s, Inches(0.9), Inches(1.0), Inches(4), Inches(0.5), 'YOUR COMPANY', 14, 'A5B4FC', True)
    d.text(s, Inches(0.9), Inches(2.2), Inches(8.5), Inches(2.0), 'The simplest way for small shops to take orders online', 46, 'FFFFFF', True, spacing=1.0)
    d.text(s, Inches(0.9), Inches(4.5), Inches(7.5), Inches(0.8), 'Seed round, 2027  ·  Presented by Your Name', 18, 'C7D2FE')

    s = d.slide(); d.title(s, 'The problem', 'What our customers struggle with today')
    for i, (head, body) in enumerate([
        ('Orders get lost', 'Shops take orders by phone, chat and walk-in, and keep track of them on paper.'),
        ('Tools are too big', 'Online store builders cost too much and take weeks to set up for a shop of three people.'),
        ('Customers wait', 'Without updates, customers call again and again to ask when their order is ready.')]):
        d.card(s, Inches(0.6 + i * 4.1), Inches(2.3), Inches(3.8), Inches(2.7), head, body, size=17)
    d.footer(s, 2, 'Your Company')

    s = d.slide(); d.title(s, 'Our solution', 'One order page, set up in ten minutes')
    d.rect(s, Inches(7.2), Inches(1.9), Inches(5.5), Inches(4.6), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
    d.text(s, Inches(7.2), Inches(3.8), Inches(5.5), Inches(0.8), 'Product picture', 20, t['muted'], align=PP_ALIGN.CENTER)
    d.text(s, Inches(0.9), Inches(2.2), Inches(5.9), Inches(4), [
        'A shareable order page with the shop\'s menu or catalogue',
        'Every order in one list, from phone, chat and the page',
        'Automatic updates to the customer by message',
        'Payments by card or bank transfer'], 20, t['ink'], bullets=True, after=14)
    d.footer(s, 3, 'Your Company')

    s = d.slide(); d.title(s, 'Market', 'Small shops in our first three markets')
    for i, (num, lab, col) in enumerate([('2.4M', 'Total shops', t['dark']), ('610K', 'Already selling online', t['accent']), ('45K', 'Our 3-year target', t['accent2'])]):
        x = Inches(0.6 + i * 4.1)
        d.rect(s, x, Inches(2.3), Inches(3.8), Inches(3.4), col, MSO_SHAPE.ROUNDED_RECTANGLE)
        d.text(s, x, Inches(2.9), Inches(3.8), Inches(1.2), num, 60, 'FFFFFF', True, align=PP_ALIGN.CENTER)
        d.text(s, x, Inches(4.3), Inches(3.8), Inches(0.6), lab, 18, 'FFFFFF', align=PP_ALIGN.CENTER)
    d.text(s, Inches(0.6), Inches(6.0), Inches(12), Inches(0.5), 'Replace these figures with your own market sizing and its source.', 13, t['muted'], italic=True)
    d.footer(s, 4, 'Your Company')

    s = d.slide(); d.title(s, 'Traction', 'Monthly paying shops since launch')
    d.chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.9), Inches(7.6), Inches(4.8),
            ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], [('Paying shops', (40, 72, 118, 190, 265, 380))], [t['accent']], labels=True)
    for i, (num, lab) in enumerate([('380', 'paying shops'), ('42%', 'monthly growth'), ('3%', 'monthly churn')]):
        y = Inches(2.1 + i * 1.5)
        d.text(s, Inches(8.8), y, Inches(3.9), Inches(0.8), num, 40, t['accent'], True)
        d.text(s, Inches(8.8), y + Inches(0.75), Inches(3.9), Inches(0.5), lab, 16, t['muted'])
    d.footer(s, 5, 'Your Company')

    s = d.slide(); d.title(s, 'Team')
    for i, (name, role) in enumerate([('Name Surname', 'CEO, ten years in retail'), ('Name Surname', 'CTO, built two payment apps'), ('Name Surname', 'Head of sales, 300 shops signed')]):
        x = Inches(0.9 + i * 4.0)
        d.rect(s, x + Inches(0.8), Inches(2.0), Inches(1.8), Inches(1.8), t['card'], MSO_SHAPE.OVAL)
        d.text(s, x, Inches(4.1), Inches(3.4), Inches(0.5), name, 20, t['ink'], True, align=PP_ALIGN.CENTER)
        d.text(s, x, Inches(4.6), Inches(3.4), Inches(0.8), role, 15, t['muted'], align=PP_ALIGN.CENTER)
    d.footer(s, 6, 'Your Company')

    s = d.slide(t['dark'])
    d.text(s, Inches(0.9), Inches(1.3), Inches(6), Inches(0.5), 'THE ASK', 14, 'A5B4FC', True)
    d.text(s, Inches(0.9), Inches(1.9), Inches(11), Inches(1.2), 'Raising 1.5M to reach 5,000 shops', 44, 'FFFFFF', True)
    for i, (pct, use) in enumerate([('50%', 'Product and engineering'), ('35%', 'Sales and marketing'), ('15%', 'Operations')]):
        x = Inches(0.9 + i * 4.0)
        d.text(s, x, Inches(3.8), Inches(3.6), Inches(0.9), pct, 40, 'A5B4FC', True)
        d.text(s, x, Inches(4.7), Inches(3.6), Inches(0.6), use, 18, 'E0E7FF')
    d.text(s, Inches(0.9), Inches(6.2), Inches(11), Inches(0.5), 'you@yourcompany.com  ·  yourcompany.com', 16, 'C7D2FE')
    d.save('pitch-deck')


# ------------------------------------------------------------------ Business review
def review():
    t = dict(bg='FFFFFF', ink='111827', muted='4B5563', accent='0E7490', accent2='F59E0B', card='F0F9FF', line='E5E7EB', title='Quarterly business review')
    d = Deck(t)
    s = d.slide('0E7490')
    d.rect(s, 0, Inches(5.6), W, Inches(1.9), '155E75')
    d.text(s, Inches(0.9), Inches(1.6), Inches(11), Inches(0.6), 'Q3 2026', 22, 'A5F3FC', True)
    d.text(s, Inches(0.9), Inches(2.3), Inches(11), Inches(1.4), 'Quarterly business review', 52, 'FFFFFF', True)
    d.text(s, Inches(0.9), Inches(3.7), Inches(11), Inches(0.6), 'Sales and operations', 22, 'CFFAFE')
    d.text(s, Inches(0.9), Inches(6.2), Inches(11), Inches(0.5), 'Team name  ·  Date', 16, 'CFFAFE')

    s = d.slide(); d.title(s, 'Agenda')
    for i, item in enumerate(['Results at a glance', 'Revenue by month', 'What went well, what did not', 'Priorities for next quarter']):
        y = Inches(1.9 + i * 1.15)
        d.rect(s, Inches(0.9), y, Inches(0.8), Inches(0.8), t['accent'], MSO_SHAPE.OVAL)
        d.text(s, Inches(0.9), y, Inches(0.8), Inches(0.8), str(i + 1), 24, 'FFFFFF', True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        d.text(s, Inches(2.0), y, Inches(9), Inches(0.8), item, 24, t['ink'], anchor=MSO_ANCHOR.MIDDLE)
    d.footer(s, 2, 'Q3 2026 business review')

    s = d.slide(); d.title(s, 'Results at a glance', 'Against the plan for the quarter')
    for i, (num, lab, delta, good) in enumerate([('$4.2M', 'Revenue', '+8% vs plan', True), ('1,280', 'New customers', '+12% vs plan', True), ('38%', 'Gross margin', '-2 pts vs plan', False), ('94%', 'On-time delivery', 'On plan', True)]):
        x = Inches(0.6 + i * 3.1)
        d.rect(s, x, Inches(2.2), Inches(2.85), Inches(3.2), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
        d.text(s, x, Inches(2.6), Inches(2.85), Inches(1.0), num, 44, t['accent'], True, align=PP_ALIGN.CENTER)
        d.text(s, x, Inches(3.7), Inches(2.85), Inches(0.5), lab, 18, t['ink'], align=PP_ALIGN.CENTER)
        d.text(s, x, Inches(4.4), Inches(2.85), Inches(0.5), delta, 15, '15803D' if good else 'B91C1C', True, align=PP_ALIGN.CENTER)
    d.footer(s, 3, 'Q3 2026 business review')

    s = d.slide(); d.title(s, 'Revenue by month', 'In thousands, actual against plan')
    d.chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.9), Inches(8.4), Inches(4.9),
            ['Jul', 'Aug', 'Sep'], [('Plan', (1250, 1300, 1350)), ('Actual', (1310, 1390, 1500))], ['CBD5E1', t['accent']])
    d.rect(s, Inches(9.4), Inches(2.1), Inches(3.3), Inches(4.4), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
    d.text(s, Inches(9.7), Inches(2.4), Inches(2.8), Inches(4), ['September was our best month so far.', 'The new reseller channel added $140K.'], 17, t['ink'], after=16)
    d.footer(s, 4, 'Q3 2026 business review')

    s = d.slide(); d.title(s, 'What went well, what did not')
    for i, (head, col, items) in enumerate([
        ('Went well', '15803D', ['Reseller channel ahead of target', 'Support replies within 2 hours', 'Two large renewals signed early']),
        ('To improve', 'B91C1C', ['Shipping costs rose 9%', 'Hiring for sales is behind', 'Margin below plan in September'])]):
        x = Inches(0.6 + i * 6.2)
        d.rect(s, x, Inches(1.9), Inches(5.9), Inches(0.12), col)
        d.text(s, x, Inches(2.15), Inches(5.9), Inches(0.6), head, 24, col, True)
        d.text(s, x, Inches(2.9), Inches(5.9), Inches(3.5), items, 19, t['ink'], bullets=True, after=14)
    d.footer(s, 5, 'Q3 2026 business review')

    s = d.slide(); d.title(s, 'Priorities for next quarter')
    d.table(s, Inches(0.6), Inches(1.9), Inches(12.1), [
        ['Priority', 'Owner', 'Target', 'Due'],
        ['Bring shipping costs back to plan', 'Operations', '-6% cost per order', '15 Dec'],
        ['Hire three account managers', 'Sales', '3 hires', '30 Nov'],
        ['Launch the annual plan', 'Product', '200 sign-ups', '1 Dec'],
        ['Grow the reseller channel', 'Partnerships', '$500K revenue', '31 Dec']], [5.2, 2.3, 2.6, 2.0], 16, row_h=0.7)
    d.footer(s, 6, 'Q3 2026 business review')
    d.save('business-review')


# ------------------------------------------------------------------ Project status
def status():
    t = dict(bg='FFFFFF', ink='1F2937', muted='6B7280', accent='059669', accent2='2563EB', card='F3F4F6', line='E5E7EB', title='Project status')
    d = Deck(t)
    s = d.slide('F9FAFB')
    d.rect(s, 0, 0, Inches(0.35), H, t['accent'])
    d.text(s, Inches(1.0), Inches(2.0), Inches(10), Inches(0.5), 'PROJECT STATUS', 16, t['accent'], True)
    d.text(s, Inches(1.0), Inches(2.6), Inches(11), Inches(1.4), 'New website launch', 54, t['ink'], True)
    d.text(s, Inches(1.0), Inches(4.0), Inches(11), Inches(0.6), 'Weekly update  ·  Week 12 of 20', 22, t['muted'])

    s = d.slide(); d.title(s, 'Summary')
    for i, (lab, state, col) in enumerate([('Overall', 'On track', '16A34A'), ('Schedule', 'At risk', 'D97706'), ('Budget', 'On track', '16A34A'), ('Scope', 'On track', '16A34A')]):
        x = Inches(0.6 + i * 3.1)
        d.rect(s, x, Inches(1.9), Inches(2.85), Inches(1.5), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
        d.rect(s, x + Inches(0.3), Inches(2.35), Inches(0.36), Inches(0.36), col, MSO_SHAPE.OVAL)
        d.text(s, x + Inches(0.8), Inches(2.1), Inches(1.9), Inches(0.45), lab, 15, t['muted'])
        d.text(s, x + Inches(0.8), Inches(2.5), Inches(1.9), Inches(0.6), state, 21, t['ink'], True)
    d.text(s, Inches(0.6), Inches(3.9), Inches(12), Inches(2.6), [
        'Design is signed off and the first ten pages are built.',
        'Content from the sales team is a week late, which puts testing at risk.',
        'Budget used: 58% at 60% of the schedule.'], 19, t['ink'], bullets=True, after=12)
    d.footer(s, 2, 'New website launch')

    s = d.slide(); d.title(s, 'Timeline', 'Milestones and where we are')
    y = Inches(3.9)
    d.rect(s, Inches(0.9), y, Inches(11.5), Inches(0.08), t['line'])
    d.rect(s, Inches(0.9), y, Inches(6.9), Inches(0.08), t['accent'])
    for i, (lab, when, done) in enumerate([('Kick-off', 'Week 1', True), ('Design', 'Week 5', True), ('Build', 'Week 11', True), ('Testing', 'Week 15', False), ('Launch', 'Week 20', False)]):
        x = Inches(0.9 + i * 2.85)
        d.rect(s, x - Inches(0.22), y - Inches(0.18), Inches(0.44), Inches(0.44), t['accent'] if done else 'FFFFFF', MSO_SHAPE.OVAL, line=None if done else t['muted'])
        d.text(s, x - Inches(1.2), y - Inches(1.2), Inches(2.4), Inches(0.5), lab, 18, t['ink'], True, PP_ALIGN.CENTER)
        d.text(s, x - Inches(1.2), y + Inches(0.5), Inches(2.4), Inches(0.5), when, 15, t['muted'], align=PP_ALIGN.CENTER)
    d.text(s, Inches(7.4), y - Inches(2.0), Inches(1.6), Inches(0.4), 'Today', 14, t['accent'], True, PP_ALIGN.CENTER)
    d.footer(s, 3, 'New website launch')

    s = d.slide(); d.title(s, 'Risks and issues')
    d.table(s, Inches(0.6), Inches(1.8), Inches(12.1), [
        ['Risk', 'Impact', 'What we are doing', 'Owner'],
        ['Late content from sales', 'High', 'Daily check-in; drafts written by marketing', 'Name'],
        ['Payment provider sign-off', 'Medium', 'Contract sent; call booked for Friday', 'Name'],
        ['Two testers on leave in week 16', 'Low', 'Extra tester booked from the agency', 'Name']], [3.6, 1.6, 5.1, 1.8], 16, row_h=0.8)
    d.footer(s, 4, 'New website launch')

    s = d.slide(); d.title(s, 'Next steps')
    for i, (head, body) in enumerate([('This week', 'Finish the product pages and send the content list to sales.'), ('Next week', 'Start testing on phones and tablets; set up the payment test account.'), ('Decisions needed', 'Approve the launch date and the press release by Friday.')]):
        d.card(s, Inches(0.6 + i * 4.1), Inches(2.0), Inches(3.8), Inches(2.7), head, body, color=[t['accent'], t['accent2'], 'D97706'][i], size=17)
    d.footer(s, 5, 'New website launch')
    d.save('project-status')


# ------------------------------------------------------------------ Training
def training():
    t = dict(bg='FFFBF5', ink='292524', muted='57534E', accent='EA580C', accent2='0D9488', card='FFEDD5', line='E7E5E4', title='Training session')
    d = Deck(t)
    s = d.slide('FFF7ED')
    d.rect(s, Inches(8.3), 0, Inches(5.03), H, 'FB923C')
    d.rect(s, Inches(9.3), Inches(1.6), Inches(3.0), Inches(3.0), 'FED7AA', MSO_SHAPE.OVAL)
    d.text(s, Inches(0.9), Inches(1.8), Inches(7), Inches(0.5), 'TRAINING', 16, t['accent'], True)
    d.text(s, Inches(0.9), Inches(2.4), Inches(7.2), Inches(2.2), 'Getting started with customer service', 46, t['ink'], True, spacing=1.0)
    d.text(s, Inches(0.9), Inches(4.7), Inches(7), Inches(0.6), 'Session 1 of 4  ·  Trainer name', 20, t['muted'])

    s = d.slide(); d.title(s, 'By the end of this session you can')
    for i, item in enumerate(['Greet a customer and find out what they need', 'Look up an order and explain where it is', 'Handle a complaint calmly, step by step', 'Know when to pass a case to a team lead']):
        y = Inches(1.9 + i * 1.15)
        d.rect(s, Inches(0.9), y + Inches(0.1), Inches(0.55), Inches(0.55), t['accent2'], MSO_SHAPE.ROUNDED_RECTANGLE)
        d.text(s, Inches(0.9), y + Inches(0.1), Inches(0.55), Inches(0.55), u'✓', 20, 'FFFFFF', True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        d.text(s, Inches(1.8), y, Inches(10), Inches(0.75), item, 22, t['ink'], anchor=MSO_ANCHOR.MIDDLE)
    d.footer(s, 2, 'Customer service training')

    s = d.slide('EA580C')
    d.text(s, Inches(0.9), Inches(2.4), Inches(3), Inches(1.2), '01', 72, 'FED7AA', True)
    d.text(s, Inches(0.9), Inches(3.7), Inches(11), Inches(1.0), 'Handling a complaint', 44, 'FFFFFF', True)

    s = d.slide(); d.title(s, 'Four steps for a complaint')
    for i, (head, body) in enumerate([('Listen', 'Let the customer finish. Take notes.'), ('Agree', 'Say what went wrong in your own words.'), ('Fix', 'Offer what you can do now, and when.'), ('Follow up', 'Check back the next day that it is solved.')]):
        x = Inches(0.6 + i * 3.1)
        d.rect(s, x, Inches(2.0), Inches(2.85), Inches(3.3), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
        d.text(s, x + Inches(0.3), Inches(2.25), Inches(1.0), Inches(0.9), str(i + 1), 40, t['accent'], True)
        d.text(s, x + Inches(0.3), Inches(3.2), Inches(2.3), Inches(0.6), head, 22, t['ink'], True)
        d.text(s, x + Inches(0.3), Inches(3.85), Inches(2.3), Inches(1.9), body, 17, t['muted'])
    d.footer(s, 4, 'Customer service training')

    s = d.slide(); d.title(s, 'Practice', 'Work in pairs: one is the customer, one answers')
    d.rect(s, Inches(0.6), Inches(2.1), Inches(12.1), Inches(3.8), 'FFFFFF', MSO_SHAPE.ROUNDED_RECTANGLE, line=t['line'])
    d.text(s, Inches(1.0), Inches(2.4), Inches(11.3), Inches(3.4), [
        u'“I ordered a lamp two weeks ago and it still has not arrived. Nobody answers my emails.”',
        'Use the four steps. Swap roles after five minutes.'], 22, t['ink'], after=22)
    d.footer(s, 5, 'Customer service training')

    s = d.slide(); d.title(s, 'Recap')
    d.text(s, Inches(0.9), Inches(1.9), Inches(7), Inches(4), ['Listen first, then agree on the problem', 'Offer a fix with a clear time', 'Always follow up', 'Ask a team lead when the fix is above your limit'], 21, t['ink'], bullets=True, after=16)
    d.rect(s, Inches(8.4), Inches(1.9), Inches(4.3), Inches(4.2), t['accent2'], MSO_SHAPE.ROUNDED_RECTANGLE)
    d.text(s, Inches(8.7), Inches(2.2), Inches(3.7), Inches(0.6), 'Next session', 20, 'CCFBF1', True)
    d.text(s, Inches(8.7), Inches(2.9), Inches(3.7), Inches(2.8), 'Orders and returns in the order system', 24, 'FFFFFF', True)
    d.footer(s, 6, 'Customer service training')
    d.save('training')


# ------------------------------------------------------------------ Product launch
def launch():
    t = dict(bg='0B1120', ink='F8FAFC', muted='94A3B8', accent='F43F5E', accent2='A855F7', card='1E293B', line='334155', title='Product launch')
    d = Deck(t)
    s = d.slide()
    d.rect(s, Inches(-1.5), Inches(4.2), Inches(6), Inches(6), '4C0519', MSO_SHAPE.OVAL)
    d.rect(s, Inches(9.5), Inches(-2.0), Inches(6), Inches(6), '3B0764', MSO_SHAPE.OVAL)
    d.text(s, Inches(0), Inches(2.1), W, Inches(0.6), 'INTRODUCING', 18, 'FDA4AF', True, PP_ALIGN.CENTER)
    d.text(s, Inches(0), Inches(2.8), W, Inches(1.4), 'Product Name 2.0', 64, 'FFFFFF', True, PP_ALIGN.CENTER)
    d.text(s, Inches(0), Inches(4.3), W, Inches(0.7), 'Faster, simpler and ready for teams', 24, t['muted'], align=PP_ALIGN.CENTER)

    s = d.slide(); d.title(s, "What's new")
    for i, (head, body, col) in enumerate([('Twice as fast', 'Pages open in under a second, even with large projects.', t['accent']), ('Shared spaces', 'Invite your team and work on the same files together.', t['accent2']), ('Works offline', 'Keep working on a plane; changes sync when you are back.', '22D3EE')]):
        d.card(s, Inches(0.6 + i * 4.1), Inches(2.0), Inches(3.8), Inches(2.9), head, body, col, size=17, head_size=22)
    d.footer(s, 2, 'Product Name 2.0')

    s = d.slide(); d.title(s, 'Faster where it matters', 'Seconds to open a large project, lower is better')
    d.chart(s, XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(1.9), Inches(12.1), Inches(4.8),
            ['Search', 'Export', 'Open project'], [('Version 1', (2.1, 6.5, 4.8)), ('Version 2', (0.6, 2.4, 0.9))], ['475569', t['accent']])
    d.footer(s, 3, 'Product Name 2.0')

    s = d.slide(); d.title(s, 'Plans')
    for i, (name, price, items, hot) in enumerate([('Free', '$0', ['1 user', '3 projects', 'Community help'], False), ('Team', '$12', ['Up to 20 users', 'Unlimited projects', 'Shared spaces'], True), ('Business', '$24', ['Unlimited users', 'Single sign-on', 'Priority help'], False)]):
        x = Inches(0.9 + i * 4.0)
        d.rect(s, x, Inches(1.9), Inches(3.6), Inches(4.6), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE, line=t['accent'] if hot else None)
        d.text(s, x + Inches(0.35), Inches(2.2), Inches(2.9), Inches(0.5), name, 20, t['accent'] if hot else t['muted'], True)
        d.text(s, x + Inches(0.35), Inches(2.75), Inches(2.9), Inches(0.9), price, 44, t['ink'], True)
        d.text(s, x + Inches(0.35), Inches(3.6), Inches(2.9), Inches(0.4), 'per user, per month', 13, t['muted'])
        d.text(s, x + Inches(0.35), Inches(4.25), Inches(2.9), Inches(2.0), items, 17, t['ink'], bullets=True, after=8)
    d.footer(s, 4, 'Product Name 2.0')

    s = d.slide(); d.title(s, 'Launch plan')
    for i, (when, what) in enumerate([('1 Dec', 'Early access for current customers'), ('8 Dec', 'Public launch and press release'), ('15 Dec', 'Webinar and partner offers'), ('January', 'Team plan in all regions')]):
        y = Inches(1.9 + i * 1.2)
        d.text(s, Inches(0.9), y, Inches(2.4), Inches(0.8), when, 24, t['accent'], True, anchor=MSO_ANCHOR.MIDDLE)
        d.rect(s, Inches(3.3), y + Inches(0.25), Inches(0.3), Inches(0.3), t['accent2'], MSO_SHAPE.OVAL)
        d.text(s, Inches(3.9), y, Inches(8.5), Inches(0.8), what, 22, t['ink'], anchor=MSO_ANCHOR.MIDDLE)
    d.footer(s, 5, 'Product Name 2.0')

    s = d.slide()
    d.text(s, Inches(0), Inches(2.4), W, Inches(1.2), 'Try it today', 60, 'FFFFFF', True, PP_ALIGN.CENTER)
    d.text(s, Inches(0), Inches(3.8), W, Inches(0.7), 'yourproduct.com/new', 26, 'FDA4AF', align=PP_ALIGN.CENTER)
    d.save('product-launch')


# ------------------------------------------------------------------ Simple
def simple():
    t = dict(bg='FFFFFF', ink='18181B', muted='71717A', accent='2563EB', accent2='18181B', card='F4F4F5', line='E4E4E7', title='Simple presentation')
    d = Deck(t)
    s = d.slide()
    d.rect(s, Inches(0.9), Inches(2.3), Inches(1.2), Inches(0.1), t['accent'])
    d.text(s, Inches(0.9), Inches(2.6), Inches(11.5), Inches(1.3), 'Presentation title', 54, t['ink'], True)
    d.text(s, Inches(0.9), Inches(3.9), Inches(11.5), Inches(0.7), 'Subtitle, your name and the date', 22, t['muted'])

    s = d.slide(); d.title(s, 'Slide with points')
    d.text(s, Inches(0.9), Inches(1.9), Inches(11.5), Inches(4.5), ['First point, in one short sentence', 'Second point', 'Third point, with a detail people should remember', 'Keep it to five points or fewer'], 24, t['ink'], bullets=True, after=18)
    d.footer(s, 2, 'Presentation title')

    s = d.slide(); d.title(s, 'Two columns')
    for i, head in enumerate(['Before', 'After']):
        x = Inches(0.9 + i * 6.0)
        d.text(s, x, Inches(1.9), Inches(5.4), Inches(0.6), head, 24, t['accent'], True)
        d.text(s, x, Inches(2.6), Inches(5.4), Inches(3.5), ['A point about this side', 'Another point', 'A third point'], 20, t['ink'], bullets=True, after=12)
    d.footer(s, 3, 'Presentation title')

    s = d.slide(); d.title(s, 'Picture and text')
    d.rect(s, Inches(0.9), Inches(1.9), Inches(6.2), Inches(4.6), t['card'], MSO_SHAPE.ROUNDED_RECTANGLE)
    d.text(s, Inches(0.9), Inches(3.9), Inches(6.2), Inches(0.6), 'Your picture here', 18, t['muted'], align=PP_ALIGN.CENTER)
    d.text(s, Inches(7.6), Inches(1.9), Inches(4.9), Inches(4.6), 'Say what the picture shows and why it matters. Two or three sentences are enough.', 20, t['ink'])
    d.footer(s, 4, 'Presentation title')

    s = d.slide(t['card'])
    d.text(s, Inches(1.4), Inches(1.9), Inches(10.5), Inches(2.6), u'“A short quote that sums up the idea of this talk.”', 40, t['ink'], True, spacing=1.05)
    d.text(s, Inches(1.4), Inches(4.7), Inches(10.5), Inches(0.6), u'Name, role', 20, t['muted'])

    s = d.slide()
    d.text(s, Inches(0), Inches(2.6), W, Inches(1.2), 'Thank you', 60, t['ink'], True, PP_ALIGN.CENTER)
    d.text(s, Inches(0), Inches(3.9), W, Inches(0.7), 'Questions?  ·  you@example.com', 22, t['muted'], align=PP_ALIGN.CENTER)
    d.save('simple')


if __name__ == '__main__':
    for make in (simple, pitch, review, status, training, launch):
        make()
    print('done')
