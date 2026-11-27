# Sample files for the website screenshots: an annual report (Word) and a budget (Excel, with a chart).
# Run: python3 scripts/site-shots/make-samples.py <out-dir>
import sys, os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

out = sys.argv[1] if len(sys.argv) > 1 else '.'
os.makedirs(out, exist_ok=True)

d = Document()
st = d.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(11)
for s, size in (('Heading 1', 18), ('Heading 2', 13), ('Title', 26)):
    d.styles[s].font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A); d.styles[s].font.size = Pt(size)
d.add_paragraph('Riverside Libraries', style='Title')
p = d.add_paragraph('Annual report 2025'); p.runs[0].font.size = Pt(14); p.runs[0].font.color.rgb = RGBColor(0x47, 0x55, 0x69)
d.add_heading('1  A year of growth', level=1)
d.add_paragraph('Visits rose for the third year in a row, and the two branches that opened in spring are already among the busiest. '
                'Members borrowed more e-books than ever, while weekend opening drew families who could not visit during the week.')
q = d.add_paragraph(); r = q.add_run('KEY RECOMMENDATION  '); r.bold = True; r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
q.add_run('Extend weekend hours to the two newest branches in 2026, funded by the reserve set aside for opening costs.')
q.paragraph_format.left_indent = Cm(0.6)
d.add_heading('2  The year in numbers', level=1)
d.add_paragraph('All figures are for the calendar year and include the two new branches from the month they opened.')
rows = [('Measure', '2024', '2025', 'Change'), ('Visits', '361,500', '412,300', '+14%'), ('Items lent', '978,000', '1,124,000', '+15%'),
        ('E-book loans', '121,000', '228,000', '+88%'), ('Active members', '48,200', '54,900', '+14%'), ('Events held', '1,240', '1,610', '+30%')]
t = d.add_table(rows=len(rows), cols=4); t.style = 'Light List Accent 1'
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        c = t.cell(i, j); c.text = v
        if j: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
cap = d.add_paragraph('Table 1. Main measures, 2024 and 2025'); cap.runs[0].italic = True; cap.runs[0].font.size = Pt(9)
d.add_heading('3  Programmes and outreach', level=1)
d.add_paragraph('Our programmes reached more people than ever, helped by partnerships with schools, community centres and two local councils.')
d.add_heading('3.1  Children and families', level=2)
d.add_paragraph('Story time now runs four times a week at every branch. The summer reading challenge had 5,800 children taking part, up from 4,100, and 71% finished all six books.')
d.add_heading('3.2  Digital skills', level=2)
d.add_paragraph('Volunteers ran 420 classes on using smartphones, online banking and staying safe online. Nine in ten people who came said they felt more confident afterwards.')
d.add_heading('4  Looking ahead', level=1)
for line in ('Weekend opening at every branch', 'A mobile library for the northern villages', 'Free printing for job seekers'):
    d.add_paragraph(line, style='List Bullet')
d.save(os.path.join(out, 'Annual report 2025.docx'))

wb = Workbook(); ws = wb.active; ws.title = 'Budget'
head = Font(bold=True, color='FFFFFF'); fill = PatternFill('solid', fgColor='059669'); thin = Side(style='thin', color='D1D5DB')
ws.append(['Household budget 2026']); ws['A1'].font = Font(bold=True, size=16, color='065F46')
ws.append([])
ws.append(['Category', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total'])
for c in ws[3]: c.font = head; c.fill = fill; c.alignment = Alignment(horizontal='center')
data = [('Rent', 1800), ('Groceries', 950), ('Transport', 420), ('Utilities', 310), ('School', 600), ('Savings', 1000), ('Leisure', 380)]
import random; random.seed(4)
for i, (name, base) in enumerate(data, start=4):
    vals = [round(base * random.uniform(0.9, 1.12)) for _ in range(6)] if name not in ('Rent', 'Savings') else [base] * 6
    ws.append([name, *vals, f'=SUM(B{i}:G{i})'])
n = 3 + len(data)
ws.append(['Total', *[f'=SUM({c}4:{c}{n})' for c in 'BCDEFG'], f'=SUM(H4:H{n})'])
for c in ws[n + 1]: c.font = Font(bold=True)
for row in ws.iter_rows(min_row=3, max_row=n + 1):
    for c in row:
        c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        if c.column > 1: c.number_format = '#,##0'
ws.column_dimensions['A'].width = 16
for col in 'BCDEFGH': ws.column_dimensions[col].width = 11
ch = BarChart(); ch.type = 'col'; ch.title = 'Spending by month'; ch.grouping = 'stacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=7, min_row=4, max_row=n), from_rows=True, titles_from_data=False)
ch.set_categories(Reference(ws, min_col=2, max_col=7, min_row=3))
ch.height = 7; ch.width = 17; ch.legend = None
ws.add_chart(ch, 'A13')
wb.save(os.path.join(out, 'Household budget 2026.xlsx'))
print('ok')
