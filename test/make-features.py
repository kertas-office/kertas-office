# Regenerates test/fixtures/features.xlsx the way a third-party writer (openpyxl) makes it.
# Run: python3 test/make-features.py
import os
HERE = os.path.dirname(os.path.abspath(__file__))
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.worksheet.filters import FilterColumn, Filters, CustomFilters, CustomFilter
wb=Workbook(); ws=wb.active; ws.title='Data'
ws.append(['Item','Qty','Price','Region'])
rows=[('Pen',3,1.5,'North'),('Book',8,12,'South'),('Bag',1,40,'North'),('Pen',5,1.5,'East')]
for r in rows: ws.append(r)
ws.auto_filter.ref='A1:D5'
fc=FilterColumn(colId=3, filters=Filters(filter=['North']))
ws.auto_filter.filterColumn.append(fc)
fc2=FilterColumn(colId=1, customFilters=CustomFilters(customFilter=[CustomFilter(operator='greaterThan', val='2')]))
ws.auto_filter.filterColumn.append(fc2)
ws.row_dimensions[3].hidden=True  # South filtered out
red=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
ws.conditional_formatting.add('B2:B5', CellIsRule(operator='greaterThan', formula=['4'], fill=red, font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add('A2:A5', FormulaRule(formula=['$B2>4'], fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
ws.conditional_formatting.add('C2:C5', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
ws.conditional_formatting.add('B2:B5', DataBarRule(start_type='min', end_type='max', color='638EC6'))
ws.conditional_formatting.add('C2:C5', IconSetRule('3Arrows', 'percent', [0, 33, 67]))
dxf=DifferentialStyle(font=Font(italic=True))
r=Rule(type='containsText', operator='containsText', text='Pe', dxf=dxf); r.formula=['NOT(ISERROR(SEARCH("Pe",A2)))']
ws.conditional_formatting.add('A2:A5', r)
r2=Rule(type='top10', rank=2, dxf=DifferentialStyle(font=Font(underline='single'))); ws.conditional_formatting.add('C2:C5', r2)
dv=DataValidation(type='list', formula1='"North,South,East,West"', allow_blank=True); dv.add('D2:D20'); ws.add_data_validation(dv)
dv2=DataValidation(type='whole', operator='between', formula1='1', formula2='100', showErrorMessage=True, error='1 to 100 only'); dv2.add('B2:B20'); ws.add_data_validation(dv2)
dv3=DataValidation(type='list', formula1='Lists!$A$1:$A$3'); dv3.add('F2:F5'); ws.add_data_validation(dv3)
dv4=DataValidation(type='date', operator='greaterThan', formula1='46023'); dv4.add('G2:G5'); ws.add_data_validation(dv4)
ws['A2'].comment=Comment('Best seller','Kertas')
ws['E1']='Site'; ws['E2']='Kertas'; ws['E2'].hyperlink='https://example.com/kertas'
ws['E3']='Go to lists'; ws['E3'].hyperlink='#Lists!A1'
img=Image(os.path.join(HERE, 'fixtures', 'logo.png')); img.width=64; img.height=64; ws.add_image(img,'H2')
l=wb.create_sheet('Lists'); l['A1']='One'; l['A2']='Two'; l['A3']='Three'
wb.save(os.path.join(HERE, 'fixtures', 'features.xlsx'))
