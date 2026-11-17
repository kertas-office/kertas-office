"""Writes test/fixtures/equations.docx: equations as Word saves them, one on its own line
and one inside a sentence. Run: python3 test/make-equations.py"""
import docx
from docx.oxml import parse_xml
M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d = docx.Document()
d.add_paragraph('The area of a circle:')
p = d.add_paragraph()
p._p.append(parse_xml(f'<m:oMathPara {M}><m:oMath><m:r><m:t>𝐴</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>𝜋</m:t></m:r>'
                      '<m:sSup><m:e><m:r><m:t>𝑟</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>'))
p = d.add_paragraph('Each side has length ')
p._p.append(parse_xml(f'<m:oMath {M}><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>2</m:t></m:r></m:e></m:rad></m:oMath>'))
p.add_run(' metres.')
d.save('test/fixtures/equations.docx')
