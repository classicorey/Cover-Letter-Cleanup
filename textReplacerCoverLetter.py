import re
from docx import Document
from docx2pdf import convert
from datetime import datetime

filename = r"C:\Users\Corey Crooks\OneDrive\zJobSearch\2026CoverLetter.docx"
document = Document(filename)

x ="{Company}"
t = re.compile(x)
r = input('Enter the name of the Company:   ')

newSaveFilename = r"C:\Users\Corey Crooks\OneDrive\zJobSearch\2026CoverLetter" + r