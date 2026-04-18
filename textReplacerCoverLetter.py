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

myreplace(document, t, r)

x ="{Position}"
t = re.compile(x)

r = input('Enter the position you are applying for:   ')

myreplace(document, t, r)


document.save(newSaveFilename + ".docx")
convert(newSaveFilename+ ".docx", newSaveFilename + ".pdf")
print('Text is replaced, and document is saved.')