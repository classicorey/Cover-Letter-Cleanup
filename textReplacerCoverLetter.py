import re
from docx import Document
from docx2pdf import convert
from datetime import datetime

filename = r"C:\Users\Corey Crooks\OneDrive\zJobSearch\2026CoverLetter.docx"
document = Document(filename)

def myreplace(file, regex, replace):
    for p in file.paragraphs:
        if regex.search(p.text):
            inline=p.runs

            for i in range (len(inline)):
                if regex.search(inline[i].text):
                    text=regex.sub(replace, inline[i].text)
                    inline[i].text=text

    for table in file.tables:
        for row in table.rows:
            for cell in row.cells:
                myreplace(cell,regex,replace)

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