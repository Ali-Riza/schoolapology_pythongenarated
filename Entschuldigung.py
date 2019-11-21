from docx import Document
import subprocess
from win32com import client
import time
import os
from datetime import date



document = Document("C:/Users/Ali Riza Bagci/Desktop/AliRiza/AliRiza_Scripts/x1.docx")

today = date.today()
DATER = input("Datum des Termins: ")
PARENT = input("Name des Erziehungsberechtigten ")
font = document.styles['Normal'].font
font.name = "Arial"

#ändere die "fetten" Wörter

for paragraph in document.paragraphs:
    if 'Dietzenbach' in paragraph.text:
        paragraph.runs[1].text = today.strftime("%d.%m.%Y")
    if 'Unterricht' in paragraph.text:
        paragraph.runs[1].text = DATER
    if 'Emine Bagci' in paragraph.text:
        paragraph.text = PARENT


document.save("C:/Users/Ali Riza Bagci/Desktop/AliRiza/AliRiza_Scripts/x2.docx")

################################################ print
  

os.startfile("C:/Users/Ali Riza Bagci/Desktop/AliRiza/AliRiza_Scripts/x2.docx", "print")

