from docx import Document
import os
import subprocess
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

def page_formatting(document):
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(.5)
        section.bottom_margin = Inches(.5)
        section.left_margin = Inches(.4)
        section.right_margin = Inches(.4)
        section.footer_distance = Inches(0.2)
def add_horizontal_line(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    hr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    hr.append(bottom)
    p._element.get_or_add_pPr().append(hr)
def add_content(document, books,chapters):
    for book in books:
        add_horizontal_line(document)
        document.add_paragraph("")
        for chapter in range(1, chapters[book] + 1):
            path = f'bom-english/{book}/{chapter}.txt'
            with open(path, 'r', encoding='utf-8') as file:
                verses = [line.strip() for line in file.readlines() if line.strip()]

            for verse in verses:
                p = document.add_paragraph()
                run = p.add_run(verse)
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
                p.paragraph_format.first_line_indent = Pt(24)


            add_horizontal_line(document)
def add_page_numbers(document):
    sections = document.sections
    for section in sections:
        footer = section.footer
        p = footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        field = OxmlElement('w:fldSimple')
        field.set(qn('w:instr'), 'PAGE')
        run._element.append(field)


def main():
    books = ["1-nephi"]
    chapters = {
        "1-nephi": 22
    }

    document = Document()
    page_formatting(document)
    add_content(document,books,chapters)
    add_page_numbers(document)

    document.save("readersEdition.docx")
    subprocess.call(['powershell.exe', 'Start-Process', 'readersEdition.docx'])

main()
