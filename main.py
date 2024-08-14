from docx import Document
import os
import subprocess  # For opening document on WSL Ubuntu
from docx.shared import Inches, Pt, RGBColor  # Styling headings
from docx.oxml.ns import qn  # Page numbers
from docx.oxml import OxmlElement  # Horizontal line, borders
from docx.enum.text import WD_ALIGN_PARAGRAPH  # For justification

def main():
    books = ["title","introduction", "three","eight","js", "1-nephi"]
    chapters = {
        "title":1,"introduction":1,"three":1,"eight":1,"js":1,
        "1-nephi": 1
    }

    document = Document()  # Create a new Word document

    # Set smaller margins
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(.5)  # Smaller margin for a more book-like appearance
        section.bottom_margin = Inches(.5)  # Smaller margin
        section.left_margin = Inches(.4)  # Smaller margin
        section.right_margin = Inches(.4)  # Smaller margin
        section.footer_distance = Inches(0.2)  # Make footer with page number smaller

    def style_cell_text(cell, text, font_name='Times New Roman', font_size=12, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False):
        # Clear existing text
        cell.text = ''
        # Create a new run for the cell
        run = cell.paragraphs[0].add_run(text)
        # Apply the styles
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.bold = bold
        run.italic = italic
        # Set the alignment to justify
        cell.paragraphs[0].alignment = alignment
        # Adjust paragraph spacing
        paragraph_format = cell.paragraphs[0].paragraph_format
        paragraph_format.space_before = Pt(0)  # No space before paragraph
        paragraph_format.space_after = Pt(4)  # Small space after the paragraph
        paragraph_format.line_spacing = Pt(12)  # Adjusted line spacing

        # Apply cell borders
        tc_pr = cell._element.get_or_add_tcPr()
        borders = tc_pr.find(qn('w:tcBorders'))
        if borders is None:
            borders = OxmlElement('w:tcBorders')
            tc_pr.append(borders)
        
        for border in ['top', 'left', 'bottom', 'right']:
            b = OxmlElement(f'w:{border}')
            b.set(qn('w:val'), 'nil')
            b.set(qn('w:space'), '0')
            borders.append(b)

        # Set specific borders for the columns
        if cell._element.getparent().index(cell._element) % 2 == 0:  # First column
            right_border = OxmlElement('w:right')
            right_border.set(qn('w:val'), 'single')
            right_border.set(qn('w:sz'), '4')
            right_border.set(qn('w:space'), '0')
            borders.append(right_border)
        else:  # Second column
            left_border = OxmlElement('w:left')
            left_border.set(qn('w:val'), 'single')
            left_border.set(qn('w:sz'), '4')
            left_border.set(qn('w:space'), '0')
            borders.append(left_border)

    def add_horizontal_line(doc):
        p = doc.add_paragraph()
        run = p.add_run()
        hr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '12')  # Border size
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'auto')
        hr.append(bottom)
        p._element.get_or_add_pPr().append(hr)

    def add_title_page(doc):
        # Title page
        doc.add_paragraph("\n\n\n")  # Add spacing before the title
        # Add the main title in large, bold font
        main_title = doc.add_paragraph()
        main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Center alignment
        run = main_title.add_run("The Book of Mormon") #could prolly take this from first line of title/1.txt
        run.font.name = 'Times New Roman'
        run.font.size = Pt(36)
        run.bold = True
        # Add subtitle in a slightly smaller font and italics
        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Center alignment
        run = subtitle.add_run("Another testament of Jesus Christ")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(24)
        run.italic = True    # Page break
        doc.add_page_break()

    add_title_page(document) # Add title page


    for book in books:
        add_horizontal_line(document)  # Line after book title
        document.add_paragraph("")  # Space after book title
        for chapter in range(1, chapters[book] + 1):
            path = f'bom-english/{book}/{chapter}.txt'
            with open(path, 'r', encoding='utf-8') as file:
                verses = [line.strip() for line in file.readlines() if line.strip()]  # Removes new line characters
            # Create a table with two columns
            table = document.add_table(rows=0, cols=2)
            # Add verses to the table with verse numbers
            for i in range(1, len(verses)): #NOTE that english verses should be the same length as spanish verses
                row_cells = table.add_row().cells
                style_cell_text(row_cells[0], f"{i} {verses[i].strip()}")

            add_horizontal_line(document)  # Line after chapter
            

    # Page numbers
    def add_page_numbers(document):
        sections = document.sections
        for section in sections:
            footer = section.footer
            p = footer.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Center-align the page number
            run = p.add_run()
            field = OxmlElement('w:fldSimple')
            field.set(qn('w:instr'), 'PAGE')  # PAGE is the instruction for page number
            run._element.append(field)

    # Add page numbers to the footer
    add_page_numbers(document)

    # Save the document
    document.save("bom.docx")
    subprocess.call(['powershell.exe', 'Start-Process', 'bom.docx'])  # Opens document


main()