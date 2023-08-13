from docx import Document as WordDocument
from docx.document import Document
from docx.section import Section
from docx.shared import Cm

MARGIN = Cm(1.27)


def document_setup():
    doc: Document = WordDocument()
    section: Section = doc.sections[0]
    # set doc size to A4
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    # set margins
    section.left_margin = Cm(2.01)
    section.right_margin = Cm(1.50)
    section.top_margin = Cm(1.40)
    section.bottom_margin = Cm(1.50)

    return doc
