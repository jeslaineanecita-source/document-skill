from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


class Fonts:
    BODY = "Calibri"
    HEADING = "Calibri"
    MONO = "Consolas"


class Sizes:
    TITLE = Pt(28)
    HEADING_1 = Pt(20)
    HEADING_2 = Pt(16)
    HEADING_3 = Pt(14)
    BODY = Pt(11)
    SMALL = Pt(9)
    TABLE_HEADER = Pt(10)
    TABLE_CELL = Pt(10)


class Colors:
    PRIMARY = RGBColor(0x1F, 0x4E, 0x79)
    SECONDARY = RGBColor(0x2E, 0x75, 0xB6)
    ACCENT = RGBColor(0x4C, 0xAF, 0x50)
    DARK = RGBColor(0x33, 0x33, 0x33)
    LIGHT = RGBColor(0xF5, 0xF5, 0xF5)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    GRAY = RGBColor(0x66, 0x66, 0x66)
    TABLE_BORDER = RGBColor(0xBF, 0xBF, 0xBF)


class Margins:
    NORMAL = (Inches(1), Inches(1), Inches(1), Inches(1))
    NARROW = (Inches(0.5), Inches(0.5), Inches(0.5), Inches(0.5))
    WIDE = (Inches(1.5), Inches(1), Inches(1.5), Inches(1))


def set_cell_shading(cell, color):
    shading = cell._tc.get_or_add_tcPr()
    shading_elm = shading.makeelement(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd",
        {
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill": color,
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val": "clear",
        },
    )
    shading.append(shading_elm)
