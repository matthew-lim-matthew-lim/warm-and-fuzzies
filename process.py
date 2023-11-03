from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, RGBColor, Inches
import random
import json

from emojis import EMOJIS

WARM_AND_FUZZIES = {}

# INDEXES
TIMESTAMP = 0
MESSAGE = 17

workbook = load_workbook(filename="responses.xlsx")
sheet = workbook.active

# Go through each row until hit an empty row
for row in sheet.iter_rows(min_row=2, max_col=18, max_row=sheet.max_row):
    for col in range(2, 17):
        if row[col].value != None:
            name = row[col].value.strip().upper().replace(" (CAMP LEADER)", "")
            
    print(name)
    if name not in WARM_AND_FUZZIES:
        WARM_AND_FUZZIES[name] = []

    WARM_AND_FUZZIES[name].append(row[MESSAGE].value)

# print(json.dumps(WARM_AND_FUZZIES, indent=2))

# print(len(WARM_AND_FUZZIES), "recipients...")

# Dump each person into a LaTeX file
for name in WARM_AND_FUZZIES:
    doc = Document()

    doc.styles["Heading 1"].font.name = "CMU Serif"
    doc.styles["Heading 1"].font.size = Pt(24)
    doc.styles["Heading 1"].font.color.rgb = RGBColor(0, 0, 0)

    doc.styles["Normal"].font.name = "CMU Serif"
    doc.styles["Normal"].paragraph_format.space_before = Pt(12)

    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run("")
    run.add_picture("slacklogo.png", width=Inches(1.5))

    doc.add_paragraph()

    t = doc.add_paragraph()
    t.alignment = 1
    run = t.add_run(name)
    run.font.size = Pt(24)
    run.bold = True

    h = doc.add_paragraph()
    h.alignment = 1
    h.add_run("CSESoc First Year Camp 2023").italic = True

    for message in WARM_AND_FUZZIES[name]:
        divider = doc.add_paragraph(random.choice(EMOJIS))
        divider.alignment = 1
        doc.add_paragraph(message.strip().strip("\n"))

    file_name = name + ".docx"

    doc.save(file_name)
