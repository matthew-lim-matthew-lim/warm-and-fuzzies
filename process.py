from pylatex import Document, Command, LineBreak
from pylatex.utils import italic, NoEscape
from openpyxl import load_workbook
import json

WARM_AND_FUZZIES = {}

# INDEXES
TIMESTAMP = 0
FIRST_NAME = 1
LAST_NAME = 2
MESSAGE = 3

workbook = load_workbook(filename="Warm and Fuzzies (Responses).xlsx")
sheet = workbook.active

# Go through each row until hit an empty row
for row in sheet.iter_rows(min_row=2, max_col=4, max_row=sheet.max_row):
    name = f"{row[FIRST_NAME].value} {row[LAST_NAME].value}".upper()

    # MAKE MANUAL FIXES HERE (for typos)
    # e.g. if HAYES CHOI set name to HAYES CHOY

    if name not in WARM_AND_FUZZIES:
        WARM_AND_FUZZIES[name] = []

    WARM_AND_FUZZIES[name].append(row[MESSAGE].value)

print(json.dumps(WARM_AND_FUZZIES, indent=2))

# Dump each person into a LaTeX file
for name in WARM_AND_FUZZIES:
    doc = Document(indent=False)

    doc.preamble.append(Command('title', name))
    doc.preamble.append(Command('author', 'Co-op Soc 2020'))
    doc.append(NoEscape(r'\maketitle'))

    for message in WARM_AND_FUZZIES[name]:
        doc.append(message + "\n\n")

    doc.generate_pdf(name)