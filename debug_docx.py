from docx import Document
from pathlib import Path
import re
f = Path('Characters/Ada — BAN240 Manga Character Sheet.docx')
doc = Document(f)
lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
print('Enugu lines:')
for line in lines:
    if 'Enugu' in line:
        print(line)
print('Contains Enugu in joined text?', bool(re.search(r'Enugu', '\n'.join(lines), re.I)))
print('Section headings and first 60 lines:')
for idx, line in enumerate(lines[:80], 1):
    print(idx, line)
