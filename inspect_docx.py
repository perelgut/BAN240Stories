from docx import Document
from pathlib import Path
files = [Path('Characters/Stephen-sensei — BAN240 Manga Character Sheet.docx'), Path('Characters/Ada — BAN240 Manga Character Sheet.docx')]
for f in files:
    print('FILE:', f)
    doc = Document(f)
    for i, p in enumerate(doc.paragraphs[:80], 1):
        text = p.text.strip()
        if text:
            print(i, repr(text))
    print('---')
