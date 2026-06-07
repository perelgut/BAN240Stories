from docx import Document
from pathlib import Path
import re
root = Path('Characters')
docs = sorted(root.glob('*— BAN240 Manga Character Sheet.docx'))
for f in docs:
    name = f.stem.replace(' — BAN240 Manga Character Sheet', '')
    print('\n===', name, '===')
    doc = Document(f)
    text = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for i, line in enumerate(text[:120], 1):
        if re.search(r'^(02|05|06|07)\b|^Character:|Birthday|born|Birth|Home|Hometown|city|from|\bDOB\b', line, re.I):
            print(i, line)
    print('---')
