from docx import Document
from pathlib import Path
import re

linefile = Path('Characters/Ada — BAN240 Manga Character Sheet.docx')
doc = Document(linefile)
lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

hometown_patterns = [
    (re.compile(r'Enugu', re.I), 'Enugu, Nigeria'),
    (re.compile(r'Lahori|Lahore', re.I), 'Lahore, Pakistan'),
    (re.compile(r'Daegu', re.I), 'Daegu, South Korea'),
    (re.compile(r'Karachi', re.I), 'Karachi, Pakistan'),
    (re.compile(r'Bandra West', re.I), 'Bandra West, Mumbai, India'),
    (re.compile(r'Mumbai', re.I), 'Mumbai, India'),
    (re.compile(r'Manila', re.I), 'Manila, Philippines'),
    (re.compile(r'Scarborough', re.I), 'Scarborough, Toronto, Canada'),
    (re.compile(r'Mississauga', re.I), 'Mississauga, Ontario, Canada'),
    (re.compile(r'Texas|Austin|Dallas|Houston|New York|Los Angeles|Chicago', re.I), None),
]

for line in lines:
    if re.search(r'came to Canada from ([A-Za-z\-\s]+)', line, re.I):
        city = re.search(r'came to Canada from ([A-Za-z\-\s]+)', line, re.I).group(1).strip()
        print('came to Canada from', city)
    if re.search(r'came from ([A-Za-z\-\s]+)', line, re.I):
        city = re.search(r'came from ([A-Za-z\-\s]+)', line, re.I).group(1).strip()
        print('came from', city)
    if re.search(r'from ([A-Za-z\-\s]+)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I):
        city = re.search(r'from ([A-Za-z\-\s]+)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I).group(1).strip()
        if len(city) < 30:
            print('from', city)

print('pattern matches:')
for pattern, location in hometown_patterns:
    if pattern.search('\n'.join(lines)):
        print(pattern.pattern, '=>', location)
