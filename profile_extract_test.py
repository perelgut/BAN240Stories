from docx import Document
from pathlib import Path
import re

section_map = {
    'HOBBIES': 'hobbies',
    'SPORTS': 'sports',
    'FANDOMS': 'fandoms',
    'DEMOGRAPHICS': 'demographics',
    'BASIC INFORMATION': 'demographics',
    'PERSONALITY': 'personality',
}

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

def text_lines(doc):
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

def extract_birthday(lines):
    for line in lines:
        if re.search(r'\bBirthday\b|\bborn\b|\bDOB\b|\bBirth\b', line, re.I):
            return line
    for line in lines:
        m = re.search(r'(\d{1,2})[- ]year[- ]old', line, re.I)
        if m:
            return f'{m.group(1)} years old'
    return 'Not specified'

def extract_hometown(lines):
    for line in lines:
        if re.search(r'came to Canada from ([A-Za-z\-\s]+)', line, re.I):
            city = re.search(r'came to Canada from ([A-Za-z\-\s]+)', line, re.I).group(1).strip()
            return f'{city}, Canada'
        if re.search(r'came from ([A-Za-z\-\s]+)', line, re.I):
            city = re.search(r'came from ([A-Za-z\-\s]+)', line, re.I).group(1).strip()
            return city
        if re.search(r'from ([A-Za-z\-\s]+)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I):
            city = re.search(r'from ([A-Za-z\-\s]+)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I).group(1).strip()
            if len(city) < 30:
                return city
    for pattern, location in hometown_patterns:
        if pattern.search('\n'.join(lines)):
            return location or pattern.pattern
    return 'Not specified'

p = Path('Characters/Ada — BAN240 Manga Character Sheet.docx')
doc = Document(p)
lines = text_lines(doc)
print('Demographics lines:')
sections = {'intro': []}
current = 'intro'
for line in lines:
    heading = None
    m = re.match(r'^(\d{1,2})\s*[·\-]+\s*(.*)$', line)
    if m:
        title = m.group(2).upper()
        for key, section in section_map.items():
            if key in title:
                heading = section
                break
    if heading is None:
        for key, section in section_map.items():
            if key in line.upper():
                heading = section
                break
    if heading:
        current = heading
        sections.setdefault(current, [])
        continue
    sections.setdefault(current, []).append(line)

print('demographics:', sections.get('demographics'))
print('birthday from demographics:', extract_birthday(sections.get('demographics', [])))
print('hometown from demographics:', extract_hometown(sections.get('demographics', [])))
print('birthday from all lines:', extract_birthday(lines))
print('hometown from all lines:', extract_hometown(lines))
print('lines containing Enugu:', [l for l in lines if 'Enugu' in l])
