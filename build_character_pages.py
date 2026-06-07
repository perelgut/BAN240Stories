from docx import Document
from pathlib import Path
import re

root = Path('Characters')
docx_files = sorted(root.glob('*BAN240 Manga Character Sheet.docx'))

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


def extract_full_name(lines, file_base):
    for line in lines[:6]:
        m = re.search(r'Character\s*:\s*(.+?)(?:\s*\|\s*Course|$)', line, re.I)
        if m:
            return m.group(1).strip()
    # fallback from first heading or file name
    if lines:
        first = lines[0]
        if '—' in first:
            return first.split('—', 1)[0].strip()
        if 'Character:' in first:
            return first.split('Character:',1)[1].strip()
    return file_base


def extract_position(lines):
    for line in lines[:6]:
        m = re.search(r'Role\s*:\s*([^|]+)(?:\||$)', line, re.I)
        if m:
            return m.group(1).strip()
        m = re.search(r'Position\s*:\s*([^|]+)(?:\||$)', line, re.I)
        if m:
            return m.group(1).strip()
    for line in lines[:6]:
        m = re.search(r'Year\s*:\s*([^|]+)(?:\||$)', line, re.I)
        if m:
            return m.group(1).strip()
    return 'Manga Character'


def extract_team(lines):
    for line in lines[:6]:
        m = re.search(r'Course\s*:\s*([^|]+)(?:\||$)', line, re.I)
        if m:
            return m.group(1).strip()
        m = re.search(r'University\s*:\s*([^|]+)(?:\||$)', line, re.I)
        if m:
            return m.group(1).strip()
    return 'BAN240 Stories'


def clean_highlight_text(text, max_length=360):
    text = re.sub(r'\s+', ' ', text).strip()
    if len(text) > max_length:
        text = text[:max_length].rsplit(' ', 1)[0] + '...'
    return text


def extract_highlights(sections):
    personality = sections.get('personality', [])
    if personality:
        return clean_highlight_text(' '.join(personality[:2]))
    intro = sections.get('intro', [])
    if intro:
        return clean_highlight_text(' '.join(intro[:2]))
    return 'Highlights unavailable.'


def extract_age(line):
    m = re.search(r'(\d{1,2})[- ]year[- ]old', line, re.I)
    if m:
        return f'{m.group(1)} years old'
    return None


def extract_birthday(lines):
    for line in lines:
        if re.match(r'^(Birthday|DOB|Date of Birth|Born)\b', line, re.I):
            return line
        if re.search(r'\b(born|Birth(?:day)?|DOB|Date of Birth)\b', line, re.I):
            if re.search(r'\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|\b\w+ \d{1,2},? \d{4}\b', line):
                return line
    for line in lines:
        m = re.search(r'\b(\d{1,2})[- ]year[- ]old\b', line, re.I)
        if m:
            return f'Age: {m.group(1)}'
    for line in lines:
        m = re.search(r'\bat\s+(\d{1,2})\b', line, re.I)
        if m and 'present at' in line.lower():
            return f'Age: {m.group(1)}'
    return 'Not specified'


def extract_hometown(lines):
    for line in lines:
        m = re.search(r'came to Canada from ([A-Za-z][A-Za-z\-\s]+?)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I)
        if m:
            return f"{m.group(1).strip()}, Canada"
        m = re.search(r'came from ([A-Za-z][A-Za-z\-\s]+?)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I)
        if m:
            candidate = m.group(1).strip()
            if not re.search(r'\b(father|mother|home|apartment|roommate|family|house)\b', candidate, re.I):
                return candidate
        m = re.search(r'\bfrom ([A-Za-z][A-Za-z\-\s]+?)(?:,|\.|\bCanada\b|\bUniversity\b)', line, re.I)
        if m:
            candidate = m.group(1).strip()
            if not re.search(r'\b(father|mother|home|apartment|roommate|family|house)\b', candidate, re.I):
                if len(candidate) < 40:
                    return candidate
    for pattern, location in hometown_patterns:
        if pattern.search('\n'.join(lines)):
            return location or pattern.pattern
    return 'Not specified'


def collect_section(lines):
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
            # also catch headings without bullets
            for key, section in section_map.items():
                if key in line.upper():
                    heading = section
                    break
        if heading:
            current = heading
            sections.setdefault(current, [])
            continue
        sections.setdefault(current, []).append(line)
    return sections


def extract_titles(lines, max_items=4):
    titles = []
    for line in lines:
        if not line:
            continue
        if '·' in line:
            part = line.split('·', 1)[1].strip()
            if part and len(part) < 80:
                titles.append(part)
                continue
        if len(line) < 80 and len(titles) < max_items:
            # ignore section names
            if not re.match(r'^(\d{1,2})\s*[:\-]', line):
                titles.append(line)
    return titles[:max_items]


def build_profile(file_path):
    doc = Document(file_path)
    lines = text_lines(doc)
    base = re.sub(r'\s+[–—-]\s+BAN240 Manga Character Sheet$', '', file_path.stem)
    full_name = extract_full_name(lines, base)
    position = extract_position(lines)
    team = extract_team(lines)
    sections = collect_section(lines)
    demographics = sections.get('demographics', [])
    hobbies = extract_titles(sections.get('hobbies', []))
    sports = extract_titles(sections.get('sports', []))
    fandoms = extract_titles(sections.get('fandoms', []))
    highlights = extract_highlights(sections)
    birthday = extract_birthday(demographics)
    if birthday == 'Not specified':
        birthday = extract_birthday(lines)
    hometown = extract_hometown(demographics)
    if hometown == 'Not specified':
        hometown = extract_hometown(lines)

    if not hobbies:
        # fallback: take list paragraph titles from personality or intro
        hobbies = extract_titles(sections.get('personality', []))
    if not sports:
        sports = ['No sports listed']
    if not fandoms:
        fandoms = ['No fandoms listed']
    return {
        'file': file_path,
        'base': base,
        'full_name': full_name,
        'position': position,
        'team': team,
        'birthday': birthday,
        'hometown': hometown,
        'highlights': highlights,
        'hobbies': hobbies or ['No hobbies listed'],
        'sports': sports,
        'fandoms': fandoms,
        'docx_name': file_path.name,
        'image_name': f'{base} image.png',
    }

html_template = '''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{full_name} | BAN240 Stories</title>
  <style>
    :root {{
      color-scheme: dark;
      color: #f8fafc;
      background: #081121;
      font-family: Inter, Arial, sans-serif;
    }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; min-height: 100vh; display: grid; place-items: center; padding: 1rem; }}
    .card {{
      width: min(900px, 100%);
      display: grid;
      grid-template-rows: auto 1fr auto auto;
      gap: 0;
      border-radius: 26px;
      overflow: hidden;
      background: #ffffff;
      box-shadow: 0 30px 80px rgba(0, 0, 0, 0.22);
    }}
    .top-banner {{
      background: #0f172a;
      color: #ffffff;
      padding: 1rem 1.4rem;
      font-size: 0.88rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.35em;
      border-bottom: 4px solid #2563eb;
    }}
    .hero {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 1rem;
      padding: 1.4rem;
      background: linear-gradient(180deg, #f8fafc 0%, #e2e8f0 100%);
    }}
    .player-name {{
      margin: 0;
      font-size: clamp(2rem, 4vw, 3.4rem);
      letter-spacing: -0.06em;
      color: #0f172a;
      line-height: 1;
    }}
    .info-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 0.85rem;
      margin-top: 1.25rem;
    }}
    .info-card {{
      background: #0f172a;
      color: #ffffff;
      border-radius: 18px;
      padding: 1rem 1rem;
      display: grid;
      gap: 0.2rem;
    }}
    .label {{
      font-size: 0.72rem;
      letter-spacing: 0.22em;
      text-transform: uppercase;
      opacity: 0.78;
    }}
    .value {{
      font-size: 1.05rem;
      font-weight: 700;
      line-height: 1.2;
    }}
    .photo {{
      background: #dbeafe;
      border-radius: 22px;
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 320px;
    }}
    .photo img {{
      width: 100%;
      height: auto;
      object-fit: cover;
      display: block;
    }}
    .highlights {{
      padding: 1.35rem 1.4rem 1.5rem;
      background: #ffffff;
      border-top: 1px solid rgba(15, 23, 42, 0.08);
      display: grid;
      gap: 0.85rem;
    }}
    .highlights h2 {{
      margin: 0;
      font-size: 0.95rem;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: #2563eb;
    }}
    .highlights p {{
      margin: 0;
      font-size: 1rem;
      line-height: 1.7;
      color: #1f2937;
    }}
    .card-footer {{
      padding: 1rem 1.5rem 1.5rem;
      background: #0f172a;
      display: flex;
      justify-content: space-between;
      gap: 1rem;
      flex-wrap: wrap;
    }}
    .profile-link,
    .return-link {{
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      color: #ffffff;
      text-decoration: none;
      font-weight: 700;
      background: #2563eb;
      padding: 0.85rem 1.1rem;
      border-radius: 999px;
    }}
    .profile-link:hover,
    .return-link:hover {{ opacity: 0.95; }}
    @media (max-width: 720px) {{
      .hero {{ grid-template-columns: 1fr; }}
      .info-grid {{ grid-template-columns: 1fr; }}
      .photo {{ min-height: 260px; }}
    }}
  </style>
</head>
<body>
  <article class="card">
    <div class="top-banner">{position}</div>
    <div class="hero">
      <div>
        <h1 class="player-name">{full_name}</h1>
        <div class="info-grid">
          <div class="info-card">
            <span class="label">Team</span>
            <span class="value">{team}</span>
          </div>
          <div class="info-card">
            <span class="label">Birthday</span>
            <span class="value">{birthday}</span>
          </div>
          <div class="info-card">
            <span class="label">Hometown</span>
            <span class="value">{hometown}</span>
          </div>
        </div>
      </div>
      <div class="photo">
        <img src="{image_name}" alt="{full_name}">
      </div>
    </div>
    <div class="highlights">
      <h2>Highlights</h2>
      <p>{highlights}</p>
    </div>
    <div class="card-footer">
      <a class="return-link" href="Characters.html">Return to Character Index</a>
      <a class="profile-link" href="{docx_name}">See the full profile</a>
    </div>
  </article>
</body>
</html>
'''

if __name__ == '__main__':
    profiles = [build_profile(p) for p in docx_files]
    for profile in profiles:
        out = html_template.format(
            full_name=profile['full_name'],
            position=profile['position'],
            team=profile['team'],
            birthday=profile['birthday'],
            hometown=profile['hometown'],
            highlights=profile['highlights'],
            image_name=profile['image_name'],
            docx_name=profile['docx_name'],
        )
        out_path = root / f"{profile['base']}.html"
        out_path.write_text(out, encoding='utf-8')
        print('Wrote', out_path)
