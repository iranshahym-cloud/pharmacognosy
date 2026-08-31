from __future__ import annotations
import glob
import os
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from docx import Document
from PIL import Image

PROJECT_ROOT = Path('.')
DOCS_DIR = PROJECT_ROOT / 'docs'
ASSETS_DIR = DOCS_DIR / 'assets'
STYLES_DIR = DOCS_DIR / 'stylesheets'

SITE_NAME = 'راهنمای آزمایشگاه فارماکوگنوزی ۱'
SITE_DESCRIPTION = 'Pharmacognosy Laboratory Manual'
SITE_AUTHOR = 'Pharmacognosy Department'

DOCS_DIR.mkdir(exist_ok=True)
ASSETS_DIR.mkdir(exist_ok=True)
STYLES_DIR.mkdir(exist_ok=True)

CSS_CONTENT = """@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;700&display=swap');

:root {
  --md-text-font: 'Vazirmatn', sans-serif;
  --md-code-font: 'Courier New', monospace;
}

html, body, .md-container, .md-content, .md-typeset {
  direction: rtl;
}

body, .md-typeset {
  font-family: var(--md-text-font);
  text-align: justify;
}

.num-ltr, .num {
  direction: ltr !important;
  unicode-bidi: isolate !important;
  display: inline-block !important;
  font-family: var(--md-text-font), Tahoma, sans-serif !important;
  text-align: left;
}

.md-typeset table:not([class]) {
  direction: rtl;
  margin-left: auto;
  margin-right: auto;
}

.md-typeset table:not([class]) th, .md-typeset table:not([class]) td {
  text-align: right;
}

.md-typeset img {
  display: block;
  margin: 1.5em auto;
  max-width: 85%;
  height: auto;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
}

.md-typeset figure {
  direction: rtl;
}

.MathJax, mjx-container, .arithmatex {
  direction: ltr !important;
  unicode-bidi: isolate !important;
}
"""

with open(STYLES_DIR / 'extra.css', 'w', encoding='utf-8') as css_file:
    css_file.write(CSS_CONTENT)

DIGIT = r'0-9\u06f0-\u06f9\u0660-\u0669'
NUMBER_PATTERN = re.compile(
    rf'(?<![\w/\.#%])[{DIGIT}]+(?::[/,\.\u066b\u066c-][{DIGIT}]+)+%?(?![\w/\.#%])',
    flags=re.UNICODE,
)

def natural_sort_key(value: str):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r'(\d+)', value)]

def run_command(command: list[str]) -> None:
    try:
        subprocess.run(command, check=True)
    except FileNotFoundError:
        raise RuntimeError(f"Could not find '{command[0]}'. Make sure it is installed and on PATH.")
    except subprocess.CalledProcessError as err:
        raise RuntimeError(f"Command failed with code {err.returncode}: {' '.join(command)}")

def convert_emf_or_wmf_to_png(source_path: Path) -> Path | None:
    png_path = source_path.with_suffix('.png')
    try:
        with Image.open(source_path) as img:
            try:
                img.load()
            except Exception:
                pass
            if img.mode in ('RGBA', 'LA', 'P'):
                rgba = img.convert('RGBA')
                bg = Image.new('RGB', rgba.size, 'white')
                bg.paste(rgba, mask=rgba.getchannel('A'))
                out = bg
            else:
                out = img.convert('RGB')
            out.save(png_path, 'PNG')
        source_path.unlink(missing_ok=True)
        return png_path
    except Exception as err:
        print(f"WARNING: Could not convert {source_path}: {err}")
        return None

def extract_docx_media_fallback(docx_path: Path, target_dir: Path) -> None:
    media_dir = target_dir / 'media'
    media_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as archive:
        for item in archive.infolist():
            if not item.filename.startswith('word/media/'):
                continue
            destination = media_dir / Path(item.filename).name
            if not destination.exists():
                with archive.open(item) as src, open(destination, 'wb') as dst:
                    shutil.copyfileobj(src, dst)

def normalize_media_files(media_dir: Path) -> None:
    if not media_dir.exists():
        return
    for media_file in list(media_dir.iterdir()):
        if media_file.suffix.lower() in {'.emf', '.wmf'}:
            convert_emf_or_wmf_to_png(media_file)

def extract_session_title(docx_path: Path, session_number: int) -> str:
    default_title = f'جلسه {session_number}'
    try:
        doc = Document(docx_path)
        for p in doc.paragraphs:
            t = p.text.strip()
            if t:
                return t.replace('#', '').strip()
    except Exception as err:
        print(f"WARNING: Could not read title from {docx_path}: {err}")
    return default_title

def image_target_path(source_ref: str, session_tag: str) -> str:
    source_ref = source_ref.strip().replace('\\', '/').split('#')[0].split('?')[0]
    filename = Path(source_ref).name
    stem = Path(filename).stem
    ext = Path(filename).suffix.lower()
    if ext in {'.emf', '.wmf'}:
        ext = '.png'
    return f'assets/{session_tag}/media/{stem}{ext}'

def convert_html_images_to_markdown(content: str, session_tag: str) -> str:
    img_tag_pat = re.compile(r'<img\b[^>]*>', flags=re.IGNORECASE | re.DOTALL)
    def repl(m: re.Match) -> str:
        tag = m.group(0)
        src_m = re.search(r'''src\s*=\s*(['"])(.*?)\1''', tag, flags=re.IGNORECASE | re.DOTALL)
        alt_m = re.search(r'''alt\s*=\s*(['"])(.*?)\1''', tag, flags=re.IGNORECASE | re.DOTALL)
        if not src_m:
            return tag
        alt = re.sub(r'\s+', ' ', alt_m.group(2)).strip() if alt_m else 'تصویر'
        alt = alt.replace('[', r'\[').replace(']', r'\]')
        target = image_target_path(src_m.group(2), session_tag)
        return f'\n\n![{alt}]({target})\n\n'
    return img_tag_pat.sub(repl, content)

def clean_markdown_images(content: str, session_tag: str) -> str:
    md_img_pat = re.compile(r'!\[([^\]]*)\]\(([^)\s]+)(?:\s+[^)]*)?\)', flags=re.DOTALL)
    def repl(m: re.Match) -> str:
        alt = re.sub(r'\s+', ' ', m.group(1)).strip()
        target = image_target_path(m.group(2), session_tag)
        return f'![{alt}]({target})'
    return md_img_pat.sub(repl, content)

def protect_rtl_numbers(text: str) -> str:
    def wrap_num(m: re.Match) -> str:
        val = m.group(0)
        if val.startswith('<span'):
            return val
        return f'<span class="num-ltr" dir="ltr">{val}</span>'

    triple_tick = chr(96) * 3
    single_tick = chr(96)
    fenced_parts = text.split(triple_tick)
    for idx in range(0, len(fenced_parts), 2):
        chunk = fenced_parts[idx]
        regex_str = f'({single_tick}[^{single_tick}]*{single_tick})|(\\[[^\\]]*\\]\\([^\\)]*\\))|(<[^>]+>)'
        prot_pat = re.compile(regex_str, flags=re.DOTALL)
        segments = []
        last = 0
        for match in prot_pat.finditer(chunk):
            norm = chunk[last:match.start()]
            segments.append(NUMBER_PATTERN.sub(wrap_num, norm))
            segments.append(match.group(0))
            last = match.end()
        segments.append(NUMBER_PATTERN.sub(wrap_num, chunk[last:]))
        fenced_parts[idx] = ''.join(segments)
    return triple_tick.join(fenced_parts)

all_docx = [Path(f) for f in glob.glob('*.docx') if not Path(f).name.startswith('~$')]
named_docx = [p for p in all_docx if re.search(r'lab_?manual\d+.*\.docx$', p.name, flags=re.IGNORECASE)]
docx_files = sorted(named_docx if named_docx else all_docx, key=lambda p: natural_sort_key(p.name))

if not docx_files:
    raise RuntimeError('No DOCX files found in project root.')

print(f'Found {len(docx_files)} session file(s):')
for p in docx_files:
    print(f'  - {p}')

sessions: list[tuple[str, str]] = []

for idx, docx_path in enumerate(docx_files, start=1):
    if not zipfile.is_zipfile(docx_path):
        print(f'WARNING: Skipping invalid DOCX: {docx_path}')
        continue

    session_tag = f'session_{idx:02d}'
    md_filename = f'{session_tag}.md'
    md_path = DOCS_DIR / md_filename
    session_title = extract_session_title(docx_path, idx)
    session_asset_dir = ASSETS_DIR / session_tag
    session_asset_dir.mkdir(parents=True, exist_ok=True)

    print(f'Converting [{idx}/{len(docx_files)}]: {docx_path.name} -> docs/{md_filename}')

    pandoc_cmd = [
        'pandoc',
        str(docx_path),
        '-f', 'docx',
        '-t', 'gfm',
        '--wrap=none',
        f'--extract-media={session_asset_dir}',
        '-o', str(md_path)
    ]
    run_command(pandoc_cmd)

    extract_docx_media_fallback(docx_path, session_asset_dir)
    normalize_media_files(session_asset_dir / 'media')

    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    content = convert_html_images_to_markdown(content, session_tag)
    content = clean_markdown_images(content, session_tag)
    content = re.sub(r'\.(?:emf|wmf)(?=[)\s"])', '.png', content, flags=re.IGNORECASE)
    content = protect_rtl_numbers(content)

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content)

    sessions.append((session_title, md_filename))

index_lines = [
    '# راهنمای جامع آزمایشگاه فارماکوگنوزی ۱',
    '',
    'به پورتال آموزش عملی و آزمایشگاهی فارماکوگنوزی خوش آمدید.',
    'برای مشاهده دستور کار هر جلسه، از منوی کناری یا لینک‌های زیر استفاده نمایید:',
    '',
    '## فهرست جلسات آزمایشگاه',
    '',
]
for title, filename in sessions:
    index_lines.append(f'- [{title}]({filename})')
index_lines.append('')

with open(DOCS_DIR / 'index.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(index_lines))

nav_lines = ['nav:', "  - 'صفحه اصلی': index.md", "  - 'جلسات آزمایشگاه':"]
for title, filename in sessions:
    safe_title = title.replace("'", "''")
    nav_lines.append(f"      - '{safe_title}': {filename}")

mkdocs_yaml = f"""site_name: "{SITE_NAME}"
site_description: "{SITE_DESCRIPTION}"
site_author: "{SITE_AUTHOR}"
docs_dir: "docs"
site_dir: "site"

theme:
  name: material
  language: fa
  direction: rtl
  palette:
    - scheme: default
      primary: teal
      accent: indigo
      toggle:
        icon: material/brightness-7
        name: حالت تاریک
    - scheme: slate
      primary: teal
      accent: indigo
      toggle:
        icon: material/brightness-4
        name: حالت روشن
  features:
    - navigation.instant
    - navigation.tracking
    - navigation.top
    - navigation.tabs
    - search.suggest
    - search.highlight
    - content.code.copy

extra_css:
  - stylesheets/extra.css

markdown_extensions:
  - admonition
  - pymdownx.details
  - pymdownx.superfences
  - pymdownx.tabbed:
      alternate_style: true
  - pymdownx.arithmatex:
      generic: true

{chr(10).join(nav_lines)}
"""

with open('mkdocs.yml', 'w', encoding='utf-8') as f:
    f.write(mkdocs_yaml)

print('\n--- Pipeline completed successfully! ---')
