#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Konwerter Markdown -> Word (.docx) dla materiałów egzaminacyjnych.
Times New Roman 12, każda teza na osobnej stronie, tabele, wyróżnienia.
"""
import os
import re
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = 'Times New Roman'
BASE = 12

# Kolory
NAVY = RGBColor(0x1F, 0x38, 0x64)      # nagłówki tez
DARKRED = RGBColor(0x8B, 0x1A, 0x1A)   # nagłówki sekcji
GREY = RGBColor(0x40, 0x40, 0x40)
SHADE_KEY = 'FFF2CC'      # zdanie-klucz (jasny amber)
SHADE_ASK = 'DEEAF6'      # czego dopyta komisja (jasny błękit)
SHADE_QUOTE = 'F2F2F2'    # cytat/wskazówka (jasny szary)
SHADE_TH = '1F3864'       # nagłówek tabeli (navy)


# ---------------------------------------------------------------- narzędzia XML
def set_run_font(run, size=BASE, bold=False, italic=False, color=None, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(attr), font)


def shade_paragraph(par, hexcolor):
    ppr = par._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    ppr.append(shd)


def par_border(par, edges=('bottom',), size=8, color='1F3864', space=2):
    ppr = par._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    for edge in edges:
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), str(space))
        el.set(qn('w:color'), color)
        pbdr.append(el)
    ppr.append(pbdr)


def shade_cell(cell, hexcolor):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcpr.append(shd)


def add_page_number_footer(section):
    footer = section.footer
    par = footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run('– ')
    set_run_font(run, 10, color=GREY)
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run2 = par.add_run()
    set_run_font(run2, 10, color=GREY)
    run2._r.append(fld_begin)
    run2._r.append(instr)
    run2._r.append(fld_end)
    run3 = par.add_run(' –')
    set_run_font(run3, 10, color=GREY)


def add_header_text(section, text):
    par = section.header.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = par.add_run(text)
    set_run_font(run, 9, italic=True, color=GREY)


# ------------------------------------------------------------- parser inline md
INLINE_RE = re.compile(r'(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`)', re.S)


def add_inline(par, text, size=BASE, base_bold=False, base_italic=False,
               color=None):
    """Dodaje tekst z obsługą **bold**, *italic*, `code`."""
    text = text.replace('\u00a0', ' ')
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            run = par.add_run(part[2:-2])
            set_run_font(run, size, bold=True, italic=base_italic, color=color)
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            run = par.add_run(part[1:-1])
            set_run_font(run, size, bold=base_bold, italic=True, color=color)
        elif part.startswith('`') and part.endswith('`') and len(part) > 2:
            run = par.add_run(part[1:-1])
            set_run_font(run, size - 1, bold=base_bold, italic=base_italic,
                         color=color, font='Consolas')
        else:
            run = par.add_run(part)
            set_run_font(run, size, bold=base_bold, italic=base_italic,
                         color=color)


# ---------------------------------------------------------------- budowa pliku
def new_document():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(BASE)
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(attr), FONT)
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)

    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.0)
    add_page_number_footer(sec)
    return doc


def p(doc, text='', size=BASE, bold=False, italic=False, color=None,
      align=None, space_before=0, space_after=6, indent=0, hanging=None,
      shade=None, keep_with_next=False):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if indent:
        pf.left_indent = Cm(indent)
    if hanging:
        pf.first_line_indent = Cm(-hanging)
    if align is not None:
        par.alignment = align
    if keep_with_next:
        pf.keep_with_next = True
    if shade:
        shade_paragraph(par, shade)
    if text:
        add_inline(par, text, size=size, base_bold=bold, base_italic=italic,
                   color=color)
    return par


def build_table(doc, rows):
    """rows: lista list komórek (pierwsza = nagłówek)."""
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=ncols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        for ci in range(ncols):
            txt = row[ci] if ci < len(row) else ''
            cell = cells[ci]
            cell.text = ''
            par = cell.paragraphs[0]
            par.paragraph_format.space_before = Pt(2)
            par.paragraph_format.space_after = Pt(2)
            par.paragraph_format.line_spacing = 1.0
            if ri == 0:
                shade_cell(cell, SHADE_TH)
                add_inline(par, txt, size=BASE - 1, base_bold=True,
                           color=RGBColor(0xFF, 0xFF, 0xFF))
            else:
                add_inline(par, txt, size=BASE - 1)
    return table


def parse_md_into(doc, lines, first_h2_no_break=False):
    i = 0
    h2_count = 0
    n = len(lines)
    while i < n:
        line = lines[i].rstrip('\n')
        stripped = line.strip()

        # pomijamy puste
        if not stripped:
            i += 1
            continue

        # tabela
        if stripped.startswith('|') and i + 1 < n and \
                re.match(r'^\s*\|[\s:\-|]+\|\s*$', lines[i + 1]):
            rows = []
            header = [c.strip() for c in stripped.strip('|').split('|')]
            rows.append(header)
            i += 2
            while i < n and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                rows.append(cells)
                i += 1
            build_table(doc, rows)
            p(doc, '', space_after=4)
            continue

        # H1
        if stripped.startswith('# '):
            par = p(doc, stripped[2:], size=17, bold=True, color=NAVY,
                    align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10,
                    keep_with_next=True)
            i += 1
            continue

        # H2 - teza (nowa strona)
        if stripped.startswith('## '):
            h2_count += 1
            par = doc.add_paragraph()
            pf = par.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(10)
            pf.keep_with_next = True
            if not (first_h2_no_break and h2_count == 1):
                pf.page_break_before = True
            add_inline(par, stripped[3:], size=14, base_bold=True, color=NAVY)
            par_border(par, edges=('bottom',), size=10, color='1F3864')
            i += 1
            continue

        # H3
        if stripped.startswith('### '):
            p(doc, stripped[4:], size=12.5, bold=True, color=DARKRED,
              space_before=8, space_after=4, keep_with_next=True)
            i += 1
            continue

        # linia pozioma -> pomijamy (podział robią nagłówki)
        if re.match(r'^-{3,}$', stripped):
            i += 1
            continue

        # cytat blokowy
        if stripped.startswith('>'):
            buf = []
            while i < n and lines[i].strip().startswith('>'):
                buf.append(lines[i].strip().lstrip('>').strip())
                i += 1
            text = ' '.join(x for x in buf if x)
            par = p(doc, text, size=BASE - 0.5, italic=False, indent=0.4,
                    space_before=4, space_after=8, shade=SHADE_QUOTE)
            par_border(par, edges=('left',), size=18, color='8B1A1A', space=6)
            continue

        # lista punktowana / numerowana
        m_bul = re.match(r'^(\s*)[-*]\s+(.*)$', line)
        m_num = re.match(r'^(\s*)(\d+[.)])\s+(.*)$', line)
        if m_bul or m_num:
            while i < n:
                cur = lines[i].rstrip('\n')
                mb = re.match(r'^(\s*)[-*]\s+(.*)$', cur)
                mn = re.match(r'^(\s*)(\d+[.)])\s+(.*)$', cur)
                if not (mb or mn):
                    if cur.strip() == '':
                        nxt = lines[i + 1].rstrip('\n') if i + 1 < n else ''
                        if re.match(r'^(\s*)([-*]|\d+[.)])\s+', nxt):
                            i += 1
                            continue
                    break
                if mb:
                    ind_sp, txt = len(mb.group(1)), mb.group(2)
                    lvl = 0 if ind_sp < 2 else (1 if ind_sp < 4 else 2)
                    marker = ['\u2022 ', '\u2013 ', '\u00b7 '][lvl]
                else:
                    ind_sp, marker, txt = len(mn.group(1)), mn.group(2) + ' ', mn.group(3)
                    lvl = 0 if ind_sp < 2 else (1 if ind_sp < 4 else 2)
                left = 0.6 + lvl * 0.55
                par = doc.add_paragraph()
                pf = par.paragraph_format
                pf.left_indent = Cm(left)
                pf.first_line_indent = Cm(-0.45)
                pf.space_before = Pt(0)
                pf.space_after = Pt(3)
                run = par.add_run(marker)
                set_run_font(run, BASE)
                add_inline(par, txt, size=BASE)
                i += 1
            p(doc, '', space_after=3)
            continue

        # akapit zwykły (z wyróżnieniami)
        shade = None
        if stripped.startswith('**Zdanie-klucz'):
            shade = SHADE_KEY
        elif stripped.startswith('**Czego dopyta'):
            shade = SHADE_ASK
        par = p(doc, stripped, size=BASE, space_after=6, shade=shade)
        if shade:
            par.paragraph_format.left_indent = Cm(0.15)
            par.paragraph_format.right_indent = Cm(0.15)
            par.paragraph_format.space_before = Pt(4)
            par.paragraph_format.space_after = Pt(8)
        i += 1
    return h2_count


def title_page(doc, title, subtitle, info_lines, checklist=None):
    p(doc, '', space_after=40)
    p(doc, title, size=24, bold=True, color=NAVY,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    par = p(doc, subtitle, size=13, italic=True, color=GREY,
            align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    par_border(par, edges=('bottom',), size=6, color='BFBFBF', space=8)
    for ln in info_lines:
        p(doc, ln, size=11.5, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3,
          color=GREY)
    if checklist:
        doc.add_page_break()
        p(doc, 'Lista kontrolna — odhaczaj opanowane tezy', size=15, bold=True,
          color=NAVY, space_after=4)
        p(doc, 'Zaznacz kwadrat, gdy potrafisz omówić tezę na głos przez 5 minut '
               'bez zaglądania do notatek.', size=11, italic=True, color=GREY,
          space_after=10)
        for item in checklist:
            par = doc.add_paragraph()
            pf = par.paragraph_format
            pf.left_indent = Cm(0.8)
            pf.first_line_indent = Cm(-0.8)
            pf.space_after = Pt(4)
            run = par.add_run('\u2610  ')
            set_run_font(run, 13, font='Segoe UI Symbol')
            add_inline(par, item, size=BASE)


def collect_h2(lines):
    return [l.strip()[3:] for l in lines if l.strip().startswith('## ')]


def build(md_paths, out_path, title, subtitle, info_lines, with_checklist=True):
    doc = new_document()
    all_lines = []
    for pth in md_paths:
        with open(pth, encoding='utf-8') as f:
            all_lines.append(f.read().split('\n'))
    checklist = []
    if with_checklist:
        for lines in all_lines:
            for h in collect_h2(lines):
                # do listy kontrolnej trafiają tylko tezy (O1..O10, K1..K20)
                if re.match(r'^[OK]\d+\.', h) and h not in checklist:
                    checklist.append(h)
    add_header_text(doc.sections[0], title)
    title_page(doc, title, subtitle, info_lines,
               checklist=checklist if with_checklist else None)
    for idx, lines in enumerate(all_lines):
        # pierwszy H1 z pliku pomijamy na rzecz strony tytułowej tylko
        # gdy dokument dotyczy jednego pliku
        doc.add_page_break()
        parse_md_into(doc, lines, first_h2_no_break=False)
    doc.save(out_path)
    print('OK  ->', out_path)


if __name__ == '__main__':
    # katalog, w którym leży ten skrypt — pliki .md muszą być obok niego
    B = os.path.join(os.path.dirname(os.path.abspath(__file__)), '')
    OG = B + 'tezy-ogolne-skrot.md'
    KI = B + 'tezy-kierunkowe-skrot.md'
    PL = B + 'plan-nauki-5-dni.md'
    MX = B + 'tezy-maks-skrot.md'

    # --- NAJWAŻNIEJSZY dokument: maks-skrót 30 tez + plan 5 dni -------------
    build([MX], B + 'Tezy-MAKS-SKROT.docx',
          'MAKS-SKRÓT — 30 tez',
          'Egzamin magisterski — pedagogika przedszkolna i wczesnoszkolna',
          ['Tylko najważniejsze informacje: definicja, nazwiska, rdzeń, zdanie-klucz',
           'Egzamin: środa 23 września · ustny, 2 wylosowane pytania z 30 tez',
           'Times New Roman 12 · każda teza na osobnej stronie'])

    build([PL, MX], B + 'DO-NAUKI-maks-skrot-i-plan.docx',
          'Do nauki w 5 dni',
          'Plan nauki + maks-skrót 30 tez',
          ['Plan na 5 dni: 18–22 września (niedziela 20.09 — 8 godzin)',
           'Egzamin: środa 23 września · ustny, 2 wylosowane pytania z 30 tez',
           'Times New Roman 12 · do druku dwustronnego'])

    build([OG], B + 'Tezy-OGOLNE-skrot.docx',
          'Zagadnienia ogólne',
          'Egzamin magisterski — pedagogika przedszkolna i wczesnoszkolna',
          ['10 tez ogólnych — skrót do nauki',
           'Format egzaminu: ustny, 2 wylosowane pytania z 30 tez',
           'Times New Roman 12 · każda teza na osobnej stronie'])

    build([KI], B + 'Tezy-KIERUNKOWE-skrot.docx',
          'Zagadnienia kierunkowe',
          'Egzamin magisterski — pedagogika przedszkolna i wczesnoszkolna',
          ['20 tez kierunkowych — skrót do nauki',
           'Format egzaminu: ustny, 2 wylosowane pytania z 30 tez',
           'Times New Roman 12 · każda teza na osobnej stronie'])

    build([PL], B + 'Plan-nauki-5-dni.docx',
          'Plan nauki na 5 dni',
          'Egzamin magisterski — pedagogika przedszkolna i wczesnoszkolna',
          ['Egzamin: środa 23 września · nauka 18–22 września',
           '4 dni po 4 godziny + niedziela 20.09 osiem godzin',
           'Szkielet odpowiedzi ustnej · liczby do wykucia'],
          with_checklist=False)

    build([PL, MX, OG, KI], B + 'KOMPLET-tezy-i-plan.docx',
          'Materiały do egzaminu magisterskiego',
          'Pedagogika przedszkolna i wczesnoszkolna — 30 tez + plan nauki',
          ['Plan nauki 18–22 września · maks-skrót 30 tez · pełne wersje tez',
           'Egzamin: środa 23 września · ustny, 2 wylosowane pytania',
           'Times New Roman 12 · do druku dwustronnego'])
