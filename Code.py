import os
import re
import copy
import glob
import json
import time
import zipfile
import tempfile
import shutil
import ipywidgets as widgets
from IPython.display import display, clear_output, HTML
import win32com.client as win32

from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ─── CONFIGURACIÓN DE GRADOS Y NIVELES ─────────────────────────

GRADES_INFO = {
    "Prejardín": {"level": "Preescolar", "multicourse": False},
    "Jardín": {"level": "Preescolar", "multicourse": False},
    "Transición": {"level": "Preescolar", "multicourse": False},
    "1°": {"level": "Básica Primaria", "multicourse": False},
    "2°": {"level": "Básica Primaria", "multicourse": False},
    "3°": {"level": "Básica Primaria", "multicourse": False},
    "4°": {"level": "Básica Primaria", "multicourse": False},
    "5°": {"level": "Básica Primaria", "multicourse": False},
    "6°": {"level": "Básica Secundaria", "multicourse": True},
    "7°": {"level": "Básica Secundaria", "multicourse": True},
    "8°": {"level": "Básica Secundaria", "multicourse": True},
    "9°": {"level": "Básica Secundaria", "multicourse": True},
    "10°": {"level": "Media Vocacional", "multicourse": True},
    "11°": {"level": "Media Vocacional", "multicourse": True}
}

MONTHS_LIST = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC']


SHORTCODES = {
    "grade": "Grado (ej. 2°)",
    "period": "Periodo (ej. P3)",
    "session": "Código de subsesión (ej. S1.1)",
    "year": "Año (ej. 2026)",
    "level": "Nivel educativo (ej. Media Vocacional)",
}


def expand_template(template, context):
    return re.sub(r'\{(\w+)\}', lambda m: str(context.get(m.group(1), m.group(0))), template)


# ─── helpers ───────────────────────────────────────────────────

CM_TO_PT = 72.0 / 2.54  # 1 cm = 28.3465 pt


def add_dynamic_page_number_to_footer(paragraph):
    font_name = "Century Gothic"
    font_size = Pt(11)
    run_text1 = paragraph.add_run("Página ")
    run_text1.font.name = font_name; run_text1.font.size = font_size
    run_page = paragraph.add_run()
    run_page.font.name = font_name; run_page.font.size = font_size; run_page.bold = True
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr1 = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    run_page._r.extend([fld1, instr1, fld2, fld3])
    run_text2 = paragraph.add_run(" de ")
    run_text2.font.name = font_name; run_text2.font.size = font_size
    run_numpages = paragraph.add_run()
    run_numpages.font.name = font_name; run_numpages.font.size = font_size; run_numpages.bold = True
    fld4 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr2 = parse_xml(r'<w:instrText %s xml:space="preserve"> NUMPAGES </w:instrText>' % nsdecls('w'))
    fld5 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld6 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    run_numpages._r.extend([fld4, instr2, fld5, fld6])


def setup_footer_page_number(footer):
    for p in footer.paragraphs: p.text = ""
    for table in footer.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: p.text = ""
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_dynamic_page_number_to_footer(fp)


def force_single_column(section):
    sectPr = section._sectPr
    cols = sectPr.xpath('w:cols')
    if cols:
        cols[0].set(qn('w:num'), '1')
        for attr in ['w:space', 'w:equalWidth']:
            if qn(attr) in cols[0].attrib:
                del cols[0].attrib[qn(attr)]
    else:
        cols_xml = parse_xml(f'<w:cols {nsdecls("w")} w:num="1"/>')
        sectPr.append(cols_xml)


def apply_page_setup(section):
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    section.header_distance = Cm(0.3)
    force_single_column(section)


def get_all_paragraphs(doc):
    from docx.text.paragraph import Paragraph
    return [Paragraph(p, doc) for p in doc.element.body.xpath('.//w:p')]


def _prepend_run(paragraph, text, bold=False):
    """Add a run at the BEGINNING of the paragraph."""
    escaped = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    new_r = parse_xml(f'<w:r {nsdecls("w")}><w:t xml:space="preserve">{escaped}</w:t></w:r>')
    first = paragraph._element.find(qn('w:r'))
    if first is not None:
        first.addprevious(new_r)
    else:
        paragraph._element.append(new_r)
    from docx.text.run import Run
    r = Run(new_r, paragraph)
    if bold:
        r.bold = True
    r.font.name = "Century Gothic"
    r.font.size = Pt(11)
    return r


def _remove_numpr(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    np = pPr.find(qn('w:numPr'))
    if np is not None:
        pPr.remove(np)


def _strip_prefix_from_runs(paragraph, length):
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for r_elem in paragraph._element.xpath('.//w:r'):
        if length <= 0:
            break
        t_elems = r_elem.findall(f'{{{wns}}}t')
        for t_elem in t_elems:
            if length <= 0:
                break
            t = t_elem.text or ""
            if not t:
                continue
            if len(t) <= length:
                length -= len(t)
                t_elem.text = ""
            else:
                t_elem.text = t[length:]
                length = 0


def _replace_paragraph_text_preserve_drawings(paragraph, new_text):
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    t_elems = paragraph._element.xpath('.//w:t')
    if not t_elems:
        return
    t_elems[0].text = new_text
    for t_elem in t_elems[1:]:
        t_elem.text = ""


def apply_renumbering_and_ranges(doc, start_number):
    all_paras = get_all_paragraphs(doc)
    mapping = {}
    q_paras = []
    cur = start_number
    for p in all_paras:
        text = p.text.strip()
        if not text:
            continue
        m = re.match(r'^(\s*[\(\[\{]?)(\d+)(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+)(?![\d\.])', text)
        if m:
            orig = int(m.group(2))
            if orig > 200:
                continue
            mapping[orig] = cur
            q_paras.append((p, orig, cur))
            cur += 1

    already_correct = q_paras and all(orig == new for _, orig, new in q_paras)

    if not already_correct:
        for p, orig_val, new_val in q_paras:
            mp = re.match(r'^(\s*[\(\[\{]?\d+(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+))(?![\d\.])', p.text)
            if mp:
                _strip_prefix_from_runs(p, len(mp.group(1)))
                _prepend_run(p, f"{new_val}. ", bold=True)
                _remove_numpr(p)
    else:
        for p, orig_val, new_val in q_paras:
            text = p.text
            mp = re.match(r'^(\s*[\(\[\{]?\d+(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+))(?![\d\.])', text)
            if mp:
                first = p.runs[0] if p.runs else None
                if first is None or not first.bold:
                    _strip_prefix_from_runs(p, len(mp.group(1)))
                    _prepend_run(p, f"{new_val}. ", bold=True)
                _remove_numpr(p)

    def replace_refs(t):
        if not mapping:
            return t
        t = re.sub(
            r'(?i)(de\s+la[s]?\s+(?:pregunta[s]?\s+)?)(\d+)(\s+(?:a|hasta)\s+la[s]?\s+(?:pregunta[s]?\s+)?)(\d+)',
            lambda m: m.group(1) + str(mapping.get(int(m.group(2)), int(m.group(2)))) + m.group(3) + str(mapping.get(int(m.group(4)), int(m.group(4)))), t)
        t = re.sub(
            r'(?i)(del\s+(?:pregunta[s]?\s+)?)(\d+)(\s+(?:al|hasta\s+el)\s+(?:pregunta[s]?\s+)?)(\d+)',
            lambda m: m.group(1) + str(mapping.get(int(m.group(2)), int(m.group(2)))) + m.group(3) + str(mapping.get(int(m.group(4)), int(m.group(4)))), t)
        t = re.sub(
            r'(?i)(pregunta[s]?\s+)(\d+)(\s+(?:a|al|hasta|y)\s+(?:pregunta[s]?\s+)?)(\d+)',
            lambda m: m.group(1) + str(mapping.get(int(m.group(2)), int(m.group(2)))) + m.group(3) + str(mapping.get(int(m.group(4)), int(m.group(4)))), t)
        return t
    for p in all_paras:
        old = p.text
        if not old.strip():
            continue
        new = replace_refs(old)
        if old != new:
            _replace_paragraph_text_preserve_drawings(p, new)
    return cur


def _resolve_autonumbering(doc):
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    try:
        numbering_part = doc.part.numbering_part
        if numbering_part is None:
            return
        num_xml = numbering_part.element
    except:
        return

    abstract_levels = {}
    for abstract_num in num_xml.findall(f'{{{wns}}}abstractNum'):
        aid = int(abstract_num.get(qn('w:abstractNumId')))
        levels = {}
        for lvl in abstract_num.findall(f'{{{wns}}}lvl'):
            ilvl = int(lvl.get(qn('w:ilvl')) or '0')
            nf = lvl.find(f'{{{wns}}}numFmt')
            fmt = nf.get(qn('w:val')) if nf is not None else 'decimal'
            lt = lvl.find(f'{{{wns}}}lvlText')
            tpl = lt.get(qn('w:val')) if lt is not None else '%1.'
            st = lvl.find(f'{{{wns}}}start')
            start = int(st.get(qn('w:val'))) if st is not None else 1
            levels[ilvl] = (fmt, tpl, start)
        abstract_levels[aid] = levels

    num_to_abstract = {}
    for num_elem in num_xml.findall(f'{{{wns}}}num'):
        nid = int(num_elem.get(qn('w:numId')))
        aid_elem = num_elem.find(f'{{{wns}}}abstractNumId')
        if aid_elem is not None:
            num_to_abstract[nid] = int(aid_elem.get(qn('w:val')))

    counters = {}

    def level_value(num_id, ilvl, levels):
        k = (num_id, ilvl)
        if k not in counters:
            info = levels.get(ilvl)
            return info[2] if info else 1
        return counters[k]

    def fmt_num(val, fmt_name):
        if fmt_name in ('decimal', 'ordinal'):
            return str(val)
        elif fmt_name in ('upperLetter', 'upperAlpha'):
            return chr(ord('A') + val - 1) if 1 <= val <= 26 else str(val)
        elif fmt_name in ('lowerLetter', 'lowerAlpha'):
            return chr(ord('a') + val - 1) if 1 <= val <= 26 else str(val)
        return str(val)

    for p in get_all_paragraphs(doc):
        pPr = p._element.get_or_add_pPr()
        np = pPr.find(qn('w:numPr'))
        if np is None:
            continue

        nid_elem = np.find(qn('w:numId'))
        ilvl_elem = np.find(qn('w:ilvl'))
        if nid_elem is None:
            continue

        num_id = int(nid_elem.get(qn('w:val')))
        ilvl = int(ilvl_elem.get(qn('w:val'))) if ilvl_elem is not None else 0

        text = p.text.strip()
        if re.match(r'^(\d+|[a-eA-E])\s*[\.\)\]\}\-\:\/]', text):
            continue

        abstract_id = num_to_abstract.get(num_id)
        if abstract_id is None:
            continue
        levels = abstract_levels.get(abstract_id, {})
        level_info = levels.get(ilvl)
        if level_info is None:
            continue

        fmt, tpl, start_val = level_info

        cur_key = (num_id, ilvl)
        if cur_key not in counters:
            counters[cur_key] = start_val
        else:
            counters[cur_key] += 1

        for sub_ilvl in range(ilvl + 1, 10):
            sub_key = (num_id, sub_ilvl)
            if sub_key in counters:
                info = levels.get(sub_ilvl)
                counters[sub_key] = info[2] if info else 1

        def resolve_ph(m):
            idx = int(m.group(1))
            target_ilvl = idx - 1
            val = level_value(num_id, target_ilvl, levels)
            info = levels.get(target_ilvl)
            lvl_fmt = info[0] if info else 'decimal'
            return fmt_num(val, lvl_fmt)

        prefix = re.sub(r'%(\d+)', resolve_ph, tpl) + ' '
        _prepend_run(p, prefix, bold=True)
        _remove_numpr(p)


def remove_indents(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        pPr.remove(ind)
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.right_indent = Pt(0)
    paragraph.paragraph_format.first_line_indent = Pt(0)


def _clear_paragraph_style_level(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    style_elem = pPr.find(qn('w:pStyle'))
    if style_elem is not None:
        pPr.remove(style_elem)
    outline_elem = pPr.find(qn('w:outlineLvl'))
    if outline_elem is not None:
        pPr.remove(outline_elem)


def format_paragraph(paragraph, doc_ref):
    text = paragraph.text
    if not text.strip():
        return

    # Strip all indents by default for clean flush-left alignment
    remove_indents(paragraph)

    if re.search(r'(NOMBRE|APELLIDO|ESTUDIANTE|ALUMNO|CURSO|GRADO|FECHA|CÓDIGO|CODIGO|NAME|DATE|CODE)\s*:\s*[_]{2,}', text, re.IGNORECASE):
        if _has_drawing(paragraph._element):
            for t in paragraph._element.xpath('.//w:t'):
                t.text = ""
        else:
            paragraph.text = ""
        return
    mc = re.match(r'^(Componente|Competencia)\s*:\s*(.*)$', text, re.IGNORECASE)
    if mc:
        label = mc.group(1).capitalize()
        content = mc.group(2).strip()
        if _has_drawing(paragraph._element):
            t_elems = paragraph._element.xpath('.//w:t')
            if t_elems:
                t_elems[0].text = f"{label}: {content}"
                for t in t_elems[1:]:
                    t.text = ""
        else:
            paragraph.text = ""
            rb = paragraph.add_run(f"{label}: "); rb.bold = True
            rt = paragraph.add_run(content); rt.bold = False
        return
    # Option letter A–E — format prefix
    mo = re.match(r'^(\s*[\(\[\{]?([a-eA-E])\s*[\.\)\]\}\-\:\/]\s)(?!\d)', text)
    if mo:
        runs = paragraph.runs
        first = runs[0] if runs else None
        already_good = (
            first is not None
            and first.text.strip().upper() == mo.group(2).upper()
            and first.bold
        )
        if not already_good:
            _strip_prefix_from_runs(paragraph, len(mo.group(1)))
            _prepend_run(paragraph, f"{mo.group(2).upper()}. ", bold=True)
            _remove_numpr(paragraph)
        remove_indents(paragraph)
        return
    # Question number — only bold if not already
    mn = re.match(r'^(\s*[\(\[\{]?(\d+)(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+))(?![\d\.])', text)
    if mn:
        runs = paragraph.runs
        first = runs[0] if runs else None
        already_bold = first is not None and first.bold
        if not already_bold:
            _strip_prefix_from_runs(paragraph, len(mn.group(1)))
            _prepend_run(paragraph, f"{mn.group(2)}. ", bold=True)
        _remove_numpr(paragraph)
        remove_indents(paragraph)


def reset_letter_sequence(doc):
    paras = get_all_paragraphs(doc)
    letter_index = 0

    for p in paras:
        text = p.text.strip()
        if not text:
            continue

        mn = re.match(r'^(\s*[\(\[\{]?)(\d+)(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+)(?![\d\.])', text)
        if mn and int(mn.group(2)) <= 200:
            letter_index = 0
            continue

        mo = re.match(r'^(\s*[\(\[\{]?)([a-eA-E])(\s*[\.\)\]\}\-\:\/]\s)', text)
        if mo:
            current = mo.group(2).upper()
            expected = chr(ord('A') + letter_index) if letter_index < 5 else None

            if expected and current != expected:
                full_prefix = mo.group(0)
                _strip_prefix_from_runs(p, len(full_prefix))
                _prepend_run(p, f"{expected}. ", bold=True)
                _remove_numpr(p)

            letter_index += 1
            continue


def set_single_line_spacing(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = parse_xml(f'<w:spacing {nsdecls("w")} w:line="240" w:lineRule="auto" w:before="0" w:after="0"/>')
        pPr.append(sp)
    else:
        sp.set(qn('w:line'), '240')
        sp.set(qn('w:lineRule'), 'auto')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '0')


def _has_drawing(para_element):
    """Check if a paragraph element contains any drawing, picture, shape or embedded object."""
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    vns = 'urn:schemas-microsoft-com:vml'
    return (
        len(para_element.findall(f'.//{{{wns}}}drawing')) > 0 or
        len(para_element.findall(f'.//{{{wns}}}pict')) > 0 or
        len(para_element.findall(f'.//{{{vns}}}shape')) > 0 or
        len(para_element.findall(f'.//{{{vns}}}imagedata')) > 0 or
        len(para_element.findall(f'.//{{{wns}}}object')) > 0
    )


def inject_list_definitions(doc, start_number=1):
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    try:
        num_part = doc.part.numbering_part
    except NotImplementedError:
        return None
    
    if num_part is None:
        return None

    num_xml = num_part.element
    if not num_xml.xpath('w:abstractNum[@w:abstractNumId="9000"]'):
        # Decimal list abstractNum (Questions)
        abs_dec = parse_xml(f'''
            <w:abstractNum {nsdecls('w')} w:abstractNumId="9000">
                <w:multiLevelType w:val="hybridMultilevel"/>
                <w:lvl w:ilvl="0">
                    <w:start w:val="1"/>
                    <w:numFmt w:val="decimal"/>
                    <w:lvlText w:val="%1."/>
                    <w:lvlJc w:val="left"/>
                    <w:suff w:val="space"/>
                    <w:pPr>
                        <w:ind w:left="360" w:hanging="360"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Century Gothic" w:hAnsi="Century Gothic" w:cs="Century Gothic"/>
                        <w:sz w:val="22"/>
                        <w:szCs w:val="22"/>
                        <w:b w:val="1"/>
                    </w:rPr>
                </w:lvl>
            </w:abstractNum>
        ''')
        
        # UpperLetter list abstractNum (Options - same left indent as Questions)
        abs_alpha = parse_xml(f'''
            <w:abstractNum {nsdecls('w')} w:abstractNumId="9001">
                <w:multiLevelType w:val="hybridMultilevel"/>
                <w:lvl w:ilvl="0">
                    <w:start w:val="1"/>
                    <w:numFmt w:val="upperLetter"/>
                    <w:lvlText w:val="%1."/>
                    <w:lvlJc w:val="left"/>
                    <w:suff w:val="space"/>
                    <w:pPr>
                        <w:ind w:left="360" w:hanging="360"/>
                    </w:pPr>
                    <w:rPr>
                        <w:rFonts w:ascii="Century Gothic" w:hAnsi="Century Gothic" w:cs="Century Gothic"/>
                        <w:sz w:val="22"/>
                        <w:szCs w:val="22"/>
                        <w:b w:val="1"/>
                    </w:rPr>
                </w:lvl>
            </w:abstractNum>
        ''')
        
        first_num = num_xml.find(qn('w:num'))
        if first_num is not None:
            first_num.addprevious(abs_dec)
            first_num.addprevious(abs_alpha)
        else:
            num_xml.append(abs_dec)
            num_xml.append(abs_alpha)
            
        num_dec = parse_xml(f'''
            <w:num {nsdecls('w')} w:numId="9000">
                <w:abstractNumId w:val="9000"/>
                <w:lvlOverride w:ilvl="0">
                    <w:startOverride w:val="{start_number}"/>
                </w:lvlOverride>
            </w:num>
        ''')
        num_xml.append(num_dec)
    else:
        # Update startOverride on existing num 9000
        existing_num = num_xml.xpath('w:num[@w:numId="9000"]')
        if existing_num:
            n_elem = existing_num[0]
            lvl_ovr = n_elem.xpath('w:lvlOverride')
            if not lvl_ovr:
                lvl_ovr_elem = parse_xml(f'<w:lvlOverride {nsdecls("w")} w:ilvl="0"><w:startOverride w:val="{start_number}"/></w:lvlOverride>')
                n_elem.append(lvl_ovr_elem)
            else:
                st = lvl_ovr[0].xpath('w:startOverride')
                if st:
                    st[0].set(qn('w:val'), str(start_number))
                else:
                    st_elem = parse_xml(f'<w:startOverride {nsdecls("w")} w:val="{start_number}"/>')
                    lvl_ovr[0].append(st_elem)

    return True


def strip_leading_tabs(para):
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    for t_elem in para._element.xpath('.//w:t'):
        if not t_elem.text:
            continue
        cleaned = t_elem.text.lstrip('\t')
        if cleaned != t_elem.text:
            t_elem.text = cleaned
        if t_elem.text:
            break


def apply_native_lists_to_final_doc(final_doc, start_offset=0):
    start_num = start_offset + 1
    inject_list_definitions(final_doc, start_number=start_num)
    q_num_id = 9000
    o_num_id_base = 9100
    o_num_id = o_num_id_base
    
    try:
        num_xml = final_doc.part.numbering_part.element
    except (NotImplementedError, AttributeError):
        num_xml = None
        
    paras = get_all_paragraphs(final_doc)
    for p in paras:
        strip_leading_tabs(p)
        text = p.text.strip()
        
        # Check Option
        mo = re.match(r'^(\s*[\(\[\{]?([a-eA-E])\s*[\.\)\]\}\-\:\/]\s)(?!\d)', text)
        if mo:
            _strip_prefix_from_runs(p, len(mo.group(1)))
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                pPr.remove(numPr)
            numPr = parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{o_num_id}"/></w:numPr>')
            pPr.append(numPr)
            strip_leading_tabs(p)
            continue
            
        # Check Question
        mn = re.match(r'^(\s*[\(\[\{]?(\d+)(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+))(?![\d\.])', text)
        if mn:
            _strip_prefix_from_runs(p, len(mn.group(1)))
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                pPr.remove(numPr)
            numPr = parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{q_num_id}"/></w:numPr>')
            pPr.append(numPr)
            strip_leading_tabs(p)
            
            # Restart options for this question
            if num_xml is not None:
                o_num_id += 1
                new_num = parse_xml(f'''
                    <w:num {nsdecls('w')} w:numId="{o_num_id}">
                        <w:abstractNumId w:val="9001"/>
                        <w:lvlOverride w:ilvl="0">
                            <w:startOverride w:val="1"/>
                        </w:lvlOverride>
                    </w:num>
                ''')
                num_xml.append(new_num)
            continue


def apply_formatting_to_document(doc):
    font_name = "Century Gothic"
    font_size = Pt(11)
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
    
    # Process all paragraphs in the document body
    for p_elem in doc.element.body.xpath('.//w:p'):
        para = Paragraph(p_elem, doc)
        
        if not _has_drawing(p_elem):
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
        _remove_numpr(para)
        _clear_paragraph_style_level(para)
        format_paragraph(para, doc)
        set_single_line_spacing(para)
        strip_leading_tabs(para)
        
        # Ensure the paragraph mark gets the correct font size and name
        pPr = para._element.get_or_add_pPr()
        rPr = pPr.find(qn('w:rPr'))
        if rPr is None:
            rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
            pPr.append(rPr)
        
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
            rPr.append(rFonts)
        else:
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)
            
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="22"/>')
            rPr.append(sz)
        else:
            sz.set(qn('w:val'), '22')
            
        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = parse_xml(f'<w:szCs {nsdecls("w")} w:val="22"/>')
            rPr.append(szCs)
        else:
            szCs.set(qn('w:val'), '22')

        # Format all runs, including those nested in hyper-links or other fields
        for r_elem in p_elem.xpath('.//w:r'):
            run = Run(r_elem, para)
            run.font.name = font_name
            run.font.size = font_size


def replace_xml_text(elements_list, replacements):
    wns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    dns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for el in elements_list:
        for t in el.iter(f'{{{wns}}}t'):
            if t.text:
                for k, v in replacements.items():
                    if k in t.text:
                        t.text = t.text.replace(k, str(v))
        for t in el.iter(f'{{{dns}}}t'):
            if t.text:
                for k, v in replacements.items():
                    if k in t.text:
                        t.text = t.text.replace(k, str(v))


def count_questions_in_doc(doc):
    c = 0
    for para in doc.paragraphs:
        if re.match(r'^(\d+)(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+)(?![\d\.])', para.text.strip()):
            c += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if re.match(r'^(\d+)(?:\s*[\.\)\]\}\-\:\/]+\s*|\s+)(?![\d\.])', para.text.strip()):
                        c += 1
    return c


def make_header_tables_inline(header):
    for table in header.tables:
        tblPr = table._element.xpath('w:tblPr')
        if tblPr:
            tblpPr = tblPr[0].xpath('w:tblpPr')
            if tblpPr:
                tblPr[0].remove(tblpPr[0])


def clear_header_completely(header):
    header.is_linked_to_previous = False
    for p in list(header.paragraphs):
        p.text = ""
    for t in list(header.tables):
        try:
            t._element.getparent().remove(t._element)
        except:
            pass
    header._element.clear()


def apply_section0_page_setup(section):
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Cm(5.0)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)
    force_single_column(section)


def apply_subsequent_page_setup(section):
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.8)
    force_single_column(section)


def replace_in_all(tpl, replacements_map):
    """Replace placeholders in the entire template (body + headers)."""
    # Replace in body (user already moved header content here)
    replace_xml_text([tpl.element.body], replacements_map)
    # Also replace in any leftover headers (defensive)
    for s in tpl.sections:
        for h in [s.header, s.first_page_header]:
            replace_xml_text(list(h._element), replacements_map)


# ── MAIN MERGE ───────────────────────────────────────────────

def merge_docx_with_guaranteed_header(template_path, file_list, output_path, config_data, start_offset=0):
    replacements_map = {
        "(EDU_LEVEL)": config_data['level'].upper(),
        "(GRADE)": config_data['grade_clean'].upper(),
        "(TERM)": config_data['period'].upper(),
        "(SESSION)": config_data['session_code'].upper(),
        "(DATE)": config_data['date'].upper(),
        "(P_C)": config_data['p_c_value'],
    }

    title_context = {k: config_data.get(k, '') for k in ['grade_clean', 'period', 'session_code', 'year', 'level']}
    title_context['grade'] = title_context['grade_clean']
    title_context['session'] = title_context['session_code']
    title_template = config_data.get('title_template', '')
    if not title_template:
        title_template = "Evaluación de Suficiencia Académica - {grade_clean} - {period} - {session_code} - {year}"

    # ── Pre-process each sub-doc ───────────────────────────────
    cur = start_offset + 1
    tmp_dir = tempfile.mkdtemp(prefix='gesa_sub_')
    temp_subs = []
    for fp in file_list:
        if not os.path.exists(fp):
            continue
        sd = Document(fp)
        _resolve_autonumbering(sd)
        cur = apply_renumbering_and_ranges(sd, cur)
        apply_formatting_to_document(sd)
        reset_letter_sequence(sd)

        # Clear headers/footers from sub-documents so they inherit the template's
        for sec in sd.sections:
            force_single_column(sec)
            sectPr = sec._sectPr
            for ref in list(sectPr.findall(qn('w:headerReference'))):
                sectPr.remove(ref)
            for ref in list(sectPr.findall(qn('w:footerReference'))):
                sectPr.remove(ref)
            try:
                titlePg = sectPr.find(qn('w:titlePg'))
                if titlePg is not None:
                    sectPr.remove(titlePg)
            except:
                pass

        tp = os.path.join(tmp_dir, os.path.basename(fp) + '.tmp.docx')
        sd.save(tp)
        temp_subs.append(tp)
    if not temp_subs:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise RuntimeError('No hay sub-documentos para procesar.')

    # ── Pre-process template: replace placeholders everywhere ──
    # The user has already placed the header content as normal body
    # content in the template.  Just replace tokens and merge.
    tpl = Document(template_path)
    replace_in_all(tpl, replacements_map)
    for sec in tpl.sections:
        force_single_column(sec)
    prepped = os.path.join(tmp_dir, 'template_prepped.docx')
    tpl.save(prepped)

    # ── Word COM: insert sub-documents at end ─────────────────
    word = None
    doc = None
    try:
        word = win32.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        time.sleep(0.3)

        doc = word.Documents.Open(
            os.path.abspath(prepped),
            ConfirmConversions=False, ReadOnly=False,
            AddToRecentFiles=False)
        time.sleep(0.2)

        # Compact margins and set single column
        sec1 = doc.Sections(1)
        sec1.PageSetup.TopMargin = CM_TO_PT * 1.2
        sec1.PageSetup.BottomMargin = CM_TO_PT * 1.2
        sec1.PageSetup.LeftMargin = CM_TO_PT * 1.2
        sec1.PageSetup.RightMargin = CM_TO_PT * 1.2
        sec1.PageSetup.HeaderDistance = CM_TO_PT * 0.5
        sec1.PageSetup.FooterDistance = CM_TO_PT * 0.8
        for s_idx in range(1, doc.Sections.Count + 1):
            try:
                doc.Sections(s_idx).PageSetup.TextColumns.SetCount(1)
            except:
                pass

        # Insert sub-documents
        for i, tp in enumerate(temp_subs):
            rng = doc.Range()
            rng.Collapse(0)
            rng.InsertFile(os.path.abspath(tp), ConfirmConversions=False)
            if i < len(temp_subs) - 1:
                rng = doc.Range()
                rng.Collapse(0)
                rng.InsertBreak(2)

        # Force single column for all sections after insertion
        for s_idx in range(1, doc.Sections.Count + 1):
            try:
                doc.Sections(s_idx).PageSetup.TextColumns.SetCount(1)
            except:
                pass

        # Safety: Find/Replace leftover placeholders
        try:
            doc.Content.Find.ClearFormatting()
            for k, v in replacements_map.items():
                doc.Content.Find.Execute(FindText=k, ReplaceWith=str(v), Replace=2)
        except:
            pass
        for i in range(1, doc.Shapes.Count + 1):
            try:
                shape = doc.Shapes.Item(i)
                if shape.Type == 6:
                    for j in range(1, shape.GroupItems.Count + 1):
                        try:
                            gi = shape.GroupItems.Item(j)
                            if gi.TextFrame.HasText:
                                gi.TextFrame.TextRange.Find.ClearFormatting()
                                for k, v in replacements_map.items():
                                    gi.TextFrame.TextRange.Find.Execute(
                                        FindText=k, ReplaceWith=str(v), Replace=2)
                        except:
                            pass
                elif shape.TextFrame.HasText:
                    shape.TextFrame.TextRange.Find.ClearFormatting()
                    for k, v in replacements_map.items():
                        shape.TextFrame.TextRange.Find.Execute(
                            FindText=k, ReplaceWith=str(v), Replace=2)
            except:
                pass

        # --- Save ---
        merged = output_path + '.word_merged.docx'
        doc.SaveAs2(os.path.abspath(merged), AddToRecentFiles=False)
        doc.Close(False); doc = None
        word.Quit(); word = None

        # ── Post-process (python-docx) ──────────────────────
        final = Document(merged)
        grade_clean = config_data.get('grade_clean', '')
        period = config_data.get('period', '')
        session_code_val = config_data.get('session_code', '')
        title_context = {
            'grade_clean': grade_clean,
            'period': period,
            'session_code': session_code_val,
            'year': config_data.get('year', ''),
            'level': config_data.get('level', ''),
        }
        title_context['grade'] = title_context['grade_clean']
        title_context['session'] = title_context['session_code']
        expanded_title = expand_template(title_template, title_context)
        final.core_properties.title = expanded_title
        final.core_properties.category = "Evaluación de Suficiencia Académica"
        final.core_properties.content_status = period

        for i, sec in enumerate(final.sections):
            # Enforce single column
            force_single_column(sec)
            # Strip all header references from every section (defensive)
            sectPr = sec._sectPr
            for ref in list(sectPr.findall(qn('w:headerReference'))):
                sectPr.remove(ref)
            titlePg = sectPr.find(qn('w:titlePg'))
            if titlePg is not None:
                sectPr.remove(titlePg)
            # Compact margins
            apply_subsequent_page_setup(sec)
            # Footer page number
            sec.footer.is_linked_to_previous = False
            setup_footer_page_number(sec.footer)

        # Apply native numbering to the fully merged document
        apply_native_lists_to_final_doc(final, start_offset=start_offset)

        if os.path.exists(output_path):
            for _ in range(5):
                try:
                    os.remove(output_path)
                    break
                except:
                    time.sleep(0.5)
        final.save(output_path)
        del final
        import gc; gc.collect()
        for _ in range(5):
            try:
                os.remove(merged)
                break
            except:
                time.sleep(0.5)
    except Exception:
        raise
    finally:
        if doc is not None:
            try: doc.Close(False)
            except: pass
        if word is not None:
            try: word.Quit()
            except: pass
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return cur - 1
