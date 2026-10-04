"""Build and audit DOCX with the user's Chinese academic formatting rules."""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W, 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
PROFILES = {
    'body': ('宋体', 12, False, 200, 360),
    'title': ('宋体', 16, True, 0, 360),
    'h1': ('宋体', 16, True, 0, 360),
    'h2': ('宋体', 15, True, 0, 360),
    'h3': ('宋体', 14, True, 200, 360),
    'h4': ('宋体', 12, True, 200, 360),
    'h5': ('宋体', 12, False, 200, 360),
    'table': ('黑体', 10.5, False, 0, 276),
    'caption': ('黑体', 10.5, False, 0, 276),
}
PATTERNS = [r'^[一二三四五六七八九十百零]+、', r'^（[一二三四五六七八九十百零]+）',
            r'^\d+、', r'^（\d+）', r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]']
FLAGS = ('keepNext', 'keepLines', 'pageBreakBefore')
P_ORDER = ('pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers '
           'pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE '
           'autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents '
           'suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle '
           'rPr sectPr pPrChange').split()
R_ORDER = ('rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint '
           'noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u '
           'effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange').split()


def canonical_properties(root):
    """Keep core OOXML property order; preserve unknown extension order at the end."""
    for node, order in (('pPr', P_ORDER), ('rPr', R_ORDER)):
        properties = ([root] if E.QName(root).localname == node else []) + root.findall('.//w:' + node, NS)
        ranks = {tag(name): i for i, name in enumerate(order)}
        for prop in properties:
            prop[:] = sorted(prop, key=lambda child: ranks.get(child.tag, len(ranks)))


def tag(name):
    return '{' + W + '}' + name


def sub(parent, name, **attrs):
    el = parent.find(tag(name))
    if el is None:
        el = E.SubElement(parent, tag(name))
    for key, value in attrs.items():
        el.set(tag(key), str(value))
    return el


def enabled(value):
    return value is not None and value not in ('0', 'false', 'off')


def no_overwrite(source, target):
    target = Path(target)
    if source is not None and Path(source).resolve() == target.resolve():
        raise ValueError('输入输出不能是同一文件')
    if target.exists():
        raise FileExistsError('输出已存在，请使用新路径: ' + str(target))
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


class Package:
    def __init__(self, path):
        with ZipFile(path) as z:
            self.parts = {name: z.read(name) for name in z.namelist()}
        self.doc = E.fromstring(self.parts['word/document.xml'])
        self.styles = E.fromstring(self.parts['word/styles.xml'])
        self.by_id = {s.get(tag('styleId')): s for s in self.styles.findall('w:style', NS)}
        self.default = next((sid for sid, s in self.by_id.items()
                             if s.get(tag('default')) == '1' and s.get(tag('type')) == 'paragraph'), 'Normal')
        self.defaults = self.styles.find('w:docDefaults', NS)
        self.numbering = E.fromstring(self.parts['word/numbering.xml']) if 'word/numbering.xml' in self.parts else None
        self.theme = E.fromstring(self.parts['word/theme/theme1.xml']) if 'word/theme/theme1.xml' in self.parts else None
        self.paragraphs = self.doc.xpath('//w:body//w:p', namespaces=NS)

    def chain(self, sid):
        chain, seen = [], set()
        while sid in self.by_id and sid not in seen:
            seen.add(sid)
            s = self.by_id[sid]
            chain.insert(0, s)
            base = s.find('w:basedOn', NS)
            sid = base.get(tag('val')) if base is not None else None
        return chain

    def style_id(self, p):
        s = p.find('w:pPr/w:pStyle', NS)
        return s.get(tag('val')) if s is not None else self.default

    def property_layers(self, p, run=None):
        kind = 'rPr' if run is not None else 'pPr'
        layers = []
        if self.defaults is not None:
            layers.append(self.defaults.find('w:' + kind + 'Default/w:' + kind, NS))
        layers.extend(s.find('w:' + kind, NS) for s in self.chain(self.style_id(p)))
        if run is None:
            layers.append(p.find('w:pPr', NS))
        else:
            rs = run.find('w:rPr/w:rStyle', NS)
            if rs is not None:
                layers.extend(s.find('w:rPr', NS) for s in self.chain(rs.get(tag('val'))))
            layers.append(run.find('w:rPr', NS))
        return [x for x in layers if x is not None]

    def value(self, p, name, attr='val', run=None):
        value = None
        for layer in self.property_layers(p, run):
            el = layer.find('w:' + name, NS)
            if el is not None:
                if el.get(tag(attr)) is not None:
                    value = el.get(tag(attr))
                elif attr == 'val':
                    value = '1'
        return value

    def bold(self, p, run):
        # b is an OOXML toggle in styles, but absolute in direct run formatting.
        value = False
        for layer in self.property_layers(p, run):
            el = layer.find('w:b', NS)
            if el is None:
                continue
            state = enabled(el.get(tag('val'), '1'))
            if layer.getparent() is not None and E.QName(layer.getparent()).localname == 'r':
                value = state
            elif state:
                value = not value
        return value

    def font(self, p, run, channel):
        # At each cascade level theme font takes precedence over literal font.
        theme_attr = {'eastAsia': 'eastAsiaTheme', 'ascii': 'asciiTheme', 'hAnsi': 'hAnsiTheme'}[channel]
        choice = None
        for layer in self.property_layers(p, run):
            fonts = layer.find('w:rFonts', NS)
            if fonts is None:
                continue
            if fonts.get(tag(theme_attr)) is not None:
                choice = ('theme', fonts.get(tag(theme_attr)))
            elif fonts.get(tag(channel)) is not None:
                choice = ('literal', fonts.get(tag(channel)))
        if choice is None:
            return None
        if choice[0] == 'literal':
            return choice[1]
        if self.theme is None:
            return 'unresolved-theme:' + choice[1]
        group = 'majorFont' if choice[1].startswith('major') else 'minorFont'
        base = self.theme.find('.//a:fontScheme/a:' + group, NS)
        if base is None:
            return 'unresolved-theme:' + choice[1]
        node = base.find('a:ea' if channel == 'eastAsia' else 'a:latin', NS)
        if node is not None and node.get('typeface'):
            return node.get('typeface')
        if channel == 'eastAsia':
            node = next((x for x in base.findall('a:font', NS) if x.get('script') == 'Hans'), None)
            if node is not None:
                return node.get('typeface')
        return 'unresolved-theme:' + choice[1]

    def number_format(self, p):
        num_id = self.value(p, 'numPr/w:numId')
        if num_id is None or num_id == '0':
            return None
        if self.numbering is None:
            return 'unresolved'
        level = self.value(p, 'numPr/w:ilvl') or '0'
        num = next((x for x in self.numbering.findall('w:num', NS) if x.get(tag('numId')) == num_id), None)
        if num is None:
            return 'unresolved'
        override = next((x for x in num.findall('w:lvlOverride', NS) if x.get(tag('ilvl')) == level), None)
        lv = override.find('w:lvl', NS) if override is not None else None
        if lv is None:
            aid = num.find('w:abstractNumId', NS)
            abstract = next((x for x in self.numbering.findall('w:abstractNum', NS)
                             if aid is not None and x.get(tag('abstractNumId')) == aid.get(tag('val'))), None)
            if abstract is not None:
                lv = next((x for x in abstract.findall('w:lvl', NS) if x.get(tag('ilvl')) == level), None)
        fmt = lv.find('w:numFmt', NS) if lv is not None else None
        return fmt.get(tag('val')) if fmt is not None else 'unresolved'

    def role(self, p):
        text = ''.join(p.xpath('.//w:t/text()', namespaces=NS)).strip()
        sid = self.style_id(p)
        s = self.by_id.get(sid)
        name_el = s.find('w:name', NS) if s is not None else None
        name = name_el.get(tag('val'), '') if name_el is not None else ''
        lower = name.lower()
        if p.xpath('ancestor::w:tc', namespaces=NS):
            return 'table'
        if re.match(r'toc|目录', lower):
            return 'toc'
        if lower == 'title' or name == '标题':
            return 'title'
        if re.match(r'^[表图]\s*\d', text) or 'caption' in lower or name == '题注':
            return 'caption'
        if p.xpath('.//w:drawing|.//w:pict', namespaces=NS):
            return 'image'
        if not text:
            return 'empty'
        match = re.match(r'heading\s*([1-5])|标题\s*([1-5])', lower)
        if match:
            return 'h' + (match.group(1) or match.group(2))
        for level, pattern in enumerate(PATTERNS, 1):
            if len(text) <= 90 and re.match(pattern, text):
                return 'candidate-h' + str(level)
        if re.match(r'^第.+章', text) or '【写作方向】' in text or '【建议字数】' in text:
            return 'special'
        if self.value(p, 'jc') == 'center' and len(text) < 80:
            return 'special'
        return 'body'

    def save(self, target):
        self.parts['word/document.xml'] = E.tostring(self.doc, encoding='UTF-8', xml_declaration=True, standalone=True)
        with ZipFile(target, 'w', ZIP_DEFLATED) as z:
            for name, value in self.parts.items():
                z.writestr(name, value)


def audit(path):
    pkg = Package(path)
    issues, stats = [], Counter()
    def add(index, code, message, severity='error'):
        issues.append({'paragraph': index, 'code': code, 'severity': severity, 'message': message})
    for index, p in enumerate(pkg.paragraphs):
        text = ''.join(p.xpath('.//w:t/text()', namespaces=NS)).strip()
        role = pkg.role(p)
        stats[role] += 1
        for flag in FLAGS:
            if enabled(pkg.value(p, flag)):
                stats['effective_' + flag] += 1
                if role == 'body':
                    add(index, 'body_pagination', flag + '作用于正文候选，需确认是否是未建样式的小标题', 'warning')
        numfmt = pkg.number_format(p)
        if numfmt:
            stats['numbered_paragraphs'] += 1
            add(index, 'list', '列表格式为' + numfmt + '，确认是否是用户要求的列表', 'warning')
        if text.startswith(('•', '●', '▪', '■', '·')):
            add(index, 'literal_bullet', '段首实际黑点字符，需判断是否误加', 'warning')
        if re.match(r'^#{1,6}\s', text) or '**' in text:
            add(index, 'markdown', '疑似Markdown残留', 'warning')
        if '【写作方向】' in text or '【建议字数】' in text:
            add(index, 'draft_note', '有写作指导标签，正式成稿需整理', 'warning')
        if role.startswith('candidate-h'):
            add(index, 'heading_style', '疑似标题尚未使用对应Heading样式', 'warning')
            role = role.removeprefix('candidate-')
        if role not in PROFILES or not text:
            if role == 'special':
                add(index, 'role_review', '封面、章节标题或指导段落需人工确认角色', 'warning')
            continue
        font, size, bold, indent, line = PROFILES[role]
        if role.startswith('h'):
            level = int(role[1])
            if not re.match(PATTERNS[level - 1], text) and numfmt is None:
                add(index, 'heading_number', '编号与师门标题级别不匹配；既有章节体系需人工映射', 'warning')
            if pkg.value(p, 'outlineLvl') != str(level - 1):
                add(index, 'outline_level', '标题缺少对应大纲级别')
        firstchars = pkg.value(p, 'ind', 'firstLineChars')
        twip = pkg.value(p, 'ind', 'firstLine') or '0'
        if firstchars is not None:
            ok = firstchars == str(indent)
        else:
            ok = twip == str(round(size * 40) if indent else 0)
        if not ok:
            add(index, 'indent', '首行缩进与角色要求不符')
        if enabled(pkg.value(p, 'ind', 'hanging')) or enabled(pkg.value(p, 'ind', 'hangingChars')):
            add(index, 'hanging_indent', '存在冲突悬挂缩进')
        # Body/table line spacing is source-mandated; heading/caption spacing is a default.
        if role in ('body', 'table'):
            if pkg.value(p, 'spacing', 'line') != str(line) or pkg.value(p, 'spacing', 'lineRule') not in (None, 'auto'):
                add(index, 'line_spacing', '行距不是规定的倍数行距')
        if enabled(pkg.value(p, 'snapToGrid')):
            add(index, 'grid', '段落与网格对齐可能覆盖倍数行距', 'warning')
        seen = set()
        for run in p.xpath('.//w:r', namespaces=NS):
            rt = ''.join(run.xpath('.//w:t/text()', namespaces=NS))
            if not rt.strip():
                continue
            checks = []
            actual_size = pkg.value(p, 'sz', run=run)
            if actual_size != str(round(size * 2)):
                checks.append(('font_size', '部分文字字号不符'))
            if re.search(r'[\u3400-\u9fff]', rt) and pkg.font(p, run, 'eastAsia') != font:
                checks.append(('chinese_font', '部分中文字体不符或无法解析'))
            if re.search(r'[A-Za-z0-9]', rt) and pkg.font(p, run, 'ascii') != 'Times New Roman':
                checks.append(('latin_font', '部分英文数字字体不符或无法解析'))
            if role.startswith('h') or role == 'title':
                if pkg.bold(p, run) != bold:
                    checks.append(('bold', '标题加粗设置不符'))
            for code, message in checks:
                if code not in seen:
                    seen.add(code)
                    add(index, code, message)
    for i, tbl in enumerate(pkg.doc.findall('.//w:tbl', NS)):
        jc = tbl.find('w:tblPr/w:jc', NS)
        if jc is None or jc.get(tag('val')) != 'center':
            issues.append({'table': i, 'code': 'table_alignment', 'severity': 'error', 'message': '表格未显式居中，需排除表样式继承后确认'})
        for h in tbl.findall('.//w:trHeight', NS):
            if h.get(tag('hRule')) == 'exact':
                issues.append({'table': i, 'code': 'fixed_height', 'severity': 'warning', 'message': '固定行高可能截断内容'})
    return {'scope': '主文档结构审计，未做视觉验收；自动角色必须复核', 'paragraphs': len(pkg.paragraphs),
            'tables': len(pkg.doc.findall('.//w:tbl', NS)), 'stats': dict(stats),
            'errors': sum(x['severity'] == 'error' for x in issues),
            'warnings': sum(x['severity'] == 'warning' for x in issues), 'issues': issues}


def style_profile(style, role):
    from docx.shared import Pt, RGBColor
    font, size, bold, indent, line = PROFILES[role]
    style.font.name = 'Times New Roman'
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    fonts = sub(rpr, 'rFonts', ascii='Times New Roman', hAnsi='Times New Roman', eastAsia=font, cs='Times New Roman')
    for attr in list(fonts.attrib):
        if 'theme' in E.QName(attr).localname.lower():
            del fonts.attrib[attr]
    sub(rpr, 'szCs', val=round(size * 2))
    sub(rpr, 'bCs', val='1' if bold else '0')
    ppr = style.element.get_or_add_pPr()
    ind = sub(ppr, 'ind')
    ind.attrib.clear()
    ind.set(tag('firstLineChars'), str(indent))
    ind.set(tag('firstLine'), str(round(size * 40) if indent else 0))
    sub(ppr, 'spacing', line=line, lineRule='auto', before='120' if role.startswith('h') else '0', after='0')
    sub(ppr, 'snapToGrid', val='0')
    sub(ppr, 'widowControl', val='1')
    sub(ppr, 'keepNext', val='1' if role.startswith('h') else '0')
    sub(ppr, 'keepLines', val='1' if role.startswith('h') or role == 'title' else '0')
    sub(ppr, 'pageBreakBefore', val='0')
    sub(ppr, 'jc', val='center' if role in ('title', 'caption') else 'both' if role == 'body' else 'left')
    if role.startswith('h'):
        sub(ppr, 'outlineLvl', val=int(role[1]) - 1)


def build(source, target):
    from docx import Document
    from docx.shared import Cm
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    source = Path(source)
    data = json.loads(source.read_text(encoding='utf-8'))
    doc = Document()
    for section in doc.sections:
        section.page_width, section.page_height = Cm(21), Cm(29.7)
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    names = {'body': 'Normal', 'title': 'Title', **{'h' + str(i): 'Heading ' + str(i) for i in range(1, 6)},
             'table': 'Shimen Table Text', 'caption': 'Caption'}
    for role, name in names.items():
        style = doc.styles[name] if name in doc.styles else doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if role != 'body':
            style.base_style = doc.styles['Normal']
        style_profile(style, role)
        style.next_paragraph_style = doc.styles['Normal']
    if data.get('title'):
        doc.add_paragraph(data['title'], 'Title')
    for block in data.get('blocks', []):
        kind = block['type']
        if kind == 'heading':
            level = block['level']
            if not isinstance(level, int) or not 1 <= level <= 5:
                raise ValueError('标题级别必须在1到5之间')
            doc.add_paragraph(block['text'], 'Heading ' + str(level))
        elif kind == 'paragraph':
            doc.add_paragraph(block['text'], 'Normal')
        elif kind == 'table':
            rows = block['rows']
            if not rows or not rows[0] or any(len(r) != len(rows[0]) for r in rows):
                raise ValueError('表格行不能为空且列数必须一致')
            if not block.get('caption'):
                raise ValueError('表格必须提供caption')
            caption = doc.add_paragraph(block['caption'], 'Caption')
            caption.paragraph_format.keep_with_next = True
            table = doc.add_table(rows=len(rows), cols=len(rows[0]))
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for row, values in zip(table.rows, rows):
                for cell, value in zip(row.cells, values):
                    cell.text = str(value)
                    cell.paragraphs[0].style = doc.styles['Shimen Table Text']
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        elif kind == 'image':
            if not block.get('caption'):
                raise ValueError('图片必须提供caption')
            width = block.get('width_cm', 12)
            if not 0 < width <= 15.92:
                raise ValueError('图片宽度必须在0到15.92cm之间')
            p = doc.add_paragraph()
            p.paragraph_format.alignment = 1
            p.paragraph_format.first_line_indent = Cm(0)
            sub(p._p.get_or_add_pPr(), 'ind', firstLineChars=0, firstLine=0)
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(str(source.parent / block['path']), width=Cm(width))
            doc.add_paragraph(block['caption'], 'Caption')
        else:
            raise ValueError('不支持的块类型: ' + kind)
    target = no_overwrite(source, target)
    canonical_properties(doc.element)
    canonical_properties(doc.styles.element)
    doc.save(target)


def repair(source, target, indices_path=None):
    pkg = Package(source)
    target = no_overwrite(source, target)
    if indices_path:
        indices = json.loads(Path(indices_path).read_text(encoding='utf-8'))
        if not isinstance(indices, list) or any(not isinstance(x, int) or x < 0 or x >= len(pkg.paragraphs) for x in indices):
            raise ValueError('paragraphs必须是有效段落索引数组')
    else:
        indices = [i for i, p in enumerate(pkg.paragraphs) if pkg.role(p) == 'body'
                   and not p.xpath('.//w:fldChar|.//w:instrText|.//w:hyperlink|.//w:drawing|.//w:pict', namespaces=NS)]
    for i in indices:
        p = pkg.paragraphs[i]
        if pkg.role(p) in ('table', 'toc', 'image', 'empty', 'title', 'caption') or pkg.role(p).startswith('h'):
            raise ValueError('索引指向受保护角色: ' + str(i))
        pp = p.find('w:pPr', NS)
        if pp is None:
            pp = E.Element(tag('pPr'))
            p.insert(0, pp)
        for flag in FLAGS:
            # Write false rather than delete, so style inheritance cannot restore true.
            sub(pp, flag, val='0')
        canonical_properties(pp)
    pkg.save(target)
    return {'candidate_paragraphs_processed': len(indices), 'scope': '仅正文分页标志，未统一其他格式'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    for command in ('build', 'audit', 'repair-body-flags'):
        p = commands.add_parser(command)
        p.add_argument('input')
        p.add_argument('--out', required=True)
        if command == 'repair-body-flags':
            p.add_argument('--paragraphs')
    args = parser.parse_args()
    if args.command == 'build':
        build(args.input, args.out)
        print('Created:', args.out)
    elif args.command == 'audit':
        result = audit(args.input)
        target = no_overwrite(args.input, args.out)
        target.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({k: v for k, v in result.items() if k != 'issues'}, ensure_ascii=False))
        raise SystemExit(1 if result['errors'] else 0)
    else:
        print(json.dumps(repair(args.input, args.out, args.paragraphs), ensure_ascii=False))


if __name__ == '__main__':
    main()
