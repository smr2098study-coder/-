"""Audit native Word citation structure for shimen citation modes A and B."""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'w': W, 'r': R, 'pr': PR, 'ct': CT}


def qn(name):
    return '{' + W + '}' + name


def load_package(path):
    with ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}


def xml(parts, name):
    return E.fromstring(parts[name]) if name in parts else None


def normal_notes(root, kind):
    if root is None:
        return []
    return [n for n in root.findall('w:' + kind, NS) if n.get(qn('type')) is None]


def text_of(node):
    return ''.join(node.xpath('.//w:t/text()', namespaces=NS))


def complex_fields(root, part):
    fields, stack = [], []
    if root is None:
        return fields, []
    for node in root.iter():
        local = E.QName(node).localname
        if local == 'fldChar':
            kind = node.get(qn('fldCharType'))
            if kind == 'begin':
                stack.append({'part': part, 'begin': True, 'separate': False, 'end': False,
                              'instruction': '', 'result': ''})
            elif kind == 'separate' and stack:
                stack[-1]['separate'] = True
            elif kind == 'end' and stack:
                field = stack.pop()
                field['end'] = True
                fields.append(field)
        elif local == 'instrText' and stack:
            stack[-1]['instruction'] += node.text or ''
        elif local == 't' and stack and stack[-1]['separate']:
            stack[-1]['result'] += node.text or ''
    return fields, stack


def bookmark_map(root, part):
    starts, ends = {}, {}
    if root is None:
        return starts
    nodes = list(root.iter())
    for position, node in enumerate(nodes):
        local = E.QName(node).localname
        if local == 'bookmarkStart':
            starts[node.get(qn('name'))] = {'id': node.get(qn('id')), 'start': position,
                                            'part': part, 'target_has_endnote_reference': False}
        elif local == 'bookmarkEnd':
            ends[node.get(qn('id'))] = position
    for value in starts.values():
        end = ends.get(value['id'])
        value['end'] = end
        if end is not None:
            value['target_has_endnote_reference'] = any(
                E.QName(node).localname == 'endnoteReference'
                for node in nodes[value['start'] + 1:end]
            )
    return starts


def parse_noteref(instruction):
    normalized = re.sub(r'\s+', ' ', instruction).strip()
    match = re.search(r'\bNOTEREF\s+"?([^\s"\\]+)"?', normalized, re.I)
    return (normalized, match.group(1) if match else None,
            bool(re.search(r'(?:^|\s)\\h(?:\s|$)', normalized, re.I)))


def run_has_vert_align(run, value):
    align = run.find('w:rPr/w:vertAlign', NS) if run is not None else None
    return align is not None and align.get(qn('val')) == value


def citation_like_hyperlinks(root):
    if root is None:
        return []
    found = []
    for link in root.findall('.//w:hyperlink', NS):
        label = text_of(link).strip()
        if link.get(qn('anchor')) and re.fullmatch(r'\[\s*\d+(?:\s*[-,–]\s*\d+)*\s*\]', label):
            found.append({'anchor': link.get(qn('anchor')), 'text': label})
    return found


def static_citation_texts(root):
    if root is None:
        return []
    found = []
    for run in root.findall('.//w:r', NS):
        label = text_of(run).strip()
        if (run_has_vert_align(run, 'superscript') and
                re.fullmatch(r'\[\s*\d+(?:\s*[-,–]\s*\d+)*\s*\]', label) and
                not run.findall('.//w:fldChar', NS) and not run.findall('.//w:endnoteReference', NS)):
            found.append(label)
    return found


def content_type_and_relationship_ok(parts, kind):
    content = xml(parts, '[Content_Types].xml')
    rels = xml(parts, 'word/_rels/document.xml.rels')
    part_name = '/word/' + kind + '.xml'
    override = bool(content is not None and content.xpath(
        './ct:Override[@PartName=$part]', namespaces=NS, part=part_name))
    relationship = False
    if rels is not None:
        relationship = any(
            rel.get('Type', '').endswith('/' + kind) and rel.get('Target', '').endswith(kind + '.xml')
            for rel in rels.findall('pr:Relationship', NS)
        )
    return override, relationship


def endnote_settings(parts):
    settings = xml(parts, 'word/settings.xml')
    document = xml(parts, 'word/document.xml')
    properties = []
    for root in (settings, document):
        if root is not None:
            properties.extend(root.findall('.//w:endnotePr', NS))
    values = []
    for node in properties:
        values.append({
            'pos': node.xpath('string(w:pos/@w:val)', namespaces=NS) or None,
            'numFmt': node.xpath('string(w:numFmt/@w:val)', namespaces=NS) or None,
            'numStart': node.xpath('string(w:numStart/@w:val)', namespaces=NS) or None,
            'numRestart': node.xpath('string(w:numRestart/@w:val)', namespaces=NS) or None,
        })
    return values


def audit(path, mode='B', word_validation=None):
    parts = load_package(path)
    document = xml(parts, 'word/document.xml')
    footnotes = xml(parts, 'word/footnotes.xml')
    endnotes = xml(parts, 'word/endnotes.xml')
    errors, warnings = [], []
    counts = Counter()

    word_roots = []
    for part, data in parts.items():
        if part.startswith('word/') and part.endswith('.xml'):
            try:
                word_roots.append((part, E.fromstring(data)))
            except E.XMLSyntaxError:
                continue

    fields, open_fields = [], []
    bookmarks = {}
    for part, root in word_roots:
        part_fields, unclosed = complex_fields(root, part)
        fields.extend(part_fields)
        open_fields.extend(unclosed)
        bookmarks.update(bookmark_map(root, part))

    simple_noteref = []
    for part, root in word_roots:
        for field in root.findall('.//w:fldSimple', NS):
            instruction = field.get(qn('instr'), '')
            if re.search(r'\bNOTEREF\b', instruction, re.I):
                simple_noteref.append({'part': part, 'instruction': instruction})

    noterefs = []
    for field in fields:
        normalized, target, hyperlink = parse_noteref(field['instruction'])
        if target is None:
            continue
        item = dict(field)
        item.update({'instruction': normalized, 'target': target, 'hyperlink_switch': hyperlink})
        target_data = bookmarks.get(target)
        item['bookmark_exists'] = target_data is not None
        item['bookmark_complete'] = bool(target_data and target_data.get('end') is not None)
        item['bookmark_targets_endnote'] = bool(target_data and target_data['target_has_endnote_reference'])
        item['stable_bookmark'] = not bool(re.fullmatch(r'_?ref_?\d+', target, re.I))
        item['valid_structure'] = all([
            item['begin'], item['separate'], item['end'], bool(item['result'].strip()),
            item['bookmark_complete'], item['bookmark_targets_endnote'], item['stable_bookmark']
        ])
        noterefs.append(item)

    ordinary_footnotes = normal_notes(footnotes, 'footnote')
    ordinary_endnotes = normal_notes(endnotes, 'endnote')
    footnote_refs = document.findall('.//w:footnoteReference', NS) if document is not None else []
    endnote_refs = document.findall('.//w:endnoteReference', NS) if document is not None else []
    endnote_ids = [n.get(qn('id')) for n in ordinary_endnotes]
    endnote_ref_ids = [n.get(qn('id')) for n in endnote_refs]
    counts.update({
        'true_footnote_reference_count': len(footnote_refs),
        'true_endnote_reference_count': len(endnote_refs),
        'ordinary_footnote_count': len(ordinary_footnotes),
        'ordinary_endnote_count': len(ordinary_endnotes),
        'noteref_field_count': len(noterefs),
        'simple_noteref_count': len(simple_noteref),
        'citation_hyperlink_count': len(citation_like_hyperlinks(document)),
        'static_superscript_citation_count': len(static_citation_texts(document)),
        'broken_noteref_count': sum(not item['valid_structure'] for item in noterefs),
        'unclosed_field_count': len(open_fields),
    })

    if mode.upper() == 'A':
        if footnotes is None:
            errors.append('Mode A 缺少 word/footnotes.xml')
        if len(footnote_refs) != len(ordinary_footnotes):
            errors.append('Mode A 正文脚注引用数与普通脚注数不一致')
        if len(set(n.get(qn('id')) for n in footnote_refs)) != len(footnote_refs):
            errors.append('Mode A 存在共享脚注ID，未做到每次引用独立脚注')
        if ordinary_endnotes or noterefs:
            errors.append('Mode A 引用部分仍有Endnote或NOTEREF')
        dynamic_structure = not errors
        navigation = dynamic_structure
    else:
        override, relationship = content_type_and_relationship_ok(parts, 'endnotes')
        if endnotes is None:
            errors.append('Mode B 缺少 word/endnotes.xml')
        if not override or not relationship:
            errors.append('Mode B 的endnotes内容类型或文档关系不完整')
        if footnote_refs or ordinary_footnotes:
            errors.append('Mode B 仍有普通Footnote或正文footnoteReference')
        if not ordinary_endnotes:
            errors.append('Mode B 没有普通真实Endnote')
        if set(endnote_ids) != set(endnote_ref_ids) or len(endnote_ids) != len(endnote_ref_ids):
            errors.append('普通Endnote与正文首次真实endnoteReference不是一一对应')
        if simple_noteref:
            errors.append('NOTEREF必须使用含begin/separate/result/end的真实复合域')
        if open_fields:
            errors.append('存在未闭合的Word复合域')
        if counts['broken_noteref_count']:
            errors.append('存在broken或结构不完整的NOTEREF')
        if counts['citation_hyperlink_count']:
            errors.append('发现静态引用文字加internal hyperlink的退化实现')
        if counts['static_superscript_citation_count']:
            errors.append('发现不含原生引用或域的静态上标引用')
        if any(not item['hyperlink_switch'] for item in noterefs):
            errors.append('NOTEREF缺少\\h，无法同时保证原生点击跳转')
        settings = endnote_settings(parts)
        counts['endnote_settings_count'] = len(settings)
        if not any(item['pos'] == 'sectEnd' for item in settings):
            errors.append('Endnote未明确设置在节末')
        if not any(item['numFmt'] == 'decimal' for item in settings):
            errors.append('Endnote未明确使用阿拉伯数字')
        if not any(item['numRestart'] == 'continuous' for item in settings):
            errors.append('Endnote未明确设置连续编号')
        special_types = Counter(n.get(qn('type')) for n in endnotes.findall('w:endnote', NS)) if endnotes is not None else Counter()
        if not special_types['separator'] or not special_types['continuationSeparator']:
            errors.append('Endnote缺少合法separator或continuationSeparator特殊节点')
        for note in ordinary_endnotes:
            refs = note.findall('.//w:endnoteRef', NS)
            tabs = note.findall('.//w:tab', NS)
            if len(refs) != 1 or not tabs:
                errors.append('普通Endnote必须含一个动态endnoteRef及其后的真实Tab')
                continue
            ref_run = refs[0].getparent()
            if not run_has_vert_align(ref_run, 'baseline'):
                errors.append('文末动态endnoteRef未显式设置为baseline')
            paragraph = refs[0].getparent().getparent()
            indent = paragraph.find('w:pPr/w:ind', NS) if paragraph is not None else None
            if indent is None or not (indent.get(qn('hanging')) or indent.get(qn('hangingChars'))):
                errors.append('Endnote参考文献段落缺少悬挂缩进')
        navigation = (not counts['citation_hyperlink_count'] and
                      all(item['hyperlink_switch'] and item['bookmark_targets_endnote'] for item in noterefs) and
                      set(endnote_ids) == set(endnote_ref_ids))
        dynamic_structure = (endnotes is not None and bool(ordinary_endnotes) and
                             not counts['simple_noteref_count'] and not counts['broken_noteref_count'] and
                             not counts['static_superscript_citation_count'] and
                             set(endnote_ids) == set(endnote_ref_ids))

    validation_status = 'not_tested'
    dynamic_regression = False
    dynamic_delete = False
    save_reopen = False
    if word_validation:
        validation_status = word_validation.get('word_field_update_validation', 'not_tested')
        dynamic_regression = word_validation.get('dynamic_insert_test_passed') is True
        dynamic_delete = word_validation.get('dynamic_delete_test_passed') is True
        save_reopen = word_validation.get('save_close_reopen_passed') is True
        if word_validation.get('broken_noteref_count', 0) != 0:
            errors.append('Word原生验收报告仍有broken NOTEREF')
    dynamic_field_ok = (dynamic_structure and validation_status == 'passed' and
                        dynamic_regression and dynamic_delete and save_reopen)
    status = 'passed' if not errors and (mode.upper() == 'A' or dynamic_field_ok) else (
        'pending_word_validation' if not errors and mode.upper() == 'B' else 'failed')
    if mode.upper() == 'B' and validation_status != 'passed':
        warnings.append('word_field_update_validation = not_tested；不能宣告Mode B成功')

    return {
        'mode': mode.upper(),
        'status': status,
        'navigation_ok': bool(navigation),
        'dynamic_field_structure_ok': bool(dynamic_structure),
        'dynamic_field_ok': bool(dynamic_field_ok) if validation_status != 'not_tested' else None,
        'word_field_update_validation': validation_status,
        'dynamic_insert_test_passed': dynamic_regression,
        'dynamic_delete_test_passed': dynamic_delete,
        'save_close_reopen_passed': save_reopen,
        'counts': dict(counts),
        'noteref_fields': noterefs,
        'errors': errors,
        'warnings': warnings,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('--mode', choices=('A', 'B'), required=True)
    parser.add_argument('--out', required=True)
    parser.add_argument('--word-validation-report')
    parser.add_argument('--static-only', action='store_true')
    args = parser.parse_args()
    validation = None
    if args.word_validation_report:
        validation = json.loads(Path(args.word_validation_report).read_text(encoding='utf-8'))
    result = audit(args.input, args.mode, validation)
    Path(args.out).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: v for k, v in result.items() if k != 'noteref_fields'}, ensure_ascii=False))
    acceptable = result['status'] == 'passed' or (args.static_only and result['status'] == 'pending_word_validation')
    raise SystemExit(0 if acceptable else 1)


if __name__ == '__main__':
    main()
