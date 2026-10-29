#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""仅清理有11列表头的交易表前、同页的多余段落；不修改列宽或图片。"""
from __future__ import annotations
import re
import sys
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HEADERS = ['No.', 'Bk.D.', 'Val.D.', 'Type', 'Vou.',
           'Vou.No./Trans.No./Details', 'DebitAmount', 'CreditAmount',
           'Balance', 'ReferenceNo.', 'Notes']
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def norm(s):
    return re.sub(r'[^a-z0-9]', '', s.lower())


def text_of_cell(tc):
    return ''.join(n.text or '' for n in tc.iter(W+'t'))


def rows_of(tbl):
    return [c for c in tbl if c.tag == W+'tr']


def cells_of(row):
    return [c for c in row if c.tag == W+'tc']


def header_match(row):
    cells = cells_of(row)
    if len(cells) != 11:
        return False
    got = [norm(text_of_cell(c)) for c in cells]
    expected = [norm(h) for h in HEADERS]
    # 前三列表头有时变形；必须严格确认剩余八列的位置
    count = sum(a == b for a, b in zip(got, expected))
    return count >= 8 and all(got[i] == expected[i] for i in range(3, 11))


def find_header_row(tbl):
    for row in rows_of(tbl)[:12]:
        if header_match(row):
            return row
    return None


def is_section_or_page_boundary(par):
    if par.tag != W+'p':
        return False
    if par.findall('.//' + W+'sectPr'):
        return False  # 分节符单独保护；不将其当作安全删除边界
    return bool(par.findall('.//' + W+'br') and any(
        x.get(W+'type') == 'page' for x in par.findall('.//' + W+'br')) or
        par.findall('.//' + W+'lastRenderedPageBreak') or
        par.findall('.//' + W+'pageBreakBefore'))


def standalone_boundary(par):
    if not is_section_or_page_boundary(par):
        return False
    # 只接受独立分页段落，避免误操作一个段落内部的分页
    if any((t.text or '').strip() for t in par.iter(W+'t')):
        return False
    if has_objects(par):
        return False
    return True


def has_objects(elem):
    critical = ('drawing', 'pict', 'object', 'fldChar', 'instrText', 'footnoteReference',
                'endnoteReference', 'bookmarkStart', 'bookmarkEnd', 'commentRangeStart',
                'commentRangeEnd', 'hyperlink', 'sdt', 'sectPr')
    return any(elem.findall('.//' + W+x) for x in critical)


def safe_to_remove(elem):
    if elem.tag != W+'p':
        return False
    if has_objects(elem):
        return False
    if any(br.get(W+'type') == 'page' for br in elem.findall('.//' + W+'br')):
        return False
    if elem.findall('.//' + W+'lastRenderedPageBreak'):
        return False
    if elem.findall('.//' + W+'pageBreakBefore'):
        return False
    return True


def set_header_auto(row):
    trPr = row.find(W+'trPr')
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        row.insert(0, trPr)
    for node in list(trPr):
        if node.tag == W+'trHeight':
            trPr.remove(node)
    # 不设置固定行高，Word/WPS 按表头文字自然确定高度
    for tc in cells_of(row):
        for p in tc.findall('.//' + W+'p'):
            pPr = p.find(W+'pPr')
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p.insert(0, pPr)
            spacing = pPr.find(W+'spacing')
            if spacing is None:
                spacing = OxmlElement('w:spacing')
                pPr.append(spacing)
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '0')
            spacing.set(qn('w:lineRule'), 'auto')
            spacing.set(qn('w:line'), '240')
            # 取消表头段落的行高锁定和分页控制，不改变字体/字号


def is_empty_par(elem):
    return elem.tag == W+'p' and not ''.join((t.text or '') for t in elem.iter(W+'t')).strip() and not has_objects(elem)


def new_blank_par():
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    s = OxmlElement('w:spacing')
    s.set(qn('w:before'), '0')
    s.set(qn('w:after'), '0')
    pPr.append(s)
    p.append(pPr)
    return p


def process(doc):
    body = doc.element.body
    report = []
    seen, cleaned, removed_total, resized = 0, 0, 0, 0
    # 收集目标元素；正文变更时重新索引
    targets = [(el, find_header_row(el)) for el in list(body) if el.tag == W+'tbl']
    targets = [(table, head) for table, head in targets if head is not None]
    for tindex, (tbl, header) in enumerate(targets, 1):
        seen += 1
        set_header_auto(header)
        resized += 1
        # 如果银行账户信息是本表顶部的合并行，也视为已有银行信息，不清除。
        earlier_rows = rows_of(tbl)[:rows_of(tbl).index(header)]
        top_text = norm(' '.join(text_of_cell(c) for r in earlier_rows for c in cells_of(r)))
        if any(k in top_text for k in ('accountno', 'accountname', 'bankname')):
            report.append(f'银行表 {tindex}: 表内顶部已有银行信息，保留此前内容；表头高度自动')
            continue
        children = list(body)
        idx = children.index(tbl)
        # 向前找本页最近的明确分页标记
        boundary_i = None
        for j in range(idx - 1, -1, -1):
            if standalone_boundary(children[j]):
                boundary_i = j
                break
            if is_section_or_page_boundary(children[j]):
                break  # 页面边界在复杂内容中，不猜测
        if boundary_i is None:
            # 文档第一张表之前，若从文档起点到此处均没有分页/表格，可确认是首页。
            prefix = children[:idx]
            if any(x.tag == W+'tbl' or is_section_or_page_boundary(x) for x in prefix):
                report.append(f'银行表 {tindex}: 表头行高已设自动；未找到安全的本页起点，跳过清理')
                continue
            preceding = prefix
            first_page = True
        else:
            preceding = children[boundary_i+1:idx]
            first_page = False
        if any(x.tag == W+'tbl' for x in preceding):
            report.append(f'银行表 {tindex}: 本页之前已有表格（可能为银行信息表），不删除')
            continue
        if any(not safe_to_remove(x) for x in preceding):
            report.append(f'银行表 {tindex}: 前方存在图片/复杂对象/分页标记，跳过删除')
            continue
        for x in preceding:
            body.remove(x)
        removed_total += len(preceding)
        # 保留边界段落作为页首的唯一空段落：已自带分页符，不能再另插空段落。
        # 若界标是 lastRenderedPageBreak，无真正分页控制，按约定在表格前保留一个空段落。
        marker = None if first_page else children[boundary_i]
        if first_page:
            body.insert(list(body).index(tbl), new_blank_par())
            left = '1个首页空白段落'
        elif marker.findall('.//' + W+'lastRenderedPageBreak') and not any(
            b.get(W+'type') == 'page' for b in marker.findall('.//' + W+'br')):
            body.insert(list(body).index(tbl), new_blank_par())
            left = '1个新空白段落（分页仅为上次渲染记录）'
        else:
            left = '1个原分页段落'
        cleaned += 1
        report.append(f'银行表 {tindex}: 删除前置普通段落 {len(preceding)} 个；保留{left}；表头行高自动')
    return seen, cleaned, removed_total, resized, report


def main():
    cwd = Path.cwd()
    candidates = sorted(p for p in cwd.glob('*.docx') if not p.name.startswith('~$') and '_页前清理' not in p.stem)
    if not candidates:
        print('当前目录没有可处理的 DOCX 文件：', cwd)
        return 1
    print('当前目录：', cwd)
    for i, p in enumerate(candidates, 1):
        print(f'  {i}. {p.name}')
    choice = input('输入文件编号（q退出）：').strip()
    if choice.lower() == 'q':
        return 0
    if not choice.isdigit() or not 1 <= int(choice) <= len(candidates):
        print('编号无效。')
        return 1
    source = candidates[int(choice)-1]
    output = cwd / (source.stem + '_页前清理.docx')
    if output.exists():
        print('目标已存在，为避免覆盖，未处理：', output.name)
        return 1
    document = Document(str(source))
    seen, cleaned, deleted, resized, lines = process(document)
    if seen == 0:
        print('没有检测到符合11列标准表头的表格，未输出。')
        return 1
    document.save(str(output))
    log_path = output.with_suffix('.log.txt')
    log_path.write_text('仅处理11列表头前方内容和表头行高\n' + '\n'.join(lines) +
                        f'\n目标表格:{seen};清理表格:{cleaned};删除段落:{deleted};表头自动行高:{resized}\n', encoding='utf-8')
    print('\n输出：', output.name)
    print('日志：', log_path.name)
    print(f'识别 {seen} 张表，清理 {cleaned} 张表前内容，共删除 {deleted} 段；表头自动高度 {resized} 张。')
    if cleaned < seen:
        print('注意：部分页面无法确认安全页首或存在其他表格，已保留原内容。查看日志。')
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except ImportError:
        print('请在当前 Python 环境安装：python3 -m pip install python-docx')
        raise