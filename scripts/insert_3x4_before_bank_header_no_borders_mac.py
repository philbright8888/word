#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""仅给缺少表头上方信息行的标准11列银行表新增无边框的3行×4个逻辑单元格。
不删除内容，不改变现有表格网格/列宽、图片或交易数据。
运行：python3 insert_3x4_before_bank_header_no_borders_mac.py
依赖：python3 -m pip install python-docx
"""
from __future__ import annotations
import re
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

LABELS = ["No.", "Bk.D.", "Val.D.", "Type", "Vou.", "Vou.No./Trans.No./Details",
          "DebitAmount", "CreditAmount", "Balance", "ReferenceNo.", "Notes"]


def compact(s):
    return re.sub(r'\s+', '', s or '').lower()


def is_header(cells):
    if len(cells) != 11:
        return False
    texts = [compact(c.text) for c in cells]
    # Support both canonical and PDF conversion duplicates while preventing ordinary data rows.
    indicators = (
        ('serialno', 'no.'), ('bookeddate', 'bk.d.'), ('valuedate', 'val.d.'),
        ('transactiontype', 'type'), ('debitvou', 'vou.'),
        ('trans.no', 'details'), ('debitamount',), ('creditamount',),
        ('balance',), ('referenceno', 'institution/cashier'), ('notes',)
    )
    matches = sum(any(token in texts[i] for token in choices) for i, choices in enumerate(indicators))
    return matches >= 8 and all(k in texts[i] for i,k in [(6,'debit'),(7,'credit'),(8,'balance'),(10,'notes')])


def preceding_table_immediately(tbl):
    # A bank-information table immediately preceding this table (possibly separated by blanks)
    # means a new 3x3 area is not required. Any nonempty paragraph breaks adjacency.
    p = tbl.getprevious()
    while p is not None:
        if p.tag == qn('w:tbl'):
            return True
        if p.tag == qn('w:p'):
            if ''.join(t.text or '' for t in p.iter(qn('w:t'))).strip():
                return False
            if any(True for _ in p.iter(qn('w:drawing'))) or any(True for _ in p.iter(qn('w:br'))):
                return False
        p = p.getprevious()
    return False


def grid_widths(tbl):
    grid = tbl.find(qn('w:tblGrid'))
    if grid is None:
        return None
    widths = []
    for c in grid.findall(qn('w:gridCol')):
        v = c.get(qn('w:w'))
        if not v:
            return None
        widths.append(int(v))
    return widths


def header_grid_ranges(tr, count):
    trpr = tr.find(qn('w:trPr'))
    before = 0
    if trpr is not None:
        x = trpr.find(qn('w:gridBefore'))
        if x is not None:
            before = int(x.get(qn('w:val')) or 0)
    index = before
    ranges = []
    for tc in tr.findall(qn('w:tc')):
        pr = tc.find(qn('w:tcPr'))
        span = pr.find(qn('w:gridSpan')) if pr is not None else None
        n = int(span.get(qn('w:val'))) if span is not None else 1
        ranges.append((index, index+n))
        index += n
    if len(ranges) != 11 or index > count:
        return None
    return ranges


def add_three_by_four(tbl, tr, widths, ranges):
    # Leave tblGrid and all original rows untouched.
    # Choose three existing logical-column boundaries to produce four near-even cells.
    from itertools import combinations
    s, e = ranges[0][0], ranges[-1][1]
    boundaries = sorted(set([a for a, _ in ranges] + [e]))
    candidates = [v for v in boundaries if s < v < e]
    if len(candidates) < 3:
        return False
    total = sum(widths[s:e])
    def score(cuts):
        points = [s, *cuts, e]
        widths4 = [sum(widths[a:b]) for a, b in zip(points, points[1:])]
        return sum((w - total/4) ** 2 for w in widths4)
    k1, k2, k3 = min(combinations(candidates, 3), key=score)
    segments = [(s,k1),(k1,k2),(k2,k3),(k3,e)]
    for _ in range(3):
        new = OxmlElement('w:tr')
        trpr = OxmlElement('w:trPr')
        if s:
            x = OxmlElement('w:gridBefore'); x.set(qn('w:val'),str(s)); trpr.append(x)
        if len(widths)-e:
            x = OxmlElement('w:gridAfter'); x.set(qn('w:val'),str(len(widths)-e)); trpr.append(x)
        if len(trpr): new.append(trpr)
        for a,b in segments:
            tc = OxmlElement('w:tc')
            tcpr = OxmlElement('w:tcPr')
            tcw = OxmlElement('w:tcW')
            tcw.set(qn('w:type'),'dxa')
            tcw.set(qn('w:w'),str(sum(widths[a:b])))
            tcpr.append(tcw)
            if b-a>1:
                gs = OxmlElement('w:gridSpan'); gs.set(qn('w:val'),str(b-a)); tcpr.append(gs)
            borders = OxmlElement('w:tcBorders')
            for edge in ('top','left','bottom','right','insideH','insideV','start','end'):
                border = OxmlElement('w:' + edge)
                border.set(qn('w:val'),'nil')
                borders.append(border)
            tcpr.append(borders)
            tc.append(tcpr)
            tc.append(OxmlElement('w:p'))
            new.append(tc)
        tr.addprevious(new)
    return True


def process(doc):
    stats = {'detected':0,'inserted':0,'skipped_existing_rows':0,'skipped_previous_table':0,'skipped_unsafe':0}
    for t in doc.tables:
        hdr = next((i for i,row in enumerate(t.rows) if is_header(row.cells)), None)
        if hdr is None: continue
        stats['detected'] += 1
        # Rows already present above the header in the SAME table count as bank info.
        if hdr > 0:
            stats['skipped_existing_rows'] += 1
            continue
        if preceding_table_immediately(t._tbl):
            stats['skipped_previous_table'] += 1
            continue
        widths = grid_widths(t._tbl)
        rows = t._tbl.findall(qn('w:tr'))
        if not widths or not rows:
            stats['skipped_unsafe'] += 1
            continue
        ranges = header_grid_ranges(rows[0],len(widths))
        if ranges is None or not add_three_by_four(t._tbl,rows[0],widths,ranges):
            stats['skipped_unsafe'] += 1
            continue
        stats['inserted'] += 1
    return stats


def main():
    cwd = Path.cwd()
    candidates = sorted(x for x in cwd.glob('*.docx') if not x.name.startswith('~$') and not x.stem.endswith('_插入3行4列'))
    if not candidates:
        print('当前目录没有可处理的 DOCX 文件：',cwd)
        return
    print('当前目录：',cwd)
    for i,f in enumerate(candidates,1): print(f'{i:>2}. {f.name}')
    choice = input('输入文件编号（q退出）：').strip()
    if choice.lower()=='q':return
    if not choice.isdigit() or not 1<=int(choice)<=len(candidates):
        print('编号无效');return
    source=candidates[int(choice)-1]
    dest=source.with_name(source.stem+'_插入3行4列.docx')
    if dest.exists():
        print('输出已存在，为避免覆盖已停止：',dest.name);return
    doc=Document(str(source))
    stats=process(doc)
    doc.save(str(dest))
    print('\n输出：',dest.name)
    print('识别标准11列表格：{detected}；新增3行×4列：{inserted}；已有表头前行跳过：{skipped_existing_rows}；前方有独立表格跳过：{skipped_previous_table}；结构异常跳过：{skipped_unsafe}'.format(**stats))
    print('原文件保持不变。')

if __name__=='__main__':
    try: main()
    except ImportError as exc:
        print('请先安装：python3 -m pip install python-docx')
        raise