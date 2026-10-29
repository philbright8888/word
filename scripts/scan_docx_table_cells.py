# -*- coding: utf-8 -*-
"""只读扫描 DOCX 中每个表格/单元格的实际 Word XML 网格结构。
运行: python scan_docx_table_cells.py
输出: *_table_scan.txt (易读报告), *_table_cells.json (完整结构)
不打开 Word，不修改原 DOCX，不依赖第三方库。
"""
from __future__ import annotations
import json
import re
import zipfile
from pathlib import Path
from collections import Counter
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
TWIPS_PER_CM = 1440 / 2.54
TARGETS = [
    ('No.', r'(?:serial\s*no\.?\s*no\.?|^no\.$)'),
    ('Bk.D.', r'(?:booked\s*date\s*bk\.?d\.?|^bk\.?d\.?)'),
    ('Val.D.', r'(?:value\s*date\s*vou\.?|^val\.?d\.?)'),
    ('Type', r'(?:transaction\s*type\s*type|^type$)'),
    ('Vou.', r'(?:debit\s*vou\.?|^vou\.$)'),
    ('Vou.No./Trans.No./Details', r'(?:vou\.?\s*no\.?\s*/\s*trans\.?\s*no\.?\s*/\s*details)'),
    ('DebitAmount', r'(?:debit\s*amount)'),
    ('CreditAmount', r'(?:credit\s*amount)'),
    ('Balance', r'(?:^balance(?:\s+balance)?$)'),
    ('ReferenceNo.', r'(?:institution\s*/\s*cashier\s*/\s*sequence\s*reference\s*no\.?|^reference\s*no\.?)'),
    ('Notes', r'(?:^notes(?:\s+notes)?$)'),
]

def attr(el, local):
    return el.get(W+local) if el is not None else None

def cm(twips):
    try: return round(float(twips) / TWIPS_PER_CM, 3)
    except (TypeError, ValueError): return None

def normalize(s):
    return re.sub(r'\s+', ' ', s.replace('\u00a0', ' ').replace('\u3000', ' ')).strip()

def get_text(cell):
    # Only direct paragraphs: nested table text belongs to the nested table.
    chunks = []
    for p in cell.findall('./w:p', NS):
        pieces = []
        for x in p.iter():
            if x.tag == W + 't' and x.text:
                pieces.append(x.text)
            elif x.tag == W + 'tab':
                pieces.append(' ')
            elif x.tag in (W + 'br', W + 'cr'):
                pieces.append(' ')
        if pieces: chunks.append(''.join(pieces))
    return normalize(' '.join(chunks))

def scan_table(tbl, index, depth):
    tblpr = tbl.find('./w:tblPr', NS)
    tblw = tblpr.find('./w:tblW', NS) if tblpr is not None else None
    layout = tblpr.find('./w:tblLayout', NS) if tblpr is not None else None
    grid = tbl.find('./w:tblGrid', NS)
    columns = []
    if grid is not None:
        for i, gc in enumerate(grid.findall('./w:gridCol', NS), 1):
            tw = attr(gc,'w')
            columns.append({'grid_column':i, 'width_twips':int(tw) if tw and tw.isdigit() else tw, 'width_cm':cm(tw)})
    rows = []
    labels = []
    for ri, tr in enumerate(tbl.findall('./w:tr', NS), 1):
        trpr = tr.find('./w:trPr', NS)
        before_el = trpr.find('./w:gridBefore', NS) if trpr is not None else None
        after_el = trpr.find('./w:gridAfter', NS) if trpr is not None else None
        before = int(attr(before_el,'val') or 0)
        after = int(attr(after_el,'val') or 0)
        col = before
        cells = []
        for ci, tc in enumerate(tr.findall('./w:tc', NS),1):
            tcpr = tc.find('./w:tcPr', NS)
            spanel = tcpr.find('./w:gridSpan', NS) if tcpr is not None else None
            span = int(attr(spanel,'val') or 1)
            vm = tcpr.find('./w:vMerge', NS) if tcpr is not None else None
            vmerge = 'none' if vm is None else (attr(vm,'val') or 'continue')
            w = tcpr.find('./w:tcW', NS) if tcpr is not None else None
            wtyp = attr(w,'type')
            wraw = attr(w,'w')
            width_cm = cm(wraw) if wtyp in (None,'dxa') else None
            data = {
                'row':ri,'cell_in_row':ci,
                'grid_start':col + 1, 'grid_end':col + span,
                'grid_span':span,'vertical_merge':vmerge,
                'tcW_type':wtyp,'tcW_value':wraw,
                'tcW_cm':width_cm,
                'text':get_text(tc),
                'nested_tables':len(tc.findall('./w:tbl', NS)),
            }
            cells.append(data)
            if ri <= 5:
                txt = data['text'].lower()
                matches=[name for name,pat in TARGETS if re.search(pat,txt,re.IGNORECASE)]
                if matches:
                    labels.append({'row':ri,'cell':ci,'grid_start':col+1,'grid_end':col+span,'labels':matches,'text':data['text'][:140]})
            col += span
        rows.append({'row':ri,'grid_before':before,'grid_after':after,'grid_covered':col+after,
                     'physical_cells':len(cells),'cells':cells})
    return {
        'table_index':index, 'nesting_depth':depth,
        'grid_columns':len(columns),'tblGrid':columns,
        'tblW_type':attr(tblw,'type'),'tblW_value':attr(tblw,'w'),
        'tblW_cm':cm(attr(tblw,'w')) if attr(tblw,'type') in ('dxa',None) else None,
        'table_layout':attr(layout,'type') or 'inherited',
        'row_count':len(rows),'header_candidates':labels,'rows':rows,
    }

def scan_document(path):
    with zipfile.ZipFile(path,'r') as z:
        if 'word/document.xml' not in z.namelist():
            raise ValueError('不是有效的 Word DOCX 文件：缺少 word/document.xml')
        root = ET.fromstring(z.read('word/document.xml'))
    tables = []
    def walk(element, depth=0):
        for child in element:
            if child.tag == W + 'tbl':
                tables.append(scan_table(child,len(tables)+1,depth))
                walk(child,depth+1)
            else:
                walk(child,depth)
    body = root.find('./w:body',NS)
    if body is not None: walk(body)
    return {'source':str(path.resolve()),'table_count':len(tables),'tables':tables}

def report(scan):
    lines=[f'SOURCE: {scan["source"]}',f'TOTAL TABLES: {scan["table_count"]}',
           'READ ONLY: original DOCX is not changed.',
           'NOTE: w:tblGrid/tcW record Word XML widths, not pixel-exact on-screen widths.',
           'NOTE: grid_start/end are *physical grid* positions, not always 11 logical bank columns.',
           '']
    for t in scan['tables']:
        lines.append('='*90)
        lines.append(f'TABLE {t["table_index"]} | nested depth={t["nesting_depth"]} | '
                     f'rows={t["row_count"]} | physical grid columns={t["grid_columns"]} | '
                     f'layout={t["table_layout"]}')
        lines.append('tblW='+str(t['tblW_value'])+' ('+str(t['tblW_cm'])+' cm) '
                     + ' | grid widths (cm)='+', '.join(str(x['width_cm']) for x in t['tblGrid']))
        lines.append('Header candidates (first five rows):')
        if t['header_candidates']:
            for h in t['header_candidates']:
                lines.append(f'  row {h["row"]:>3} cell {h["cell"]:>2} grid {h["grid_start"]}-{h["grid_end"]}: '
                             f'{"/".join(h["labels"])} | {h["text"]}')
        else: lines.append('  (none recognized)')
        lines.append('Row shape histogram: '+str(dict(Counter((r['physical_cells'],r['grid_covered']) for r in t['rows']))))
        lines.append('Each row: row number / cell index / occupied grid / span / width / merge / text')
        for r in t['rows']:
            lines.append(f'  ROW {r["row"]} | {r["physical_cells"]} actual cells | covered grid={r["grid_covered"]} '
                         f'| gridBefore={r["grid_before"]}, gridAfter={r["grid_after"]}')
            for c in r['cells']:
                text=c['text'].replace('\n',' ')
                if len(text)>115: text=text[:112]+'...'
                lines.append(f'    cell {c["cell_in_row"]:>2} | grid {c["grid_start"]:>2}-{c["grid_end"]:<2} | '
                             f'span={c["grid_span"]:<2} | w={str(c["tcW_cm"]):>6} cm | '
                             f'vMerge={c["vertical_merge"]:<8} | {text}')
        lines.append('')
    return '\n'.join(lines)+'\n'

def main():
    folder=Path(__file__).resolve().parent
    files=sorted((p for p in folder.glob('*.docx') if not p.name.startswith('~$')),key=lambda x:x.name.lower())
    if not files:
        print('当前脚本目录没有找到 DOCX 文件:', folder)
        return
    print('请选择 DOCX 文件（只读扫描，不会修改）：')
    for i,p in enumerate(files,1): print(f'{i:>2}. {p.name}')
    while True:
        choice=input('请输入文件编号（q退出）：').strip()
        if choice.lower()=='q':return
        if choice.isdigit() and 1<=int(choice)<=len(files):break
        print('编号无效，请重试。')
    path=files[int(choice)-1]
    try:
        result=scan_document(path)
        base=path.with_suffix('')
        text_path=base.with_name(base.name+'_table_scan.txt')
        json_path=base.with_name(base.name+'_table_cells.json')
        text_path.write_text(report(result),encoding='utf-8-sig')
        json_path.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        print('\n扫描完成，原文件没有修改。')
        print(f'发现表格 {result["table_count"]} 张，单元格 {sum(len(r["cells"]) for t in result["tables"] for r in t["rows"])} 个。')
        print('易读报告:',text_path)
        print('完整结构:',json_path)
        print('请把 *_table_scan.txt 发给我，我就能据此定位 Vou./Details/Notes 的真实列边界。')
    except Exception as exc:
        print('扫描失败：',repr(exc))
        raise

if __name__=='__main__':
    main()