# -*- coding: utf-8 -*-
"""DOCX 17张银行表格：交易区规范成11个真实底层网格，同时应用上一版本11列统一宽度。
依赖：仅 Python 标准库。必须使用对应的脱敏 private_table_cells(1).json。
方式：将不规则的顶部信息行及最后一行与交易区分离成独立表格；保留其原始结构。
"""
from __future__ import annotations
import copy
import json
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
Q = lambda name: '{%s}%s' % (W, name)
ET.register_namespace('w', W)
# Retain common namespace prefixes when rewriting document.xml.
for prefix, uri in [('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
                    ('wp','http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),
                    ('a','http://schemas.openxmlformats.org/drawingml/2006/main'),
                    ('pic','http://schemas.openxmlformats.org/drawingml/2006/picture'),
                    ('v','urn:schemas-microsoft-com:vml'),
                    ('o','urn:schemas-microsoft-com:office:office'),
                    ('mc','http://schemas.openxmlformats.org/markup-compatibility/2006')]:
    ET.register_namespace(prefix, uri)

# 完全沿用上一个13表调整脚本的11列目标宽度（twips）
TARGET_WIDTHS = [455, 580, 580, 745, 414, 2110, 1574, 1574, 1740, 1988, 3980]
BASE_WIDTHS = [455, 580, 580, 745, 414, 3811, 1574, 1574, 1740, 1988, 2279]

LABELS = ['No.','Bk.D.','Val.D.','Type','Vou.',
          'Vou.No./Trans.No./Details','DebitAmount','CreditAmount',
          'Balance','ReferenceNo.','Notes']


def child(parent, name):
    return parent.find(Q(name))


def direct_rows(tbl):
    return [x for x in tbl if x.tag == Q('tr')]


def direct_cells(row):
    return [x for x in row if x.tag == Q('tc')]


def number_prop(parent, name, default=0):
    node = child(parent, name)
    return int(node.get(Q('val'))) if node is not None else default


def validate_row(row, record, table_index):
    cells = direct_cells(row)
    if len(cells) != 11 or len(record['cells']) != 11:
        raise ValueError(f'表{table_index} 行{record["row"]}不是11个单元格')
    trpr = child(row,'trPr')
    before = number_prop(trpr,'gridBefore') if trpr is not None else 0
    after = number_prop(trpr,'gridAfter') if trpr is not None else 0
    if before != record['grid_before'] or after != record['grid_after']:
        raise ValueError(f'表{table_index} 行{record["row"]}左右空网格不匹配')
    cur = before + 1
    for i, (tc, c) in enumerate(zip(cells,record['cells'])):
        props = child(tc,'tcPr')
        span = number_prop(props,'gridSpan',1) if props is not None else 1
        if (cur,cur+span-1) != (c['grid_start'],c['grid_end']):
            raise ValueError(f'表{table_index} 行{record["row"]} 单元格{i+1}跨度不匹配')
        if c['logical_column'] != LABELS[i]:
            raise ValueError(f'表{table_index} 行{record["row"]} 第{i+1}列映射异常')
        cur += span
    return cells


def make_tbl(original, rows, grid_widths, normalized):
    tbl = ET.Element(Q('tbl'))
    pr = child(original,'tblPr')
    if pr is not None:
        tbl.append(copy.deepcopy(pr))
    grid = ET.SubElement(tbl,Q('tblGrid'))
    for tw in grid_widths:
        ET.SubElement(grid,Q('gridCol'),{Q('w'):str(tw)})
    for row in rows:
        rr=copy.deepcopy(row)
        if normalized:
            trpr = child(rr,'trPr')
            if trpr is not None:
                for name in ('gridBefore','gridAfter','wBefore','wAfter'):
                    for node in list(trpr):
                        if node.tag == Q(name): trpr.remove(node)
            for tc in direct_cells(rr):
                tcpr = child(tc,'tcPr')
                if tcpr is not None:
                    for node in list(tcpr):
                        if node.tag == Q('gridSpan'):tcpr.remove(node)
                if tcpr is None: raise ValueError('标准交易单元格缺少 tcPr')
                tcw=child(tcpr,'tcW')
                if tcw is None: raise ValueError('标准交易单元格缺少 tcW')
                tcw.set(Q('w'), str(TARGET_WIDTHS[direct_cells(rr).index(tc)]))
                tcw.set(Q('type'), 'dxa')
        tbl.append(rr)
    return tbl


def para_separator():
    # Prevent Word from joining adjacent table nodes into one table.
    p=ET.Element(Q('p'))
    pr=ET.SubElement(p,Q('pPr'))
    spacing=ET.SubElement(pr,Q('spacing'))
    spacing.set(Q('before'),'0');spacing.set(Q('after'),'0')
    ET.SubElement(pr,Q('rPr'))
    return p


def normalize(doc_bytes, report):
    root=ET.fromstring(doc_bytes)
    body=child(root,'body')
    if body is None:raise ValueError('无法找到Word文档正文')
    tbls = list(body.iter(Q('tbl')))
    if len(tbls)!=report['total_tables_in_document'] or len(tbls)!=18:
        raise ValueError(f'文档表格数量不匹配：实际{len(tbls)}，报告{report["total_tables_in_document"]}')
    targets=report['tables']
    if sorted(t['table_index'] for t in targets)!=list(range(1,18)):
        raise ValueError('扫描报告不是预期的第1～17张银行表格')
    replacements=[]
    for t in targets:
        idx=t['table_index']
        orig=tbls[idx-1]
        # Do not alter nested tables; scan says all are top-level.
        if orig not in list(body):
            raise ValueError(f'表{idx}不是正文直接表格，停止')
        rows=direct_rows(orig)
        if len(rows)!=t['total_rows_in_table']:
            raise ValueError(f'表{idx}行数不匹配')
        grid=child(orig,'tblGrid')
        cols=list(grid) if grid is not None else []
        if len(cols)!=t['table_grid_column_count']:
            raise ValueError(f'表{idx}网格数不匹配')
        widths=[int(c.get(Q('w'))) for c in cols]
        # Match scanned widths within ~1 twip to avoid using a different version of the input.
        for i,(a,b) in enumerate(zip(widths,t['table_grid_widths'])):
            if abs(a-round(b['width_cm']*1440/2.54))>2:
                raise ValueError(f'表{idx}网格{i+1}宽度与扫描报告不符')
        bank_records=[r for r in t['row_data'] if r['cell_count']==11]
        if idx<=16:
            if len(bank_records)!=21 or [r['row'] for r in bank_records]!=list(range(5,26)):
                raise ValueError(f'表{idx}交易行范围不符合预期')
            start,end=5,25
        else:
            if len(bank_records)!=1 or bank_records[0]['row']!=6:
                raise ValueError('表17交易行范围不符合预期')
            start,end=6,6
        for rec in bank_records:
            validate_row(rows[rec['row']-1],rec,idx)
        # 读取原交易行列宽用于校验；所有规范化交易行统一使用上一版目标宽度。
        first_cells=direct_cells(rows[start-1])
        column_widths=[]
        for tc in first_cells:
            tcpr=child(tc,'tcPr')
            tcw=child(tcpr,'tcW') if tcpr is not None else None
            if tcw is None or tcw.get(Q('type')) != 'dxa':
                raise ValueError(f'表{idx}缺少固定列宽，停止')
            column_widths.append(int(tcw.get(Q('w'))))
        # 不再保留原宽度：严格统一到上一版13表脚本的11列目标宽度。
        if any(abs(a-b)>4 for a,b in zip(column_widths,BASE_WIDTHS)):
            raise ValueError(f'表{idx}原交易列宽与上一版参考表不一致，已停止防止错改')
        groups=[]
        if start>1:groups.append((False,rows[:start-1],widths))
        groups.append((True,rows[start-1:end],TARGET_WIDTHS))
        if end<len(rows):groups.append((False,rows[end:],widths))
        out=[]
        for normalized,subset,ww in groups:
            if out:out.append(para_separator())
            out.append(make_tbl(orig,subset,ww,normalized))
        replacements.append((orig,out,idx,len(bank_records)))
    for orig,sequence,idx,count in replacements:
        pos=list(body).index(orig)
        body.remove(orig)
        for k,element in enumerate(sequence):body.insert(pos+k,element)
    return ET.tostring(root,encoding='utf-8',xml_declaration=True),replacements


def select_docx():
    docs=[p for p in Path.cwd().glob('*.docx') if not p.name.startswith('~$') and '_11cols' not in p.stem]
    if not docs:raise ValueError('当前文件夹没有待处理的 DOCX')
    print('\n当前文件夹 DOCX：')
    for i,p in enumerate(docs,1):print(f' {i}. {p.name}')
    s=input('输入文件编号：').strip()
    if not s.isdigit() or not 1<=int(s)<=len(docs):raise ValueError('编号无效')
    return docs[int(s)-1]


def main():
    original=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else select_docx()
    report=Path(sys.argv[2]).resolve() if len(sys.argv)>2 else Path.cwd()/'private_table_cells(1).json'
    if not report.exists():raise FileNotFoundError(f'没有找到脱敏扫描JSON：{report}')
    meta=json.loads(report.read_text(encoding='utf-8'))
    if meta.get('matched_bank_tables')!=17:raise ValueError('报告必须识别17张银行表格')
    with zipfile.ZipFile(original) as zin:
        original_doc=zin.read('word/document.xml')
        new_doc,changes=normalize(original_doc,meta)
        dest=original.with_name(original.stem+'_11cols_same_width.docx')
        if dest.exists():raise FileExistsError(f'输出文件已存在，请先重命名：{dest}')
        with zipfile.ZipFile(dest,'w') as zout:
            for entry in zin.infolist():
                data=new_doc if entry.filename=='word/document.xml' else zin.read(entry.filename)
                zout.writestr(entry,data)
    with zipfile.ZipFile(dest) as check:
        ET.fromstring(check.read('word/document.xml'))
        if check.testzip():raise ValueError('输出ZIP校验失败')
    print('\n完成：',dest)
    print('17张银行表格的标准11列交易区已规范为11个物理网格。')
    print('其他不规则行（顶部信息行及第17张最后一行）已单独保留；第18张原表未改。')
    print('统一列宽（cm）：0.803, 1.023, 1.023, 1.314, 0.730, 3.722, 2.776, 2.776, 3.069, 3.507, 7.020')
    print('只修改标准交易区的网格与单元格宽度，其它行及第18张表保留。')
    print('注意：表格拆分可能造成Word分页/边框变化，请打开副本检查。')

if __name__=='__main__':
    try:main()
    except Exception as e:
        print('\n未修改原始文件。错误：',e)
        sys.exit(1)