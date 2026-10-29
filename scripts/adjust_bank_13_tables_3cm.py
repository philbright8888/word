# -*- coding: utf-8 -*-
"""仅调整脱敏扫描报告确认的13张银行表格的第6、11列宽度。

无需第三方库，Python 3.9+；不上传或输出任何银行数据。
用法：python adjust_bank_13_tables_3cm.py [原始文件.docx]
"""
from __future__ import annotations
import os
import sys
import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'
# 来自用户上传的脱敏扫描报告；顺序是 document.xml 中表格编号（从 1 开始）。
TARGETS = {i: (15 if i == 26 else 21) for i in range(2, 27, 2)}
EXPECTED_GRID_TWIPS = [455, 580, 580, 745, 414, 3811, 1574, 1574, 1740, 1988, 2279]
MOVE_TWIPS = round(3 / 2.54 * 1440)  # 3cm，约 1701 twips
DOCUMENT_PART = 'word/document.xml'


def rows_of(tbl):
    return tbl.findall(f'{W}tr')


def cells_of(row):
    return row.findall(f'{W}tc')


def require_safe_table(tbl, index, expected_rows):
    rows = rows_of(tbl)
    if len(rows) != expected_rows:
        return f'行数不匹配：{len(rows)} != {expected_rows}'
    grid = tbl.find(f'{W}tblGrid')
    if grid is None:
        return '缺少 tblGrid'
    cols = grid.findall(f'{W}gridCol')
    if len(cols) != 11:
        return f'底层网格不是11列：{len(cols)}'
    for i, (col, expected) in enumerate(zip(cols, EXPECTED_GRID_TWIPS), 1):
        try:
            actual = int(col.get(W + 'w'))
        except (ValueError, TypeError):
            return f'网格第{i}列缺少有效宽度'
        if abs(actual - expected) > 4:
            return f'网格第{i}列宽度与扫描报告不一致：{actual} != {expected}'
    for r, row in enumerate(rows, 1):
        cells = cells_of(row)
        if len(cells) != 11:
            return f'第{r}行单元格数量不是11：{len(cells)}'
        trpr = row.find(f'{W}trPr')
        if trpr is not None and trpr.find(f'{W}gridBefore') is not None:
            return f'第{r}行含 gridBefore'
        if trpr is not None and trpr.find(f'{W}gridAfter') is not None:
            return f'第{r}行含 gridAfter'
        for c, cell in enumerate(cells, 1):
            tcpr = cell.find(f'{W}tcPr')
            if tcpr is None:
                return f'第{r}行第{c}格无 tcPr'
            if tcpr.find(f'{W}gridSpan') is not None or tcpr.find(f'{W}vMerge') is not None:
                return f'第{r}行第{c}格存在合并标记'
            tcw = tcpr.find(f'{W}tcW')
            if tcw is None or tcw.get(W + 'type') != 'dxa':
                return f'第{r}行第{c}格宽度类型不支持'
            try:
                cell_width = int(tcw.get(W + 'w'))
            except (ValueError, TypeError):
                return f'第{r}行第{c}格宽度无效'
            if abs(cell_width - EXPECTED_GRID_TWIPS[c - 1]) > 4:
                return f'第{r}行第{c}格宽度与扫描报告不一致'
    return None


def adjust_table(tbl):
    grid = tbl.find(f'{W}tblGrid')
    grid_cols = grid.findall(f'{W}gridCol')
    grid_cols[5].set(W + 'w', str(int(grid_cols[5].get(W + 'w')) - MOVE_TWIPS))
    grid_cols[10].set(W + 'w', str(int(grid_cols[10].get(W + 'w')) + MOVE_TWIPS))
    for row in rows_of(tbl):
        cells = cells_of(row)
        for idx, delta in ((5, -MOVE_TWIPS), (10, MOVE_TWIPS)):
            tcw = cells[idx].find(f'{W}tcPr/{W}tcW')
            tcw.set(W + 'w', str(int(tcw.get(W + 'w')) + delta))


def choose_input():
    if len(sys.argv) > 1:
        p = Path(sys.argv[1].strip('"')).expanduser().resolve()
        return p
    folder = Path(__file__).resolve().parent
    matches = sorted(p for p in folder.glob('*.docx') if not p.name.startswith('~$') and not p.stem.endswith('_13tables_adjusted'))
    if not matches:
        print('未找到 DOCX。请将原始 DOCX 放在此脚本同一文件夹，或在命令后指定路径。')
        return None
    print('请选择原始 DOCX：')
    for i, p in enumerate(matches, 1):
        print(f'  {i}. {p.name}')
    try:
        num = int(input('输入文件编号：').strip())
        return matches[num - 1] if 1 <= num <= len(matches) else None
    except (ValueError, IndexError, EOFError):
        return None


def main():
    src = choose_input()
    if not src or not src.is_file() or src.suffix.lower() != '.docx':
        print('输入文件无效。')
        return 1
    out = src.with_name(src.stem + '_13tables_adjusted.docx')
    if out.exists():
        print('输出文件已存在，为避免覆盖，程序停止：', out)
        return 1
    with zipfile.ZipFile(src, 'r') as zin:
        if zin.testzip() is not None or DOCUMENT_PART not in zin.namelist():
            print('DOCX ZIP 损坏或不包含 word/document.xml')
            return 1
        original_xml = zin.read(DOCUMENT_PART)
        root = ET.fromstring(original_xml)
        body = root.find(f'{W}body')
        if body is None:
            print('DOCX 缺少 document body')
            return 1
        tables = list(body.iter(W + 'tbl'))
        print(f'文档扫描表格总数：{len(tables)}')
        # 严格对齐报告编号，以免目标表编号偏移导致误改。
        if len(tables) != 26:
            print('安全检查未通过：扫描报告对应26张表格，实际数量不同。未修改任何文件。')
            return 1
        issues = []
        for table_index, expected_rows in TARGETS.items():
            why = require_safe_table(tables[table_index - 1], table_index, expected_rows)
            if why:
                issues.append(f'表格 {table_index}: {why}')
        if issues:
            print('安全检查未通过，未修改任何表格：')
            for issue in issues:
                print(' -', issue)
            return 1
        for table_index in TARGETS:
            adjust_table(tables[table_index - 1])
        # 仅更新 document.xml；所有其他压缩包条目直接复制。
        ET.register_namespace('w', NS['w'])
        # 保留文档中其它命名空间的常见前缀，减少无意义 XML 改写。
        for event, data in ET.iterparse(__import__('io').BytesIO(original_xml), events=('start-ns',)):
            prefix, uri = data
            if prefix not in ('xml', 'xmlns'):
                try:
                    ET.register_namespace(prefix, uri)
                except ValueError:
                    pass
        patched_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        fd, tmp = tempfile.mkstemp(suffix='.docx', dir=str(src.parent))
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, 'w') as zout:
                for info in zin.infolist():
                    zout.writestr(info, patched_xml if info.filename == DOCUMENT_PART else zin.read(info.filename))
            with zipfile.ZipFile(tmp, 'r') as chk:
                if chk.testzip() is not None:
                    raise RuntimeError('输出 ZIP 校验失败')
                check_root = ET.fromstring(chk.read(DOCUMENT_PART))
                check_tables = list(check_root.find(f'{W}body').iter(W + 'tbl'))
                for ti in TARGETS:
                    t = check_tables[ti - 1]
                    gc = t.find(f'{W}tblGrid').findall(f'{W}gridCol')
                    if int(gc[5].get(W+'w')) != EXPECTED_GRID_TWIPS[5] - MOVE_TWIPS:
                        raise RuntimeError(f'输出表格 {ti} 校验失败')
                # 所有未调整的表，其 XML 内容结构保持一致。
                for ti in range(1, 27, 2):
                    if ET.tostring(check_tables[ti-1]) != ET.tostring(tables[ti-1]):
                        raise RuntimeError(f'非目标表格 {ti} 被意外修改')
            os.replace(tmp, out)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
    print('\n调整成功！')
    print('修改表格：', ', '.join(map(str, TARGETS)))
    print('未修改表格：', ', '.join(map(str, range(1,27,2))))
    print(f'Details: 6.722 cm → 约 {((EXPECTED_GRID_TWIPS[5]-MOVE_TWIPS)*2.54/1440):.3f} cm')
    print(f'Notes: 4.020 cm → 约 {((EXPECTED_GRID_TWIPS[10]+MOVE_TWIPS)*2.54/1440):.3f} cm')
    print('输出文件：', out)
    print('原文件未修改；请使用 Word 检查版式。')
    return 0

if __name__ == '__main__':
    sys.exit(main())