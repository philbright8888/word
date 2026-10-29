# -*- coding: utf-8 -*-
"""Fast, Word-free removal of two types of bank-header TABLE ROWS from DOCX.

No third-party packages required; Python 3.12 compatible.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import re
import sys
import traceback
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TO_DATE = re.compile(r'\bto\s*[（(]\s*\d{8}\s*[)）]', re.IGNORECASE)
ACCOUNT_NO = re.compile(r'account\s*no\s*\.?', re.IGNORECASE)
ACCOUNT_NAME = re.compile(r'account\s*name', re.IGNORECASE)
BANK_NAME = re.compile(r'bank\s*name', re.IGNORECASE)


def text_in_row(row):
    """Collect visible text from cells, ignoring any nested tables."""
    pieces = []
    for cell in row.findall(W + 'tc'):
        for child in cell:
            if child.tag == W + 'tbl':
                continue
            # paragraph nodes can have runs split across many w:t elements
            for el in child.iter():
                if el.tag == W + 't':
                    pieces.append(el.text or '')
                elif el.tag in (W + 'tab', W + 'br'):
                    pieces.append(' ')
        pieces.append(' ')
    return re.sub(r'\s+', ' ', ''.join(pieces).replace('\u00a0', ' ')).strip()


def process_xml(xml_bytes):
    root = ET.fromstring(xml_bytes)
    n_to, n_account = 0, 0
    samples = []
    # Iterate table containers, including nested tables. Remove direct child rows only.
    for table in root.iter(W + 'tbl'):
        for row in list(table):
            if row.tag != W + 'tr':
                continue
            content = text_in_row(row)
            to_hit = bool(TO_DATE.search(content))
            account_hit = all(p.search(content) for p in (ACCOUNT_NO, ACCOUNT_NAME, BANK_NAME))
            if to_hit or account_hit:
                table.remove(row)
                if to_hit:
                    n_to += 1
                else:
                    n_account += 1
                if len(samples) < 30:
                    samples.append(('To(date)' if to_hit else 'Account fields', content[:200]))
    if not (n_to or n_account):
        return xml_bytes, n_to, n_account, samples
    return ET.tostring(root, encoding='utf-8', xml_declaration=True), n_to, n_account, samples


def run_one(src):
    dst = src.with_name(src.stem + '_rows_removed.docx')
    log = src.with_name(src.stem + '_rows_removed.log.txt')
    if dst.exists():
        print('输出文件已存在，请先改名或移动，避免覆盖：', dst)
        return
    # Rewrite the ZIP package, changing ONLY word/document.xml.
    with ZipFile(src, 'r') as z_in:
        if 'word/document.xml' not in z_in.namelist():
            raise ValueError('不是有效的 DOCX：找不到 word/document.xml')
        new_xml, n_to, n_account, samples = process_xml(z_in.read('word/document.xml'))
        with ZipFile(dst, 'w') as z_out:
            for info in z_in.infolist():
                payload = new_xml if info.filename == 'word/document.xml' else z_in.read(info.filename)
                z_out.writestr(info, payload)
    lines = [f'SOURCE: {src}', f'OUTPUT: {dst}',
             f'Rows deleted by To(8-digit-date): {n_to}',
             f'Rows deleted by Account No. + Account Name + Bank Name: {n_account}',
             f'TOTAL ROWS DELETED: {n_to + n_account}', '',
             'Matching row previews:']
    lines.extend(f'{kind}: {preview}' for kind, preview in samples)
    log.write_text('\n'.join(lines), encoding='utf-8')
    print('\n处理完成：')
    print('To(日期) 所在行删除：', n_to)
    print('账户三字段所在行删除：', n_account)
    print('输出：', dst.name)
    print('日志：', log.name)
    if n_to + n_account == 0:
        print('没有匹配到表格行；请检查字段是否分布在不同行或文本框中。')


def main():
    folder = Path(__file__).resolve().parent
    files = sorted((p for p in folder.glob('*.docx')
                    if not p.name.startswith('~$')
                    and not p.stem.endswith('_rows_removed')),
                   key=lambda p: p.name.lower())
    if not files:
        print('当前脚本目录没有可处理的 DOCX：', folder)
        return
    print('当前目录：', folder, '\n')
    for i, f in enumerate(files, 1):
        print(f'{i}. {f.name}')
    choice = input('\n请输入文件编号（q 退出）：').strip()
    if choice.lower() == 'q':
        return
    if not choice.isdecimal() or not 1 <= int(choice) <= len(files):
        print('编号无效')
        return
    run_one(files[int(choice) - 1])


if __name__ == '__main__':
    try:
        main()
    except Exception:
        traceback.print_exc()
    if sys.stdin.isatty():
        input('\n按 Enter 退出...')