#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 银行表格排版 V4：17张11逻辑列表统一列宽，隐私本地处理。
依赖: python-docx。可选 Pillow（提高宽度估算精度）。
原件绝不修改；输出 原文件名_格式处理.docx；如果存在则拒绝覆盖。
"""
import re, sys, math, copy
from pathlib import Path
from collections import Counter
try:
    from docx import Document
    from docx.table import _Cell
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    print('缺少 python-docx。Mac 安装：python3 -m pip install python-docx')
    print('如果禁止直接安装：python3 -m venv .venv && source .venv/bin/activate && python -m pip install python-docx')
    sys.exit(1)
try:
    from PIL import ImageFont
except ImportError:
    ImageFont = None

LABELS=['No.','Bk.D.','Val.D.','Type','Vou.','Vou.No./Trans.No./Details',
        'DebitAmount','CreditAmount','Balance','ReferenceNo.','Notes']
RAW=[['Serial No. No.','No.'],['Booked Date Bk.D.','Bk.D.'],['Value Date Vou.','Val.D.'],
     ['Transaction Type Type','Type'],['Debit Vou.','Vou.'],
     ['Vou. No./Trans. No./Details Vou. No./Trans. No./Details','Vou.No./Trans.No./Details','Vou. No./Trans. No./Details'],
     ['Debit Amount Debit Amount','DebitAmount'],['Credit Amount Credit Amount','CreditAmount'],
     ['Balance Balance','Balance'],['Institution/Cashier/Sequence Reference No.','ReferenceNo.','Reference No.'],
     ['Notes Notes','Notes']]
def norm(s):return re.sub(r'[^a-z0-9]','',s.casefold())
LOOKUP={norm(v):i for i, vv in enumerate(RAW) for v in vv}
ALIGN=[WD_ALIGN_PARAGRAPH.CENTER]*5+[WD_ALIGN_PARAGRAPH.LEFT]+[WD_ALIGN_PARAGRAPH.RIGHT]*3+[WD_ALIGN_PARAGRAPH.CENTER,WD_ALIGN_PARAGRAPH.LEFT]
COUNT=Counter(); WARN=[]
TWIPS_PER_CM=1440/2.54

def cm2tw(x):return int(round(x*TWIPS_PER_CM))
def tw2cm(x):return x/TWIPS_PER_CM

def cells(row):return [_Cell(tc,row.table) for tc in row._tr.tc_lst]
def text(c):return ' '.join(p.text for p in c.paragraphs).strip()
def unsafe(p):
    return bool(p._p.xpath('.//w:drawing | .//w:pict | .//w:object | .//w:fldChar | .//w:instrText | .//w:hyperlink | .//w:bookmarkStart | .//w:bookmarkEnd | .//w:sdt | .//w:footnoteReference | .//w:endnoteReference | .//w:commentRangeStart | .//w:commentRangeEnd'))
def child(parent,name):
    n=parent.find(qn(name))
    if n is None:n=OxmlElement(name);parent.append(n)
    return n

def plain(c):
    if c.tables or not c.paragraphs:return False
    for p in c.paragraphs:
        if unsafe(p):return False
        if any(x.tag not in (qn('w:pPr'),qn('w:r')) for x in p._p):return False
        for r in p.runs:
            if any(x.tag not in (qn('w:rPr'),qn('w:t'),qn('w:tab'),qn('w:br'),qn('w:cr')) for x in r._r):return False
            if any(x.get(qn('w:type')) not in (None,'textWrapping') for x in r._r.xpath('.//w:br')):return False
    return True

def clean(s):
    s=s.replace('\u00a0',' ').replace('\u3000',' ').replace('\u200b','')
    return re.sub(r' {2,}',' ',re.sub(r'[\t\n\r\v\f]+',' ',s)).strip()

def content_for_measure(c):
    return clean(' '.join(p.text for p in c.paragraphs))

def join_cell(c, header=None):
    if not plain(c):COUNT['complex_cells_skipped_text_cleanup']+=1;return
    ps=c.paragraphs
    if header is None:
        parts=[clean(p.text) for p in ps];parts=[p for p in parts if p]
        out=''
        for part in parts:
            if not out:out=part
            elif re.search(r'[\u3400-\u9fff]$',out) or re.match(r'^[\u3400-\u9fff]',part):out+=part
            else:out+=' '+part
        out=clean(out)
    else:out=header
    if len(ps)==1 and ps[0].text==out and len(ps[0].runs)<=1:return
    first=ps[0]
    for run in first.runs:first._p.remove(run._r)
    if out:first.add_run(out)
    for p in ps[1:]:p._p.getparent().remove(p._p)
    COUNT['cleaned_cells']+=1

def format_run(run):
    if run._r.xpath('.//w:drawing | .//w:pict | .//w:object'):return
    run.font.size=Pt(8)
    pr=run._r.get_or_add_rPr();fonts=child(pr,'w:rFonts')
    for k,v in [('ascii','Times New Roman'),('hAnsi','Times New Roman'),('eastAsia','SimSun'),('cs','Times New Roman')]:
        fonts.set(qn('w:'+k),v)
    for k in ('asciiTheme','hAnsiTheme','eastAsiaTheme','csTheme'):fonts.attrib.pop(qn('w:'+k),None)
    for k,v in [('w:w','100'),('w:spacing','0'),('w:position','0'),('w:szCs','16')]:child(pr,k).set(qn('w:val'),v)

def format_para(p, alignment):
    if unsafe(p):COUNT['protected_image_or_complex_paragraphs']+=1;return
    f=p.paragraph_format
    f.left_indent=Pt(0);f.right_indent=Pt(0);f.first_line_indent=Pt(0)
    f.space_before=Pt(0);f.space_after=Pt(0)
    f.line_spacing=Pt(8);f.line_spacing_rule=WD_LINE_SPACING.EXACTLY
    p.alignment=alignment
    for run in p.runs:format_run(run)

def set_cell_lr_margin_zero(c):
    # Only non-standard cells: tcMar left/right are internal padding, not table widths.
    pr=c._tc.get_or_add_tcPr();mar=child(pr,'w:tcMar')
    for edge in ('left','right','start','end'):
        e=child(mar,'w:'+edge);e.set(qn('w:w'),'0');e.set(qn('w:type'),'dxa')

def identify(table):
    best={};bestrow=None
    for ix,row in enumerate(table.rows[:min(14,len(table.rows))]):
        found={}
        for j,c in enumerate(cells(row)):
            k=LOOKUP.get(norm(text(c)))
            if k is not None and k not in found:found[k]=j
        if len(found)>len(best):best=found;bestrow=ix
    if len(best)!=11 or [best.get(i) for i in range(11)]!=list(range(11)):
        return None
    return bestrow

def grid_info(table):
    grid=table._tbl.tblGrid
    cols=list(grid.gridCol_lst)
    ws=[int(c.get(qn('w:w'),'0')) for c in cols]
    return cols,ws

def row_cells_with_grid(row):
    trpr=row._tr.trPr
    before=0
    if trpr is not None:
        n=trpr.find(qn('w:gridBefore'))
        if n is not None:before=int(n.get(qn('w:val'),'0'))
    ix=before;out=[]
    for c in cells(row):
        pr=c._tc.tcPr;gs=pr.find(qn('w:gridSpan')) if pr is not None else None
        span=int(gs.get(qn('w:val'),'1')) if gs is not None else 1
        out.append((c,ix,ix+span));ix+=span
    return out

def map_ranges(table,header_ix):
    row=list(table.rows)[header_ix]
    mapping=row_cells_with_grid(row)
    if len(mapping)!=11:return None
    cols,old=grid_info(table)
    if not cols or any(a>=b or b>len(cols) for _,a,b in mapping):return None
    return [(a,b) for _,a,b in mapping]

# Use font measurement if available, fallback conservative approximation in cm.
FONT=None
if ImageFont is not None:
    for f in ['/System/Library/Fonts/Supplemental/Times New Roman.ttf',
              '/Library/Fonts/Times New Roman.ttf',
              '/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf']:
        try:FONT=ImageFont.truetype(f,80);break
        except Exception:pass

def measure_cm(s):
    if not s:return 0.0
    if FONT is not None:
        try:
            # 80px simulated at 8 pt => 0.28222cm / 80px
            return FONT.getlength(s)* (8*2.54/72)/80
        except Exception:pass
    units=0
    for c in s:
        if '\u3400'<=c<='\u9fff':units+=1
        elif c in 'il1.,:;!|\'`':units+=0.27
        elif c in 'MW@#%':units+=0.87
        elif c==' ':units+=0.26
        else:units+=0.50
    return units*(8*2.54/72)

def wrap_line_count(s,width):
    # Estimator for optimization only; actual line breaking belongs to Word.
    usable=max(.30,width-.18)
    if not s:return 1
    # Approximate word wrapping; very long tokens can break at characters.
    tokens=re.findall(r'\S+\s*',s)
    lines=1;used=0.0
    for token in tokens:
        w=measure_cm(token)
        if w<=usable:
            if used+w>usable and used>0:lines+=1;used=0
            used+=w
        else:
            if used>0:lines+=1;used=0
            lines+=max(0,math.ceil(w/usable)-1)
            used=w%usable
    return lines

def page_width_cm(doc):
    vals=[]
    for s in doc.sections:
        width=tw2cm(int(s.page_width.twips-s.left_margin.twips-s.right_margin.twips))
        vals.append(width)
    # One uniform width must fit every section.
    return min(vals) if vals else 27.7

def safe_total_width(doc, banks):
    # 页面实际版心宽度，横向 A4 也必须扣掉两侧页边距。
    printable = page_width_cm(doc)
    # Word 的少数 PDF 转换表在 11 个逻辑列之外还有左右辅助网格。
    guards = 0.0
    for _, table, hi, mapping in banks:
        _, old = grid_info(table)
        first, last = mapping[0][0], mapping[-1][1]
        guards = max(guards, tw2cm(sum(old[:first]) + sum(old[last:])))
    safety = 0.25   # 右侧预留 2.5mm，避免边框被裁切
    limit = printable - guards - safety
    if limit < 15:
        raise RuntimeError(f'当前页面可用宽度 {printable:.3f} cm，扣除辅助网格与安全余量后仅 {limit:.3f} cm，已停止。')
    width = min(27.764, limit)
    if width < 27.764:
        WARN.append(f'页面可用宽度 {printable:.3f}cm，辅助网格最多 {guards:.3f}cm；11列宽自动由27.764缩至{width:.3f}cm。')
    return width, printable, guards

def optimize(doc,banks):
    # 以原来的 27.764cm 为上限；实际页面装不下则自动收窄。
    avail, printable, guards = safe_total_width(doc,banks)
    values=[[] for _ in range(11)]
    records=[]
    for _,t,hi,_ in banks:
        for ri,row in enumerate(t.rows):
            if ri<hi:continue
            cc=cells(row)
            if len(cc)!=11:continue
            rowtxt=[content_for_measure(c) for c in cc]
            if ri==hi:rowtxt=LABELS[:]
            for i,tt in enumerate(rowtxt):
                if tt:values[i].append(tt)
            if ri!=hi:records.append((rowtxt[5],rowtxt[10],rowtxt[3]))
    # Columns 0..10: 6/11 are elastic; type allowed two lines.
    fixed={}
    for i in (0,1,2,4,6,7,8,9):
        longest=max([measure_cm(s) for s in values[i]] or [0])
        fixed[i]=max(.65 if i==4 else .78,longest+.28)
    # Type 2 lines, avoid extreme narrowness; basic header protection.
    typeword=max([measure_cm(w) for s in values[3] for w in s.split()] or [.5])
    fixed[3]=max(.88,min(1.6,typeword+.22))
    fixedsum=sum(fixed.values())
    remaining=avail-fixedsum
    if remaining<4.0:
        raise RuntimeError(f'固定目标宽度 {avail:.2f}cm，固定列估算占 {fixedsum:.2f}cm，Details+Notes 不足4cm；为避免破坏排版已停止。')
    # Notes 指定为 7.55cm；总宽不增加，差额由 Details 承担。
    # 固定列仍按全文件的最大内容紧凑计算。
    min_details=1.7
    notes_target=7.55
    if remaining-notes_target < min_details:
        raise RuntimeError(
            f'页面可用总宽 {avail:.3f}cm；其他9列需 {fixedsum:.3f}cm，'
            f'Notes=7.55cm 后 Details 只有 {remaining-notes_target:.3f}cm。'
            '为避免超出右边界，已停止；请核对页面边距或减少其他列宽。'
        )
    low=remaining-notes_target;high=low
    best=None
    for j in range(101):
        details=low+(high-low)*j/100
        notes=remaining-details
        cost=0
        for d,n,ty in records:
            cost+=max(wrap_line_count(d,details),wrap_line_count(n,notes),min(2,wrap_line_count(ty,fixed[3])))
        # Tie breaker balances the allocation; small preference for wider Notes if same cost.
        tie=abs(details-notes)*0.0001
        candidate=(cost+tie,details,notes)
        if best is None or candidate<best:best=candidate
    _,d,n=best
    widths=[fixed.get(i,d if i==5 else n) for i in range(11)]
    # Round in twips and assign rounding residue to Notes to exactly fill available width.
    tw=[cm2tw(w) for w in widths]
    tw[10]+=cm2tw(avail)-sum(tw)
    return tw,avail

def set_tc_width(c,tw):
    pr=c._tc.get_or_add_tcPr();w=child(pr,'w:tcW')
    w.set(qn('w:w'),str(tw));w.set(qn('w:type'),'dxa')

def align_table_to_text_area(table):
    # 去掉 PDF 转换可能留下的整表左侧缩进，避免总宽不超却右边框出纸。
    pr=table._tbl.tblPr
    ind=child(pr,'w:tblInd');ind.set(qn('w:w'),'0');ind.set(qn('w:type'),'dxa')
    jc=child(pr,'w:jc');jc.set(qn('w:val'),'left')

def adjust_grid(table,hi,target):
    cols,old=grid_info(table);ranges=map_ranges(table,hi)
    if ranges is None:raise ValueError('无法映射表头11个逻辑列')
    new=old[:]
    # Every original internal grid inside a logical column is proportionally redistributed.
    for j,(a,b) in enumerate(ranges):
        orig=old[a:b];total=sum(orig)
        if total<=0:raise ValueError('网格宽度为0')
        alloc=[round(target[j]*v/total) for v in orig]
        alloc[-1]+=target[j]-sum(alloc)
        if min(alloc)<=0:raise ValueError('网格拆分后出现0宽度')
        new[a:b]=alloc
    # Keep small omitted left/right guard grids; actual table total may differ by ~0.05cm.
    for c,w in zip(cols,new):c.set(qn('w:w'),str(w))
    # All rows, including merged account info and irregular last row, update tcW according
    # to the grid cells they span. Never delete, move or change gridSpan/row structures.
    for row in table.rows:
        for c,a,b in row_cells_with_grid(row):
            if b>len(new):raise ValueError('单元格超出表格网格')
            set_tc_width(c,sum(new[a:b]))
    table.autofit=False
    align_table_to_text_area(table)
    COUNT['uniform_width_tables']+=1

def process_table(table,index,header_idx):
    for ri,row in enumerate(table.rows):
        cc=cells(row)
        standard=header_idx is not None and ri>=header_idx and len(cc)==11
        for ci,c in enumerate(cc):
            is_header=standard and ri==header_idx
            h=LABELS[ci] if is_header and LOOKUP.get(norm(text(c)))==ci else None
            join_cell(c,h)
            # Set internal left and right cell margins to zero on non-standard cells.
            if not standard:set_cell_lr_margin_zero(c)
            align=(WD_ALIGN_PARAGRAPH.CENTER if is_header else ALIGN[ci]) if standard else WD_ALIGN_PARAGRAPH.LEFT
            for p in c.paragraphs:format_para(p,align)

def safe_blank(p):
    if p.tag!=qn('w:p'):return False
    if p.xpath('.//w:drawing | .//w:pict | .//w:object | .//w:br | .//w:lastRenderedPageBreak | .//w:sectPr | .//w:fldChar | .//w:instrText | .//w:fldSimple | .//w:hyperlink | .//w:bookmarkStart | .//w:bookmarkEnd | .//w:sdt | .//w:footnoteReference | .//w:endnoteReference | .//w:commentRangeStart | .//w:commentRangeEnd'):return False
    return not any((t.text or '').strip() for t in p.xpath('.//w:t'))

def clean_outside(doc):
    body=doc.element.body;nodes=list(body)
    for i,p in enumerate(nodes):
        if not safe_blank(p):continue
        j=i-1
        while j>=0 and safe_blank(nodes[j]):j-=1
        k=i+1
        while k<len(nodes) and safe_blank(nodes[k]):k+=1
        prev=nodes[j] if j>=0 else None;next_=nodes[k] if k<len(nodes) else None
        if ((prev is not None and prev.tag==qn('w:tbl')) or (next_ is not None and next_.tag==qn('w:tbl')) or
            (i>0 and safe_blank(nodes[i-1]))):
            if p.getparent() is body:body.remove(p);COUNT['outside_blank_paragraphs_removed']+=1

def choose_file():
    base=Path.cwd();fs=sorted([f for f in base.glob('*.docx') if not f.name.startswith('~$') and not f.stem.endswith('_格式处理')],key=lambda x:x.name.casefold())
    if not fs:print('当前目录没有待处理 Word 文件：',base);return None
    print('当前目录：',base)
    for i,f in enumerate(fs,1):print(f'{i:2}. {f.name}')
    while True:
        s=input('输入 DOCX 编号（q 退出）：').strip()
        if s.lower()=='q':return None
        if s.isdigit() and 1<=int(s)<=len(fs):return fs[int(s)-1]
        print('输入的编号无效。')

def main():
    source=choose_file()
    if source is None:return
    dest=source.with_name(source.stem+'_格式处理.docx')
    if dest.exists():print('输出文件已存在，不覆盖：',dest);return
    doc=Document(source)
    banks=[]
    for i,t in enumerate(doc.tables,1):
        hi=identify(t)
        if hi is not None:
            mapping=map_ranges(t,hi)
            if mapping is None:WARN.append(f'表格{i}表头网格异常，不参与统一列宽')
            else:banks.append((i,t,hi,mapping))
    print(f'识别到 {len(banks)} 张11列银行表格，其他表格仅做文字清理。')
    if not banks:print('没有识别到标准11列表格，为防止误处理已停止。');return
    widths,available=optimize(doc,banks)
    print(f'标准11列目标总宽：{available:.3f}cm（Notes 固定 7.55cm，超过页边距则自动收窄）')
    print(f'页面最窄可用宽度：{page_width_cm(doc):.3f}cm；保留辅助网格和右侧安全余量')
    print('统一列宽（cm）：')
    for k,w in zip(LABELS,widths):print(f'  {k:<29} {tw2cm(w):.3f}')
    # Preflight on cloned XML before modifying original in-memory document.
    for index,table,hi,_ in banks:
        temp=copy.deepcopy(table._tbl)
        class Dummy:pass
        # preflight is based on all real row grid spans and both guard widths
        n=len(grid_info(table)[0])
        for row in table.rows:
            for _,a,b in row_cells_with_grid(row):
                if a<0 or b>n or a>=b:raise RuntimeError(f'表格{index}结构异常，停止处理')
    bank_headers={id(t._tbl):hi for _,t,hi,_ in banks}
    for index,t in enumerate(doc.tables,1):
        hi=bank_headers.get(id(t._tbl))
        process_table(t,index,hi)
    for index,t,hi,_ in banks:
        adjust_grid(t,hi,widths)
    clean_outside(doc)
    # Atomic create: exclusive output path, never overwrite input/output.
    with open(dest,'xb') as f:doc.save(f)
    log=dest.with_suffix('.log.txt')
    log.write_text('Word 格式处理 V4；原 DOCX 不覆盖；图片及对象不主动修改。\n'
                   +f'标准银行表格：{len(banks)}，11列目标总宽：{available:.3f}cm（页面约束）\n'
                   +'统一列宽：\n'+'\n'.join(f'{k}: {tw2cm(v):.3f} cm' for k,v in zip(LABELS,widths))
                   +'\n处理统计：\n'+'\n'.join(f'{k}: {v}' for k,v in sorted(COUNT.items()))
                   +'\n警告：\n'+'\n'.join(WARN or ['无'])+'\n',encoding='utf-8')
    print('完成：',dest,'\n日志：',log)
    if len(banks)!=17:print('注意：识别数量不是预期17张，请核对输出。')
if __name__=='__main__':
    try:main()
    except Exception as e:
        print('处理失败（原文件未修改）：',e)
        sys.exit(1)