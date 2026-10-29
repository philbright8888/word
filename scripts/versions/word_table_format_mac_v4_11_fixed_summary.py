#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 银行表格排版 V4.11：四组带金额汇总固定制表位且尽量单行；保留V4.10其他功能。
依赖: python-docx。可选 Pillow（提高宽度估算精度）。
原件绝不修改；输出 原文件名_格式处理.docx；如果存在则拒绝覆盖。
"""
import re, sys, math, copy
from pathlib import Path
from collections import Counter
try:
    from docx import Document
    from docx.table import _Cell
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
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
    # 只接受“11 个单元格且按顺序至少 7 个表头匹配”的行；其余必须核验位置。
    # 适配 PDF/WPS 转 DOCX 后各页表头存在空格、换行、文字变体的情况。
    best_score=-1
    bestrow=None
    for ix,row in enumerate(table.rows[:min(18,len(table.rows))]):
        cc=cells(row)
        if len(cc)!=11:
            continue
        hits=[]
        for j,c in enumerate(cc):
            t=norm(text(c))
            k=LOOKUP.get(t)
            if k is None and t:
                # 允许英文表头中多出的重复前后缀，但限制为原始已知表头的包含关系。
                matches=[i for i,variants in enumerate(RAW)
                         if any(len(norm(v))>=8 and norm(v) in t for v in variants)]
                if len(set(matches))==1:k=matches[0]
            hits.append(k==j)
        score=sum(hits)
        anchors=sum(hits[j] for j in (0,1,2,5,6,7,8,9,10))
        if score>=7 and anchors>=5 and score>best_score:
            best_score=score
            bestrow=ix
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
    # Physical page width; do not deduct Word's old left/right margins.
    return min((tw2cm(s.page_width.twips) for s in doc.sections), default=29.7)

def use_full_page_width(doc):
    # User requested to ignore old page margins. Horizontal margins become zero.
    # Keep paper size, landscape orientation, and vertical margins unchanged.
    for s in doc.sections:
        s.left_margin=Pt(0)
        s.right_margin=Pt(0)
    COUNT['sections_horizontal_margins_zeroed']=len(doc.sections)

def safe_total_width(doc,banks):
    paper = page_width_cm(doc)
    guards = 0.0
    for _, table, hi, mapping in banks:
        _, old = grid_info(table)
        first, last = mapping[0][0], mapping[-1][1]
        guards = max(guards, tw2cm(sum(old[:first]) + sum(old[last:])))
    # 4 mm buffer on each edge, including all auxiliary grid strips.
    safety = 0.8
    width = paper - guards - safety
    if width < 15:
        raise RuntimeError(f'纸张宽度 {paper:.3f}cm，扣除辅助网格和两侧安全空间后剩余 {width:.3f}cm，已停止。')
    WARN.append(f'左右页边距设为0；物理纸张宽{paper:.3f}cm，辅助网格合计最多{guards:.3f}cm，两侧共保留{safety:.3f}cm。')
    return width, paper, guards

def optimize(doc,banks):
    # Use the physical page width, not the former printable area.
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
    # Both elastic columns may wrap. Details should be slightly wider than Notes.
    # Search allocations in which Details exceeds Notes by 0.35-1.20 cm.
    if remaining < 4.8:
        raise RuntimeError(f'Details + Notes 只剩 {remaining:.2f}cm，无法安全分配。')
    best=None
    for j in range(86):
        delta=0.35 + j*0.01
        details=(remaining+delta)/2
        notes=remaining-details
        if notes < 2.0 or details < 2.35:continue
        cost=0
        for d,n,ty in records:
            cost+=max(wrap_line_count(d,details),wrap_line_count(n,notes),min(2,wrap_line_count(ty,fixed[3])))
        # Prefer a difference near 0.65 cm if estimated rows tie.
        candidate=(cost,abs(delta-0.65),details,notes)
        if best is None or candidate<best:best=candidate
    if best is None:raise RuntimeError('无法在纸张内分配 Details 和 Notes 的最小宽度。')
    _,_,d,n=best
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
    jc=child(pr,'w:jc');jc.set(qn('w:val'),'center')

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
    # Synchronize table preferred width with full underlying grid (guards included).
    pr=table._tbl.tblPr
    tw=child(pr,'w:tblW');tw.set(qn('w:w'),str(sum(new)));tw.set(qn('w:type'),'dxa')
    align_table_to_text_area(table)
    COUNT['uniform_width_tables']+=1

# 四组动态金额：用全文件统一的制表位定位，禁止用多空格撑宽。
SUMMARY_KEYS = ['Debit Total', 'Credit Total', 'This Page Balance', 'Closing Balance of This Reconciliation']
_SUMMARY_NUM = r'[-+]?\(?\d[\d,]*(?:\.\d+)?\)?'

def parse_amount_summary(s):
    raw = re.sub(r'[\u00a0\u3000\t\r\n]+', ' ', s)
    found=[]
    pos=0
    for key in SUMMARY_KEYS:
        pat=re.compile(r'\s*'+r'\s+'.join(re.escape(k) for k in key.split())+r'\s*('+_SUMMARY_NUM+r')',re.I)
        m=pat.search(raw,pos)
        if m is None or raw[pos:m.start()].strip():return None
        # Retain the exact captured numeric string; only normalize the fixed label.
        found.append((key,m.group(1)))
        pos=m.end()
    if raw[pos:].strip():return None
    return found

def summary_group_text(key,amount):
    # Nonbreaking spaces prevent the label and its amount being broken across lines.
    return key.replace(' ','\u00a0')+'\u00a0'+amount

def summary_positions(doc, width_cm):
    summaries=[]
    for p in doc.paragraphs:
        if not unsafe(p):
            match=parse_amount_summary(p.text)
            if match:summaries.append(match)
    if not summaries:return None
    # Max group display demand across every page, so the 4 start points are identical.
    demands=[max(measure_cm(summary_group_text(m[i][0],m[i][1])) for m in summaries)+.28 for i in range(4)]
    # Word tabs need enough space for the previous group's text before the next tab.
    needed=sum(demands)
    if needed>width_cm-.12:
        WARN.append(f'汇总行四组内容在8磅下估算需要{needed:.2f}cm，标准表宽{width_cm:.2f}cm；不强行换行或缩小字号，可能超出。')
        COUNT['summary_width_warning']+=1
    gap=max(0,(width_cm-needed)/3)
    return [0,demands[0]+gap,demands[0]+gap+demands[1]+gap,demands[0]+gap+demands[1]+gap+demands[2]+gap]

def compact_amount_summary(s):
    found=parse_amount_summary(s)
    return None if found is None else '\t'.join(summary_group_text(*x) for x in found)

def format_summary_paragraph(p, positions=None, total_cm=None, paper_cm=None):
    if unsafe(p):return False
    found=parse_amount_summary(p.text)
    if found is None:return False
    # Preserve all four amounts exactly, rebuild only this plain-text paragraph.
    for r in list(p.runs):p._p.remove(r._r)
    p.add_run('\t'.join(summary_group_text(*x) for x in found))
    format_para(p,WD_ALIGN_PARAGRAPH.LEFT)
    if positions is not None and total_cm is not None and paper_cm is not None:
        side=max(0,(paper_cm-total_cm)/2)
        f=p.paragraph_format
        f.left_indent=Cm(side)
        f.right_indent=Cm(side)
        f.first_line_indent=Pt(0)
        f.tab_stops.clear_all()
        for x in positions[1:]: f.tab_stops.add_tab_stop(Cm(x),WD_TAB_ALIGNMENT.LEFT)
        # Explicitly prevent full-line paragraph from splitting at spaces inside each group.
        COUNT['summary_fixed_tab_rows']+=1
    else:
        COUNT['summary_in_cell_compacted_only']+=1
    COUNT['amount_summary_paragraphs_compacted']+=1
    return True

def compact_amount_summaries(doc, total_cm, paper_cm):
    positions=summary_positions(doc,total_cm)
    for p in doc.paragraphs:
        if positions is not None:format_summary_paragraph(p,positions,total_cm,paper_cm)
    # Summary in a cell is not full page width: preserve cell structure and only normalize.
    for t in doc.tables:
        for row in t.rows:
            for c in cells(row):
                if len(c.paragraphs)==1 and format_summary_paragraph(c.paragraphs[0]):
                    COUNT['amount_summary_table_cells_compacted']+=1
    if positions is not None:
        WARN.append('四组汇总起始位置（距标准表左边框）：'+', '.join(f'{x:.2f}cm' for x in positions))

def content_driven_info_grid(table, total_tw):
    # Constraints: nonzero grid, all row spans remain untouched; assign width in proportion
    # to measured text demand rather than stretching the old grid proportions.
    cols, old=grid_info(table)
    n=len(cols)
    if not n or not old or any(w<=0 for w in old):raise ValueError('银行信息表网格无效')
    min_tw=max(8,cm2tw(.12))
    if min_tw*n >= total_tw:raise ValueError('空间不足，无法为信息表分配列宽')
    demands=[.12]*n
    for row in table.rows:
        for c,a,b in row_cells_with_grid(row):
            if a<0 or b>n or a>=b:raise ValueError('单元格跨度异常')
            # Require one-line display, including minimal cell padding.
            need=max(.16, measure_cm(content_for_measure(c))+.20)
            current=sum(demands[a:b])
            if need>current:
                extra=(need-current)/(b-a)
                for k in range(a,b):demands[k]+=extra
    # Preserve minimum width and distribute leftover according to content demand.
    weights=[max(.12,x) for x in demands]
    remaining=total_tw-min_tw*n
    weight_total=sum(weights)
    out=[min_tw+int(remaining*w/weight_total) for w in weights]
    out[-1]+=total_tw-sum(out)
    if min(out)<=0:raise ValueError('计算得到无效列宽')
    # Report genuine impossibility; avoid claiming a guarantee.
    for row in table.rows:
        for c,a,b in row_cells_with_grid(row):
            real=tw2cm(sum(out[a:b]))
            need=measure_cm(content_for_measure(c))+.20
            if need>real+.02:
                COUNT['bank_info_cells_possible_wrap']+=1
    return out

def align_bank_info_both_edges(doc, banks):
    """让独立银行信息表与邻近标准交易表的外侧左右边界重合。

    注意：不能仅设置 tblInd 为标准表的半宽偏移；需要同步总宽、网格、单元格宽。
    保留所有原有单元格、文字和合并关系。
    """
    if not banks:
        return
    standard_idx = {index for index, _, _, _ in banks}
    refs = []
    for index, table, _, _ in banks:
        _, widths = grid_info(table)
        refs.append((index, sum(widths)))

    for index, table in enumerate(doc.tables, 1):
        if index in standard_idx:
            continue
        sample = ' '.join(text(c) for row in list(table.rows)[:min(6, len(table.rows))]
                          for c in cells(row))
        key = norm(sample)
        markers = ('accountno', 'accountname', 'bankname',
                   'accounttype', 'previouspagebalance', 'currency')
        if sum(m in key for m in markers) < 2:
            continue

        ref_index, ref_total = min(refs, key=lambda x: abs(x[0]-index))
        cols, old = grid_info(table)
        if not cols or not old or any(w <= 0 for w in old):
            WARN.append(f'银行信息表{index}网格无效，跳过左右边界对齐')
            continue
        old_total = sum(old)
        try:
            new = content_driven_info_grid(table, ref_total)
        except Exception as exc:
            WARN.append(f'银行信息表{index}内容宽度分配失败，跳过：{exc}')
            continue
        if min(new) <= 0:
            WARN.append(f'银行信息表{index}网格重算存在无效宽度，跳过')
            continue
        # 修改前先检查每行的跨度，避免生成部分更新。
        mapped = []
        try:
            for row in table.rows:
                for c, aa, bb in row_cells_with_grid(row):
                    if aa < 0 or bb > len(new) or aa >= bb:
                        raise ValueError('单元格网格范围无效')
                    mapped.append((c, aa, bb))
        except Exception as exc:
            WARN.append(f'银行信息表{index}结构检查失败，跳过：{exc}')
            continue

        for col, width in zip(cols, new):
            col.set(qn('w:w'), str(width))
        for c, aa, bb in mapped:
            set_tc_width(c, sum(new[aa:bb]))
        table.autofit = False
        pr = table._tbl.tblPr
        tw = child(pr, 'w:tblW')
        tw.set(qn('w:w'), str(sum(new)))
        tw.set(qn('w:type'), 'dxa')
        # 与标准表完全相同的居中规则，不能用独立的左偏移近似。
        ind = child(pr, 'w:tblInd')
        ind.set(qn('w:w'), '0')
        ind.set(qn('w:type'), 'dxa')
        jc = child(pr, 'w:jc')
        jc.set(qn('w:val'), 'center')
        COUNT['bank_info_tables_both_edges_aligned'] += 1
        WARN.append(f'银行信息表{index}与标准表{ref_index}左右边框对齐：'
                    f'原宽{tw2cm(old_total):.3f}cm → {tw2cm(ref_total):.3f}cm，'
                    '两表均居中，内部列宽按最长内容分配；保留单元格和合并关系')

def process_table(table,index,header_idx):
    for ri,row in enumerate(table.rows):
        cc=cells(row)
        standard=header_idx is not None and ri>=header_idx and len(cc)==11
        for ci,c in enumerate(cc):
            is_header=standard and ri==header_idx
            h=LABELS[ci] if is_header else None
            if is_header:COUNT['standard_headers_targeted']+=1
            join_cell(c,h)
            # Set internal left and right cell margins to zero on non-standard cells.
            if not standard:set_cell_lr_margin_zero(c)
            align=(WD_ALIGN_PARAGRAPH.CENTER if is_header else ALIGN[ci]) if standard else WD_ALIGN_PARAGRAPH.LEFT
            for p in c.paragraphs:format_para(p,align)

# 严格删除三类没有实际业务数据的重复标题。
# 不按“包含关键词”模糊删除：四个字段必须对应四个独立的非空单元格，
# 段落必须仅包含四个汇总标题，不能有金额等额外内容。
def match_key(s):
    return re.sub(r"\s+", "", s or "").casefold().replace("\u00a0", "")

TO_PATTERN = re.compile(r"to(?:\(?(?:\d{8}|yyyymmdd)\)?)reconciliationperiod", re.I)
FROM_PATTERN = re.compile(r"from(?:\(?(?:\d{8}|yyyymmdd)\)?)page\d+of\d+", re.I)
SUMMARY_PATTERN = re.compile(
    r"debittotalcredittotalcurrentpagebalancebalanceattheendoftheperiod", re.I)

def title_row_kind(row):
    # 必须是 4 个非空、无图片/字段的独立单元格，允许额外的空单元格。
    cc=cells(row)
    if any(not plain(c) for c in cc):return None
    vals=[match_key(text(c)) for c in cc if match_key(text(c))]
    if len(vals)!=4:return None
    if vals[:3]==['currency','accounttype','previouspagebalance'] and TO_PATTERN.fullmatch(vals[3]):
        return 'currency_header'
    if vals[:3]==['accountno.','accountname','bankname'] and FROM_PATTERN.fullmatch(vals[3]):
        return 'account_header'
    return None

def remove_duplicate_table_rows(doc):
    for table in doc.tables:
        for row in list(table.rows):
            kind=title_row_kind(row)
            if kind:
                row._tr.getparent().remove(row._tr)
                COUNT['deleted_'+kind+'_rows']+=1

def remove_summary_title_paragraphs(doc):
    body=doc.element.body
    for p in list(body):
        if p.tag!=qn('w:p'):continue
        if p.xpath('.//w:drawing | .//w:pict | .//w:object | .//w:br | .//w:sectPr | .//w:fldChar | .//w:instrText | .//w:fldSimple | .//w:hyperlink | .//w:sdt | .//w:bookmarkStart | .//w:bookmarkEnd'):
            continue
        txt=''.join((x.text or '') for x in p.xpath('.//w:t'))
        if SUMMARY_PATTERN.fullmatch(match_key(txt)):
            body.remove(p)
            COUNT['deleted_summary_title_paragraphs']+=1

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
    # 先删除重复标题行，避免常规空白清理和表头格式化改变其原始文本。
    remove_duplicate_table_rows(doc)
    remove_summary_title_paragraphs(doc)
    print('已删除重复 Currency 标题行：',COUNT['deleted_currency_header_rows'])
    print('已删除重复 Account 标题行：',COUNT['deleted_account_header_rows'])
    print('已删除无金额汇总标题段落：',COUNT['deleted_summary_title_paragraphs'])
    banks=[]
    for i,t in enumerate(doc.tables,1):
        hi=identify(t)
        if hi is not None:
            mapping=map_ranges(t,hi)
            if mapping is None:WARN.append(f'表格{i}表头网格异常，不参与统一列宽')
            else:banks.append((i,t,hi,mapping))
    print(f'识别到 {len(banks)} 张11列银行表格，其他表格仅做文字清理。')
    print('识别的表格编号及表头所在行：'+', '.join(f'{i}号(第{hi+1}行)' for i,_,hi,_ in banks))
    if len(banks)!=17:WARN.append(f'本文件仅识别到 {len(banks)} 张银行表格（预期17张），请核对处理日志。')
    if not banks:print('没有识别到标准11列表格，为防止误处理已停止。');return
    use_full_page_width(doc)
    widths,available=optimize(doc,banks)
    compact_amount_summaries(doc, available, page_width_cm(doc))
    print('已固定制表位排版的四组金额汇总行：',COUNT['summary_fixed_tab_rows'])
    print(f'标准11列目标总宽：{available:.3f}cm（左右页边距0；Details 略宽于 Notes）')
    print(f'物理页面宽度：{page_width_cm(doc):.3f}cm；两侧预留0.4cm及辅助网格；表格居中')
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
    # 使用表格在文档中的固定序号，不依赖 Word 表格 wrapper 的 Python 对象 id。
    bank_headers={index:hi for index,_,hi,_ in banks}
    for index,t in enumerate(doc.tables,1):
        process_table(t,index,bank_headers.get(index))
    for index,t,hi,_ in banks:
        adjust_grid(t,hi,widths)
        COUNT['standard_bank_tables_processed']+=1
    align_bank_info_both_edges(doc,banks)
    clean_outside(doc)
    expected=11*len(banks)
    if COUNT['standard_headers_targeted']!=expected:
        raise RuntimeError(f'标准表头应处理 {expected} 个单元格，实际 {COUNT["standard_headers_targeted"]}，为避免错误未保存文件')
    # Atomic create: exclusive output path, never overwrite input/output.
    with open(dest,'xb') as f:doc.save(f)
    log=dest.with_suffix('.log.txt')
    log.write_text('Word 格式处理 V4.11；带金额四组汇总固定制表位，原金额保留；沿用 V4.7 删除规则，银行信息表左右对齐并按内容分配内部列宽；带金额汇总行压缩空格；原件不覆盖。\n'
                   +f'标准银行表格：{len(banks)}，11列目标总宽：{available:.3f}cm（物理纸张约束，左右页边距0，居中）\n'
                   +'统一列宽：\n'+'\n'.join(f'{k}: {tw2cm(v):.3f} cm' for k,v in zip(LABELS,widths))
                   +'\n处理统计：\n'+'\n'.join(f'{k}: {v}' for k,v in sorted(COUNT.items()))
                   +'\n警告：\n'+'\n'.join(WARN or ['无'])+'\n',encoding='utf-8')
    print('已对齐独立银行信息表：',COUNT['bank_info_tables_both_edges_aligned'])
    print('完成：',dest,'\n日志：',log)
    if len(banks)!=17:print('注意：识别数量不是预期17张，请核对输出。')
if __name__=='__main__':
    try:main()
    except Exception as e:
        print('处理失败（原文件未修改）：',e)
        sys.exit(1)