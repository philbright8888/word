# 标准11列表头前多余段落清理工具

**源码：** [word_cleanup_before_11_header_standalone.py](../../scripts/word_cleanup_before_11_header_standalone.py)

## 用途
针对识别出的标准11列表格，清理表格前符合脚本条件的多余段落。

## 适用场景
只想清理标准11列表格前的段落，不想执行全文排版、列宽调整或金额汇总行处理时使用。

## 不会做什么
- 不调整标准11列宽度。
- 不执行全文字体统一。
- 不应假定它能准确识别 DOCX/WPS 的真实分页；请查看程序跳过/处理统计。
- 不应将“遇到 Account Number 删除整页”等未在源码中明确实现的规则当作已有功能。

## 运行
需要 Python 3 与 python-docx：
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/word_cleanup_before_11_header_standalone.py
```
按提示选择当前目录 DOCX。

## 验收
确认被清理的是目标表格前的多余段落，银行信息、图片、交易内容和列宽均保持。分页定位不确定时应保留原件并人工检查。
