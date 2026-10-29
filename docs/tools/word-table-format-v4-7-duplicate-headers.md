# Word 格式处理 V4.7（重复标题清理版本）

**源码：** [word_table_format_mac_v4_7_remove_duplicate_headers.py](../../scripts/versions/word_table_format_mac_v4_7_remove_duplicate_headers.py)

## 用途
保存 V4.7 中重复表头/标题清理的历史实现。

## 运行
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/versions/word_table_format_mac_v4_7_remove_duplicate_headers.py
```

## 注意
清理规则与后续版本可能不同。运行前确认哪些标题会被删除，尤其要保证带实际金额的汇总行不会被误删。

## 验收
对照原文逐项检查删除行，确认交易金额、图片、表格列宽和分页未意外变化。
