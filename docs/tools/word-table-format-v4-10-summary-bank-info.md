# Word 格式处理 V4.10（汇总行与银行信息版本）

**源码：** [word_table_format_mac_v4_10_summary_bank_info.py](../../scripts/versions/word_table_format_mac_v4_10_summary_bank_info.py)

## 用途
保存 V4.10 的独立历史版本，包含当时的银行信息与金额汇总行排版实现。

## 运行
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/versions/word_table_format_mac_v4_10_summary_bank_info.py
```

## 注意
这是历史版本，不能默认具备 V4.11 的所有修正。适合对比或复现 V4.10 行为；新文档应先检查其规则是否匹配。

## 验收
核对金额未被改写、银行信息表边界和列宽、图片及分页。原件不覆盖。
