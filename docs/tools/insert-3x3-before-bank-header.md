# 表头前插入3行×3列区域（保留边框版本）

**源码：** [insert_3x3_before_bank_header_mac.py](../../scripts/insert_3x3_before_bank_header_mac.py)

## 用途
这是一个独立的3行×3列插入版本，与无边框版分开保存。它用于需要插入区域但希望保留该版本原有边框设置的场景。

## 不会做什么
不执行全文格式统一、银行信息表对齐、汇总行排版或内容清理。

## 运行
依赖 Python 3 和 python-docx：
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/insert_3x3_before_bank_header_mac.py
```
在包含 DOCX 的目录运行并按编号选择。

## 验收
运行前阅读脚本行为；运行后检查新增区域的边框、表头结构和原有内容。若需要新增区域完全无边框，请使用独立的 no-borders 版本。
