# Word 格式处理 V4.11（固定汇总行版本）

**源码：** [word_table_format_mac_v4_11_fixed_summary.py](../../scripts/versions/word_table_format_mac_v4_11_fixed_summary.py)

## 用途
V4.11 固定汇总行版本，专注四组动态金额汇总项的固定位置和尽量单行显示，同时保留该版本原有银行表格排版逻辑。

## 运行
需要 Python 3 与 python-docx；Pillow 为可选依赖。
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/versions/word_table_format_mac_v4_11_fixed_summary.py
```

## 不会做什么
不应与 V4.11 23张表兼容版视为同一文件；两者是独立脚本，目标文档假设和规则可能不同。

## 验收
确认四组标题与动态金额原样保留、汇总行实际显示、银行信息表边界、列宽、分页及图片位置。
