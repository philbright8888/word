# 银行对账单格式处理 V4.11（23张表兼容版）

**源码：** [word_table_format_mac_v4_11_23tables.py](../../scripts/word_table_format_mac_v4_11_23tables.py)

## 用途
针对识别到的银行对账单标准11列表格进行格式处理，并处理四组动态金额汇总项目的固定位置/单行排版。

## 适用场景
目标文档结构与本版本假设相符，尤其是扫描到多张标准11列表格的文档。

## 主要规则
- 识别标准11列表头及常见重复表头。
- 按脚本中的规则处理表格格式、银行信息及汇总行。
- 汇总金额按原文动态提取，不应写死为示例数字。
- 原件不覆盖；输出名为“原文件名_格式处理.docx”，同名输出存在时拒绝覆盖。

## 不会做什么
这不是通用所有 DOCX 的格式化器；不应对结构差异明显、未扫描确认的文档直接批量运行。

## 环境与运行
依赖 python-docx；Pillow 可选，用于提高宽度估算精度。
```bash
python3 -m pip install python-docx
# 可选：python3 -m pip install Pillow
python3 /path/to/scripts/word_table_format_mac_v4_11_23tables.py
```

## 验收
核对表格数量、11列映射、银行信息表左右边界、四组汇总金额是否原样保留，以及图片、分页、实际列宽。模拟测试不等于真实 WPS 验收。
