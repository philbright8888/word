# 表头前插入3行×3列无边框区域

**源码：** [insert_3x3_before_bank_header_no_borders_mac.py](../../scripts/insert_3x3_before_bank_header_no_borders_mac.py)

## 用途
在满足脚本识别条件的标准11列表格前插入3行×3个逻辑单元格，新增单元格不显示边框。

## 不会做什么
不删除原有内容、不统一字体、不修改交易数据，也不代替3行×4列版本。脚本识别条件不一定适用于所有合并表格。

## 运行
依赖 Python 3 和 python-docx：
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/insert_3x3_before_bank_header_no_borders_mac.py
```
在含 DOCX 的目录运行并按编号选择。

## 验收
检查新增区域行列数及无边框效果；确认原表头边框、交易内容、列宽、图片和分页未意外改变。
