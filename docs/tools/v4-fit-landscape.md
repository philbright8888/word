# V4 横向页面适配版本

**源码：** [word_table_format_mac_v4_fit_landscape.py](../../scripts/versions/word_table_format_mac_v4_fit_landscape.py)

## 用途
保存当时横向页面与表格适配逻辑。

## 适用范围
这是独立历史版本，用于复现原有行为或与其他版本比较。先确认源码内的表格结构假设与当前文档一致。

## 不会做什么
不会自动调用其他版本，也不能假定包含后续版本的全部修复。不要把该脚本当作通用的银行对账单处理程序。

## 运行
依赖与输入选择方式以脚本内说明和运行提示为准。通常在包含 DOCX 的目录运行：
```bash
python3 -m pip install python-docx
python3 /path/to/scripts/versions/word_table_format_mac_v4_fit_landscape.py
```

## 输出与验收
确认输出文件名及原文件保护方式。运行后在 Word/WPS 检查列宽、分页、交易金额、银行信息和图片位置。

## 版本注意
此页记录历史用途，不代表在当前所有 DOCX 上通过实测。先用副本测试。
