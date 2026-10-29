# Word 格式处理（不调整列宽，手动选择文件）

**源码：** [word_format_no_width_mac_select.py](../../scripts/word_format_no_width_mac_select.py)

## 用途
对选定 DOCX 执行脚本所定义的文字/段落格式处理，并避免主动调整表格列宽。

## 不会做什么
不应当作 V4.16 或 V4.11 的完全等价版本；各版本的清理规则和图片处理方式可能不同。

## 运行
```bash
python3 /path/to/scripts/word_format_no_width_mac_select.py
```
按程序提示选择文件。若提示缺少 python-docx，安装：`python3 -m pip install python-docx`。

## 验收
确认输出文件名、文字对齐、表格宽度及图片位置。先在副本上测试，不要覆盖原件。
