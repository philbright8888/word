# 银行对账单 Word 自动排版工具

本仓库整理银行对账单 Word/WPS 文档的自动排版与辅助处理脚本，主要运行环境为 macOS，部分脚本也可在安装 python-docx 的 Windows 环境运行。

## 目录

- `scripts/`：Python 工具
- `macros/`：Word VBA 宏
- `docs/`：项目规则与使用说明

## 当前独立工具：表头前插入 3 行 × 4 列无边框区域

文件：`scripts/insert_3x4_before_bank_header_no_borders_mac.py`

功能：
- 在识别到标准银行交易 11 列表头的表格前，按条件插入 3 行 × 4 个逻辑单元格。
- 新增单元格无边框。
- 不删除原内容；不主动修改原11列宽、交易数据或图片。
- 当前目录列出 DOCX，输入编号选择文件。
- 原件不覆盖；输出文件名为 `原文件名_插入3行4列.docx`。
- 已有输出文件时停止，避免覆盖。

安装依赖：

```bash
python3 -m pip install python-docx
```

运行：

```bash
python3 scripts/insert_3x4_before_bank_header_no_borders_mac.py
```

## 标准 11 列

1. No.
2. Bk.D.
3. Val.D.
4. Type
5. Vou.
6. Vou.No./Trans.No./Details
7. DebitAmount
8. CreditAmount
9. Balance
10. ReferenceNo.
11. Notes

## 重要安全原则

- 不将真实银行对账单、账号、交易金额、测试 DOCX 或包含交易数据的扫描文件提交到仓库。
- 处理程序应输出到新文件，不覆盖原文件。
- 对含复杂合并单元格、非标准网格或分页不确定的文档，应记录并跳过无法安全处理的区域。
- 自动排版后仍需在 Microsoft Word 或 WPS 中核对实际分页和边框显示。

## 项目状态

本仓库优先保存可复用的程序源码和规则说明。不同版本的格式处理脚本应在文件名中标明版本；不要把模拟测试结果当作真实文档的最终版面验证。
