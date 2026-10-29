# Word 银行对账单独立工具集

本仓库由多个**独立工具**组成。每个功能有自己的 Python 文件和使用说明，单独运行、单独输出。**不要把所有功能合并成一个程序，也不要假设每个工具都会执行所有格式规则。**

## 独立工具目录

### 插入表格区域
| 工具 | 源码 | 使用说明 |
|---|---|---|
| 插入3行×4列无边框区域 | [Python](scripts/insert_3x4_before_bank_header_no_borders_mac.py) | [说明](docs/tools/insert-3x4-before-bank-header.md) |
| 插入3行×3列无边框区域 | [Python](scripts/insert_3x3_before_bank_header_no_borders_mac.py) | [说明](docs/tools/insert-3x3-before-bank-header-no-borders.md) |
| 插入3行×3列（原边框版本） | [Python](scripts/insert_3x3_before_bank_header_mac.py) | [说明](docs/tools/insert-3x3-before-bank-header.md) |

### 格式统一与清理
| 工具 | 源码 | 使用说明 |
|---|---|---|
| V4.16 内容清理/格式统一，不改列宽 | [Python](scripts/clean_word_v4_16_no_width.py) | [说明](docs/tools/clean-word-v4-16-no-width.md) |
| V4.11 23张表兼容格式处理 | [Python](scripts/word_table_format_mac_v4_11_23tables.py) | [说明](docs/tools/word-table-format-v4-11-23tables.md) |
| 不改列宽的文件选择格式工具 | [Python](scripts/word_format_no_width_mac_select.py) | [说明](docs/tools/word-format-no-width-mac-select.md) |
| 标准11列表头前段落清理 | [Python](scripts/word_cleanup_before_11_header_standalone.py) | [说明](docs/tools/cleanup-before-11-header.md) |
| 银行表格行快速删除（谨慎使用） | [Python](scripts/delete_bank_rows_fast.py) | [说明](docs/tools/delete-bank-rows-fast.md) |
| DOCX 表格单元格扫描 | [Python](scripts/scan_docx_table_cells.py) | [说明](docs/tools/scan-docx-table-cells.md) |
| 特定13张表3 cm 调整 | [Python](scripts/adjust_bank_13_tables_3cm.py) | [说明](docs/tools/adjust-bank-13-tables-3cm.md) |

### 历史格式版本（独立保存，不互相覆盖）
| 版本 | 源码 | 使用说明 |
|---|---|---|
| V4.11 固定汇总行 | [Python](scripts/versions/word_table_format_mac_v4_11_fixed_summary.py) | [说明](docs/tools/word-table-format-v4-11-fixed-summary.md) |
| V4.10 汇总行与银行信息 | [Python](scripts/versions/word_table_format_mac_v4_10_summary_bank_info.py) | [说明](docs/tools/word-table-format-v4-10-summary-bank-info.md) |
| V4.9 银行信息左右边界 | [Python](scripts/versions/word_table_format_mac_v4_9_bank_both_edges.py) | [说明](docs/tools/word-table-format-v4-9-bank-edges.md) |
| V4.8 银行信息左边界 | [Python](scripts/versions/word_table_format_mac_v4_8_bank_left_align.py) | [说明](docs/tools/word-table-format-v4-8-bank-left.md) |
| V4.7 重复标题清理 | [Python](scripts/versions/word_table_format_mac_v4_7_remove_duplicate_headers.py) | [说明](docs/tools/word-table-format-v4-7-duplicate-headers.md) |

## 目录约定

- `scripts/`：可单独运行的 Python 工具。
- `scripts/versions/`：保留的历史版本；版本之间不自动联动。
- `docs/tools/`：每个工具独立说明页。
- [通用使用规范](docs/USAGE.md)
- [工具说明模板](docs/tools/TOOL_TEMPLATE.md)
- [项目规则](docs/PROJECT_SPEC.md)

## 使用前须知

1. 先打开对应工具说明，确认适用场景和“不会做什么”。
2. 每个工具只执行自身功能；运行插入工具不会顺便统一字体或删除段落。
3. 有些工具针对特定表格数量或结构，只能用于匹配的文档。
4. 复杂合并单元格、分页或图片定位无法安全判断时，应跳过并检查日志。
5. 模拟测试不能代替在 Microsoft Word/WPS 中检查真实输出。

## 隐私与原件保护

- 不提交真实银行对账单、账号、交易金额、测试 DOCX 或包含交易数据的扫描报告。
- 原文件不得覆盖或删除；输出新文件，已存在同名输出时应停止。
- 只把源码真实存在的工具列为可用，不根据讨论内容伪造代码。
