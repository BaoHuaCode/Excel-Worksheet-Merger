#Excel Worksheet Merger | Excel 多工作表合并工具

[English Version Below | 中文版见下文]

---

# English Description
A high-efficiency Python script designed to automate the process of merging multiple Excel files (.xlsx) into a single summary sheet. 

### Key Features
* **Natural Sorting**: Uses `natsort` to ensure files are merged in logical order (1, 2, 10) instead of (1, 10, 2).
* **Smart Row Offsets**: Precise management of `current_row` to ensure seamless data stitching without empty rows or overwriting.
* **Modern File Handling**: Built with `pathlib` for clean, cross-platform path management.

###  Tech Stack
* Python 3.x, `openpyxl`, `pathlib`, `natsort`

---

#中文说明
这是一个高效的办公自动化脚本，旨在自动将文件夹内的所有 Excel 文件汇总成一个总表。

### 核心功能
* **自然排序**：利用 `natsort` 库，完美解决文件名 `1, 2, 10` 的排序乱序问题。
* **智能行偏移**：精准管理行指针，确保多表数据无缝拼接，避免空行或覆盖。
* **自动化路径**：使用 `pathlib` 库，自动适配不同系统的路径格式。

###  技术栈
* Python 3.x, `openpyxl`, `pathlib`, `natsort`
