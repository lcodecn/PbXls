# PbXls Library v2.6



PbXls - PureBasic Excel xlsx/xlsm Library

- Author: lcode.cn
- Version: 2.6
- License: Apache 2.0
- Compiler: PureBasic 6.40 (Windows - x86)

***

## Introduction

PbXls is a PureBasic library for manipulating Excel files, requiring no Microsoft Office or any third-party dependencies to create and read Excel xlsx/xlsm files.

This library is based on the Python openpyxl project, implemented using PureBasic's built-in XML and Packer (ZIP compression) libraries.

## Features

- Create Excel files: Create xlsx/xlsm files compliant with Office Open XML standard from scratch
- Read Excel files: Parse existing Excel file contents (partially implemented)
- Cell operations: Read and write string, numeric, formula, boolean, date and other data types
- Multiple worksheets: Support creating and managing multiple worksheets in the same workbook
- Cell styles: Support font, fill, border, alignment, number format and other style settings
- Merge cells: Support cell merging and unmerging
- Row/Column settings: Support setting column widths and row heights
- Column insert/delete: Support inserting and deleting columns at specified positions, automatically updating cells, column widths, and merged cell ranges
- Data validation: Support dropdown lists, integer validation, decimal validation and other validation rules
- Conditional formatting: Support cell value comparison, formula conditions, color scales and other conditional formatting rules
- Shared strings table: Automatically optimize string storage using shared strings table to reduce file size

## Requirements

- This project compiles with PureBasic 6.40 (Windows x86). Other environments have not been tested.

## Quick Start

For detailed documentation, see: docs\PbXls_Help_en.html

### Creating an Excel Workbook

```purebasic
XIncludeFile "PbXls.pb"

; Create a new workbook
wbId = PbXls_CreateWorkbook()

; Get the active worksheet
wsId = PbXls_GetSheetByIndex(wbId, 0)

; Write cell data
PbXls_SetCell(wsId, 1, 1, "Hello PbXls!")
PbXls_SetCell(wsId, 1, 2, "123.45")
PbXls_SetCell(wsId, 2, 1, "Row 2")
PbXls_SetCell(wsId, 2, 2, "20")

; Set column width
PbXls_SetColumnWidth(wsId, 1, 20.0)

; Merge cells
PbXls_MergeCells(wsId, "A1:B1")

; Save file
PbXls_SaveWorkbook(wbId, "output.xlsx")

; Clean up resources
PbXls_Free()
```

### Reading an Excel Workbook

```purebasic
XIncludeFile "PbXls.pb"

; Load an existing workbook
wbId = PbXls_LoadWorkbook("input.xlsx")

; Get worksheet
wsId = PbXls_GetSheetByIndex(wbId, 0)

; Clean up resources
PbXls_Free()
```

## API Reference

### Workbook Operations

| Function | Description |
| ---------------------------------------- | --------------------- |
| `PbXls_CreateWorkbook()` | Create a new workbook, returns workbook ID |
| `PbXls_LoadWorkbook(filename.s)` | Load an existing Excel workbook, returns workbook ID |
| `PbXls_SaveWorkbook(wbId.i, filename.s)` | Save workbook to specified file path |
| `PbXls_GetSheetCount(wbId.i)` | Get the number of worksheets in the workbook |
| `PbXls_GetSheetByIndex(wbId.i, index.i)` | Get worksheet by index, returns worksheet pointer |

### Worksheet Operations

| Function | Description |
| ------------------------------------------------ | ------------------- |
| `PbXls_AddWorksheet(wbId.i, title.s)` | Add a new worksheet to the workbook, returns worksheet ID |
| `PbXls_DeleteWorksheet(wbId.i, index.i)` | Delete worksheet at specified index |
| `PbXls_GetSheetTitle(wsId.i)` | Get worksheet name |
| `PbXls_SetColumnWidth(wsId.i, col.i, width.f)` | Set the width of specified column |
| `PbXls_SetRowHeight(wsId.i, row.i, height.f)` | Set the height of specified row |
| `PbXls_MergeCells(wsId.i, rangeString.s)` | Merge cells in specified range |
| `PbXls_UnmergeCells(wsId.i, rangeString.s)` | Unmerge cells in specified range |
| `PbXls_AppendRow(wsId.i, List values.s())` | Append a row of data at the end of the worksheet |
| `PbXls_InsertColumns(wsId.i, colIdx.i, count.i)` | Insert columns at specified position |
| `PbXls_DeleteColumns(wsId.i, colIdx.i, count.i)` | Delete columns at specified position |

### Cell Operations

| Function | Description |
| ------------------------------------------------------- | ----------- |
| `PbXls_SetCell(wsId.i, row.i, col.i, value.s)` | Set the value of specified cell |
| `PbXls_SetCellFormula(wsId.i, row.i, col.i, formula.s)` | Set cell formula |
| `PbXls_SetCellType(wsId.i, row.i, col.i, dataType.i)` | Set cell data type |

### Utility Functions

| Function | Description |
| ------------------------------------------ | --------------------- |
| `PbXls_GetColumnLetter(colNum.i)` | Convert column number to column letter (e.g. 1 -> "A") |
| `PbXls_ColumnIndexFromString(colLetter.s)` | Convert column letter to column number (e.g. "A" -> 1) |
| `PbXls_EscapeXML(str.s)` | Escape XML special characters |
| `PbXls_GetCurrentDateTime()` | Get current date-time string |

### Data Validation

| Function | Description |
| ----------------------------------------------------------------------- | -------- |
| `PbXls_CreateDataValidation(type, sqref, formula1, formula2, operator)` | Create data validation rule |
| `PbXls_SetValidationPrompt(id, title, message)` | Set input prompt |
| `PbXls_SetValidationError(id, title, message)` | Set error prompt |

### Conditional Formatting

| Function | Description |
| --------------------------------------------------------------------------------------------------- | -------- |
| `PbXls_CreateConditionalFormat(type, sqref, formula1, formula2, operator)` | Create conditional formatting rule |
| `PbXls_SetConditionalFormatDxf(id, fontColor, fillColor, fontBold, fontItalic)` | Set differential style |
| `PbXls_SetConditionalFormatColorScale(id, minColor, midColor, maxColor, minType, midType, maxType)` | Set color scale parameters |

### Charts

| Function | Description |
| --------------------------------------------------------- | -------- |
| `PbXls_CreateChart(type, title, anchorRef)` | Create chart |
| `PbXls_AddChartSeries(chartId, name, values, categories)` | Add chart data series |

## File Structure

The PbXls.pb file is organized into the following sections by functionality:

| Section | Content |
| ----- | ---------------------------------- |
| Section 1 | Constants (Excel specs, file paths, XML namespaces, MIME types, etc.) |
| Section 2 | Enumerations (data types, worksheet states, borders, alignments, fills, etc.) |
| Section 3 | Structure definitions and global data storage |
| Section 4 | Utility functions (coordinate conversion, string processing, date-time, XML/ZIP helpers) |
| Section 5 | XML constants module |
| Section 6 | Style module (fonts, fills, borders, alignments, number formats) |
| Section 6.5 | Data validation module |
| Section 6.6 | Conditional formatting module |
| Section 6.7 | Chart module |
| Section 7 | Cell module |
| Section 8 | Worksheet module |
| Section 9 | Workbook module |
| Section 10 | XML writer (generates XML for each part of the Excel file) |
| Section 11 | XML reader |
| Section 12 | Advanced features (column insert/delete) |
| Section 13 | Public API |
| Section 14 | Initialization and cleanup |

## Version History

### v2.6 (2026-04-22)

- [New] Data validation feature (PbXls\_CreateDataValidation), supporting dropdown lists, integer validation, decimal validation, etc.
- [New] Conditional formatting feature (PbXls\_CreateConditionalFormat), supporting cell value comparison, formula conditions, color scales
- [New] Chart feature interface (PbXls\_CreateChart, PbXls\_AddChartSeries)
- [New] Complete test suite (FeaturesTest, etc.)

### v2.5 (2026-04-22)

- [New] Column insert feature (PbXls\_InsertColumns), supporting inserting multiple columns at specified position
- [New] Column delete feature (PbXls\_DeleteColumns), supporting deleting columns within specified range
- [New] Automatic update of cell positions, column widths, and merged cell ranges when inserting/deleting columns
- [Fix] PbXls\_Free missing cleanup of newly added global lists (data validation, conditional formatting, charts)
- [New] Complete test suite (ColumnTest, etc.)

### v2.4 (2026-04-21)

- [New] Convenient style API (PbXls\_SetCellStyleWS)
- [Fix] Bug where comments and code were on the same line during border XML writing (causing invalid XML nodes)
- [Fix] fillId offset issue (Excel reserves the first two fills, user fills start from index 2)
- [Optimize] All test files converted to UTF-8 BOM format for PureBasic compatibility
- [New] Complete test suite (StyleStepTest, FullStyleTest, etc.)

### v2.3 (2026-04-20)

- [Refactor] Project renamed to PbXls (formerly PureXL)
- [Fix] Fixed UTF-8 encoding buffer overflow causing FreeMemory crash
- [Fix] Fixed string type cell data not being written to XML
- [Fix] Fixed Map access syntax error (correct PureBasic Map access method)
- [Fix] Fixed List parameter declaration syntax
- [New] Added complete test code (10 test items)
- [New] Added detailed code comments
- [New] Generated HTML help documentation
- [New] Added README.md documentation

### v2.2 (2026-04-14)

- [Optimize] Optimized XML generation performance, reduced memory usage
- [Optimize] Optimized shared strings table construction algorithm, improved large data processing speed
- [Fix] Fixed prefix error when setting XML node attributes
- [Fix] Fixed workbook relationship file (workbook.xml.rels) generation logic
- [New] Added inline string (inlineStr) support for cells

### v2.1 (2026-04-04)

- [New] Cell style module (fonts, fills, borders, alignments, number formats)
- [New] Font style settings (font name, size, color, bold, italic, underline)
- [New] Fill style settings (pattern fill, foreground color, background color)
- [New] Border style settings (top, bottom, left, right border styles and colors)
- [New] Alignment style settings (horizontal alignment, vertical alignment, text wrapping)
- [New] Number format settings (built-in number formats, custom number formats)
- [New] Style table XML generation (styles.xml)

### v2.0 (2026-03-25)

- [New] Formula cell support (SetCellFormula function)
- [New] Boolean cell support (true/false)
- [New] Date-time cell support
- [New] Error cell support
- [New] Cell data type auto-inference
- [New] Cell type enumeration definitions
- [Optimize] Refactored data structure design, using global Map/List instead of nested structures
- [Fix] Fixed formula cell XML writing logic

### v1.9 (2026-03-15)

- [New] Workbook metadata support (core.xml, app.xml)
- [New] Content type file generation (\[Content\_Types].xml)
- [New] Root relationship file generation (\_rels/.rels)
- [New] Workbook relationship file generation (xl/\_rels/workbook.xml.rels)
- [New] Document property settings (creator, title, description, modification time, etc.)
- [Optimize] Improved Office Open XML standard compatibility

### v1.8 (2026-03-05)

- [New] Multiple worksheet support (AddWorksheet function)
- [New] Worksheet deletion (DeleteWorksheet function)
- [New] Get worksheet by index (GetSheetByIndex function)
- [New] Get worksheet count (GetSheetCount function)
- [New] Get worksheet name (GetSheetTitle function)
- [New] Worksheet relationship management
- [Optimize] Worksheet ID management system

### v1.7 (2026-02-23)

- [New] Worksheet XML generator (worksheet.xml)
- [New] Workbook XML generator (workbook.xml)
- [New] Shared strings table XML generator (sharedStrings.xml)
- [New] XML helper functions (node creation, attribute setting, text setting, etc.)
- [New] XML save to string function
- [New] XML namespace constant definitions
- [Optimize] Modularized XML generation code

### v1.6 (2026-02-13)

- [New] Shared strings table support (Shared Strings Table)
- [New] String deduplication optimization to reduce file size
- [New] String index mapping function
- [New] Shared string count statistics
- [Optimize] String storage method, using global List instead of structure members

### v1.5 (2026-02-03)

- [New] Merge cells feature (MergeCells function)
- [New] Unmerge cells feature (UnmergeCells function)
- [New] Merged cells XML node generation (mergeCells)
- [New] Merged cell range parsing
- [Optimize] Worksheet data structure to support merged cell storage

### v1.4 (2026-01-24)

- [New] Set column width feature (SetColumnWidth function)
- [New] Set row height feature (SetRowHeight function)
- [New] Column width XML node generation (cols/col)
- [New] Row height XML attribute generation (row ht attribute)
- [New] Row/column dimension data storage (global Map)

### v1.3 (2026-01-14)

- [New] Numeric cell support
- [New] Auto-detection of numbers (integers, floating-point)
- [New] Cell data type enumeration (string, numeric, etc.)
- [New] Numeric cell XML writing logic (v node)
- [Optimize] Cell value storage method

### v1.2 (2026-01-04)

- [New] Batch append row data feature (AppendRow function)
- [New] Worksheet current row tracking
- [New] Worksheet max row/column auto-update
- [New] Cell data batch writing optimization
- [Optimize] Cell access performance using Map for fast lookup

### v1.1 (2025-12-25)

- [New] Cell write feature (SetCell function)
- [New] String type cell support
- [New] Cell data structure definition
- [New] Worksheet data structure definition
- [New] XML special character escaping (EscapeXML function)
- [New] Column number and letter conversion (GetColumnLetter, ColumnIndexFromString)

### v1.0 (2025-12-15)

- [New] Initial project creation, based on openpyxl
- [New] Workbook creation feature (CreateWorkbook function)
- [New] Default worksheet auto-creation
- [New] Excel specification constant definitions (max rows/columns, etc.)
- [New] File path constant definitions
- [New] MIME type constant definitions
- [New] XML namespace constant definitions
- [New] Built-in number format constant definitions
- [New] Cell type enumeration definitions
- [New] Worksheet state enumeration definitions
- [New] ZIP packaging support (UseZipPacker, CreatePack, etc.)
- [New] Initialization and cleanup functions (Init, Free)

## License

This library is licensed under the Apache 2.0 License.

```
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```

The openpyxl project referenced by this library is licensed under the MIT License.

```
This software is under the MIT Licence
======================================

Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

```

## Acknowledgements

- Thanks to the openpyxl project for providing an excellent reference implementation
- Thanks to the PureBasic QQ community for their support

## Donate

If PbXls is helpful to you, donations are welcome to support the continued development and maintenance of this project. Thank you for your generosity!

- **PayPal**: [https://www.paypal.me/lcodecn](https://www.paypal.me/lcodecn)
- **WeChat**: #Pay:lcodecn(Business_lcodecn)/openlib/003
