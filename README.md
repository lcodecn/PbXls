# PbXls Library v2.3

**PbXls** - PureBasic Excel xlsx/xlsm 操作库

- **作者**: lcode.cn
- **版本**: 2.3
- **许可证**: Apache 2.0
- **编译器**: PureBasic 6.40 (Windows - x86)

***

## 简介

PbXls 是一个纯 PureBasic 实现的 Excel 文件操作库，无需安装 Microsoft Office 或任何第三方依赖，即可创建和读取 Excel xlsx/xlsm 文件。

该库基于 Python 的 openpyxl 项目编写，使用 PureBasic内置的 XML 和 Packer（ZIP压缩）库实现。

## 主要功能

- **创建Excel文件**: 从零创建符合 Office Open XML 标准的 xlsx/xlsm 文件
- **读取Excel文件**: 解析现有 Excel 文件内容（部分实现）
- **单元格操作**: 读写字符串、数值、公式、布尔值、日期等类型数据
- **多工作表**: 支持在同一工作簿中创建和管理多个工作表
- **单元格样式**: 支持字体、填充、边框、对齐、数字格式等样式设置
- **合并单元格**: 支持单元格合并和取消合并
- **行列设置**: 支持设置列宽和行高
- **共享字符串表**: 自动优化字符串存储，使用共享字符串表减少文件大小

## 系统要求

- 本项目在PureBasic 6.40 （Windows x86）中编译通过，其他环境请自行测试。

## 快速开始
具体可参考开发文档：docs\PbXls_Help.html
### 创建Excel工作簿

```purebasic
XIncludeFile "PbXls.pb"

; 创建新工作簿
wbId = PbXls_CreateWorkbook()

; 获取当前工作表
wsId = PbXls_GetSheetByIndex(wbId, 0)

; 写入单元格数据
PbXls_SetCell(wsId, 1, 1, "Hello PbXls!")
PbXls_SetCell(wsId, 1, 2, "123.45")
PbXls_SetCell(wsId, 2, 1, "行2")
PbXls_SetCell(wsId, 2, 2, "20")

; 设置列宽
PbXls_SetColumnWidth(wsId, 1, 20.0)

; 合并单元格
PbXls_MergeCells(wsId, "A1:B1")

; 保存文件
PbXls_SaveWorkbook(wbId, "output.xlsx")

; 清理资源
PbXls_Free()
```

### 读取Excel工作簿

```purebasic
XIncludeFile "PbXls.pb"

; 加载现有工作簿
wbId = PbXls_LoadWorkbook("input.xlsx")

; 获取工作表
wsId = PbXls_GetSheetByIndex(wbId, 0)

; 清理资源
PbXls_Free()
```

## API 文档

### 工作簿操作

| 函数                                       | 说明                    |
| ---------------------------------------- | --------------------- |
| `PbXls_CreateWorkbook()`                 | 创建新的工作簿，返回工作簿ID       |
| `PbXls_LoadWorkbook(filename.s)`         | 加载现有的Excel工作簿，返回工作簿ID |
| `PbXls_SaveWorkbook(wbId.i, filename.s)` | 保存工作簿到指定文件路径          |
| `PbXls_GetSheetCount(wbId.i)`            | 获取工作簿中的工作表数量          |
| `PbXls_GetSheetByIndex(wbId.i, index.i)` | 根据索引获取工作表，返回工作表指针     |

### 工作表操作

| 函数                                             | 说明                  |
| ---------------------------------------------- | ------------------- |
| `PbXls_AddWorksheet(wbId.i, title.s)`          | 在工作簿中添加新工作表，返回工作表ID |
| `PbXls_DeleteWorksheet(wbId.i, index.i)`       | 删除指定索引的工作表          |
| `PbXls_GetSheetTitle(wsId.i)`                  | 获取工作表名称             |
| `PbXls_SetColumnWidth(wsId.i, col.i, width.f)` | 设置指定列的宽度            |
| `PbXls_SetRowHeight(wsId.i, row.i, height.f)`  | 设置指定行的高度            |
| `PbXls_MergeCells(wsId.i, rangeString.s)`      | 合并指定范围的单元格          |
| `PbXls_UnmergeCells(wsId.i, rangeString.s)`    | 取消合并指定范围的单元格        |
| `PbXls_AppendRow(wsId.i, List values.s())`     | 在工作表末尾追加一行数据        |

### 单元格操作

| 函数                                                      | 说明          |
| ------------------------------------------------------- | ----------- |
| `PbXls_SetCell(wsId.i, row.i, col.i, value.s)`          | 设置指定单元格的值   |
| `PbXls_SetCellFormula(wsId.i, row.i, col.i, formula.s)` | 设置单元格公式     |
| `PbXls_SetCellType(wsId.i, row.i, col.i, dataType.i)`   | 设置单元格数据类型类型 |

### 工具函数

| 函数                                         | 说明                    |
| ------------------------------------------ | --------------------- |
| `PbXls_GetColumnLetter(colNum.i)`          | 将列号转换为列字母（如 1 -> "A"） |
| `PbXls_ColumnIndexFromString(colLetter.s)` | 将列字母转换为列号（如 "A" -> 1） |
| `PbXls_EscapeXML(str.s)`                   | 转义XML特殊字符             |
| `PbXls_GetCurrentDateTime()`               | 获取当前日期时间字符串           |

## 文件结构

PbXls.pb 文件按照功能模块分为以下分区：

| 分区     | 内容                                 |
| ------ | ---------------------------------- |
| 分区1    | 常量定义（Excel规范、文件路径、XML命名空间、MIME类型等） |
| 分区2    | 枚举定义（数据类型、工作表状态、边框、对齐、填充等）         |
| 分区3    | 结构体定义和全局数据存储                       |
| 分区4    | 工具函数（坐标转换、字符串处理、日期时间、XML/ZIP辅助）    |
| 分区5    | XML常量模块                            |
| 分区6    | 样式模块（字体、填充、边框、对齐、数字格式）             |
| 分区7    | 单元格模块                              |
| 分区8    | 工作表模块                              |
| 分区9    | 工作簿模块                              |
| 分区10   | XML写入器（生成Excel文件各部分XML）            |
| 分区11   | XML读取器（预留）                         |
| 分区12   | 高级功能（预留）                           |
| 分区13   | 公共API                              |
| 分区14   | 初始化和清理                             |

## 版本历史

### v2.3 (2026-04-20)

- \[重构] 项目重命名为PbXls，原PureXL更名为PbXls
- \[修复] 修复UTF-8编码缓冲区溢出导致FreeMemory崩溃的问题
- \[修复] 修复字符串类型单元格数据未写入XML的问题
- \[修复] 修复Map访问语法错误（PureBasic Map正确访问方式）
- \[修复] 修复List参数声明语法
- \[新增] 添加完整的测试代码（10个测试项）
- \[新增] 添加详细的代码注释
- \[新增] 生成HTML帮助文档
- \[新增] 添加README.md文档

### v2.2 (2026-04-14)

- \[优化] 优化XML生成性能，减少内存占用
- \[优化] 优化共享字符串表构建算法，提升大数据量处理速度
- \[修复] 修复XML节点属性设置时的前缀错误
- \[修复] 修复工作簿关系文件（workbook.xml.rels）生成逻辑
- \[新增] 添加单元格内联字符串（inlineStr）支持

### v2.1 (2026-04-04)

- \[新增] 单元格样式模块（字体、填充、边框、对齐、数字格式）
- \[新增] 字体样式设置（字体名称、大小、颜色、粗体、斜体、下划线）
- \[新增] 填充样式设置（图案填充、前景色、背景色）
- \[新增] 边框样式设置（上、下、左、右边框样式和颜色）
- \[新增] 对齐样式设置（水平对齐、垂直对齐、文本换行）
- \[新增] 数字格式设置（内置数字格式、自定义数字格式）
- \[新增] 样式表XML生成（styles.xml）

### v2.0 (2026-03-25)

- \[新增] 公式单元格支持（SetCellFormula函数）
- \[新增] 布尔类型单元格支持（true/false）
- \[新增] 日期时间类型单元格支持
- \[新增] 错误类型单元格支持
- \[新增] 单元格数据类型自动推断功能
- \[新增] 单元格类型枚举定义
- \[优化] 重构数据结构设计，使用全局Map/List替代结构体内嵌套
- \[修复] 修复公式单元格XML写入逻辑

### v1.9 (2026-03-15)

- \[新增] 工作簿元数据支持（core.xml、app.xml）
- \[新增] 内容类型文件生成（\[Content\_Types].xml）
- \[新增] 根关系文件生成（\_rels/.rels）
- \[新增] 工作簿关系文件生成（xl/\_rels/workbook.xml.rels）
- \[新增] 文档属性设置（创建者、标题、描述、修改时间等）
- \[优化] 完善Office Open XML标准兼容性

### v1.8 (2026-03-05)

- \[新增] 多工作表支持（AddWorksheet函数）
- \[新增] 工作表删除功能（DeleteWorksheet函数）
- \[新增] 按索引获取工作表（GetSheetByIndex函数）
- \[新增] 获取工作表数量功能（GetSheetCount函数）
- \[新增] 获取工作表名称功能（GetSheetTitle函数）
- \[新增] 工作表关系管理
- \[优化] 工作表ID管理系统

### v1.7 (2026-02-23)

- \[新增] 工作表XML生成器（worksheet.xml）
- \[新增] 工作簿XML生成器（workbook.xml）
- \[新增] 共享字符串表XML生成器（sharedStrings.xml）
- \[新增] XML辅助函数（节点创建、属性设置、文本设置等）
- \[新增] XML保存为字符串功能
- \[新增] XML命名空间常量定义
- \[优化] XML生成代码模块化

### v1.6 (2026-02-13)

- \[新增] 共享字符串表支持（Shared Strings Table）
- \[新增] 字符串去重优化，减少文件大小
- \[新增] 字符串索引映射功能
- \[新增] 共享字符串计数统计
- \[优化] 字符串存储方式，使用全局List替代结构体成员

### v1.5 (2026-02-03)

- \[新增] 合并单元格功能（MergeCells函数）
- \[新增] 取消合并单元格功能（UnmergeCells函数）
- \[新增] 合并单元格XML节点生成（mergeCells）
- \[新增] 合并单元格范围解析
- \[优化] 工作表数据结构，支持合并单元格存储

### v1.4 (2026-01-24)

- \[新增] 设置列宽功能（SetColumnWidth函数）
- \[新增] 设置行高功能（SetRowHeight函数）
- \[新增] 列宽XML节点生成（cols/col）
- \[新增] 行高XML属性生成（row ht属性）
- \[新增] 行列尺寸数据存储（全局Map）

### v1.3 (2026-01-14)

- \[新增] 数值类型单元格支持
- \[新增] 数值自动识别（整数、浮点数）
- \[新增] 单元格数据类型枚举（字符串、数值等）
- \[新增] 数值单元格XML写入逻辑（v节点）
- \[优化] 单元格值存储方式

### v1.2 (2026-01-04)

- \[新增] 批量追加行数据功能（AppendRow函数）
- \[新增] 工作表当前行跟踪
- \[新增] 工作表最大行/列自动更新
- \[新增] 单元格数据批量写入优化
- \[优化] 单元格访问性能，使用Map快速查找

### v1.1 (2025-12-25)

- \[新增] 单元格写入功能（SetCell函数）
- \[新增] 字符串类型单元格支持
- \[新增] 单元格数据结构定义
- \[新增] 工作表数据结构定义
- \[新增] XML特殊字符转义功能（EscapeXML函数）
- \[新增] 列号与列字母互转功能（GetColumnLetter、ColumnIndexFromString）

### v1.0 (2025-12-15)

- \[新增] 初始项目创建，参考openpyxl 编写
- \[新增] 工作簿创建功能（CreateWorkbook函数）
- \[新增] 默认工作表自动创建
- \[新增] Excel规范常量定义（最大行列数等）
- \[新增] 文件路径常量定义
- \[新增] MIME类型常量定义
- \[新增] XML命名空间常量定义
- \[新增] 内置数字格式常量定义
- \[新增] 单元格类型枚举定义
- \[新增] 工作表状态枚举定义
- \[新增] ZIP打包支持（UseZipPacker、CreatePack等）
- \[新增] 初始化和清理函数（Init、Free）

## 许可证

本库采用 Apache 2.0 许可证。

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

本库参考的openpyxl 项目采用MIT许可证。

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

## 致谢

- 感谢 openpyxl 项目提供了优秀的参考实现
- 感谢 PureBasic QQ群的支持

