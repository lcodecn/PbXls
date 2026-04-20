; ***************************************************************************************
; PbXls Library - Excel xlsx/xlsm 操作库
; 版本: 2.3
; 作者: lcode.cn
; 许可证: Apache 2.0
;
; 说明: 无依赖操作Excel xlsx/xlsm文件的PureBASIC库
;       使用PureBasic内置XML和Packer库
;       无需安装Microsoft Office或任何第三方依赖
;
; 主要功能:
;   - 创建和读取Excel xlsx/xlsm文件
;   - 读写单元格数据（字符串、数值、公式、布尔、日期等）
;   - 多工作表支持
;   - 单元格样式（字体、填充、边框、对齐、数字格式）
;   - 合并单元格
;   - 行列宽高设置
;   - 共享字符串表优化
;   - 完整的ZIP压缩支持
;
; 文件结构:
;   分区1: 常量定义（Excel规范、文件路径、XML命名空间、MIME类型等）
;   分区2: 枚举定义（数据类型、工作表状态、边框、对齐、填充等）
;   分区3: 结构体定义和全局数据存储
;   分区4: 工具函数（坐标转换、字符串处理、日期时间、XML/ZIP辅助）
;   分区5: XML常量模块
;   分区6: 样式模块（字体、填充、边框、对齐、数字格式）
;   分区7: 单元格模块
;   分区8: 工作表模块
;   分区9: 工作簿模块
;   分区10: XML写入器（生成Excel文件各部分XML）
;   分区11: XML读取器（预留）
;   分区12: 高级功能（预留）
;   分区13: 公共API
;   分区14: 初始化和清理
;   分区15: 测试代码
; ***************************************************************************************

; 启用ZIP打包库（PureBASIC的XML库是内置的，无需初始化）
UseZipPacker()

; ***************************************************************************************
; 分区1: 常量定义
; ***************************************************************************************

; 1.1 Excel规范常量（定义Excel的最大行列数等限制）
#PbXls_MaxRow = 1048576
#PbXls_MaxColumn = 16384
#PbXls_MinRow = 1
#PbXls_MinColumn = 1

; 1.2 文件路径常量（定义xlsx文件内部各XML文件的路径）
#PbXls_ARCContentTypes$ = "[Content_Types].xml"
#PbXls_ARCRootRels$ = "_rels/.rels"
#PbXls_ARCWorkbookRels$ = "xl/_rels/workbook.xml.rels"
#PbXls_ARCCore$ = "docProps/core.xml"
#PbXls_ARCApp$ = "docProps/app.xml"
#PbXls_ARCCustom$ = "docProps/custom.xml"
#PbXls_ARCWorkbook$ = "xl/workbook.xml"
#PbXls_ARCStyles$ = "xl/styles.xml"
#PbXls_ARCTheme$ = "xl/theme/theme1.xml"
#PbXls_ARCSharedStrings$ = "xl/sharedStrings.xml"

#PbXls_PackageProps$ = "docProps"
#PbXls_PackageXL$ = "xl"
#PbXls_PackageRels$ = "_rels"
#PbXls_PackageTheme$ = "xl/theme"
#PbXls_PackageWorksheets$ = "xl/worksheets"
#PbXls_PackageChartsheets$ = "xl/chartsheets"
#PbXls_PackageDrawings$ = "xl/drawings"
#PbXls_PackageCharts$ = "xl/charts"
#PbXls_PackageImages$ = "xl/media"

; 1.3 XML命名空间常量（定义Excel XML文件中使用的各种XML命名空间URI）
;    用于正确生成符合Office Open XML标准的XML文件
#PbXls_XML_NS$ = "http://www.w3.org/XML/1998/namespace"
#PbXls_DCORE_NS$ = "http://purl.org/dc/elements/1.1/"
#PbXls_DCTERMS_NS$ = "http://purl.org/dc/terms/"
#PbXls_DOC_NS$ = "http://schemas.openxmlformats.org/officeDocument/2006/"
#PbXls_REL_NS$ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#PbXls_VTYPES_NS$ = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
#PbXls_PKG_NS$ = "http://schemas.openxmlformats.org/package/2006/"
#PbXls_PKG_REL_NS$ = "http://schemas.openxmlformats.org/package/2006/relationships"
#PbXls_COREPROPS_NS$ = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
#PbXls_CONTYPES_NS$ = "http://schemas.openxmlformats.org/package/2006/content-types"
#PbXls_XSI_NS$ = "http://www.w3.org/2001/XMLSchema-instance"
#PbXls_SHEET_MAIN_NS$ = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
#PbXls_DRAWING_NS$ = "http://schemas.openxmlformats.org/drawingml/2006/main"
#PbXls_SHEET_DRAWING_NS$ = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
#PbXls_CHART_NS$ = "http://schemas.openxmlformats.org/drawingml/2006/chart"

; 1.4 MIME类型常量（定义xlsx文件中各内容类型的MIME标识）
#PbXls_WORKBOOK_MACRO$ = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
#PbXls_WORKBOOK$ = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
#PbXls_WORKBOOK_TEMPLATE$ = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
#PbXls_WORKBOOK_MACRO_TEMPLATE$ = "application/vnd.ms-excel.template.macroEnabled.main+xml"
#PbXls_SPREADSHEET$ = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
#PbXls_SHARED_STRINGS$ = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
#PbXls_STYLES_TYPE$ = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
#PbXls_COMMENTS_TYPE$ = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
#PbXls_DRAWING_TYPE$ = "application/vnd.openxmlformats-officedocument.drawing+xml"
#PbXls_CHART_TYPE$ = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
#PbXls_THEME_TYPE$ = "application/vnd.openxmlformats-officedocument.theme+xml"

#PbXls_XLSX$ = "xlsx"
#PbXls_XLSM$ = "xlsm"
#PbXls_XLTX$ = "xltx"
#PbXls_XLTM$ = "xltm"

; 1.5 单元格类型常量（定义单元格数据类型枚举值）
;    包括字符串、数值、公式、布尔、日期、空值、内联字符串、错误等类型
#PbXls_TypeString = 0
#PbXls_TypeNumeric = 1
#PbXls_TypeBoolean = 2
#PbXls_TypeFormula = 3
#PbXls_TypeDate = 4
#PbXls_TypeNull = 5
#PbXls_TypeInline = 6
#PbXls_TypeError = 7

; 1.6 内置数字格式常量（定义Excel内置的数字格式ID）
;    包括通用格式、数值格式、百分比、科学计数法、分数、日期时间等
#PbXls_NumFmtGeneral = 0
#PbXls_NumFmt0 = 1
#PbXls_NumFmt0_00 = 2
#PbXls_NumFmt_0 = 3
#PbXls_NumFmt_0_00 = 4
#PbXls_NumFmtPercent = 9
#PbXls_NumFmtPercent_00 = 10
#PbXls_NumFmtScientific = 11
#PbXls_NumFmtFraction = 12
#PbXls_NumFmtDate = 14
#PbXls_NumFmtDate2 = 15
#PbXls_NumFmtDate3 = 16
#PbXls_NumFmtDate4 = 17
#PbXls_NumFmtTime1 = 18
#PbXls_NumFmtTime2 = 19
#PbXls_NumFmtDateTime = 22
#PbXls_NumFmt_0_000 = 38
#PbXls_NumFmt_0_0000 = 39
#PbXls_NumFmtMaxBuiltin = 163

; ***************************************************************************************
; 分区2: 枚举定义
;    定义库中使用的各种枚举类型，包括数据类型、工作表状态、边框样式、
;    对齐方式、填充模式和错误类型等
; ***************************************************************************************

Enumeration
  #PbXls_DataTypeString
  #PbXls_DataTypeNumeric
  #PbXls_DataTypeBoolean
  #PbXls_DataTypeFormula
  #PbXls_DataTypeDate
  #PbXls_DataTypeNull
  #PbXls_DataTypeInline
  #PbXls_DataTypeError
EndEnumeration

Enumeration
  #PbXls_SheetVisible
  #PbXls_SheetHidden
  #PbXls_SheetVeryHidden
EndEnumeration

Enumeration
  #PbXls_BorderNone
  #PbXls_BorderThin
  #PbXls_BorderMedium
  #PbXls_BorderThick
  #PbXls_BorderDouble
  #PbXls_BorderDashed
  #PbXls_BorderDotted
  #PbXls_BorderHair
  #PbXls_BorderDashDot
  #PbXls_BorderDashDotDot
  #PbXls_BorderSlantDashDot
  #PbXls_BorderMediumDashed
  #PbXls_BorderMediumDashDot
  #PbXls_BorderMediumDashDotDot
EndEnumeration

Enumeration
  #PbXls_AlignGeneral
  #PbXls_AlignLeft
  #PbXls_AlignCenter
  #PbXls_AlignRight
  #PbXls_AlignFill
  #PbXls_AlignJustify
  #PbXls_AlignCenterContinuous
EndEnumeration

Enumeration
  #PbXls_ValignTop
  #PbXls_ValignCenter
  #PbXls_ValignBottom
  #PbXls_ValignJustify
  #PbXls_ValignDistributed
EndEnumeration

Enumeration
  #PbXls_FillNone
  #PbXls_FillSolid
  #PbXls_FillGray125
  #PbXls_FillLightGray
  #PbXls_FillMediumGray
  #PbXls_FillDarkGray
  #PbXls_FillLightHorizontal
  #PbXls_FillLightVertical
  #PbXls_FillLightDown
  #PbXls_FillLightUp
  #PbXls_FillLightGrid
  #PbXls_FillLightTrellis
EndEnumeration

Enumeration
  #PbXls_ErrorNull
  #PbXls_ErrorDiv0
  #PbXls_ErrorValue
  #PbXls_ErrorRef
  #PbXls_ErrorName
  #PbXls_ErrorNum
  #PbXls_ErrorNA
EndEnumeration

; ***************************************************************************************
; 分区3: 结构体定义
; ***************************************************************************************

Structure PbXls_Cell
  row.i
  column.i
  value.s
  dataType.i
  styleId.i
  formula.s
  richText.s
  comment.s
  hyperlink.s
EndStructure

Structure PbXls_Worksheet
  id.i
  title.s
  sheetState.s
  currentRow.i
  maxRow.i
  maxColumn.i
  parent.i
  printArea.s
  printTitles.s
  orientation.s
  paperSize.i
  autoFilter.s
  freezePanes.s
EndStructure

Structure PbXls_Workbook
  id.i
  activeSheetIndex.i
  path.s
  fontsId.i
  fillsId.i
  bordersId.i
  cellStylesId.i
  title.s
  subject.s
  creator.s
  description.s
  created.s
  modified.s
  keywords.s
  category.s
  isReadOnly.b
  isWriteOnly.b
EndStructure

Structure PbXls_Font
  name.s
  size.f
  bold.b
  italic.b
  underline.b
  strike.b
  color.s
  vertAlign.s
  charset.i
  family.i
  scheme.s
EndStructure

Structure PbXls_Fill
  patternType.s
  fgColor.s
  bgColor.s
EndStructure

Structure PbXls_BorderItem
  style.s
  color.s
EndStructure

Structure PbXls_Border
  left.PbXls_BorderItem
  right.PbXls_BorderItem
  top.PbXls_BorderItem
  bottom.PbXls_BorderItem
  diagonal.PbXls_BorderItem
  diagonalDown.b
  diagonalUp.b
EndStructure

Structure PbXls_Alignment
  horizontal.s
  vertical.s
  wrapText.b
  shrinkToFit.b
  indent.i
  textRotation.i
EndStructure

Structure PbXls_CellStyle
  fontId.i
  fillId.i
  borderId.i
  numFmtId.i
  numFmt.s
  alignment.PbXls_Alignment
  hidden.b
  locked.b
  quotePrefix.b
EndStructure

Structure PbXls_Color
  rgb.s
  indexed.i
  auto.b
  theme.i
  tint.f
EndStructure

; 全局数据存储
Global NewMap PbXls_AllCells.PbXls_Cell()
Global NewMap PbXls_ColumnWidths.f()
Global NewMap PbXls_RowHeights.f()
Global NewMap PbXls_MergedCells.s()
Global NewMap PbXls_MergedCellCount.i()
Global NewList PbXls_Fonts.PbXls_Font()
Global NewList PbXls_Fills.PbXls_Fill()
Global NewList PbXls_Borders.PbXls_Border()
Global NewList PbXls_CellStyles.PbXls_CellStyle()
Global NewList PbXls_AllWorksheets.PbXls_Worksheet()
Global NewList PbXls_SharedStrings.s()
Global NewList PbXls_Workbooks.PbXls_Workbook()
Global NewMap PbXls_WorkbookSheetCount.i()
Global NewMap PbXls_WorkbookSharedStrings.i()

; ***************************************************************************************
; 分区4: 工具函数
; ***************************************************************************************

; GetColumnLetter - 列号转字母 (1->"A", 27->"AA")
Procedure.s PbXls_GetColumnLetter(colIdx.i)
  If colIdx < 1 Or colIdx > 18278
    ProcedureReturn ""
  EndIf
  
  Define result.s = ""
  Define remainder.i
  Define tempIdx.i = colIdx
  
  While tempIdx > 0
    tempIdx - 1
    remainder = tempIdx % 26
    result = Chr(65 + remainder) + result
    tempIdx / 26
  Wend
  
  ProcedureReturn result
EndProcedure

; ColumnIndexFromString - 字母转列号 ("A"->1, "AA"->27)
Procedure.i PbXls_ColumnIndexFromString(col.s)
  col = Trim(UCase(col))
  Define length.i = Len(col)
  If length < 1 Or length > 3
    ProcedureReturn 0
  EndIf
  
  Define result.i = 0
  Define i.i, charCode.i
  
  For i = 1 To length
    charCode = Asc(Mid(col, i, 1)) - 64
    If charCode < 1 Or charCode > 26
      ProcedureReturn 0
    EndIf
    result = result * 26 + charCode
  Next i
  
  ProcedureReturn result
EndProcedure

; CoordinateToTuple - 坐标转行列元组 ("A1"->(1,1))
Procedure.i PbXls_CoordinateToTuple(coord.s, *row.Integer, *col.Integer)
  coord = Trim(UCase(coord))
  Define length.i = Len(coord)
  Define colStr.s = ""
  Define rowStr.s = ""
  Define i.i
  
  For i = 1 To length
    Define ch.c = Asc(Mid(coord, i, 1))
    If (ch >= 65 And ch <= 90)
      colStr + Mid(coord, i, 1)
    ElseIf (ch >= 48 And ch <= 57)
      rowStr + Mid(coord, i, 1)
    Else
      ProcedureReturn #False
    EndIf
  Next i
  
  If colStr = "" Or rowStr = ""
    ProcedureReturn #False
  EndIf
  
  *row\i = Val(rowStr)
  *col\i = PbXls_ColumnIndexFromString(colStr)
  
  ProcedureReturn #True
EndProcedure

; RangeBoundaries - 范围解析 ("A1:D5"->(1,1,4,5))
Procedure.i PbXls_RangeBoundaries(rangeStr.s, *minCol.Integer, *minRow.Integer, *maxCol.Integer, *maxRow.Integer)
  rangeStr = Trim(UCase(rangeStr))
  Define colonPos.i = FindString(rangeStr, ":", 1)
  
  If colonPos = 0
    Define coord.s = rangeStr
    Define row.i, col.i
    If PbXls_CoordinateToTuple(coord, @row, @col) = #False
      ProcedureReturn #False
    EndIf
    *minCol\i = col
    *minRow\i = row
    *maxCol\i = col
    *maxRow\i = row
  Else
    Define startCoord.s = Mid(rangeStr, 1, colonPos - 1)
    Define endCoord.s = Mid(rangeStr, colonPos + 1)
    Define startRow.i, startCol.i, endRow.i, endCol.i
    If PbXls_CoordinateToTuple(startCoord, @startRow, @startCol) = #False
      ProcedureReturn #False
    EndIf
    If PbXls_CoordinateToTuple(endCoord, @endRow, @endCol) = #False
      ProcedureReturn #False
    EndIf
    *minCol\i = startCol
    *minRow\i = startRow
    *maxCol\i = endCol
    *maxRow\i = endRow
  EndIf
  
  ProcedureReturn #True
EndProcedure

; CoordinateFromRowCol - 行列转坐标 (1,1->"A1")
Procedure.s PbXls_CoordinateFromRowCol(row.i, col.i)
  ProcedureReturn PbXls_GetColumnLetter(col) + Str(row)
EndProcedure

; RangeString - 生成范围字符串
Procedure.s PbXls_RangeString(minRow.i, minCol.i, maxRow.i, maxCol.i)
  ProcedureReturn PbXls_CoordinateFromRowCol(minRow, minCol) + ":" + PbXls_CoordinateFromRowCol(maxRow, maxCol)
EndProcedure

; QuoteSheetName - 为包含特殊字符的工作表名添加引号
Procedure.s PbXls_QuoteSheetName(sheetName.s)
  If FindString(sheetName, " ", 1) Or FindString(sheetName, "'", 1)
    sheetName = ReplaceString(sheetName, "'", "''")
    ProcedureReturn "'" + sheetName + "'"
  EndIf
  ProcedureReturn sheetName
EndProcedure

; EscapeXML - 转义XML特殊字符
Procedure.s PbXls_EscapeXML(text.s)
  text = ReplaceString(text, "&", "&amp;")
  text = ReplaceString(text, "<", "&lt;")
  text = ReplaceString(text, ">", "&gt;")
  text = ReplaceString(text, ~"\"", "&quot;")
  text = ReplaceString(text, "'", "&apos;")
  ProcedureReturn text
EndProcedure

; UnescapeXML - 反转义XML特殊字符
Procedure.s PbXls_UnescapeXML(text.s)
  text = ReplaceString(text, "&amp;", "&")
  text = ReplaceString(text, "&lt;", "<")
  text = ReplaceString(text, "&gt;", ">")
  text = ReplaceString(text, "&quot;", ~"\"")
  text = ReplaceString(text, "&apos;", "'")
  ProcedureReturn text
EndProcedure

; IsNumeric - 检查是否为数字
Procedure.b PbXls_IsNumeric(text.s)
  text = Trim(text)
  If text = ""
    ProcedureReturn #False
  EndIf
  Define testVal.f = ValF(text)
  Define i.i
  Define hasDot.b = #False
  Define hasDigit.b = #False
  Define start.i = 1
  If Mid(text, 1, 1) = "+" Or Mid(text, 1, 1) = "-"
    start = 2
  EndIf
  For i = start To Len(text)
    Define ch.c = Asc(Mid(text, i, 1))
    If ch = 46
      If hasDot
        ProcedureReturn #False
      EndIf
      hasDot = #True
    ElseIf ch >= 48 And ch <= 57
      hasDigit = #True
    Else
      ProcedureReturn #False
    EndIf
  Next i
  ProcedureReturn hasDigit
EndProcedure

; IsBoolean - 检查是否为布尔值
Procedure.b PbXls_IsBoolean(text.s)
  text = Trim(LCase(text))
  If text = "true" Or text = "false" Or text = "0" Or text = "1"
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

; IsDate - 检查是否为日期格式
Procedure.b PbXls_IsDate(text.s)
  If Len(text) >= 8
    Define date.i = ParseDate("%yyyy-%mm-%dd", text)
    If date = 0
      date = ParseDate("%mm/%dd/%yyyy", text)
    EndIf
    If date = 0
      date = ParseDate("%dd-%mm-%yyyy", text)
    EndIf
    If date <> 0
      ProcedureReturn #True
    EndIf
  EndIf
  ProcedureReturn #False
EndProcedure

; IsFormula - 检查是否为公式
Procedure.b PbXls_IsFormula(text.s)
  text = Trim(text)
  If Len(text) > 1 And Mid(text, 1, 1) = "="
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

; DateToExcel - PureBASIC日期转Excel日期数字
Procedure.f PbXls_DateToExcel(pbDate.i)
  Define baseDate.i = ParseDate("%yyyy-%mm-%dd %hh:%nn:%ss", "1899-12-30 00:00:00")
  If baseDate = 0
    ProcedureReturn 0.0
  EndIf
  Define diff.i = pbDate - baseDate
  Define days.f = diff / 86400.0
  If days >= 60
    days + 1
  EndIf
  ProcedureReturn days
EndProcedure

; ExcelToDate - Excel日期数字转PureBASIC日期
Procedure.i PbXls_ExcelToDate(excelDate.f)
  Define baseDate.i = ParseDate("%yyyy-%mm-%dd %hh:%nn:%ss", "1899-12-30 00:00:00")
  If baseDate = 0
    ProcedureReturn 0
  EndIf
  Define days.f = excelDate
  If days >= 60
    days - 1
  EndIf
  Define diff.i = days * 86400
  Define resultDate.i = baseDate + diff
  ProcedureReturn resultDate
EndProcedure

; GetCurrentDateTime - 获取当前日期时间
Procedure.s PbXls_GetCurrentDateTime()
  Define now.i = Date()
  ProcedureReturn FormatDate("%yyyy-%mm-%ddT%hh:%nn:%ss", now)
EndProcedure

; XML辅助函数
Procedure.i PbXls_XMLCreateDocument()
  ProcedureReturn CreateXML(#PB_Any)
EndProcedure

Procedure.i PbXls_XMLAddNode(xmlId.i, parentNode.i, nodeName.s)
  ProcedureReturn CreateXMLNode(parentNode, nodeName)
EndProcedure

Procedure.b PbXls_XMLSetAttribute(node.i, attrName.s, attrValue.s)
  SetXMLAttribute(node, attrName, attrValue)
EndProcedure

Procedure.b PbXls_XMLSetText(node.i, text.s)
  SetXMLNodeText(node, text)
EndProcedure

Procedure.s PbXls_XMLGetAttribute(node.i, attrName.s)
  ProcedureReturn GetXMLAttribute(node, attrName)
EndProcedure

Procedure.s PbXls_XMLGetText(node.i)
  ProcedureReturn GetXMLNodeText(node)
EndProcedure

Procedure.s PbXls_XMLSaveToString(xmlId.i)
  Define tempFile.s = GetTemporaryDirectory() + "PbXls_temp_" + Str(Random(999999)) + ".xml"
  If SaveXML(xmlId, tempFile)
    Define file.i = ReadFile(#PB_Any, tempFile)
    If file
      Define content.s = ReadString(file, #PB_UTF8)
      CloseFile(file)
      DeleteFile(tempFile)
      ProcedureReturn content
    EndIf
    DeleteFile(tempFile)
  EndIf
  ProcedureReturn ""
EndProcedure

Procedure.b PbXls_XMLSaveToFile(xmlId.i, filename.s)
  ProcedureReturn SaveXML(xmlId, filename)
EndProcedure

Procedure.i PbXls_XMLParseString(xmlString.s)
  Define xmlId.i = ParseXML(#PB_Any, xmlString)
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_XMLParseFile(filename.s)
  Define xmlId.i = LoadXML(#PB_Any, filename)
  ProcedureReturn xmlId
EndProcedure

Procedure.b PbXls_XMLFree(xmlId.i)
  FreeXML(xmlId)
EndProcedure

Procedure.i PbXls_XMLGetRoot(xmlId.i)
  ProcedureReturn RootXMLNode(xmlId)
EndProcedure

Procedure.i PbXls_XMLGetChild(node.i)
  ProcedureReturn ChildXMLNode(node)
EndProcedure

Procedure.i PbXls_XMLGetNext(node.i)
  ProcedureReturn NextXMLNode(node)
EndProcedure

Procedure.s PbXls_XMLGetNodeName(node.i)
  ProcedureReturn GetXMLNodeName(node)
EndProcedure

; ZIP辅助函数
Procedure.i PbXls_ZIPCreate(filename.s)
  ProcedureReturn CreatePack(#PB_Any, filename)
EndProcedure

Procedure.i PbXls_ZIPOpen(filename.s)
  ProcedureReturn OpenPack(#PB_Any, filename)
EndProcedure

Procedure.b PbXls_ZIPAddFile(packId.i, sourceFile.s, archivePath.s)
  If AddPackFile(packId, sourceFile, archivePath)
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

Procedure.b PbXls_ZIPAddMemory(packId.i, *data, dataSize.i, archivePath.s)
  If AddPackMemory(packId, *data, dataSize, archivePath)
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

Procedure.b PbXls_ZIPClose(packId.i)
  ClosePack(packId)
EndProcedure

Procedure.b PbXls_ZIPExamine(packId.i)
  If ExaminePack(packId)
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

Procedure.b PbXls_ZIPNextEntry(packId.i)
  If NextPackEntry(packId)
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

Procedure.s PbXls_ZIPEntryName(packId.i)
  ProcedureReturn PackEntryName(packId)
EndProcedure

Procedure.i PbXls_ZIPEntrySize(packId.i)
  ProcedureReturn PackEntrySize(packId)
EndProcedure

Procedure.i PbXls_ZIPExtractToMemory(packId.i, *buffer, bufferSize.i)
  If UncompressPackMemory(packId, *buffer, bufferSize)
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

; ***************************************************************************************
; 分区5: XML常量模块
; ***************************************************************************************

Procedure.s PbXls_GetNamespaceMappings()
  ProcedureReturn ~"xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                  ~"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                  ~"xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                  ~"mc:Ignorable=\"x14ac xr xr2 xr3\" " +
                  ~"xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" " +
                  ~"xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" " +
                  ~"xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" " +
                  ~"xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\""
EndProcedure

Procedure.s PbXls_GetWorksheetNamespace()
  ProcedureReturn #PbXls_SHEET_MAIN_NS$
EndProcedure

Procedure.s PbXls_GetWorkbookNamespace()
  ProcedureReturn #PbXls_SHEET_MAIN_NS$
EndProcedure

Procedure.s PbXls_GetRelationshipNamespace()
  ProcedureReturn #PbXls_PKG_REL_NS$
EndProcedure

; ***************************************************************************************
; 分区6: 样式模块
; ***************************************************************************************

Procedure.i PbXls_CreateFont()
  Define fontId.i = ListSize(PbXls_Fonts())
  AddElement(PbXls_Fonts())
  PbXls_Fonts()\name = "Calibri"
  PbXls_Fonts()\size = 11.0
  PbXls_Fonts()\color = "000000"
  PbXls_Fonts()\bold = #False
  PbXls_Fonts()\italic = #False
  PbXls_Fonts()\underline = #False
  PbXls_Fonts()\strike = #False
  ProcedureReturn fontId
EndProcedure

Procedure.b PbXls_SetFont(fontId.i, name.s = "", size.f = -1, bold.b = -1, italic.b = -1, color.s = "")
  If fontId < 0 Or fontId >= ListSize(PbXls_Fonts())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_Fonts(), fontId)
  If name <> ""
    PbXls_Fonts()\name = name
  EndIf
  If size >= 0
    PbXls_Fonts()\size = size
  EndIf
  If bold >= 0
    PbXls_Fonts()\bold = bold
  EndIf
  If italic >= 0
    PbXls_Fonts()\italic = italic
  EndIf
  If color <> ""
    PbXls_Fonts()\color = color
  EndIf
  ProcedureReturn #True
EndProcedure

Procedure.i PbXls_CreateFill()
  Define fillId.i = ListSize(PbXls_Fills())
  AddElement(PbXls_Fills())
  PbXls_Fills()\patternType = "none"
  ProcedureReturn fillId
EndProcedure

Procedure.b PbXls_SetFill(fillId.i, patternType.s = "", fgColor.s = "", bgColor.s = "")
  If fillId < 0 Or fillId >= ListSize(PbXls_Fills())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_Fills(), fillId)
  If patternType <> ""
    PbXls_Fills()\patternType = patternType
  EndIf
  If fgColor <> ""
    PbXls_Fills()\fgColor = fgColor
  EndIf
  If bgColor <> ""
    PbXls_Fills()\bgColor = bgColor
  EndIf
  ProcedureReturn #True
EndProcedure

Procedure.i PbXls_CreateBorder()
  Define borderId.i = ListSize(PbXls_Borders())
  AddElement(PbXls_Borders())
  ProcedureReturn borderId
EndProcedure

Procedure.b PbXls_SetBorder(borderId.i, side.s, style.s = "thin", color.s = "000000")
  If borderId < 0 Or borderId >= ListSize(PbXls_Borders())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_Borders(), borderId)
  side = LCase(side)
  Select side
    Case "left"
      PbXls_Borders()\left\style = style
      PbXls_Borders()\left\color = color
    Case "right"
      PbXls_Borders()\right\style = style
      PbXls_Borders()\right\color = color
    Case "top"
      PbXls_Borders()\top\style = style
      PbXls_Borders()\top\color = color
    Case "bottom"
      PbXls_Borders()\bottom\style = style
      PbXls_Borders()\bottom\color = color
    Case "diagonal"
      PbXls_Borders()\diagonal\style = style
      PbXls_Borders()\diagonal\color = color
  EndSelect
  ProcedureReturn #True
EndProcedure

Procedure.i PbXls_CreateAlignment()
  Define styleId.i = ListSize(PbXls_CellStyles())
  AddElement(PbXls_CellStyles())
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  PbXls_CellStyles()\numFmtId = 0
  PbXls_CellStyles()\numFmt = "General"
  PbXls_CellStyles()\alignment\horizontal = "general"
  PbXls_CellStyles()\alignment\vertical = "bottom"
  PbXls_CellStyles()\alignment\wrapText = #False
  PbXls_CellStyles()\alignment\shrinkToFit = #False
  PbXls_CellStyles()\alignment\indent = 0
  PbXls_CellStyles()\alignment\textRotation = 0
  PbXls_CellStyles()\locked = #True
  ProcedureReturn styleId
EndProcedure

Procedure.b PbXls_SetAlignment(styleId.i, horizontal.s = "", vertical.s = "", wrapText.b = -1, indent.i = -1)
  If styleId < 0 Or styleId >= ListSize(PbXls_CellStyles())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_CellStyles(), styleId)
  If horizontal <> ""
    PbXls_CellStyles()\alignment\horizontal = horizontal
  EndIf
  If vertical <> ""
    PbXls_CellStyles()\alignment\vertical = vertical
  EndIf
  If wrapText >= 0
    PbXls_CellStyles()\alignment\wrapText = wrapText
  EndIf
  If indent >= 0
    PbXls_CellStyles()\alignment\indent = indent
  EndIf
  ProcedureReturn #True
EndProcedure

Procedure.s PbXls_GetBuiltinFormat(numFmtId.i)
  Select numFmtId
    Case 0: ProcedureReturn "General"
    Case 1: ProcedureReturn "0"
    Case 2: ProcedureReturn "0.00"
    Case 3: ProcedureReturn "#,##0"
    Case 4: ProcedureReturn "#,##0.00"
    Case 9: ProcedureReturn "0%"
    Case 10: ProcedureReturn "0.00%"
    Case 11: ProcedureReturn "0.00E+00"
    Case 12: ProcedureReturn "# ?/?"
    Case 13: ProcedureReturn "# ??/??"
    Case 14: ProcedureReturn "mm-dd-yy"
    Case 15: ProcedureReturn "d-mmm-yy"
    Case 16: ProcedureReturn "d-mmm"
    Case 17: ProcedureReturn "mmm-yy"
    Case 18: ProcedureReturn "h:mm AM/PM"
    Case 19: ProcedureReturn "h:mm:ss AM/PM"
    Case 20: ProcedureReturn "h:mm"
    Case 21: ProcedureReturn "h:mm:ss"
    Case 22: ProcedureReturn "m/d/yy h:mm"
    Case 37: ProcedureReturn "#,##0 ;(#,##0)"
    Case 38: ProcedureReturn "#,##0 ;[Red](#,##0)"
    Case 39: ProcedureReturn "#,##0.00;(#,##0.00)"
    Case 40: ProcedureReturn "#,##0.00;[Red](#,##0.00)"
    Case 44: ProcedureReturn ~"_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
    Case 45: ProcedureReturn "mm:ss"
    Case 46: ProcedureReturn "[h]:mm:ss"
    Case 47: ProcedureReturn "mmss.0"
    Case 48: ProcedureReturn "##0.0E+0"
    Case 49: ProcedureReturn "@"
    Default: ProcedureReturn "General"
  EndSelect
EndProcedure

Procedure.b PbXls_IsDateFormat(numFmtId.i)
  Select numFmtId
    Case 14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47
      ProcedureReturn #True
    Default:
      ProcedureReturn #False
  EndSelect
EndProcedure

Procedure.b PbXls_SetNumberFormat(styleId.i, numFmt.s)
  If styleId < 0 Or styleId >= ListSize(PbXls_CellStyles())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_CellStyles(), styleId)
  PbXls_CellStyles()\numFmt = numFmt
  ProcedureReturn #True
EndProcedure

Procedure.s PbXls_ParseColor(color.s)
  color = Trim(color)
  If Left(color, 1) = "#"
    ProcedureReturn Mid(color, 2)
  EndIf
  ProcedureReturn color
EndProcedure

Procedure.s PbXls_FormatColor(color.s)
  color = PbXls_ParseColor(color)
  If Len(color) = 6
    ProcedureReturn "FF" + color
  ElseIf Len(color) = 8
    ProcedureReturn color
  Else
    ProcedureReturn "FF000000"
  EndIf
EndProcedure

; ***************************************************************************************
; 分区7: 单元格模块
; ***************************************************************************************

Procedure PbXls_InitCell(*cell.PbXls_Cell, row.i, col.i)
  *cell\row = row
  *cell\column = col
  *cell\value = ""
  *cell\dataType = #PbXls_DataTypeNull
  *cell\styleId = 0
  *cell\formula = ""
  *cell\richText = ""
  *cell\comment = ""
  *cell\hyperlink = ""
EndProcedure

Procedure.b PbXls_SetCellValue(*cell.PbXls_Cell, value.s)
  If value = ""
    *cell\value = ""
    *cell\dataType = #PbXls_DataTypeNull
    ProcedureReturn #True
  EndIf
  If PbXls_IsFormula(value)
    *cell\value = value
    *cell\dataType = #PbXls_DataTypeFormula
    *cell\formula = Mid(value, 2)
    ProcedureReturn #True
  EndIf
  Define lowerValue.s = LCase(Trim(value))
  If lowerValue = "true"
    *cell\value = "1"
    *cell\dataType = #PbXls_DataTypeBoolean
    ProcedureReturn #True
  ElseIf lowerValue = "false"
    *cell\value = "0"
    *cell\dataType = #PbXls_DataTypeBoolean
    ProcedureReturn #True
  EndIf
  If PbXls_IsNumeric(value)
    *cell\value = value
    *cell\dataType = #PbXls_DataTypeNumeric
    ProcedureReturn #True
  EndIf
  *cell\value = value
  *cell\dataType = #PbXls_DataTypeString
  ProcedureReturn #True
EndProcedure

Procedure.s PbXls_GetCellValue(*cell.PbXls_Cell)
  ProcedureReturn *cell\value
EndProcedure

Procedure.i PbXls_GetCellDataType(*cell.PbXls_Cell)
  ProcedureReturn *cell\dataType
EndProcedure

Procedure.s PbXls_GetCellFormula(*cell.PbXls_Cell)
  If *cell\dataType = #PbXls_DataTypeFormula
    ProcedureReturn "=" + *cell\formula
  EndIf
  ProcedureReturn ""
EndProcedure

Procedure.b PbXls_SetCellFormula(*cell.PbXls_Cell, formula.s)
  If Left(formula, 1) = "="
    formula = Mid(formula, 2)
  EndIf
  *cell\formula = formula
  *cell\value = "=" + formula
  *cell\dataType = #PbXls_DataTypeFormula
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_SetCellStyle(*cell.PbXls_Cell, styleId.i)
  *cell\styleId = styleId
  ProcedureReturn #True
EndProcedure

Procedure.i PbXls_GetCellStyle(*cell.PbXls_Cell)
  ProcedureReturn *cell\styleId
EndProcedure

; ***************************************************************************************
; 分区8: 工作表模块
; ***************************************************************************************

Procedure.i PbXls_GetCell(*ws.PbXls_Worksheet, row.i, col.i)
  If row < 1 Or col < 1
    ProcedureReturn 0
  EndIf
  Define key.s = Str(*ws\id) + "_" + Str(row) + "," + Str(col)
  If FindMapElement(PbXls_AllCells(), key)
    ProcedureReturn @PbXls_AllCells()
  Else
    Define *newCell.PbXls_Cell = @PbXls_AllCells(key)
    PbXls_InitCell(*newCell, row, col)
    If row > *ws\maxRow
      *ws\maxRow = row
    EndIf
    If col > *ws\maxColumn
      *ws\maxColumn = col
    EndIf
    ProcedureReturn *newCell
  EndIf
EndProcedure

Procedure.b PbXls_SetCell(*ws.PbXls_Worksheet, row.i, col.i, value.s)
  Define *cell.PbXls_Cell = PbXls_GetCell(*ws, row, col)
  If *cell = 0
    ProcedureReturn #False
  EndIf
  PbXls_SetCellValue(*cell, value)
  If row > *ws\currentRow
    *ws\currentRow = row
  EndIf
  ProcedureReturn #True
EndProcedure

Procedure.s PbXls_GetCellString(*ws.PbXls_Worksheet, row.i, col.i)
  Define key.s = Str(*ws\id) + "_" + Str(row) + "," + Str(col)
  If FindMapElement(PbXls_AllCells(), key)
    ProcedureReturn PbXls_GetCellValue(PbXls_AllCells())
  EndIf
  ProcedureReturn ""
EndProcedure

Procedure.b PbXls_SetCellFormulaWS(*ws.PbXls_Worksheet, row.i, col.i, formula.s)
  Define *cell.PbXls_Cell = PbXls_GetCell(*ws, row, col)
  If *cell = 0
    ProcedureReturn #False
  EndIf
  PbXls_SetCellFormula(*cell, formula)
  If row > *ws\currentRow
    *ws\currentRow = row
  EndIf
  ProcedureReturn #True
EndProcedure

Procedure.s PbXls_GetCellFormulaStr(*ws.PbXls_Worksheet, row.i, col.i)
  Define key.s = Str(*ws\id) + "_" + Str(row) + "," + Str(col)
  If FindMapElement(PbXls_AllCells(), key)
    ProcedureReturn PbXls_GetCellFormula(PbXls_AllCells())
  EndIf
  ProcedureReturn ""
EndProcedure

Procedure.b PbXls_AppendRow(*ws.PbXls_Worksheet, List values.s())
  Define row.i = *ws\currentRow + 1
  Define col.i = 1
  ForEach values()
    PbXls_SetCell(*ws, row, col, values())
    col + 1
  Next
  *ws\currentRow = row
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_MergeCells(*ws.PbXls_Worksheet, rangeString.s)
  Define minCol.Integer, minRow.Integer, maxCol.Integer, maxRow.Integer
  If PbXls_RangeBoundaries(rangeString, @minCol, @minRow, @maxCol, @maxRow) = #False
    ProcedureReturn #False
  EndIf
  Define mcKey.s = Str(*ws\id) + "_" + UCase(rangeString)
  PbXls_MergedCells(mcKey) = UCase(rangeString)
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_UnmergeCells(*ws.PbXls_Worksheet, rangeString.s)
  Define mcKey.s = Str(*ws\id) + "_" + UCase(rangeString)
  If FindMapElement(PbXls_MergedCells(), mcKey)
    DeleteMapElement(PbXls_MergedCells(), mcKey)
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

Procedure.b PbXls_SetColumnWidth(*ws.PbXls_Worksheet, col.i, width.f)
  If col < 1 Or col > #PbXls_MaxColumn
    ProcedureReturn #False
  EndIf
  Define cwKey.s = Str(*ws\id) + "_" + Str(col)
  PbXls_ColumnWidths(cwKey) = width
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_SetRowHeight(*ws.PbXls_Worksheet, row.i, height.f)
  If row < 1 Or row > #PbXls_MaxRow
    ProcedureReturn #False
  EndIf
  Define rhKey.s = Str(*ws\id) + "_" + Str(row)
  PbXls_RowHeights(rhKey) = height
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_SetFreezePanes(*ws.PbXls_Worksheet, cellRef.s)
  *ws\freezePanes = UCase(cellRef)
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_SetAutoFilter(*ws.PbXls_Worksheet, range.s)
  *ws\autoFilter = UCase(range)
  ProcedureReturn #True
EndProcedure

; ***************************************************************************************
; 分区9: 工作簿模块
; ***************************************************************************************

Procedure.i PbXls_CreateWorkbook()
  Define wbId.i = ListSize(PbXls_Workbooks())
  AddElement(PbXls_Workbooks())
  PbXls_Workbooks()\id = wbId
  PbXls_Workbooks()\activeSheetIndex = 0
  PbXls_Workbooks()\path = ""
  PbXls_Workbooks()\creator = "PbXls Library"
  PbXls_Workbooks()\created = PbXls_GetCurrentDateTime()
  PbXls_Workbooks()\modified = PbXls_GetCurrentDateTime()
  PbXls_Workbooks()\isReadOnly = #False
  PbXls_Workbooks()\isWriteOnly = #False
  
  Define sheetCount.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = wbId
      sheetCount + 1
    EndIf
  Wend
  
  AddElement(PbXls_AllWorksheets())
  PbXls_AllWorksheets()\id = sheetCount
  PbXls_AllWorksheets()\title = "Sheet" + Str(sheetCount + 1)
  PbXls_AllWorksheets()\sheetState = "visible"
  PbXls_AllWorksheets()\currentRow = 0
  PbXls_AllWorksheets()\maxRow = 0
  PbXls_AllWorksheets()\maxColumn = 0
  PbXls_AllWorksheets()\parent = wbId
  
  Define wbsKey.s = Str(wbId)
  PbXls_WorkbookSheetCount(wbsKey) = sheetCount + 1
  ProcedureReturn wbId
EndProcedure

Procedure.i PbXls_GetWorkbookPtr(wbId.i)
  If wbId < 0 Or wbId >= ListSize(PbXls_Workbooks())
    ProcedureReturn 0
  EndIf
  SelectElement(PbXls_Workbooks(), wbId)
  ProcedureReturn @PbXls_Workbooks()
EndProcedure

Procedure.i PbXls_ActiveSheet(wbId.i)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn 0
  EndIf
  Define currentIndex.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = wbId
      If currentIndex = *wb\activeSheetIndex
        ProcedureReturn @PbXls_AllWorksheets()
      EndIf
      currentIndex + 1
    EndIf
  Wend
  ProcedureReturn 0
EndProcedure

Procedure.i PbXls_CreateSheet(wbId.i, title.s = "")
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn 0
  EndIf
  If title = ""
    Define wbKey1.s = Str(wbId)
    title = "Sheet" + Str(PbXls_WorkbookSheetCount(wbKey1) + 1)
  EndIf
  Define wbKey2.s = Str(wbId)
  Define sheetCount.i = PbXls_WorkbookSheetCount(wbKey2)
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = wbId And PbXls_AllWorksheets()\title = title
      title = title + "_" + Str(sheetCount)
      Break
    EndIf
  Wend
  AddElement(PbXls_AllWorksheets())
  PbXls_AllWorksheets()\id = sheetCount
  PbXls_AllWorksheets()\title = title
  PbXls_AllWorksheets()\sheetState = "visible"
  PbXls_AllWorksheets()\currentRow = 0
  PbXls_AllWorksheets()\maxRow = 0
  PbXls_AllWorksheets()\maxColumn = 0
  PbXls_AllWorksheets()\parent = wbId
  Define wbKey3.s = Str(wbId)
  PbXls_WorkbookSheetCount(wbKey3) = sheetCount + 1
  *wb\activeSheetIndex = sheetCount
  ProcedureReturn sheetCount
EndProcedure

Procedure.i PbXls_GetSheetByIndex(wbId.i, index.i)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn 0
  EndIf
  Define currentIndex.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = wbId
      If currentIndex = index
        ProcedureReturn @PbXls_AllWorksheets()
      EndIf
      currentIndex + 1
    EndIf
  Wend
  ProcedureReturn 0
EndProcedure

Procedure.i PbXls_GetSheetByName(wbId.i, name.s)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn 0
  EndIf
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = wbId And PbXls_AllWorksheets()\title = name
      ProcedureReturn @PbXls_AllWorksheets()
    EndIf
  Wend
  ProcedureReturn 0
EndProcedure

Procedure.i PbXls_GetSheetCount(wbId.i)
  Define wbKey4.s = Str(wbId)
  ProcedureReturn PbXls_WorkbookSheetCount(wbKey4)
EndProcedure

Procedure.s PbXls_GetSheetName(wbId.i, index.i)
  Define *ws.PbXls_Worksheet = PbXls_GetSheetByIndex(wbId, index)
  If *ws = 0
    ProcedureReturn ""
  EndIf
  ProcedureReturn *ws\title
EndProcedure

Procedure.b PbXls_SetActiveSheet(wbId.i, index.i)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn #False
  EndIf
  Define wbKey5.s = Str(wbId)
  If index >= 0 And index < PbXls_WorkbookSheetCount(wbKey5)
    *wb\activeSheetIndex = index
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

Procedure.b PbXls_RemoveSheet(wbId.i, index.i)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn #False
  EndIf
  Define wbKey6.s = Str(wbId)
  Define count.i = PbXls_WorkbookSheetCount(wbKey6)
  If count <= 1
    ProcedureReturn #False
  EndIf
  PbXls_WorkbookSheetCount(wbKey6) = count - 1
  If *wb\activeSheetIndex >= PbXls_WorkbookSheetCount(wbKey6)
    *wb\activeSheetIndex = PbXls_WorkbookSheetCount(wbKey6) - 1
  EndIf
  ProcedureReturn #True
EndProcedure

; ***************************************************************************************
; 分区10: XML写入器
; ***************************************************************************************

Procedure.i PbXls_WriteSharedStrings(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define sstNode.i = PbXls_XMLAddNode(xmlId, rootNode, "sst")
  PbXls_XMLSetAttribute(sstNode, "xmlns", #PbXls_SHEET_MAIN_NS$)
  
  Define NewMap stringMap.i()
  Define NewList uniqueStrings.s()
  
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      Define wsId.i = PbXls_AllWorksheets()\id
      ForEach PbXls_AllCells()
        Define cellKey.s = MapKey(PbXls_AllCells())
        Define prefix.s = Str(wsId) + "_"
        If Left(cellKey, Len(prefix)) = prefix
          If PbXls_AllCells()\dataType = #PbXls_DataTypeString
            Define value.s = PbXls_AllCells()\value
            If value <> ""
              If FindMapElement(stringMap(), value) = #False
                stringMap(value) = ListIndex(uniqueStrings())
                AddElement(uniqueStrings())
                uniqueStrings() = value
              EndIf
            EndIf
          EndIf
        EndIf
      Next
    EndIf
  Wend
  
  PbXls_XMLSetAttribute(sstNode, "count", Str(ListSize(uniqueStrings())))
  PbXls_XMLSetAttribute(sstNode, "uniqueCount", Str(ListSize(uniqueStrings())))
  
  ResetList(uniqueStrings())
  While NextElement(uniqueStrings())
    Define siNode.i = PbXls_XMLAddNode(xmlId, sstNode, "si")
    Define tNode.i = PbXls_XMLAddNode(xmlId, siNode, "t")
    PbXls_XMLSetText(tNode, PbXls_EscapeXML(uniqueStrings()))
  Wend
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteWorksheetXML(*ws.PbXls_Worksheet, *wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define worksheetNode.i = PbXls_XMLAddNode(xmlId, rootNode, "worksheet")
  PbXls_XMLSetAttribute(worksheetNode, "xmlns", #PbXls_SHEET_MAIN_NS$)
  PbXls_XMLSetAttribute(worksheetNode, "xmlns:r", #PbXls_REL_NS$)
  
  If *ws\maxRow > 0 And *ws\maxColumn > 0
    Define dimensionNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "dimension")
    PbXls_XMLSetAttribute(dimensionNode, "ref", PbXls_RangeString(1, 1, *ws\maxRow, *ws\maxColumn))
  EndIf
  
  Define sheetViewsNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "sheetViews")
  Define sheetViewNode.i = PbXls_XMLAddNode(xmlId, sheetViewsNode, "sheetView")
  PbXls_XMLSetAttribute(sheetViewNode, "showGridLines", "1")
  PbXls_XMLSetAttribute(sheetViewNode, "showRowColHeaders", "1")
  Define isActive.i = 0
  If *ws\id = *wb\activeSheetIndex
    isActive = 1
  EndIf
  PbXls_XMLSetAttribute(sheetViewNode, "tabSelected", Str(isActive))
  PbXls_XMLSetAttribute(sheetViewNode, "workbookViewId", "0")
  
  If *ws\freezePanes <> ""
    Define paneNode.i = PbXls_XMLAddNode(xmlId, sheetViewNode, "pane")
    PbXls_XMLSetAttribute(paneNode, "topLeftCell", *ws\freezePanes)
    PbXls_XMLSetAttribute(paneNode, "activePane", "topRight")
    PbXls_XMLSetAttribute(paneNode, "state", "frozen")
    Define freezeRow.i, freezeCol.i
    PbXls_CoordinateToTuple(*ws\freezePanes, @freezeRow, @freezeCol)
    If freezeCol > 1
      PbXls_XMLSetAttribute(paneNode, "xSplit", Str(freezeCol - 1))
    EndIf
    If freezeRow > 1
      PbXls_XMLSetAttribute(paneNode, "ySplit", Str(freezeRow - 1))
    EndIf
  EndIf
  
  Define hasCols.b = #False
  ForEach PbXls_ColumnWidths()
    Define cwKey.s = MapKey(PbXls_ColumnWidths())
    Define wsPrefix.s = Str(*ws\id) + "_"
    If Left(cwKey, Len(wsPrefix)) = wsPrefix
      hasCols = #True
      Break
    EndIf
  Next
  
  If hasCols
    Define colsNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "cols")
    ForEach PbXls_ColumnWidths()
      Define key.s = MapKey(PbXls_ColumnWidths())
      If Left(key, Len(wsPrefix)) = wsPrefix
        Define colIdx.s = Mid(key, Len(wsPrefix) + 1)
        Define colNode.i = PbXls_XMLAddNode(xmlId, colsNode, "col")
        PbXls_XMLSetAttribute(colNode, "min", colIdx)
        PbXls_XMLSetAttribute(colNode, "max", colIdx)
        PbXls_XMLSetAttribute(colNode, "width", StrF(PbXls_ColumnWidths(), 2))
        PbXls_XMLSetAttribute(colNode, "customWidth", "1")
        PbXls_XMLSetAttribute(colNode, "bestFit", "1")
      EndIf
    Next
  EndIf
  
  Define sheetDataNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "sheetData")
  
  Define NewMap rows.i()
  Define wsIdStr.s = Str(*ws\id) + "_"
  
  ForEach PbXls_AllCells()
    Define cKey.s = MapKey(PbXls_AllCells())
    If Left(cKey, Len(wsIdStr)) = wsIdStr
      Define cellRow.s = Mid(cKey, Len(wsIdStr) + 1)
      Define commaPos.i = FindString(cellRow, ",", 1)
      If commaPos > 0
        Define r.i = Val(Left(cellRow, commaPos - 1))
        If FindMapElement(rows(), Str(r)) = #False
          rows(Str(r)) = r
        EndIf
      EndIf
    EndIf
  Next
  
  Define NewList sortedRows.i()
  ForEach rows()
    AddElement(sortedRows())
    sortedRows() = rows()
  Next
  SortList(sortedRows(), #PB_Sort_Ascending)
  
  ResetList(sortedRows())
  While NextElement(sortedRows())
    Define currentRow.i = sortedRows()
    Define rowNode.i = PbXls_XMLAddNode(xmlId, sheetDataNode, "row")
    PbXls_XMLSetAttribute(rowNode, "r", Str(currentRow))
    
    Define rhKey.s = Str(*ws\id) + "_" + Str(currentRow)
    If FindMapElement(PbXls_RowHeights(), rhKey)
    PbXls_XMLSetAttribute(rowNode, "ht", StrF(PbXls_RowHeights(rhKey), 2))
      PbXls_XMLSetAttribute(rowNode, "customHeight", "1")
    EndIf
    
    ForEach PbXls_AllCells()
      Define ck.s = MapKey(PbXls_AllCells())
      If Left(ck, Len(wsIdStr)) = wsIdStr
        Define cellPart.s = Mid(ck, Len(wsIdStr) + 1)
        Define cPos.i = FindString(cellPart, ",", 1)
        If cPos > 0
          Define cr.i = Val(Left(cellPart, cPos - 1))
          If cr = currentRow
            Define *cell.PbXls_Cell = PbXls_AllCells()
            Define cellNode.i = PbXls_XMLAddNode(xmlId, rowNode, "c")
            PbXls_XMLSetAttribute(cellNode, "r", PbXls_CoordinateFromRowCol(*cell\row, *cell\column))
            
            If *cell\styleId > 0
              PbXls_XMLSetAttribute(cellNode, "s", Str(*cell\styleId))
            EndIf
            
            Select *cell\dataType
              Case #PbXls_DataTypeNumeric
                Define vNode1.i = PbXls_XMLAddNode(xmlId, cellNode, "v")
                PbXls_XMLSetText(vNode1, *cell\value)
              Case #PbXls_DataTypeBoolean
                PbXls_XMLSetAttribute(cellNode, "t", "b")
                Define vNode2.i = PbXls_XMLAddNode(xmlId, cellNode, "v")
                PbXls_XMLSetText(vNode2, *cell\value)
              Case #PbXls_DataTypeFormula
                PbXls_XMLSetAttribute(cellNode, "t", "str")
                Define fNode.i = PbXls_XMLAddNode(xmlId, cellNode, "f")
                PbXls_XMLSetText(fNode, *cell\formula)
              Case #PbXls_DataTypeString
                PbXls_XMLSetAttribute(cellNode, "t", "inlineStr")
                Define isNode.i = PbXls_XMLAddNode(xmlId, cellNode, "is")
                Define tNode2.i = PbXls_XMLAddNode(xmlId, isNode, "t")
                PbXls_XMLSetText(tNode2, PbXls_EscapeXML(*cell\value))
              Case #PbXls_DataTypeInline
                PbXls_XMLSetAttribute(cellNode, "t", "inlineStr")
                Define isNode2.i = PbXls_XMLAddNode(xmlId, cellNode, "is")
                Define tNode3.i = PbXls_XMLAddNode(xmlId, isNode2, "t")
                PbXls_XMLSetText(tNode3, PbXls_EscapeXML(*cell\value))
              Case #PbXls_DataTypeError
                PbXls_XMLSetAttribute(cellNode, "t", "e")
                Define vNode3.i = PbXls_XMLAddNode(xmlId, cellNode, "v")
                PbXls_XMLSetText(vNode3, *cell\value)
              Case #PbXls_DataTypeNull
            EndSelect
          EndIf
        EndIf
      EndIf
    Next
  Wend
  
  Define mcCount.i = 0
  ForEach PbXls_MergedCells()
    Define mKey.s = MapKey(PbXls_MergedCells())
    If Left(mKey, Len(wsIdStr)) = wsIdStr
      mcCount + 1
    EndIf
  Next
  
  If mcCount > 0
    Define mergeCellsNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "mergeCells")
    PbXls_XMLSetAttribute(mergeCellsNode, "count", Str(mcCount))
    ForEach PbXls_MergedCells()
      Define mk.s = MapKey(PbXls_MergedCells())
      If Left(mk, Len(wsIdStr)) = wsIdStr
        Define mergeCellNode.i = PbXls_XMLAddNode(xmlId, mergeCellsNode, "mergeCell")
        PbXls_XMLSetAttribute(mergeCellNode, "ref", PbXls_MergedCells())
      EndIf
    Next
  EndIf
  
  If *ws\autoFilter <> ""
    Define autoFilterNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "autoFilter")
    PbXls_XMLSetAttribute(autoFilterNode, "ref", *ws\autoFilter)
  EndIf
  
  Define pageSetupNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "pageSetup")
  If *ws\orientation <> ""
    PbXls_XMLSetAttribute(pageSetupNode, "orientation", *ws\orientation)
  Else
    PbXls_XMLSetAttribute(pageSetupNode, "orientation", "portrait")
  EndIf
  If *ws\paperSize > 0
    PbXls_XMLSetAttribute(pageSetupNode, "paperSize", Str(*ws\paperSize))
  Else
    PbXls_XMLSetAttribute(pageSetupNode, "paperSize", "9")
  EndIf
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteWorkbookXML(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define workbookNode.i = PbXls_XMLAddNode(xmlId, rootNode, "workbook")
  PbXls_XMLSetAttribute(workbookNode, "xmlns", #PbXls_SHEET_MAIN_NS$)
  PbXls_XMLSetAttribute(workbookNode, "xmlns:r", #PbXls_REL_NS$)
  
  Define bookViewsNode.i = PbXls_XMLAddNode(xmlId, workbookNode, "bookViews")
  Define workBookViewNode.i = PbXls_XMLAddNode(xmlId, bookViewsNode, "workbookView")
  PbXls_XMLSetAttribute(workBookViewNode, "activeTab", Str(*wb\activeSheetIndex))
  PbXls_XMLSetAttribute(workBookViewNode, "firstSheet", "0")
  PbXls_XMLSetAttribute(workBookViewNode, "tabRatio", "600")
  PbXls_XMLSetAttribute(workBookViewNode, "visibility", "visible")
  PbXls_XMLSetAttribute(workBookViewNode, "xWindow", "0")
  PbXls_XMLSetAttribute(workBookViewNode, "yWindow", "0")
  
  Define sheetsNode.i = PbXls_XMLAddNode(xmlId, workbookNode, "sheets")
  Define sheetIndex.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      Define sheetNode.i = PbXls_XMLAddNode(xmlId, sheetsNode, "sheet")
      PbXls_XMLSetAttribute(sheetNode, "name", PbXls_AllWorksheets()\title)
      PbXls_XMLSetAttribute(sheetNode, "sheetId", Str(sheetIndex + 1))
      PbXls_XMLSetAttribute(sheetNode, "state", PbXls_AllWorksheets()\sheetState)
      PbXls_XMLSetAttribute(sheetNode, "r:id", "rId" + Str(sheetIndex + 2))
      sheetIndex + 1
    EndIf
  Wend
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteDocPropsXML(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define cpNode.i = PbXls_XMLAddNode(xmlId, rootNode, "cp:coreProperties")
  PbXls_XMLSetAttribute(cpNode, "xmlns:cp", #PbXls_COREPROPS_NS$)
  PbXls_XMLSetAttribute(cpNode, "xmlns:dc", #PbXls_DCORE_NS$)
  PbXls_XMLSetAttribute(cpNode, "xmlns:dcterms", #PbXls_DCTERMS_NS$)
  PbXls_XMLSetAttribute(cpNode, "xmlns:dcmitype", "http://purl.org/dc/dcmitype/")
  PbXls_XMLSetAttribute(cpNode, "xmlns:xsi", #PbXls_XSI_NS$)
  
  If *wb\creator <> ""
    Define creatorNode.i = PbXls_XMLAddNode(xmlId, cpNode, "dc:creator")
    PbXls_XMLSetText(creatorNode, *wb\creator)
  EndIf
  
  If *wb\title <> ""
    Define titleNode.i = PbXls_XMLAddNode(xmlId, cpNode, "dc:title")
    PbXls_XMLSetText(titleNode, *wb\title)
  EndIf
  
  If *wb\subject <> ""
    Define subjectNode.i = PbXls_XMLAddNode(xmlId, cpNode, "dc:subject")
    PbXls_XMLSetText(subjectNode, *wb\subject)
  EndIf
  
  If *wb\description <> ""
    Define descNode.i = PbXls_XMLAddNode(xmlId, cpNode, "dc:description")
    PbXls_XMLSetText(descNode, *wb\description)
  EndIf
  
  If *wb\keywords <> ""
    Define keywordsNode.i = PbXls_XMLAddNode(xmlId, cpNode, "cp:keywords")
    PbXls_XMLSetText(keywordsNode, *wb\keywords)
  EndIf
  
  If *wb\category <> ""
    Define categoryNode.i = PbXls_XMLAddNode(xmlId, cpNode, "cp:category")
    PbXls_XMLSetText(categoryNode, *wb\category)
  EndIf
  
  If *wb\created <> ""
    Define createdNode.i = PbXls_XMLAddNode(xmlId, cpNode, "dcterms:created")
    PbXls_XMLSetAttribute(createdNode, "xsi:type", "dcterms:W3CDTF")
    PbXls_XMLSetText(createdNode, *wb\created + "Z")
  EndIf
  
  If *wb\modified <> ""
    Define modifiedNode.i = PbXls_XMLAddNode(xmlId, cpNode, "dcterms:modified")
    PbXls_XMLSetAttribute(modifiedNode, "xsi:type", "dcterms:W3CDTF")
    PbXls_XMLSetText(modifiedNode, *wb\modified + "Z")
  EndIf
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteAppPropsXML(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define propsNode.i = PbXls_XMLAddNode(xmlId, rootNode, "Properties")
  PbXls_XMLSetAttribute(propsNode, "xmlns", #PbXls_VTYPES_NS$)
  PbXls_XMLSetAttribute(propsNode, "xmlns:vt", #PbXls_VTYPES_NS$)
  
  Define appNode.i = PbXls_XMLAddNode(xmlId, propsNode, "Application")
  PbXls_XMLSetText(appNode, "PbXls Library")
  
  Define companyNode.i = PbXls_XMLAddNode(xmlId, propsNode, "Company")
  PbXls_XMLSetText(companyNode, "")
  
  Define versionNode.i = PbXls_XMLAddNode(xmlId, propsNode, "AppVersion")
  PbXls_XMLSetText(versionNode, "1.0.0")
  
  Define wbKey7.s = Str(*wb\id)
  Define sheetCount.i = PbXls_WorkbookSheetCount(wbKey7)
  Define headingPairsNode.i = PbXls_XMLAddNode(xmlId, propsNode, "HeadingPairs")
  Define vtVector1.i = PbXls_XMLAddNode(xmlId, headingPairsNode, "vt:vector")
  PbXls_XMLSetAttribute(vtVector1, "size", "2")
  PbXls_XMLSetAttribute(vtVector1, "baseType", "variant")
  
  Define vtVariant1.i = PbXls_XMLAddNode(xmlId, vtVector1, "vt:variant")
  Define vtLpwstr1.i = PbXls_XMLAddNode(xmlId, vtVariant1, "vt:lpwstr")
  PbXls_XMLSetText(vtLpwstr1, "Worksheets")
  
  Define vtVariant2.i = PbXls_XMLAddNode(xmlId, vtVector1, "vt:variant")
  Define vtI4.i = PbXls_XMLAddNode(xmlId, vtVariant2, "vt:i4")
  PbXls_XMLSetText(vtI4, Str(sheetCount))
  
  Define titlesOfPartsNode.i = PbXls_XMLAddNode(xmlId, propsNode, "TitlesOfParts")
  Define vtVector2.i = PbXls_XMLAddNode(xmlId, titlesOfPartsNode, "vt:vector")
  PbXls_XMLSetAttribute(vtVector2, "size", Str(sheetCount))
  PbXls_XMLSetAttribute(vtVector2, "baseType", "lpwstr")
  
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      Define lpwstrNode.i = PbXls_XMLAddNode(xmlId, vtVector2, "vt:lpwstr")
      PbXls_XMLSetText(lpwstrNode, PbXls_AllWorksheets()\title)
    EndIf
  Wend
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteContentTypesXML(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define typesNode.i = PbXls_XMLAddNode(xmlId, rootNode, "Types")
  PbXls_XMLSetAttribute(typesNode, "xmlns", #PbXls_CONTYPES_NS$)
  
  Define default1.i = PbXls_XMLAddNode(xmlId, typesNode, "Default")
  PbXls_XMLSetAttribute(default1, "Extension", "rels")
  PbXls_XMLSetAttribute(default1, "ContentType", "application/vnd.openxmlformats-package.relationships+xml")
  
  Define default2.i = PbXls_XMLAddNode(xmlId, typesNode, "Default")
  PbXls_XMLSetAttribute(default2, "Extension", "xml")
  PbXls_XMLSetAttribute(default2, "ContentType", "application/xml")
  
  Define override1.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
  PbXls_XMLSetAttribute(override1, "PartName", "/xl/workbook.xml")
  PbXls_XMLSetAttribute(override1, "ContentType", #PbXls_WORKBOOK$)
  
  Define override2.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
  PbXls_XMLSetAttribute(override2, "PartName", "/xl/sharedStrings.xml")
  PbXls_XMLSetAttribute(override2, "ContentType", #PbXls_SHARED_STRINGS$)
  
  Define override3.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
  PbXls_XMLSetAttribute(override3, "PartName", "/xl/styles.xml")
  PbXls_XMLSetAttribute(override3, "ContentType", #PbXls_STYLES_TYPE$)
  
  Define override4.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
  PbXls_XMLSetAttribute(override4, "PartName", "/docProps/core.xml")
  PbXls_XMLSetAttribute(override4, "ContentType", "application/vnd.openxmlformats-package.core-properties+xml")
  
  Define override5.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
  PbXls_XMLSetAttribute(override5, "PartName", "/docProps/app.xml")
  PbXls_XMLSetAttribute(override5, "ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
  
  Define wsIdx.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      Define overrideWs.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
    PbXls_XMLSetAttribute(overrideWs, "PartName", "/xl/worksheets/sheet" + Str(wsIdx + 1) + ".xml")
      PbXls_XMLSetAttribute(overrideWs, "ContentType", #PbXls_SPREADSHEET$)
      wsIdx + 1
    EndIf
  Wend
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteRootRelsXML()
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define relsNode.i = PbXls_XMLAddNode(xmlId, rootNode, "Relationships")
  PbXls_XMLSetAttribute(relsNode, "xmlns", #PbXls_PKG_REL_NS$)
  
  Define rel1.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
  PbXls_XMLSetAttribute(rel1, "Id", "rId1")
  PbXls_XMLSetAttribute(rel1, "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
  PbXls_XMLSetAttribute(rel1, "Target", "xl/workbook.xml")
  
  Define rel2.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
  PbXls_XMLSetAttribute(rel2, "Id", "rId2")
  PbXls_XMLSetAttribute(rel2, "Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
  PbXls_XMLSetAttribute(rel2, "Target", "docProps/core.xml")
  
  Define rel3.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
  PbXls_XMLSetAttribute(rel3, "Id", "rId3")
  PbXls_XMLSetAttribute(rel3, "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
  PbXls_XMLSetAttribute(rel3, "Target", "docProps/app.xml")
  
  ProcedureReturn xmlId
EndProcedure

Procedure.i PbXls_WriteWorkbookRelsXML(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define relsNode.i = PbXls_XMLAddNode(xmlId, rootNode, "Relationships")
  PbXls_XMLSetAttribute(relsNode, "xmlns", #PbXls_PKG_REL_NS$)
  
  Define rel1.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
  PbXls_XMLSetAttribute(rel1, "Id", "rId1")
  PbXls_XMLSetAttribute(rel1, "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
  PbXls_XMLSetAttribute(rel1, "Target", "sharedStrings.xml")
  
  Define wsIdx2.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      Define relWs.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
      PbXls_XMLSetAttribute(relWs, "Id", "rId" + Str(wsIdx2 + 2))
      PbXls_XMLSetAttribute(relWs, "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
      PbXls_XMLSetAttribute(relWs, "Target", "worksheets/sheet" + Str(wsIdx2 + 1) + ".xml")
      wsIdx2 + 1
    EndIf
  Wend
  
  Define relStyles.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
  PbXls_XMLSetAttribute(relStyles, "Id", "rId" + Str(wsIdx2 + 2))
  PbXls_XMLSetAttribute(relStyles, "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
  PbXls_XMLSetAttribute(relStyles, "Target", "styles.xml")
  
  ProcedureReturn xmlId
EndProcedure

; Helper procedure to add XML content to ZIP
Procedure.b PbXls_AddXMLToZIP(packId.i, xmlId.i, archivePath.s)
  If xmlId = 0
    ProcedureReturn #False
  EndIf
  
  Define xmlStr.s = PbXls_XMLSaveToString(xmlId)
  If xmlStr = ""
    ProcedureReturn #False
  EndIf
  
  ; UTF8编码每个字符最多需要4个字节，分配足够大的缓冲区
  Define bufferSize.i = Len(xmlStr) * 4 + 10
  Define *buffer = AllocateMemory(bufferSize)
  If *buffer = 0
    ProcedureReturn #False
  EndIf
  
  PokeA(*buffer, $EF)
  PokeA(*buffer + 1, $BB)
  PokeA(*buffer + 2, $BF)
  Define strLen.i = PokeS(*buffer + 3, xmlStr, bufferSize - 3, #PB_UTF8)
  Define totalSize.i = 3 + strLen
  
  Define result.b = PbXls_ZIPAddMemory(packId, *buffer, totalSize, archivePath)
  FreeMemory(*buffer)
  PbXls_XMLFree(xmlId)
  
  ProcedureReturn result
EndProcedure

; PbXls_SaveWorkbookToFile - 保存工作簿到文件
Procedure.b PbXls_SaveWorkbookToFile(wbId.i, filename.s)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    ProcedureReturn #False
  EndIf
  
  *wb\modified = PbXls_GetCurrentDateTime()
  
  Define packId.i = PbXls_ZIPCreate(filename)
  If packId = 0
    ProcedureReturn #False
  EndIf
  
  ; 1. [Content_Types].xml
  Define contentTypesXml.i = PbXls_WriteContentTypesXML(*wb)
  PbXls_AddXMLToZIP(packId, contentTypesXml, #PbXls_ARCContentTypes$)
  
  ; 2. _rels/.rels
  Define rootRelsXml.i = PbXls_WriteRootRelsXML()
  PbXls_AddXMLToZIP(packId, rootRelsXml, #PbXls_ARCRootRels$)
  
  ; 3. docProps/core.xml
  Define docPropsXml.i = PbXls_WriteDocPropsXML(*wb)
  PbXls_AddXMLToZIP(packId, docPropsXml, #PbXls_ARCCore$)
  
  ; 4. docProps/app.xml
  Define appPropsXml.i = PbXls_WriteAppPropsXML(*wb)
  PbXls_AddXMLToZIP(packId, appPropsXml, #PbXls_ARCApp$)
  
  ; 5. xl/workbook.xml
  Define workbookXml.i = PbXls_WriteWorkbookXML(*wb)
  PbXls_AddXMLToZIP(packId, workbookXml, #PbXls_ARCWorkbook$)
  
  ; 6. xl/_rels/workbook.xml.rels
  Define workbookRelsXml.i = PbXls_WriteWorkbookRelsXML(*wb)
  PbXls_AddXMLToZIP(packId, workbookRelsXml, #PbXls_ARCWorkbookRels$)
  
  ; 7. xl/sharedStrings.xml
  Define sharedStringsXml.i = PbXls_WriteSharedStrings(*wb)
  PbXls_AddXMLToZIP(packId, sharedStringsXml, #PbXls_ARCSharedStrings$)
  
  ; 8. xl/styles.xml (need to implement)
  
  ; 9. xl/worksheets/sheetN.xml
  Define sheetIdx3.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      Define worksheetXml.i = PbXls_WriteWorksheetXML(@PbXls_AllWorksheets(), *wb)
      Define archivePath.s = "xl/worksheets/sheet" + Str(sheetIdx3 + 1) + ".xml"
      PbXls_AddXMLToZIP(packId, worksheetXml, archivePath)
      sheetIdx3 + 1
    EndIf
  Wend
  
  PbXls_ZIPClose(packId)
  *wb\path = filename
  
  ProcedureReturn #True
EndProcedure

; ***************************************************************************************
; 分区13: 公共API
; ***************************************************************************************

Procedure.i PbXls_LoadWorkbook(filename.s)
  ProcedureReturn -1
EndProcedure

Procedure.b PbXls_SaveWorkbook(workbook.i, filename.s)
  ProcedureReturn PbXls_SaveWorkbookToFile(workbook, filename)
EndProcedure

Procedure.b PbXls_CloseWorkbook(workbook.i)
  ProcedureReturn #True
EndProcedure

Procedure.i PbXls_GetSheetByIndexAPI(workbook.i, index.i)
  ProcedureReturn PbXls_GetSheetByIndex(workbook, index)
EndProcedure

Procedure.b PbXls_SetActiveSheetAPI(workbook.i, index.i)
  ProcedureReturn PbXls_SetActiveSheet(workbook, index)
EndProcedure

Procedure.b PbXls_SetCellAPI(worksheet.i, row.i, col.i, value.s)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetCell(*ws, row, col, value)
EndProcedure

Procedure.s PbXls_GetCellStringAPI(worksheet.i, row.i, col.i)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn ""
  EndIf
  ProcedureReturn PbXls_GetCellString(*ws, row, col)
EndProcedure

Procedure.b PbXls_SetCellFormulaAPI(worksheet.i, row.i, col.i, formula.s)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetCellFormulaWS(*ws, row, col, formula)
EndProcedure

Procedure.s PbXls_GetCellFormulaAPI(worksheet.i, row.i, col.i)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn ""
  EndIf
  ProcedureReturn PbXls_GetCellFormulaStr(*ws, row, col)
EndProcedure

Procedure.b PbXls_MergeCellsAPI(worksheet.i, rangeString.s)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_MergeCells(*ws, rangeString)
EndProcedure

Procedure.b PbXls_SetColumnWidthAPI(worksheet.i, col.i, width.f)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetColumnWidth(*ws, col, width)
EndProcedure

Procedure.b PbXls_SetRowHeightAPI(worksheet.i, row.i, height.f)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetRowHeight(*ws, row, height)
EndProcedure

Procedure.b PbXls_SetCellStyleAPI(worksheet.i, row.i, col.i, styleId.i)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  Define *cell.PbXls_Cell = PbXls_GetCell(*ws, row, col)
  If *cell = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetCellStyle(*cell, styleId)
EndProcedure

Procedure.i PbXls_CreateFontAPI()
  ProcedureReturn PbXls_CreateFont()
EndProcedure

Procedure.b PbXls_SetFontAPI(fontId.i, name.s = "", size.f = -1, bold.b = -1, italic.b = -1, color.s = "")
  ProcedureReturn PbXls_SetFont(fontId, name, size, bold, italic, color)
EndProcedure

Procedure.i PbXls_CreateFillAPI()
  ProcedureReturn PbXls_CreateFill()
EndProcedure

Procedure.b PbXls_SetFillAPI(fillId.i, patternType.s = "", fgColor.s = "", bgColor.s = "")
  ProcedureReturn PbXls_SetFill(fillId, patternType, fgColor, bgColor)
EndProcedure

Procedure.i PbXls_CreateBorderAPI()
  ProcedureReturn PbXls_CreateBorder()
EndProcedure

Procedure.b PbXls_SetBorderAPI(borderId.i, side.s, style.s = "thin", color.s = "000000")
  ProcedureReturn PbXls_SetBorder(borderId, side, style, color)
EndProcedure

Procedure.i PbXls_CreateAlignmentAPI()
  ProcedureReturn PbXls_CreateAlignment()
EndProcedure

Procedure.b PbXls_SetAlignmentAPI(styleId.i, horizontal.s = "", vertical.s = "", wrapText.b = -1, indent.i = -1)
  ProcedureReturn PbXls_SetAlignment(styleId, horizontal, vertical, wrapText, indent)
EndProcedure

Procedure.b PbXls_SetNumberFormatAPI(styleId.i, numFmt.s)
  ProcedureReturn PbXls_SetNumberFormat(styleId, numFmt)
EndProcedure

Procedure.b PbXls_IsDateFormatAPI(numFmtId.i)
  ProcedureReturn PbXls_IsDateFormat(numFmtId)
EndProcedure

Procedure.s PbXls_GetBuiltinFormatAPI(numFmtId.i)
  ProcedureReturn PbXls_GetBuiltinFormat(numFmtId)
EndProcedure

; ***************************************************************************************
; 分区14: 初始化和清理
; ***************************************************************************************

Procedure.b PbXls_Init()
  UseZipPacker()
  ClearList(PbXls_Fonts())
  AddElement(PbXls_Fonts())
  PbXls_Fonts()\name = "Calibri"
  PbXls_Fonts()\size = 11.0
  PbXls_Fonts()\color = "000000"
  PbXls_Fonts()\family = 2
  PbXls_Fonts()\scheme = "minor"
  ClearList(PbXls_Fills())
  ClearList(PbXls_Borders())
  AddElement(PbXls_Borders())
  ClearList(PbXls_CellStyles())
  AddElement(PbXls_CellStyles())
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  PbXls_CellStyles()\numFmtId = 0
  PbXls_CellStyles()\numFmt = "General"
  ProcedureReturn #True
EndProcedure

Procedure.b PbXls_Free()
  ClearList(PbXls_Workbooks())
  ClearList(PbXls_Fonts())
  ClearList(PbXls_Fills())
  ClearList(PbXls_Borders())
  ClearList(PbXls_CellStyles())
  ClearList(PbXls_AllWorksheets())
  ClearList(PbXls_SharedStrings())
  ClearMap(PbXls_AllCells())
  ClearMap(PbXls_ColumnWidths())
  ClearMap(PbXls_RowHeights())
  ClearMap(PbXls_MergedCells())
  ClearMap(PbXls_MergedCellCount())
  ClearMap(PbXls_WorkbookSheetCount())
  ClearMap(PbXls_WorkbookSharedStrings())
  ProcedureReturn #True
EndProcedure

PbXls_Init()