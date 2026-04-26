; ***************************************************************************************
; PbXls Library - Excel xlsx/xlsm 操作库
; 版本: 2.6
; 作者 lcode.cn
; 许可证 Apache 2.0
;
; 说明: 无依赖操作Excel xlsx/xlsm文件的PureBasic库
;       使用PureBasic内置XML和Packer库
;       无需安装Microsoft Office或任何第三方依赖
;
; 捐赠支持:
;   PayPal: https://www.paypal.me/lcodecn
;   微信:   #付款:lcodecn(经营_lcodecn)/openlib/003
;
;   如果这个库对您有帮助，欢迎捐赠支持项目的持续开发和维护。感谢您的慷慨！
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
;   分区2: 枚举定义（数据类型、工作状态、边框、对齐、填充等）
;   分区3: 结构体定义和全局数据存储
;   分区4: 工具函数（坐标转换、字符串处理、日期时间、XML/ZIP辅助）
;   分区5: XML常量模块
;   分区6: 样式模块（字体、填充、边框、对齐、数字格式）
;   分区7: 单元格模块
;   分区8: 工作表模块
;   分区9: 工作簿模块
;   分区10: XML写入器（生成Excel文件各部分XML）
;   分区11: XML读取器（解析Excel文件各部分XML）
;   分区12: 高级功能
;   分区13: 公共API
;   分区14: 初始化和清理
;   分区15: 测试代码
; ***************************************************************************************

; 启用ZIP打包库（PureBASIC的XML库是内置的，无需初始化）
UseZipPacker()

; ***************************************************************************************
; 分区1: 常量定义

; 1.1 Excel规范常量（定义Excel的最大行列数等限制）
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

; 1.6 内置数字格式常量（定义Excel内置的数字格式ID）?
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
;    定义库中使用的各种枚举类型，包括数据类型、工作表状态、边框样式、?
;    对齐方式、填充模式和错误类型等?
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

; 数据验证结构体
Structure PbXls_DataValidation
  type.s          ; "list", "whole", "decimal", "textLength", "date", "time", "custom"
  operator.s      ; "between", "notBetween", "equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"
  sqref.s         ; 单元格范围, 如 "A1:A10"
  formula1.s      ; 公式或值列表(逗号分隔)
  formula2.s      ; 第二个公式(某些操作符需要)
  showErrorMessage.b
  showInputMessage.b
  allowBlank.b
  errorTitle.s
  error.s
  promptTitle.s
  prompt.s
  showDropDown.b  ; 是否隐藏下拉箭头
EndStructure

; 条件格式规则结构体
Structure PbXls_ConditionalFormatRule
  type.s          ; "cellIs", "expression", "colorScale", "dataBar", "iconSet"
  operator.s      ; 用于cellIs类型
  priority.i      ; 优先级
  sqref.s         ; 应用范围, 如 "A1:A10"
  formula1.s      ; 公式或值
  formula2.s      ; 第二个公式
  ; Dxf差异样式
  dxfFontName.s
  dxfFontSize.f
  dxfFontColor.s
  dxfFontBold.b
  dxfFontItalic.b
  dxfFillColor.s
  dxfFillPattern.s
  ; colorScale/dataBar/iconSet参数
  minType.s
  minValue.s
  maxType.s
  maxValue.s
  midType.s
  midValue.s
  minColor.s
  midColor.s
  maxColor.s
EndStructure

; 图表系列数据(使用全局Map存储)
Global NewMap PbXls_ChartSeriesName.s()
Global NewMap PbXls_ChartSeriesValues.s()
Global NewMap PbXls_ChartSeriesCategories.s()

; 图表结构体
Structure PbXls_Chart
  title.s
  type.s          ; "barChart", "lineChart", "pieChart", "scatterChart", "areaChart"
  style.i         ; 图表样式(1-48)
  x.i             ; 绘图位置X(EM单位)
  y.i             ; 绘图位置Y(EM单位)
  cx.i            ; 宽度(EM单位)
  cy.i            ; 高度(EM单位)
  anchorRef.s     ; 锚定单元格范围, 如 "A1:F20"
  seriesCount.i   ; 数据系列数量
EndStructure

; 全局数据存储
Global NewMap PbXls_AllCells.PbXls_Cell()
Global NewMap PbXls_ColumnWidths.f()
Global NewMap PbXls_RowHeights.f()
Global NewMap PbXls_MergedCells.s()
Global NewMap PbXls_MergedCellCount.i()
Global NewList PbXls_DataValidations.PbXls_DataValidation()
Global NewList PbXls_ConditionalFormats.PbXls_ConditionalFormatRule()
Global NewList PbXls_Charts.PbXls_Chart()
Global NewList PbXls_Fonts.PbXls_Font()
Global NewList PbXls_Fills.PbXls_Fill()
Global NewList PbXls_Borders.PbXls_Border()
Global NewList PbXls_CellStyles.PbXls_CellStyle()
Global NewList PbXls_AllWorksheets.PbXls_Worksheet()
Global NewList PbXls_SharedStrings.s()
Global NewList PbXls_Workbooks.PbXls_Workbook()
Global NewMap PbXls_WorkbookSheetCount.i()
Global NewMap PbXls_WorkbookSharedStrings.i()
Global PbXls_DxfCounter.i = 0
Global PbXls_ChartCounter.i = 0

; ***************************************************************************************
; 分区4: 工具函数
; ***************************************************************************************

; GetColumnLetter - 鍒楀彿杞瓧姣?(1->"A", 27->"AA")
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

; ColumnIndexFromString - 瀛楁瘝杞垪鍙?("A"->1, "AA"->27)
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

; CoordinateToTuple - 鍧愭爣杞鍒楀厓缁?("A1"->(1,1))
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

; RangeBoundaries - 鑼冨洿瑙ｆ瀽 ("A1:D5"->(1,1,4,5))
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

; CoordinateFromRowCol - 琛屽垪杞潗鏍?(1,1->"A1")
Procedure.s PbXls_CoordinateFromRowCol(row.i, col.i)
  ProcedureReturn PbXls_GetColumnLetter(col) + Str(row)
EndProcedure

; RangeString - 鐢熸垚鑼冨洿瀛楃涓?
Procedure.s PbXls_RangeString(minRow.i, minCol.i, maxRow.i, maxCol.i)
  ProcedureReturn PbXls_CoordinateFromRowCol(minRow, minCol) + ":" + PbXls_CoordinateFromRowCol(maxRow, maxCol)
EndProcedure

; QuoteSheetName - 涓哄寘鍚壒娈婂瓧绗︾殑宸ヤ綔琛ㄥ悕娣诲姞寮曞彿
Procedure.s PbXls_QuoteSheetName(sheetName.s)
  If FindString(sheetName, " ", 1) Or FindString(sheetName, "'", 1)
    sheetName = ReplaceString(sheetName, "'", "''")
    ProcedureReturn "'" + sheetName + "'"
  EndIf
  ProcedureReturn sheetName
EndProcedure

; EscapeXML - 杞箟XML鐗规畩瀛楃
Procedure.s PbXls_EscapeXML(text.s)
  text = ReplaceString(text, "&", "&amp;")
  text = ReplaceString(text, "<", "&lt;")
  text = ReplaceString(text, ">", "&gt;")
  text = ReplaceString(text, ~"\"", "&quot;")
  text = ReplaceString(text, "'", "&apos;")
  ProcedureReturn text
EndProcedure

; UnescapeXML - 鍙嶈浆涔塜ML鐗规畩瀛楃
Procedure.s PbXls_UnescapeXML(text.s)
  text = ReplaceString(text, "&amp;", "&")
  text = ReplaceString(text, "&lt;", "<")
  text = ReplaceString(text, "&gt;", ">")
  text = ReplaceString(text, "&quot;", ~"\"")
  text = ReplaceString(text, "&apos;", "'")
  ProcedureReturn text
EndProcedure

; IsNumeric - 妫€鏌ユ槸鍚︿负鏁板瓧
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

; IsDate - 妫€鏌ユ槸鍚︿负鏃ユ湡鏍煎紡
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

; IsFormula - 妫€鏌ユ槸鍚︿负鍏紡
Procedure.b PbXls_IsFormula(text.s)
  text = Trim(text)
  If Len(text) > 1 And Mid(text, 1, 1) = "="
    ProcedureReturn #True
  EndIf
  ProcedureReturn #False
EndProcedure

; DateToExcel - PureBASIC鏃ユ湡杞珽xcel鏃ユ湡鏁板瓧
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

; ExcelToDate - Excel鏃ユ湡鏁板瓧杞琍ureBASIC鏃ユ湡
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

; GetCurrentDateTime - 鑾峰彇褰撳墠鏃ユ湡鏃堕棿
Procedure.s PbXls_GetCurrentDateTime()
  Define now.i = Date()
  ProcedureReturn FormatDate("%yyyy-%mm-%ddT%hh:%nn:%ss", now)
EndProcedure

; XML杈呭姪鍑芥暟
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

Procedure.s PbXls_ReadFileToString(filename.s)
  If FileSize(filename) = -1
    ProcedureReturn ""
  EndIf
  Define file.i = ReadFile(#PB_Any, filename)
  If file = 0
    ProcedureReturn ""
  EndIf
  Define content.s = ReadString(file, #PB_UTF8)
  CloseFile(file)
  ProcedureReturn content
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

; ZIP杈呭姪鍑芥暟
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
; 分区5: XML甯搁噺妯″潡
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
; 分区6: 鏍峰紡妯″潡
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
  Define fillId.i = ListSize(PbXls_Fills()) + 2
  AddElement(PbXls_Fills())
  PbXls_Fills()\patternType = "none"
  ProcedureReturn fillId
EndProcedure

Procedure.b PbXls_SetFill(fillId.i, patternType.s = "", fgColor.s = "", bgColor.s = "")
  ; fillId 偏移2（前两个是Excel保留的默认填充）
  Define listIdx.i = fillId - 2
  If listIdx < 0 Or listIdx >= ListSize(PbXls_Fills())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_Fills(), listIdx)
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
; 分区6.5: 数据验证模块
; ***************************************************************************************

; PbXls_CreateDataValidation - 创建数据验证规则
; 参数: type="list"(下拉列表)/"whole"(整数)/"decimal"(小数)/"textLength"(文本长度)/"date"/"time"/"custom"(公式)
; 返回: 验证规则ID(列表索引)
Procedure.i PbXls_CreateDataValidation(type.s, sqref.s, formula1.s = "", formula2.s = "", operator.s = "between")
  AddElement(PbXls_DataValidations())
  PbXls_DataValidations()\type = type
  PbXls_DataValidations()\sqref = sqref
  PbXls_DataValidations()\formula1 = formula1
  PbXls_DataValidations()\formula2 = formula2
  PbXls_DataValidations()\operator = operator
  PbXls_DataValidations()\showErrorMessage = #True
  PbXls_DataValidations()\showInputMessage = #True
  PbXls_DataValidations()\allowBlank = #True
  PbXls_DataValidations()\showDropDown = #False
  ProcedureReturn ListIndex(PbXls_DataValidations())
EndProcedure

; PbXls_SetValidationPrompt - 设置输入提示
Procedure.b PbXls_SetValidationPrompt(validationId.i, title.s, message.s)
  If validationId < 0 Or validationId >= ListSize(PbXls_DataValidations())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_DataValidations(), validationId)
  PbXls_DataValidations()\promptTitle = title
  PbXls_DataValidations()\prompt = message
  ProcedureReturn #True
EndProcedure

; PbXls_SetValidationError - 设置错误提示
Procedure.b PbXls_SetValidationError(validationId.i, title.s, message.s)
  If validationId < 0 Or validationId >= ListSize(PbXls_DataValidations())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_DataValidations(), validationId)
  PbXls_DataValidations()\errorTitle = title
  PbXls_DataValidations()\error = message
  ProcedureReturn #True
EndProcedure

; PbXls_SetValidationFlags - 设置验证标志
Procedure.b PbXls_SetValidationFlags(validationId.i, allowBlank.b = -1, showErrorMessage.b = -1, showInputMessage.b = -1, showDropDown.b = -1)
  If validationId < 0 Or validationId >= ListSize(PbXls_DataValidations())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_DataValidations(), validationId)
  If allowBlank >= 0 : PbXls_DataValidations()\allowBlank = allowBlank : EndIf
  If showErrorMessage >= 0 : PbXls_DataValidations()\showErrorMessage = showErrorMessage : EndIf
  If showInputMessage >= 0 : PbXls_DataValidations()\showInputMessage = showInputMessage : EndIf
  If showDropDown >= 0 : PbXls_DataValidations()\showDropDown = showDropDown : EndIf
  ProcedureReturn #True
EndProcedure

; ***************************************************************************************
; 分区6.6: 条件格式模块
; ***************************************************************************************

; PbXls_CreateConditionalFormat - 创建条件格式规则
; 参数: type="cellIs"(单元格值比较)/"expression"(公式)/"colorScale"(颜色刻度)/"dataBar"(数据条)/"iconSet"(图标集)
; 返回: 规则ID(列表索引)
Procedure.i PbXls_CreateConditionalFormat(type.s, sqref.s, formula1.s = "", formula2.s = "", operator.s = "greaterThan")
  AddElement(PbXls_ConditionalFormats())
  PbXls_ConditionalFormats()\type = type
  PbXls_ConditionalFormats()\sqref = sqref
  PbXls_ConditionalFormats()\formula1 = formula1
  PbXls_ConditionalFormats()\formula2 = formula2
  PbXls_ConditionalFormats()\operator = operator
  PbXls_ConditionalFormats()\priority = ListSize(PbXls_ConditionalFormats())
  PbXls_ConditionalFormats()\minType = "num"
  PbXls_ConditionalFormats()\maxType = "num"
  PbXls_ConditionalFormats()\minColor = "FFFF0000"
  PbXls_ConditionalFormats()\maxColor = "FF00CC00"
  PbXls_ConditionalFormats()\midColor = "FFFFFF00"
  PbXls_ConditionalFormats()\midType = "num"
  PbXls_DxfCounter + 1
  ProcedureReturn ListIndex(PbXls_ConditionalFormats())
EndProcedure

; PbXls_SetConditionalFormatDxf - 设置条件格式的差异样式
Procedure.b PbXls_SetConditionalFormatDxf(ruleId.i, fontColor.s = "", fillColor.s = "", fontBold.b = -1, fontItalic.b = -1)
  If ruleId < 0 Or ruleId >= ListSize(PbXls_ConditionalFormats())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_ConditionalFormats(), ruleId)
  If fontColor <> "" : PbXls_ConditionalFormats()\dxfFontColor = fontColor : EndIf
  If fillColor <> "" : PbXls_ConditionalFormats()\dxfFillColor = fillColor : EndIf
  If fontBold >= 0 : PbXls_ConditionalFormats()\dxfFontBold = fontBold : EndIf
  If fontItalic >= 0 : PbXls_ConditionalFormats()\dxfFontItalic = fontItalic : EndIf
  ProcedureReturn #True
EndProcedure

; PbXls_SetConditionalFormatColorScale - 设置颜色刻度参数
Procedure.b PbXls_SetConditionalFormatColorScale(ruleId.i, minColor.s = "", midColor.s = "", maxColor.s = "", minType.s = "", midType.s = "", maxType.s = "")
  If ruleId < 0 Or ruleId >= ListSize(PbXls_ConditionalFormats())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_ConditionalFormats(), ruleId)
  If minColor <> "" : PbXls_ConditionalFormats()\minColor = minColor : EndIf
  If midColor <> "" : PbXls_ConditionalFormats()\midColor = midColor : EndIf
  If maxColor <> "" : PbXls_ConditionalFormats()\maxColor = maxColor : EndIf
  If minType <> "" : PbXls_ConditionalFormats()\minType = minType : EndIf
  If midType <> "" : PbXls_ConditionalFormats()\midType = midType : EndIf
  If maxType <> "" : PbXls_ConditionalFormats()\maxType = maxType : EndIf
  ProcedureReturn #True
EndProcedure

; ***************************************************************************************
; 分区6.7: 图表模块
; ***************************************************************************************

; PbXls_CreateChart - 创建图表
; 参数: type="barChart"(柱状图)/"lineChart"(折线图)/"pieChart"(饼图)/"scatterChart"(散点图)/"areaChart"(面积图)
; 返回: 图表ID
Procedure.i PbXls_CreateChart(type.s, title.s = "", anchorRef.s = "A1:F20")
  AddElement(PbXls_Charts())
  PbXls_Charts()\type = type
  PbXls_Charts()\title = title
  PbXls_Charts()\anchorRef = anchorRef
  PbXls_Charts()\style = 2
  PbXls_Charts()\x = 0
  PbXls_Charts()\y = 0
  PbXls_Charts()\cx = 6000000
  PbXls_Charts()\cy = 4000000
  PbXls_ChartCounter + 1
  ProcedureReturn ListIndex(PbXls_Charts())
EndProcedure

; PbXls_AddChartSeries - 添加图表数据系列
Procedure.b PbXls_AddChartSeries(chartId.i, name.s, values.s, categories.s = "")
  If chartId < 0 Or chartId >= ListSize(PbXls_Charts())
    ProcedureReturn #False
  EndIf
  SelectElement(PbXls_Charts(), chartId)
  Define seriesKey.s = Str(chartId) + "_" + Str(PbXls_Charts()\seriesCount)
  PbXls_ChartSeriesName(seriesKey) = name
  PbXls_ChartSeriesValues(seriesKey) = values
  PbXls_ChartSeriesCategories(seriesKey) = categories
  PbXls_Charts()\seriesCount + 1
  ProcedureReturn #True
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

Procedure.b PbXls_SetCellHyperlink(*cell.PbXls_Cell, url.s, tooltip.s = "")
  *cell\hyperlink = url
  If tooltip <> ""
    *cell\comment = tooltip
  EndIf
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

; PbXls_SetCellStyleWS - 设置单元格样式（便捷版本，使用worksheet指针）
Procedure.b PbXls_SetCellStyleWS(worksheet.i, row.i, col.i, styleId.i)
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

; PbXls_SetPageMargins - 设置页边距 (参考openpyxl: worksheet/page.py PageMargins)
; left, right, top, bottom 单位: 英寸 (默认 0.7, 0.7, 0.75, 0.75)
; header, footer 单位: 英寸 (默认 0.3, 0.3)
Procedure.b PbXls_SetPageMargins(*ws.PbXls_Worksheet, left.f = 0.7, right.f = 0.7, top.f = 0.75, bottom.f = 0.75, header.f = 0.3, footer.f = 0.3)
  ; 存储到工作表结构中 (需要使用Map存储，因为PbXls_Worksheet没有这些字段)
  Define marginKey.s = "margins_" + Str(*ws\id)
  PbXls_ColumnWidths(marginKey + "_left") = left
  PbXls_ColumnWidths(marginKey + "_right") = right
  PbXls_ColumnWidths(marginKey + "_top") = top
  PbXls_ColumnWidths(marginKey + "_bottom") = bottom
  PbXls_ColumnWidths(marginKey + "_header") = header
  PbXls_ColumnWidths(marginKey + "_footer") = footer
  ProcedureReturn #True
EndProcedure

; PbXls_SetHeaderFooter - 设置页眉页脚 (参考openpyxl: worksheet/header_footer.py)
; header/footer 格式: &C居中 &L左 &R右, 例如: "&C标题 &L页码"
Procedure.b PbXls_SetHeaderFooter(*ws.PbXls_Worksheet, header.s = "", footer.s = "")
  Define hfKey.s = "hf_" + Str(*ws\id)
  PbXls_MergedCells(hfKey + "_header") = header
  PbXls_MergedCells(hfKey + "_footer") = footer
  ProcedureReturn #True
EndProcedure

; PbXls_SetPrintOptions - 设置打印选项 (参考openpyxl: worksheet/_writer.py write_print)
; gridLines: 显示网格线, headings: 显示行列标题, horizontalCentered: 水平居中, verticalCentered: 垂直居中
Procedure.b PbXls_SetPrintOptions(*ws.PbXls_Worksheet, gridLines.b = #True, headings.b = #False, horizontalCentered.b = #False, verticalCentered.b = #False)
  Define poKey.s = "print_" + Str(*ws\id)
  PbXls_MergedCells(poKey + "_gridLines") = Str(gridLines)
  PbXls_MergedCells(poKey + "_headings") = Str(headings)
  PbXls_MergedCells(poKey + "_hCentered") = Str(horizontalCentered)
  PbXls_MergedCells(poKey + "_vCentered") = Str(verticalCentered)
  ProcedureReturn #True
EndProcedure

; PbXls_SetOrientation - 设置页面方向 (portrait/landscape)
Procedure.b PbXls_SetOrientation(*ws.PbXls_Worksheet, orientation.s)
  *ws\orientation = LCase(orientation)
  ProcedureReturn #True
EndProcedure

; PbXls_SetPaperSize - 设置纸张大小 (9=A4, 1=Letter, 等)
Procedure.b PbXls_SetPaperSize(*ws.PbXls_Worksheet, paperSize.i)
  *ws\paperSize = paperSize
  ProcedureReturn #True
EndProcedure

; PbXls_SetCellComment - 设置单元格注释 (参考openpyxl: comments/comments.py)
; 注意: 当前版本注释只作为提示文本存储在单元格中
Procedure.b PbXls_SetCellComment(*ws.PbXls_Worksheet, row.i, col.i, content.s, author.s = "")
  If row < 1 Or col < 1
    ProcedureReturn #False
  EndIf
  Define *cell.PbXls_Cell = PbXls_GetCell(*ws, row, col)
  If *cell = 0
    ProcedureReturn #False
  EndIf
  ; 存储注释到单元格 (如果已有超链接，不覆盖)
  If *cell\hyperlink = ""
    If author <> ""
      *cell\comment = author + ":" + content
    Else
      *cell\comment = content
    EndIf
  EndIf
  ProcedureReturn #True
EndProcedure

; PbXls_UpdateRangeRows - 鏇存柊鑼冨洿涓殑琛屽彿(鐢ㄤ簬鎻掑叆/鍒犻櫎琛屾椂更新合并单元格
Procedure.s PbXls_UpdateRangeRows(rangeStr.s, rowIdx.i, count.i, isInsert.b)
  Define minCol.i, minRow.i, maxCol.i, maxRow.i
  Define minCol2.Integer, minRow2.Integer, maxCol2.Integer, maxRow2.Integer
  If PbXls_RangeBoundaries(rangeStr, @minCol2, @minRow2, @maxCol2, @maxRow2) = #False
    ProcedureReturn ""
  EndIf
  minCol = minCol2\i
  minRow = minRow2\i
  maxCol = maxCol2\i
  maxRow = maxRow2\i
  
  If isInsert
    If minRow >= rowIdx
      minRow = minRow + count
    EndIf
    If maxRow >= rowIdx
      maxRow = maxRow + count
    EndIf
  Else
    If minRow > rowIdx + count - 1
      minRow = minRow - count
    ElseIf minRow >= rowIdx
      minRow = rowIdx
    EndIf
    If maxRow > rowIdx + count - 1
      maxRow = maxRow - count
    ElseIf maxRow >= rowIdx
      maxRow = rowIdx + count - 1
    EndIf
    If maxRow < minRow
      ProcedureReturn ""
    EndIf
  EndIf
  
  ProcedureReturn PbXls_RangeString(minRow, minCol, maxRow, maxCol)
EndProcedure

; PbXls_UpdateRangeCols - 更新合并单元格范围中的列号
; 参考PbXls_UpdateRangeRows
Procedure.s PbXls_UpdateRangeCols(rangeStr.s, colIdx.i, count.i, isInsert.b)
  Define minCol.i, minRow.i, maxCol.i, maxRow.i
  Define minCol2.Integer, minRow2.Integer, maxCol2.Integer, maxRow2.Integer
  If PbXls_RangeBoundaries(rangeStr, @minCol2, @minRow2, @maxCol2, @maxRow2) = #False
    ProcedureReturn ""
  EndIf
  minCol = minCol2\i
  minRow = minRow2\i
  maxCol = maxCol2\i
  maxRow = maxRow2\i
  
  If isInsert
    If minCol >= colIdx
      minCol = minCol + count
    EndIf
    If maxCol >= colIdx
      maxCol = maxCol + count
    EndIf
  Else
    If minCol > colIdx + count - 1
      minCol = minCol - count
    ElseIf minCol >= colIdx
      minCol = colIdx
    EndIf
    If maxCol > colIdx + count - 1
      maxCol = maxCol - count
    ElseIf maxCol >= colIdx
      maxCol = colIdx + count - 1
    EndIf
    If maxCol < minCol
      ProcedureReturn ""
    EndIf
  EndIf
  
  ProcedureReturn PbXls_RangeString(minRow, minCol, maxRow, maxCol)
EndProcedure

; PbXls_InsertColumns - 在指定列前插入空列
; 参考openpyxl: worksheet/worksheet.py insert_cols方法
Procedure.b PbXls_InsertColumns(*ws.PbXls_Worksheet, colIdx.i, count.i = 1)
  If colIdx < 1 Or count < 1
    ProcedureReturn #False
  EndIf
  
  Define wsKey.s = Str(*ws\id) + "_"
  Define wsKeyLen.i = Len(wsKey)
  
  ; 1. 收集所有需要移动的单元格（按列号降序排序以避免覆盖）
  Define NewMap cellColMap.i()
  Define NewList cellsToMove.s()
  ForEach PbXls_AllCells()
    Define ck.s = MapKey(PbXls_AllCells())
    If Left(ck, Len(wsKey)) = wsKey
      Define cellPart.s = Mid(ck, wsKeyLen + 1)
      Define commaPos.i = FindString(cellPart, ",", 1)
      If commaPos > 0
        Define c.i = Val(Mid(cellPart, commaPos + 1))
        If c >= colIdx
          cellColMap(ck) = c
        EndIf
      EndIf
    EndIf
  Next
  
  ; 按列号降序排序
  Define maxPossibleCol.i = *ws\maxColumn + count + 100
  For i = maxPossibleCol To colIdx Step -1
    ForEach cellColMap()
      If cellColMap() = i
        AddElement(cellsToMove())
        cellsToMove() = MapKey(cellColMap())
      EndIf
    Next
  Next
  
  ; 移动单元格
  ResetList(cellsToMove())
  While NextElement(cellsToMove())
    Define oldKey.s = cellsToMove()
    Define oldCellPart.s = Mid(oldKey, wsKeyLen + 1)
    Define oldCommaPos.i = FindString(oldCellPart, ",", 1)
    Define oldRow.i = Val(Left(oldCellPart, oldCommaPos - 1))
    Define oldCol.i = Val(Mid(oldCellPart, oldCommaPos + 1))
    
    Define newCol.i = oldCol + count
    Define newKey.s = wsKey + Str(oldRow) + "," + Str(newCol)
    
    ; 创建新的单元格并复制数据
    Define *newCell.PbXls_Cell = @PbXls_AllCells(newKey)
    CopyStructure(@PbXls_AllCells(oldKey), *newCell, PbXls_Cell)
    *newCell\column = newCol
    
    ; 删除旧的单元格
    DeleteMapElement(PbXls_AllCells(), oldKey)
  Wend
  
  ; 2. 更新列宽
  Define NewMap cwColMap.i()
  Define NewList colsToMove.s()
  ForEach PbXls_ColumnWidths()
    Define cwKey.s = MapKey(PbXls_ColumnWidths())
    If Left(cwKey, Len(wsKey)) = wsKey
      Define cwPart.s = Mid(cwKey, wsKeyLen + 1)
      Define c.i = Val(cwPart)
      If c >= colIdx
        cwColMap(cwKey) = c
      EndIf
    EndIf
  Next
  
  ; 按列号降序排序
  For i = maxPossibleCol To colIdx Step -1
    ForEach cwColMap()
      If cwColMap() = i
        AddElement(colsToMove())
        colsToMove() = MapKey(cwColMap())
      EndIf
    Next
  Next
  
  ResetList(colsToMove())
  While NextElement(colsToMove())
    Define oldCwKey.s = colsToMove()
    Define oldCwVal.i = Val(Mid(oldCwKey, wsKeyLen + 1))
    Define newCwKey.s = wsKey + Str(oldCwVal + count)
    PbXls_ColumnWidths(newCwKey) = PbXls_ColumnWidths(oldCwKey)
    DeleteMapElement(PbXls_ColumnWidths(), oldCwKey)
  Wend
  
  ; 3. 更新合并单元格
  ForEach PbXls_MergedCells()
    Define mk.s = MapKey(PbXls_MergedCells())
    If Left(mk, Len(wsKey)) = wsKey
      Define rangeStr.s = PbXls_MergedCells()
      Define newRange.s = PbXls_UpdateRangeCols(rangeStr, colIdx, count, #True)
      If newRange <> ""
        Define newMk.s = wsKey + UCase(newRange)
        PbXls_MergedCells(newMk) = UCase(newRange)
        DeleteMapElement(PbXls_MergedCells(), mk)
      EndIf
    EndIf
  Next
  
  ; 4. 更新工作表最大列
  *ws\maxColumn + count
  
  ProcedureReturn #True
EndProcedure

; PbXls_DeleteColumns - 删除指定列
; 参考openpyxl: worksheet/worksheet.py delete_cols方法
Procedure.b PbXls_DeleteColumns(*ws.PbXls_Worksheet, colIdx.i, count.i = 1)
  If colIdx < 1 Or count < 1
    ProcedureReturn #False
  EndIf
  
  Define wsKey.s = Str(*ws\id) + "_"
  Define wsKeyLen.i = Len(wsKey)
  Define maxCol.i = colIdx + count - 1
  
  ; 1. 收集需要删除的单元格键值
  Define NewList deleteCells.s()
  ForEach PbXls_AllCells()
    Define ck.s = MapKey(PbXls_AllCells())
    If Left(ck, Len(wsKey)) = wsKey
      Define cellPart.s = Mid(ck, wsKeyLen + 1)
      Define commaPos.i = FindString(cellPart, ",", 1)
      If commaPos > 0
        Define c.i = Val(Mid(cellPart, commaPos + 1))
        If c >= colIdx And c <= maxCol
          AddElement(deleteCells())
          deleteCells() = ck
        EndIf
      EndIf
    EndIf
  Next
  
  ; 批量删除
  ResetList(deleteCells())
  While NextElement(deleteCells())
    DeleteMapElement(PbXls_AllCells(), deleteCells())
  Wend
  
  ; 2. 移动右侧的单元格（向前移动）
  Define NewMap cellColMap.i()
  Define NewList cellsToMove.s()
  ForEach PbXls_AllCells()
    Define ck.s = MapKey(PbXls_AllCells())
    If Left(ck, Len(wsKey)) = wsKey
      Define cellPart.s = Mid(ck, wsKeyLen + 1)
      Define commaPos.i = FindString(cellPart, ",", 1)
      If commaPos > 0
        Define c.i = Val(Mid(cellPart, commaPos + 1))
        If c > maxCol
          cellColMap(ck) = c
        EndIf
      EndIf
    EndIf
  Next
  
  ; 按列号升序排序
  Define maxPossibleCol.i = *ws\maxColumn
  For i = colIdx To maxPossibleCol
    ForEach cellColMap()
      If cellColMap() = i
        AddElement(cellsToMove())
        cellsToMove() = MapKey(cellColMap())
      EndIf
    Next
  Next
  
  ResetList(cellsToMove())
  While NextElement(cellsToMove())
    Define oldKey.s = cellsToMove()
    Define oldCellPart.s = Mid(oldKey, wsKeyLen + 1)
    Define oldCommaPos.i = FindString(oldCellPart, ",", 1)
    Define oldRow.i = Val(Left(oldCellPart, oldCommaPos - 1))
    Define oldCol.i = Val(Mid(oldCellPart, oldCommaPos + 1))
    
    Define newCol.i = oldCol - count
    Define newKey.s = wsKey + Str(oldRow) + "," + Str(newCol)
    
    ; 创建新的单元格并复制数据
    Define *newCell.PbXls_Cell = @PbXls_AllCells(newKey)
    CopyStructure(@PbXls_AllCells(oldKey), *newCell, PbXls_Cell)
    *newCell\column = newCol
    
    ; 删除旧的单元格
    DeleteMapElement(PbXls_AllCells(), oldKey)
  Wend
  
  ; 3. 更新列宽
  ; 收集并删除范围内的列宽
  Define NewList deleteCw.s()
  ForEach PbXls_ColumnWidths()
    Define cwKey.s = MapKey(PbXls_ColumnWidths())
    If Left(cwKey, Len(wsKey)) = wsKey
      Define cwPart.s = Mid(cwKey, wsKeyLen + 1)
      Define c.i = Val(cwPart)
      If c >= colIdx And c <= maxCol
        AddElement(deleteCw())
        deleteCw() = cwKey
      EndIf
    EndIf
  Next
  ResetList(deleteCw())
  While NextElement(deleteCw())
    DeleteMapElement(PbXls_ColumnWidths(), deleteCw())
  Wend
  
  ; 移动右侧的列宽
  Define NewMap cwColMap.i()
  Define NewList colsToMove.s()
  ForEach PbXls_ColumnWidths()
    Define cwKey.s = MapKey(PbXls_ColumnWidths())
    If Left(cwKey, Len(wsKey)) = wsKey
      Define cwPart.s = Mid(cwKey, wsKeyLen + 1)
      Define c.i = Val(cwPart)
      If c > maxCol
        cwColMap(cwKey) = c
      EndIf
    EndIf
  Next
  
  For i = colIdx To maxPossibleCol
    ForEach cwColMap()
      If cwColMap() = i
        AddElement(colsToMove())
        colsToMove() = MapKey(cwColMap())
      EndIf
    Next
  Next
  
  ResetList(colsToMove())
  While NextElement(colsToMove())
    Define oldCwKey.s = colsToMove()
    Define oldCwVal.i = Val(Mid(oldCwKey, wsKeyLen + 1))
    Define newCwKey.s = wsKey + Str(oldCwVal - count)
    PbXls_ColumnWidths(newCwKey) = PbXls_ColumnWidths(oldCwKey)
    DeleteMapElement(PbXls_ColumnWidths(), oldCwKey)
  Wend
  
  ; 4. 更新合并单元格
  ForEach PbXls_MergedCells()
    Define mk.s = MapKey(PbXls_MergedCells())
    If Left(mk, Len(wsKey)) = wsKey
      Define rangeStr.s = PbXls_MergedCells()
      Define newRange.s = PbXls_UpdateRangeCols(rangeStr, colIdx, count, #False)
      If newRange <> ""
        Define newMk.s = wsKey + UCase(newRange)
        PbXls_MergedCells(newMk) = UCase(newRange)
        DeleteMapElement(PbXls_MergedCells(), mk)
      Else
        DeleteMapElement(PbXls_MergedCells(), mk)
      EndIf
    EndIf
  Next
  
  ; 5. 更新工作表最大列
  *ws\maxColumn - count
  If *ws\maxColumn < 1
    *ws\maxColumn = 1
  EndIf
  
  ProcedureReturn #True
EndProcedure

; PbXls_InsertRows - 鍦ㄦ寚瀹氳鍓嶆彃鍏ョ┖琛?
; 参考openpyxl: worksheet/worksheet.py insert_rows鏂规硶
Procedure.b PbXls_InsertRows(*ws.PbXls_Worksheet, rowIdx.i, count.i = 1)
  If rowIdx < 1 Or count < 1
    ProcedureReturn #False
  EndIf
  
  Define wsKey.s = Str(*ws\id) + "_"
  Define wsKeyLen.i = Len(wsKey)
  
  ; 收集所有需要移动的单元格鎸夎鍙烽檷搴忔帓搴忎互閬垮厤瑕嗙洊)
  Define NewMap cellRowMap.i()
  Define NewList cellsToMove.s()
  ForEach PbXls_AllCells()
    Define ck.s = MapKey(PbXls_AllCells())
    If Left(ck, Len(wsKey)) = wsKey
      Define cellPart.s = Mid(ck, wsKeyLen + 1)
      Define commaPos.i = FindString(cellPart, ",", 1)
      If commaPos > 0
        Define r.i = Val(Left(cellPart, commaPos - 1))
        If r >= rowIdx
          cellRowMap(ck) = r
        EndIf
      EndIf
    EndIf
  Next
  
  ; 鎸夎鍙烽檷搴忔帓搴?- 使用纯Basic方式
  Define maxPossibleRow.i = *ws\maxRow + count + 100
  For i = maxPossibleRow To rowIdx Step -1
    ForEach cellRowMap()
      If cellRowMap() = i
        AddElement(cellsToMove())
        cellsToMove() = MapKey(cellRowMap())
      EndIf
    Next
  Next
  
  ; 移动单元格
  ResetList(cellsToMove())
  While NextElement(cellsToMove())
    Define oldKey.s = cellsToMove()
    Define oldCellPart.s = Mid(oldKey, wsKeyLen + 1)
    Define oldCommaPos.i = FindString(oldCellPart, ",", 1)
    Define oldRow.i = Val(Left(oldCellPart, oldCommaPos - 1))
    Define oldCol.i = Val(Mid(oldCellPart, oldCommaPos + 1))
    
    Define newRow.i = oldRow + count
    Define newKey.s = wsKey + Str(newRow) + "," + Str(oldCol)
    
    ; 鍒涘缓鏂扮殑鍗曞厓鏍煎苟澶嶅埗鏁版嵁
    Define *newCell.PbXls_Cell = @PbXls_AllCells(newKey)
    CopyStructure(@PbXls_AllCells(oldKey), *newCell, PbXls_Cell)
    *newCell\row = newRow
    
    ; 删除旧的单元格
    DeleteMapElement(PbXls_AllCells(), oldKey)
  Wend
  
  ; 鏇存柊琛岄珮
  Define NewMap rhRowMap.i()
  Define NewList rowsToMove.s()
  ForEach PbXls_RowHeights()
    Define rhKey.s = MapKey(PbXls_RowHeights())
    If Left(rhKey, Len(wsKey)) = wsKey
      Define rhPart.s = Mid(rhKey, wsKeyLen + 1)
      Define rh.i = Val(rhPart)
      If rh >= rowIdx
        rhRowMap(rhKey) = rh
      EndIf
    EndIf
  Next
  
  ; 鎸夎鍙烽檷搴忔帓搴?
  For i = maxPossibleRow To rowIdx Step -1
    ForEach rhRowMap()
      If rhRowMap() = i
        AddElement(rowsToMove())
        rowsToMove() = MapKey(rhRowMap())
      EndIf
    Next
  Next
  
  ResetList(rowsToMove())
  While NextElement(rowsToMove())
    Define oldRhKey.s = rowsToMove()
    Define oldRhVal.i = Val(Mid(oldRhKey, wsKeyLen + 1))
    Define newRhKey.s = wsKey + Str(oldRhVal + count)
    PbXls_RowHeights(newRhKey) = PbXls_RowHeights(oldRhKey)
    DeleteMapElement(PbXls_RowHeights(), oldRhKey)
  Wend
  
  ; 更新合并单元格
  ForEach PbXls_MergedCells()
    Define mk.s = MapKey(PbXls_MergedCells())
    If Left(mk, Len(wsKey)) = wsKey
      Define rangeStr.s = PbXls_MergedCells()
      Define newRange.s = PbXls_UpdateRangeRows(rangeStr, rowIdx, count, #True)
      If newRange <> ""
        Define newMk.s = wsKey + UCase(newRange)
        PbXls_MergedCells(newMk) = UCase(newRange)
        DeleteMapElement(PbXls_MergedCells(), mk)
      EndIf
    EndIf
  Next
  
  ; 鏇存柊宸ヤ綔琛ㄦ渶澶ц
  *ws\maxRow + count
  
  ProcedureReturn #True
EndProcedure

; PbXls_DeleteRows - 删除指定行
; 参考openpyxl: worksheet/worksheet.py delete_rows鏂规硶
Procedure.b PbXls_DeleteRows(*ws.PbXls_Worksheet, rowIdx.i, count.i = 1)
  If rowIdx < 1 Or count < 1
    ProcedureReturn #False
  EndIf
  
  Define wsKey.s = Str(*ws\id) + "_"
  Define wsKeyLen.i = Len(wsKey)
  Define maxRow.i = rowIdx + count - 1
  
  ; 删除范围内的单元格
  ForEach PbXls_AllCells()
    Define ck.s = MapKey(PbXls_AllCells())
    If Left(ck, Len(wsKey)) = wsKey
      Define cellPart.s = Mid(ck, wsKeyLen + 1)
      Define commaPos.i = FindString(cellPart, ",", 1)
      If commaPos > 0
        Define r.i = Val(Left(cellPart, commaPos - 1))
        If r >= rowIdx And r <= maxRow
          DeleteMapElement(PbXls_AllCells(), ck)
        ElseIf r > maxRow
          ; 需要向下移动的单元格

          Define newR.i = r - count
          Define newKey.s = wsKey + Str(newR) + "," + Mid(cellPart, commaPos + 1)
          Define *newCell.PbXls_Cell = @PbXls_AllCells(newKey)
          CopyStructure(@PbXls_AllCells(ck), *newCell, PbXls_Cell)
          *newCell\row = newR
          DeleteMapElement(PbXls_AllCells(), ck)
        EndIf
      EndIf
    EndIf
  Next
  
  ; 鍒犻櫎/绉诲姩琛岄珮
  ForEach PbXls_RowHeights()
    Define rhKey.s = MapKey(PbXls_RowHeights())
    If Left(rhKey, Len(wsKey)) = wsKey
      Define rh.i = Val(Mid(rhKey, wsKeyLen + 1))
      If rh >= rowIdx And rh <= maxRow
        DeleteMapElement(PbXls_RowHeights(), rhKey)
      ElseIf rh > maxRow
        Define newRhKey.s = wsKey + Str(rh - count)
        PbXls_RowHeights(newRhKey) = PbXls_RowHeights(rhKey)
        DeleteMapElement(PbXls_RowHeights(), rhKey)
      EndIf
    EndIf
  Next
  
  ; 更新合并单元格
  ForEach PbXls_MergedCells()
    Define mk.s = MapKey(PbXls_MergedCells())
    If Left(mk, Len(wsKey)) = wsKey
      Define rangeStr.s = PbXls_MergedCells()
      Define newRange.s = PbXls_UpdateRangeRows(rangeStr, rowIdx, count, #False)
      If newRange = ""
        ; 鍒犻櫎瀹屽叏鍦ㄨ鍒犺寖鍥村唴鐨勫悎骞跺崟鍏冩牸
        DeleteMapElement(PbXls_MergedCells(), mk)
      Else
        Define newMk.s = wsKey + UCase(newRange)
        PbXls_MergedCells(newMk) = UCase(newRange)
        DeleteMapElement(PbXls_MergedCells(), mk)
      EndIf
    EndIf
  Next
  
  ; 鏇存柊宸ヤ綔琛ㄦ渶澶ц
  If *ws\maxRow > maxRow
    *ws\maxRow - count
  Else
    *ws\maxRow = rowIdx - 1
  EndIf
  
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
  
  ; 数据验证 - 参考openpyxl: worksheet/datavalidation.py
  Define dvCount.i = 0
  ForEach PbXls_DataValidations()
    dvCount + 1
  Next
  If dvCount > 0
    Define dataValidationsNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "dataValidations")
    PbXls_XMLSetAttribute(dataValidationsNode, "count", Str(dvCount))
    ForEach PbXls_DataValidations()
      Define dvNode.i = PbXls_XMLAddNode(xmlId, dataValidationsNode, "dataValidation")
      PbXls_XMLSetAttribute(dvNode, "sqref", PbXls_DataValidations()\sqref)
      If PbXls_DataValidations()\type <> ""
        PbXls_XMLSetAttribute(dvNode, "type", PbXls_DataValidations()\type)
      EndIf
      If PbXls_DataValidations()\operator <> ""
        PbXls_XMLSetAttribute(dvNode, "operator", PbXls_DataValidations()\operator)
      EndIf
      If PbXls_DataValidations()\allowBlank
        PbXls_XMLSetAttribute(dvNode, "allowBlank", "1")
      EndIf
      If PbXls_DataValidations()\showInputMessage
        PbXls_XMLSetAttribute(dvNode, "showInputMessage", "1")
      EndIf
      If PbXls_DataValidations()\showErrorMessage
        PbXls_XMLSetAttribute(dvNode, "showErrorMessage", "1")
      EndIf
      If PbXls_DataValidations()\showDropDown
        PbXls_XMLSetAttribute(dvNode, "showDropDown", "1")
      EndIf
      If PbXls_DataValidations()\promptTitle <> ""
        Define promptTitleNode.i = PbXls_XMLAddNode(xmlId, dvNode, "promptTitle")
        PbXls_XMLSetText(promptTitleNode, PbXls_DataValidations()\promptTitle)
      EndIf
      If PbXls_DataValidations()\prompt <> ""
        Define promptNode.i = PbXls_XMLAddNode(xmlId, dvNode, "prompt")
        PbXls_XMLSetText(promptNode, PbXls_DataValidations()\prompt)
      EndIf
      If PbXls_DataValidations()\errorTitle <> ""
        Define errorTitleNode.i = PbXls_XMLAddNode(xmlId, dvNode, "errorTitle")
        PbXls_XMLSetText(errorTitleNode, PbXls_DataValidations()\errorTitle)
      EndIf
      If PbXls_DataValidations()\error <> ""
        Define errorNode.i = PbXls_XMLAddNode(xmlId, dvNode, "error")
        PbXls_XMLSetText(errorNode, PbXls_DataValidations()\error)
      EndIf
      If PbXls_DataValidations()\formula1 <> ""
        Define f1Node.i = PbXls_XMLAddNode(xmlId, dvNode, "formula1")
        PbXls_XMLSetText(f1Node, PbXls_DataValidations()\formula1)
      EndIf
      If PbXls_DataValidations()\formula2 <> ""
        Define f2Node.i = PbXls_XMLAddNode(xmlId, dvNode, "formula2")
        PbXls_XMLSetText(f2Node, PbXls_DataValidations()\formula2)
      EndIf
    Next
  EndIf
  
  ; 条件格式 - 参考openpyxl: formatting/formatting.py
  Define cfCount.i = 0
  ForEach PbXls_ConditionalFormats()
    cfCount + 1
  Next
  If cfCount > 0
    ForEach PbXls_ConditionalFormats()
      Define cfNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "conditionalFormatting")
      PbXls_XMLSetAttribute(cfNode, "sqref", PbXls_ConditionalFormats()\sqref)
      Define cfRuleNode.i = PbXls_XMLAddNode(xmlId, cfNode, "cfRule")
      PbXls_XMLSetAttribute(cfRuleNode, "type", PbXls_ConditionalFormats()\type)
      PbXls_XMLSetAttribute(cfRuleNode, "priority", Str(PbXls_ConditionalFormats()\priority))
      If PbXls_ConditionalFormats()\operator <> ""
        PbXls_XMLSetAttribute(cfRuleNode, "operator", PbXls_ConditionalFormats()\operator)
      EndIf
      If PbXls_ConditionalFormats()\formula1 <> ""
        Define cfFormulaNode.i = PbXls_XMLAddNode(xmlId, cfRuleNode, "formula")
        PbXls_XMLSetText(cfFormulaNode, PbXls_ConditionalFormats()\formula1)
      EndIf
      ; 处理colorScale类型
      If PbXls_ConditionalFormats()\type = "colorScale"
        If PbXls_ConditionalFormats()\minColor <> ""
          Define colorScaleNode.i = PbXls_XMLAddNode(xmlId, cfRuleNode, "colorScale")
          Define cfvo1.i = PbXls_XMLAddNode(xmlId, colorScaleNode, "cfvo")
          PbXls_XMLSetAttribute(cfvo1, "type", PbXls_ConditionalFormats()\minType)
          PbXls_XMLSetAttribute(cfvo1, "val", "0")
          Define color1.i = PbXls_XMLAddNode(xmlId, colorScaleNode, "color")
          PbXls_XMLSetAttribute(color1, "rgb", PbXls_FormatColor(PbXls_ConditionalFormats()\minColor))
          If PbXls_ConditionalFormats()\midColor <> ""
            Define cfvo2.i = PbXls_XMLAddNode(xmlId, colorScaleNode, "cfvo")
            PbXls_XMLSetAttribute(cfvo2, "type", PbXls_ConditionalFormats()\midType)
            PbXls_XMLSetAttribute(cfvo2, "val", "0")
            Define color2.i = PbXls_XMLAddNode(xmlId, colorScaleNode, "color")
            PbXls_XMLSetAttribute(color2, "rgb", PbXls_FormatColor(PbXls_ConditionalFormats()\midColor))
          EndIf
          Define cfvo3.i = PbXls_XMLAddNode(xmlId, colorScaleNode, "cfvo")
          PbXls_XMLSetAttribute(cfvo3, "type", PbXls_ConditionalFormats()\maxType)
          PbXls_XMLSetAttribute(cfvo3, "val", "0")
          Define color3.i = PbXls_XMLAddNode(xmlId, colorScaleNode, "color")
          PbXls_XMLSetAttribute(color3, "rgb", PbXls_FormatColor(PbXls_ConditionalFormats()\maxColor))
        EndIf
      EndIf
    Next
  EndIf
  
  If *ws\autoFilter <> ""
    Define autoFilterNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "autoFilter")
    PbXls_XMLSetAttribute(autoFilterNode, "ref", *ws\autoFilter)
  EndIf
  
  ; 超链接支持- 参考openpyxl: worksheet/hyperlink.py
  Define hlCount.i = 0
  ForEach PbXls_AllCells()
    Define hlKey.s = MapKey(PbXls_AllCells())
    If Left(hlKey, Len(wsIdStr)) = wsIdStr
      If PbXls_AllCells()\hyperlink <> ""
        hlCount + 1
      EndIf
    EndIf
  Next
  
  If hlCount > 0
    Define hyperlinksNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "hyperlinks")
    Define hlId.i = 1
    ForEach PbXls_AllCells()
      Define hlKey2.s = MapKey(PbXls_AllCells())
      If Left(hlKey2, Len(wsIdStr)) = wsIdStr
        If PbXls_AllCells()\hyperlink <> ""
          Define hlNode.i = PbXls_XMLAddNode(xmlId, hyperlinksNode, "hyperlink")
          Define hlRef.s = PbXls_CoordinateFromRowCol(PbXls_AllCells()\row, PbXls_AllCells()\column)
          PbXls_XMLSetAttribute(hlNode, "ref", hlRef)
          PbXls_XMLSetAttribute(hlNode, "id", "rId" + Str(hlId))
          PbXls_XMLSetAttribute(hlNode, "display", PbXls_AllCells()\hyperlink)
          If PbXls_AllCells()\comment <> ""
            PbXls_XMLSetAttribute(hlNode, "tooltip", PbXls_AllCells()\comment)
          EndIf
          hlId + 1
        EndIf
      EndIf
    Next
  EndIf
  
  ; 娉ㄩ噴鏀寔 - 参考openpyxl: comments/comment_sheet.py
  Define commentCount.i = 0
  ForEach PbXls_AllCells()
    Define cmKey.s = MapKey(PbXls_AllCells())
    If Left(cmKey, Len(wsIdStr)) = wsIdStr
      If PbXls_AllCells()\comment <> "" And PbXls_AllCells()\hyperlink = ""
        commentCount + 1
      EndIf
    EndIf
  Next
  
  If commentCount > 0
    Define extLstNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "extLst")
    Define extNode.i = PbXls_XMLAddNode(xmlId, extLstNode, "ext")
    PbXls_XMLSetAttribute(extNode, "uri", "{24D24A12-3E60-46E9-9998-0616A0593193}")
    PbXls_XMLAddNode(xmlId, extNode, "x14:author")
  EndIf
  
  ; 椤甸潰璁剧疆 - 参考openpyxl: worksheet/page.py
  Define marginKey2.s = "margins_" + Str(*ws\id)
  Define ml.f, mr.f, mt.f, mb.f, mh.f, mf.f
  If FindMapElement(PbXls_ColumnWidths(), marginKey2 + "_left")
    ml = PbXls_ColumnWidths(marginKey2 + "_left")
  Else
    ml = 0.7
  EndIf
  If FindMapElement(PbXls_ColumnWidths(), marginKey2 + "_right")
    mr = PbXls_ColumnWidths(marginKey2 + "_right")
  Else
    mr = 0.7
  EndIf
  If FindMapElement(PbXls_ColumnWidths(), marginKey2 + "_top")
    mt = PbXls_ColumnWidths(marginKey2 + "_top")
  Else
    mt = 0.75
  EndIf
  If FindMapElement(PbXls_ColumnWidths(), marginKey2 + "_bottom")
    mb = PbXls_ColumnWidths(marginKey2 + "_bottom")
  Else
    mb = 0.75
  EndIf
  If FindMapElement(PbXls_ColumnWidths(), marginKey2 + "_header")
    mh = PbXls_ColumnWidths(marginKey2 + "_header")
  Else
    mh = 0.3
  EndIf
  If FindMapElement(PbXls_ColumnWidths(), marginKey2 + "_footer")
    mf = PbXls_ColumnWidths(marginKey2 + "_footer")
  Else
    mf = 0.3
  EndIf
  
  Define pageMarginsNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "pageMargins")
  PbXls_XMLSetAttribute(pageMarginsNode, "left", StrF(ml, 2))
  PbXls_XMLSetAttribute(pageMarginsNode, "right", StrF(mr, 2))
  PbXls_XMLSetAttribute(pageMarginsNode, "top", StrF(mt, 2))
  PbXls_XMLSetAttribute(pageMarginsNode, "bottom", StrF(mb, 2))
  PbXls_XMLSetAttribute(pageMarginsNode, "header", StrF(mh, 2))
  PbXls_XMLSetAttribute(pageMarginsNode, "footer", StrF(mf, 2))
  
  ; 页眉页脚
  Define hfKey2.s = "hf_" + Str(*ws\id)
  Define hasHF.b = #False
  Define hfHeader.s = "", hfFooter.s = ""
  If FindMapElement(PbXls_MergedCells(), hfKey2 + "_header")
    hfHeader = PbXls_MergedCells(hfKey2 + "_header")
    If hfHeader <> ""
      hasHF = #True
    EndIf
  EndIf
  If FindMapElement(PbXls_MergedCells(), hfKey2 + "_footer")
    hfFooter = PbXls_MergedCells(hfKey2 + "_footer")
    If hfFooter <> ""
      hasHF = #True
    EndIf
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
  
  ; 页眉页脚节点
  If hasHF
    Define hfNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "headerFooter")
    If hfHeader <> ""
      Define oddHeaderNode.i = PbXls_XMLAddNode(xmlId, hfNode, "oddHeader")
      PbXls_XMLSetText(oddHeaderNode, hfHeader)
    EndIf
    If hfFooter <> ""
      Define oddFooterNode.i = PbXls_XMLAddNode(xmlId, hfNode, "oddFooter")
      PbXls_XMLSetText(oddFooterNode, hfFooter)
    EndIf
  EndIf
  
  ; 鎵撳嵃閫夐」 - 参考openpyxl: worksheet/_writer.py write_print
  Define poKey2.s = "print_" + Str(*ws\id)
  Define gl.s = "1", hd.s = "0", hc.s = "0", vc.s = "0"
  If FindMapElement(PbXls_MergedCells(), poKey2 + "_gridLines")
    gl = PbXls_MergedCells(poKey2 + "_gridLines")
  EndIf
  If FindMapElement(PbXls_MergedCells(), poKey2 + "_headings")
    hd = PbXls_MergedCells(poKey2 + "_headings")
  EndIf
  If FindMapElement(PbXls_MergedCells(), poKey2 + "_hCentered")
    hc = PbXls_MergedCells(poKey2 + "_hCentered")
  EndIf
  If FindMapElement(PbXls_MergedCells(), poKey2 + "_vCentered")
    vc = PbXls_MergedCells(poKey2 + "_vCentered")
  EndIf
  
  Define printOptionsNode.i = PbXls_XMLAddNode(xmlId, worksheetNode, "printOptions")
  If gl = "1"
    PbXls_XMLSetAttribute(printOptionsNode, "gridLinesSet", "1")
  EndIf
  If hd = "1"
    PbXls_XMLSetAttribute(printOptionsNode, "headings", "1")
  EndIf
  If hc = "1"
    PbXls_XMLSetAttribute(printOptionsNode, "horizontalCentered", "1")
  EndIf
  If vc = "1"
    PbXls_XMLSetAttribute(printOptionsNode, "verticalCentered", "1")
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
  
  Define overrideTheme.i = PbXls_XMLAddNode(xmlId, typesNode, "Override")
  PbXls_XMLSetAttribute(overrideTheme, "PartName", "/xl/theme/theme1.xml")
  PbXls_XMLSetAttribute(overrideTheme, "ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")
  
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

; PbXls_WriteStylesXML - 鐢熸垚鏍峰紡琛╔ML (styles.xml)
; 参考openpyxl: styles/stylesheet.py 涓殑 write_stylesheet() 鍑芥暟
; styles.xml 缁撴瀯: numFmts -> fonts -> fills -> borders -> cellStyleXfs -> cellXfs -> cellStyles -> dxfs -> tableStyles
Procedure.i PbXls_WriteStylesXML(*wb.PbXls_Workbook)
  Define xmlId.i = PbXls_XMLCreateDocument()
  If xmlId = 0
    ProcedureReturn 0
  EndIf
  
  Define rootNode.i = PbXls_XMLGetRoot(xmlId)
  Define styleSheetNode.i = PbXls_XMLAddNode(xmlId, rootNode, "styleSheet")
  PbXls_XMLSetAttribute(styleSheetNode, "xmlns", #PbXls_SHEET_MAIN_NS$)
  
  ; 1. numFmts - 自定义数字格式 (从164开始)
  Define numFmtsNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "numFmts")
  PbXls_XMLSetAttribute(numFmtsNode, "count", "0")
  
  ; 2. fonts - 瀛椾綋鍒楄〃
  Define fontsNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "fonts")
  Define fontCount.i = ListSize(PbXls_Fonts())
  If fontCount = 0
    fontCount = 1 ; 纭繚鑷冲皯鏈変竴涓粯璁ゅ瓧浣
  EndIf
  PbXls_XMLSetAttribute(fontsNode, "count", Str(fontCount))
  
  If ListSize(PbXls_Fonts()) = 0
    ; 娣诲姞榛樿瀛椾綋
    Define defaultFontNode.i = PbXls_XMLAddNode(xmlId, fontsNode, "font")
    Define szNode1.i = PbXls_XMLAddNode(xmlId, defaultFontNode, "sz")
    PbXls_XMLSetAttribute(szNode1, "val", "11")
    Define nameNode1.i = PbXls_XMLAddNode(xmlId, defaultFontNode, "name")
    PbXls_XMLSetAttribute(nameNode1, "val", "Calibri")
    Define familyNode1.i = PbXls_XMLAddNode(xmlId, defaultFontNode, "family")
    PbXls_XMLSetAttribute(familyNode1, "val", "2")
    Define schemeNode1.i = PbXls_XMLAddNode(xmlId, defaultFontNode, "scheme")
    PbXls_XMLSetAttribute(schemeNode1, "val", "minor")
    Define colorNode1.i = PbXls_XMLAddNode(xmlId, defaultFontNode, "color")
    PbXls_XMLSetAttribute(colorNode1, "theme", "1")
  Else
    Define fontIdx.i = 0
    ForEach PbXls_Fonts()
      Define fontNode.i = PbXls_XMLAddNode(xmlId, fontsNode, "font")
      
      ; 瀛椾綋鍚嶇О
      If PbXls_Fonts()\name <> ""
        Define fNameNode.i = PbXls_XMLAddNode(xmlId, fontNode, "name")
        PbXls_XMLSetAttribute(fNameNode, "val", PbXls_Fonts()\name)
      EndIf
      
      ; 瀛椾綋澶у皬
      If PbXls_Fonts()\size > 0
        Define fSzNode.i = PbXls_XMLAddNode(xmlId, fontNode, "sz")
        PbXls_XMLSetAttribute(fSzNode, "val", StrF(PbXls_Fonts()\size, 1))
      EndIf
      
      ; 瀛椾綋棰滆壊
      If PbXls_Fonts()\color <> ""
        Define fColorNode.i = PbXls_XMLAddNode(xmlId, fontNode, "color")
        PbXls_XMLSetAttribute(fColorNode, "rgb", "FF" + PbXls_ParseColor(PbXls_Fonts()\color))
      EndIf
      
      ; 绮椾綋
      If PbXls_Fonts()\bold
        Define fBoldNode.i = PbXls_XMLAddNode(xmlId, fontNode, "b")
        PbXls_XMLSetAttribute(fBoldNode, "val", "1")
      EndIf
      
      ; 鏂滀綋
      If PbXls_Fonts()\italic
        Define fItalicNode.i = PbXls_XMLAddNode(xmlId, fontNode, "i")
        PbXls_XMLSetAttribute(fItalicNode, "val", "1")
      EndIf
      
      ; 下划线
      If PbXls_Fonts()\underline
        Define fUnderlineNode.i = PbXls_XMLAddNode(xmlId, fontNode, "u")
        PbXls_XMLSetAttribute(fUnderlineNode, "val", "single")
      EndIf
      
      ; 删除线
      If PbXls_Fonts()\strike
        Define fStrikeNode.i = PbXls_XMLAddNode(xmlId, fontNode, "strike")
        PbXls_XMLSetAttribute(fStrikeNode, "val", "1")
      EndIf
      
      ; 字体系列
      If PbXls_Fonts()\family > 0
        Define fFamilyNode.i = PbXls_XMLAddNode(xmlId, fontNode, "family")
        PbXls_XMLSetAttribute(fFamilyNode, "val", Str(PbXls_Fonts()\family))
      EndIf
      
      ; 字体方案
      If PbXls_Fonts()\scheme <> ""
        Define fSchemeNode.i = PbXls_XMLAddNode(xmlId, fontNode, "scheme")
        PbXls_XMLSetAttribute(fSchemeNode, "val", PbXls_Fonts()\scheme)
      EndIf
      
      fontIdx + 1
    Next
  EndIf
  
  ; 3. fills - 填充列表 (必须包含至少两个默认填充: none 和 gray125)
  Define fillsNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "fills")
  ; Excel要求: 前两个填充是保留的(none和gray125)，用户填充从索引2开始
  Define fillCount.i = ListSize(PbXls_Fills()) + 2
  PbXls_XMLSetAttribute(fillsNode, "count", Str(fillCount))
  
  ; 添加默认填充1: patternType="none"
  Define defaultFill1Node.i = PbXls_XMLAddNode(xmlId, fillsNode, "fill")
  Define pf1Node.i = PbXls_XMLAddNode(xmlId, defaultFill1Node, "patternFill")
  PbXls_XMLSetAttribute(pf1Node, "patternType", "none")
  
  ; 添加默认填充2: patternType="gray125"
  Define defaultFill2Node.i = PbXls_XMLAddNode(xmlId, fillsNode, "fill")
  Define pf2Node.i = PbXls_XMLAddNode(xmlId, defaultFill2Node, "patternFill")
  PbXls_XMLSetAttribute(pf2Node, "patternType", "gray125")
  
  ; 添加用户自定义填充
  If ListSize(PbXls_Fills()) > 0
    ForEach PbXls_Fills()
      Define fillNode.i = PbXls_XMLAddNode(xmlId, fillsNode, "fill")
      Define pfNode.i = PbXls_XMLAddNode(xmlId, fillNode, "patternFill")
      
      If PbXls_Fills()\patternType <> "" And PbXls_Fills()\patternType <> "none" And PbXls_Fills()\patternType <> "gray125"
        PbXls_XMLSetAttribute(pfNode, "patternType", PbXls_Fills()\patternType)
        
        If PbXls_Fills()\fgColor <> ""
          Define fgNode.i = PbXls_XMLAddNode(xmlId, pfNode, "fgColor")
          PbXls_XMLSetAttribute(fgNode, "rgb", PbXls_FormatColor(PbXls_Fills()\fgColor))
        EndIf
        
        If PbXls_Fills()\bgColor <> ""
          Define bgNode.i = PbXls_XMLAddNode(xmlId, pfNode, "bgColor")
          PbXls_XMLSetAttribute(bgNode, "rgb", PbXls_FormatColor(PbXls_Fills()\bgColor))
        EndIf
      EndIf
    Next
  EndIf
  
  ; 4. borders - 边框列表
  Define bordersNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "borders")
  Define borderCount.i = ListSize(PbXls_Borders())
  If borderCount = 0
    borderCount = 1 ; 确保至少有一个默认边框
  EndIf
  PbXls_XMLSetAttribute(bordersNode, "count", Str(borderCount))
  
  If ListSize(PbXls_Borders()) = 0
    ; 添加默认边框 (所有边都是none)
    Define defaultBorderNode.i = PbXls_XMLAddNode(xmlId, bordersNode, "border")
    Define leftNode1.i = PbXls_XMLAddNode(xmlId, defaultBorderNode, "left")
    Define rightNode1.i = PbXls_XMLAddNode(xmlId, defaultBorderNode, "right")
    Define topNode1.i = PbXls_XMLAddNode(xmlId, defaultBorderNode, "top")
    Define bottomNode1.i = PbXls_XMLAddNode(xmlId, defaultBorderNode, "bottom")
    Define diagNode1.i = PbXls_XMLAddNode(xmlId, defaultBorderNode, "diagonal")
  Else
    Define borderIdx.i = 0
    ForEach PbXls_Borders()
      Define borderNode.i = PbXls_XMLAddNode(xmlId, bordersNode, "border")
      
      ; 左边框
      Define bLeftNode.i = PbXls_XMLAddNode(xmlId, borderNode, "left")
      If PbXls_Borders()\left\style <> "" And PbXls_Borders()\left\style <> "none"
        PbXls_XMLSetAttribute(bLeftNode, "style", PbXls_Borders()\left\style)
        Define bLeftColor.i = PbXls_XMLAddNode(xmlId, bLeftNode, "color")
        PbXls_XMLSetAttribute(bLeftColor, "rgb", PbXls_FormatColor(PbXls_Borders()\left\color))
      EndIf
      
      ; 右边框
      Define bRightNode.i = PbXls_XMLAddNode(xmlId, borderNode, "right")
      If PbXls_Borders()\right\style <> "" And PbXls_Borders()\right\style <> "none"
        PbXls_XMLSetAttribute(bRightNode, "style", PbXls_Borders()\right\style)
        Define bRightColor.i = PbXls_XMLAddNode(xmlId, bRightNode, "color")
        PbXls_XMLSetAttribute(bRightColor, "rgb", PbXls_FormatColor(PbXls_Borders()\right\color))
      EndIf
      
      ; 上边框
      Define bTopNode.i = PbXls_XMLAddNode(xmlId, borderNode, "top")
      If PbXls_Borders()\top\style <> "" And PbXls_Borders()\top\style <> "none"
        PbXls_XMLSetAttribute(bTopNode, "style", PbXls_Borders()\top\style)
        Define bTopColor.i = PbXls_XMLAddNode(xmlId, bTopNode, "color")
        PbXls_XMLSetAttribute(bTopColor, "rgb", PbXls_FormatColor(PbXls_Borders()\top\color))
      EndIf
      
      ; 下边框
      Define bBottomNode.i = PbXls_XMLAddNode(xmlId, borderNode, "bottom")
      If PbXls_Borders()\bottom\style <> "" And PbXls_Borders()\bottom\style <> "none"
        PbXls_XMLSetAttribute(bBottomNode, "style", PbXls_Borders()\bottom\style)
        Define bBottomColor.i = PbXls_XMLAddNode(xmlId, bBottomNode, "color")
        PbXls_XMLSetAttribute(bBottomColor, "rgb", PbXls_FormatColor(PbXls_Borders()\bottom\color))
      EndIf
      
      ; 瀵硅绾胯竟妗
      Define bDiagNode.i = PbXls_XMLAddNode(xmlId, borderNode, "diagonal")
      If PbXls_Borders()\diagonal\style <> "" And PbXls_Borders()\diagonal\style <> "none"
        PbXls_XMLSetAttribute(bDiagNode, "style", PbXls_Borders()\diagonal\style)
        Define bDiagColor.i = PbXls_XMLAddNode(xmlId, bDiagNode, "color")
        PbXls_XMLSetAttribute(bDiagColor, "rgb", PbXls_FormatColor(PbXls_Borders()\diagonal\color))
        If PbXls_Borders()\diagonalUp
          PbXls_XMLSetAttribute(borderNode, "diagonalUp", "1")
        EndIf
        If PbXls_Borders()\diagonalDown
          PbXls_XMLSetAttribute(borderNode, "diagonalDown", "1")
        EndIf
      EndIf
      
      borderIdx + 1
    Next
  EndIf
  
  ; 5. cellStyleXfs - 榛樿鏍峰紡鏍煎紡 (蹇呴』鑷冲皯鏈変竴涓

  Define cellStyleXfsNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "cellStyleXfs")
  PbXls_XMLSetAttribute(cellStyleXfsNode, "count", "1")
  Define defaultXfNode.i = PbXls_XMLAddNode(xmlId, cellStyleXfsNode, "xf")
  PbXls_XMLSetAttribute(defaultXfNode, "numFmtId", "0")
  PbXls_XMLSetAttribute(defaultXfNode, "fontId", "0")
  PbXls_XMLSetAttribute(defaultXfNode, "fillId", "0")
  PbXls_XMLSetAttribute(defaultXfNode, "borderId", "0")
  
  ; 6. cellXfs - 单元格样式应用
  Define cellXfsNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "cellXfs")
  Define cellXfCount.i = ListSize(PbXls_CellStyles())
  If cellXfCount = 0
    cellXfCount = 1 ; 纭繚鑷冲皯鏈変竴涓粯璁ゆ牱寮
  EndIf
  PbXls_XMLSetAttribute(cellXfsNode, "count", Str(cellXfCount))
  
  If ListSize(PbXls_CellStyles()) = 0
    ; 娣诲姞榛樿鍗曞厓鏍兼牱寮
    Define defaultCellXf.i = PbXls_XMLAddNode(xmlId, cellXfsNode, "xf")
    PbXls_XMLSetAttribute(defaultCellXf, "numFmtId", "0")
    PbXls_XMLSetAttribute(defaultCellXf, "fontId", "0")
    PbXls_XMLSetAttribute(defaultCellXf, "fillId", "0")
    PbXls_XMLSetAttribute(defaultCellXf, "borderId", "0")
    PbXls_XMLSetAttribute(defaultCellXf, "xfId", "0")
  Else
    Define cellStyleIdx.i = 0
    ForEach PbXls_CellStyles()
      Define cellXfNode.i = PbXls_XMLAddNode(xmlId, cellXfsNode, "xf")
      PbXls_XMLSetAttribute(cellXfNode, "numFmtId", Str(PbXls_CellStyles()\numFmtId))
      PbXls_XMLSetAttribute(cellXfNode, "fontId", Str(PbXls_CellStyles()\fontId))
      PbXls_XMLSetAttribute(cellXfNode, "fillId", Str(PbXls_CellStyles()\fillId))
      PbXls_XMLSetAttribute(cellXfNode, "borderId", Str(PbXls_CellStyles()\borderId))
      PbXls_XMLSetAttribute(cellXfNode, "xfId", "0")
      
      ; 检查是否需要 applyAlignment
      If PbXls_CellStyles()\alignment\horizontal <> "general" Or
         PbXls_CellStyles()\alignment\vertical <> "bottom" Or
         PbXls_CellStyles()\alignment\wrapText Or
         PbXls_CellStyles()\alignment\shrinkToFit Or
         PbXls_CellStyles()\alignment\indent > 0 Or
         PbXls_CellStyles()\alignment\textRotation > 0
        PbXls_XMLSetAttribute(cellXfNode, "applyAlignment", "1")
        
        Define alignNode.i = PbXls_XMLAddNode(xmlId, cellXfNode, "alignment")
        If PbXls_CellStyles()\alignment\horizontal <> "general"
          PbXls_XMLSetAttribute(alignNode, "horizontal", PbXls_CellStyles()\alignment\horizontal)
        EndIf
        If PbXls_CellStyles()\alignment\vertical <> "bottom"
          PbXls_XMLSetAttribute(alignNode, "vertical", PbXls_CellStyles()\alignment\vertical)
        EndIf
        If PbXls_CellStyles()\alignment\wrapText
          PbXls_XMLSetAttribute(alignNode, "wrapText", "1")
        EndIf
        If PbXls_CellStyles()\alignment\shrinkToFit
          PbXls_XMLSetAttribute(alignNode, "shrinkToFit", "1")
        EndIf
        If PbXls_CellStyles()\alignment\indent > 0
          PbXls_XMLSetAttribute(alignNode, "indent", Str(PbXls_CellStyles()\alignment\indent))
        EndIf
        If PbXls_CellStyles()\alignment\textRotation > 0
          PbXls_XMLSetAttribute(alignNode, "textRotation", Str(PbXls_CellStyles()\alignment\textRotation))
        EndIf
      EndIf
      
      ; 检查是否需要 applyProtection
      If PbXls_CellStyles()\locked Or PbXls_CellStyles()\hidden
        PbXls_XMLSetAttribute(cellXfNode, "applyProtection", "1")
        
        Define protNode.i = PbXls_XMLAddNode(xmlId, cellXfNode, "protection")
        If PbXls_CellStyles()\locked
          PbXls_XMLSetAttribute(protNode, "locked", "1")
        EndIf
        If PbXls_CellStyles()\hidden
          PbXls_XMLSetAttribute(protNode, "hidden", "1")
        EndIf
      EndIf
      
      cellStyleIdx + 1
    Next
  EndIf
  
  ; 7. cellStyles - 命名样式 (至少需要 Normal)
  Define cellStylesNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "cellStyles")
  PbXls_XMLSetAttribute(cellStylesNode, "count", "1")
  Define normalStyleNode.i = PbXls_XMLAddNode(xmlId, cellStylesNode, "cellStyle")
  PbXls_XMLSetAttribute(normalStyleNode, "name", "Normal")
  PbXls_XMLSetAttribute(normalStyleNode, "xfId", "0")
  PbXls_XMLSetAttribute(normalStyleNode, "builtinId", "0")
  
  ; 8. dxfs - 差异格式 (条件格式用，暂时为空)
  Define dxfsNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "dxfs")
  PbXls_XMLSetAttribute(dxfsNode, "count", "0")
  
  ; 9. tableStyles - 表格样式 (暂时为空)
  Define tableStylesNode.i = PbXls_XMLAddNode(xmlId, styleSheetNode, "tableStyles")
  PbXls_XMLSetAttribute(tableStylesNode, "count", "0")
  PbXls_XMLSetAttribute(tableStylesNode, "defaultPivotStyle", "PivotStyleLight16")
  PbXls_XMLSetAttribute(tableStylesNode, "defaultTableStyle", "TableStyleMedium9")
  
  ProcedureReturn xmlId
EndProcedure

; PbXls_GetThemeXML - 获取固定主题XML字符串
; 参考openpyxl: writer/theme.py — 使用固定的theme字符串
; theme.xml 定义了Office主题的颜色方案、字体方案和格式方案
Procedure.s PbXls_GetThemeXML()
  ProcedureReturn ~"<?xml version=\"1.0\"?>" +
                  ~"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\">" +
                  ~"<a:themeElements>" +
                  ~"<a:clrScheme name=\"Office\">" +
                  ~"<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
                  ~"<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                  ~"<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>" +
                  ~"<a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
                  ~"<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>" +
                  ~"<a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2>" +
                  ~"<a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3>" +
                  ~"<a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4>" +
                  ~"<a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5>" +
                  ~"<a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>" +
                  ~"<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>" +
                  ~"<a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>" +
                  ~"</a:clrScheme>" +
                  ~"<a:fontScheme name=\"Office\">" +
                  ~"<a:majorFont>" +
                  ~"<a:latin typeface=\"Cambria\"/>" +
                  ~"<a:ea typeface=\"\"/>" +
                  ~"<a:cs typeface=\"\"/>" +
                  ~"<a:font script=\"Jpan\" typeface=\"&#xFF2D;&#xFF33; &#xFF30;&#x30B4;&#x30B7;&#x30C3;&#x30AF;\"/>" +
                  ~"<a:font script=\"Hang\" typeface=\"&#xB9D1;&#xC740; &#xACE0;&#xB515;\"/>" +
                  ~"<a:font script=\"Hans\" typeface=\"&#x5B8B;&#x4F53;\"/>" +
                  ~"<a:font script=\"Hant\" typeface=\"&#x65B0;&#x7D30;&#x660E;&#x9AD4;\"/>" +
                  ~"<a:font script=\"Arab\" typeface=\"Times New Roman\"/>" +
                  ~"<a:font script=\"Hebr\" typeface=\"Times New Roman\"/>" +
                  ~"<a:font script=\"Thai\" typeface=\"Tahoma\"/>" +
                  ~"<a:font script=\"Ethi\" typeface=\"Nyala\"/>" +
                  ~"<a:font script=\"Beng\" typeface=\"Vrinda\"/>" +
                  ~"<a:font script=\"Gujr\" typeface=\"Shruti\"/>" +
                  ~"<a:font script=\"Khmr\" typeface=\"MoolBoran\"/>" +
                  ~"<a:font script=\"Knda\" typeface=\"Tunga\"/>" +
                  ~"<a:font script=\"Guru\" typeface=\"Raavi\"/>" +
                  ~"<a:font script=\"Cans\" typeface=\"Euphemia\"/>" +
                  ~"<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>" +
                  ~"<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>" +
                  ~"<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>" +
                  ~"<a:font script=\"Thaa\" typeface=\"MV Boli\"/>" +
                  ~"<a:font script=\"Deva\" typeface=\"Mangal\"/>" +
                  ~"<a:font script=\"Telu\" typeface=\"Gautami\"/>" +
                  ~"<a:font script=\"Taml\" typeface=\"Latha\"/>" +
                  ~"<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>" +
                  ~"<a:font script=\"Orya\" typeface=\"Kalinga\"/>" +
                  ~"<a:font script=\"Mlym\" typeface=\"Kartika\"/>" +
                  ~"<a:font script=\"Laoo\" typeface=\"DokChampa\"/>" +
                  ~"<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>" +
                  ~"<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>" +
                  ~"<a:font script=\"Viet\" typeface=\"Times New Roman\"/>" +
                  ~"<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>" +
                  ~"</a:majorFont>" +
                  ~"<a:minorFont>" +
                  ~"<a:latin typeface=\"Calibri\"/>" +
                  ~"<a:ea typeface=\"\"/>" +
                  ~"<a:cs typeface=\"\"/>" +
                  ~"<a:font script=\"Jpan\" typeface=\"&#xFF2D;&#xFF33; &#xFF30;&#x30B4;&#x30B7;&#x30C3;&#x30AF;\"/>" +
                  ~"<a:font script=\"Hang\" typeface=\"&#xB9D1;&#xC740; &#xACE0;&#xB515;\"/>" +
                  ~"<a:font script=\"Hans\" typeface=\"&#x5B8B;&#x4F53;\"/>" +
                  ~"<a:font script=\"Hant\" typeface=\"&#x65B0;&#x7D30;&#x660E;&#x9AD4;\"/>" +
                  ~"<a:font script=\"Arab\" typeface=\"Arial\"/>" +
                  ~"<a:font script=\"Hebr\" typeface=\"Arial\"/>" +
                  ~"<a:font script=\"Thai\" typeface=\"Tahoma\"/>" +
                  ~"<a:font script=\"Ethi\" typeface=\"Nyala\"/>" +
                  ~"<a:font script=\"Beng\" typeface=\"Vrinda\"/>" +
                  ~"<a:font script=\"Gujr\" typeface=\"Shruti\"/>" +
                  ~"<a:font script=\"Khmr\" typeface=\"DaunPenh\"/>" +
                  ~"<a:font script=\"Knda\" typeface=\"Tunga\"/>" +
                  ~"<a:font script=\"Guru\" typeface=\"Raavi\"/>" +
                  ~"<a:font script=\"Cans\" typeface=\"Euphemia\"/>" +
                  ~"<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>" +
                  ~"<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>" +
                  ~"<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>" +
                  ~"<a:font script=\"Thaa\" typeface=\"MV Boli\"/>" +
                  ~"<a:font script=\"Deva\" typeface=\"Mangal\"/>" +
                  ~"<a:font script=\"Telu\" typeface=\"Gautami\"/>" +
                  ~"<a:font script=\"Taml\" typeface=\"Latha\"/>" +
                  ~"<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>" +
                  ~"<a:font script=\"Orya\" typeface=\"Kalinga\"/>" +
                  ~"<a:font script=\"Mlym\" typeface=\"Kartika\"/>" +
                  ~"<a:font script=\"Laoo\" typeface=\"DokChampa\"/>" +
                  ~"<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>" +
                  ~"<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>" +
                  ~"<a:font script=\"Viet\" typeface=\"Arial\"/>" +
                  ~"<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>" +
                  ~"</a:minorFont>" +
                  ~"</a:fontScheme>" +
                  ~"<a:fmtScheme name=\"Office\">" +
                  ~"<a:fillStyleLst>" +
                  ~"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                  ~"<a:gradFill rotWithShape=\"1\">" +
                  ~"<a:gsLst>" +
                  ~"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"35000\"><a:schemeClr val=\"phClr\"><a:tint val=\"37000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:tint val=\"15000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>" +
                  ~"</a:gsLst>" +
                  ~"<a:lin ang=\"16200000\" scaled=\"1\"/>" +
                  ~"</a:gradFill>" +
                  ~"<a:gradFill rotWithShape=\"1\">" +
                  ~"<a:gsLst>" +
                  ~"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:shade val=\"51000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"80000\"><a:schemeClr val=\"phClr\"><a:shade val=\"93000\"/><a:satMod val=\"130000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"94000\"/><a:satMod val=\"135000\"/></a:schemeClr></a:gs>" +
                  ~"</a:gsLst>" +
                  ~"<a:lin ang=\"16200000\" scaled=\"0\"/>" +
                  ~"</a:gradFill>" +
                  ~"</a:fillStyleLst>" +
                  ~"<a:lnStyleLst>" +
                  ~"<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" +
                  ~"<a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"95000\"/><a:satMod val=\"105000\"/></a:schemeClr></a:solidFill>" +
                  ~"<a:prstDash val=\"solid\"/>" +
                  ~"</a:ln>" +
                  ~"<a:ln w=\"25400\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" +
                  ~"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                  ~"<a:prstDash val=\"solid\"/>" +
                  ~"</a:ln>" +
                  ~"<a:ln w=\"38100\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">" +
                  ~"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                  ~"<a:prstDash val=\"solid\"/>" +
                  ~"</a:ln>" +
                  ~"</a:lnStyleLst>" +
                  ~"<a:effectStyleLst>" +
                  ~"<a:effectStyle>" +
                  ~"<a:effectLst>" +
                  ~"<a:outerShdw blurRad=\"40000\" dist=\"20000\" dir=\"5400000\" rotWithShape=\"0\">" +
                  ~"<a:srgbClr val=\"000000\"><a:alpha val=\"38000\"/></a:srgbClr>" +
                  ~"</a:outerShdw>" +
                  ~"</a:effectLst>" +
                  ~"</a:effectStyle>" +
                  ~"<a:effectStyle>" +
                  ~"<a:effectLst>" +
                  ~"<a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\">" +
                  ~"<a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr>" +
                  ~"</a:outerShdw>" +
                  ~"</a:effectLst>" +
                  ~"</a:effectStyle>" +
                  ~"<a:effectStyle>" +
                  ~"<a:effectLst>" +
                  ~"<a:outerShdw blurRad=\"40000\" dist=\"23000\" dir=\"5400000\" rotWithShape=\"0\">" +
                  ~"<a:srgbClr val=\"000000\"><a:alpha val=\"35000\"/></a:srgbClr>" +
                  ~"</a:outerShdw>" +
                  ~"</a:effectLst>" +
                  ~"<a:scene3d>" +
                  ~"<a:camera prst=\"orthographicFront\"><a:rot lat=\"0\" lon=\"0\" rev=\"0\"/></a:camera>" +
                  ~"<a:lightRig rig=\"threePt\" dir=\"t\"><a:rot lat=\"0\" lon=\"0\" rev=\"1200000\"/></a:lightRig>" +
                  ~"</a:scene3d>" +
                  ~"<a:sp3d><a:bevelT w=\"63500\" h=\"25400\"/></a:sp3d>" +
                  ~"</a:effectStyle>" +
                  ~"</a:effectStyleLst>" +
                  ~"<a:bgFillStyleLst>" +
                  ~"<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                  ~"<a:gradFill rotWithShape=\"1\">" +
                  ~"<a:gsLst>" +
                  ~"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"40000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"40000\"><a:schemeClr val=\"phClr\"><a:tint val=\"45000\"/><a:shade val=\"99000\"/><a:satMod val=\"350000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"20000\"/><a:satMod val=\"255000\"/></a:schemeClr></a:gs>" +
                  ~"</a:gsLst>" +
                  ~"<a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"-80000\" r=\"50000\" b=\"180000\"/></a:path>" +
                  ~"</a:gradFill>" +
                  ~"<a:gradFill rotWithShape=\"1\">" +
                  ~"<a:gsLst>" +
                  ~"<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"80000\"/><a:satMod val=\"300000\"/></a:schemeClr></a:gs>" +
                  ~"<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"30000\"/><a:satMod val=\"200000\"/></a:schemeClr></a:gs>" +
                  ~"</a:gsLst>" +
                  ~"<a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path>" +
                  ~"</a:gradFill>" +
                  ~"</a:bgFillStyleLst>" +
                  ~"</a:fmtScheme>" +
                  ~"</a:themeElements>" +
                  ~"<a:objectDefaults/>" +
                  ~"<a:extraClrSchemeLst/>" +
                  ~"</a:theme>"
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
  
  Define relTheme.i = PbXls_XMLAddNode(xmlId, relsNode, "Relationship")
  PbXls_XMLSetAttribute(relTheme, "Id", "rId" + Str(wsIdx2 + 3))
  PbXls_XMLSetAttribute(relTheme, "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
  PbXls_XMLSetAttribute(relTheme, "Target", "theme/theme1.xml")
  
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
  
  ; UTF8缂栫爜姣忎釜瀛楃鏈€澶氶渶瑕?涓瓧鑺傦紝鍒嗛厤瓒冲澶х殑缂撳啿鍖

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

; PbXls_SaveWorkbookToFile - 淇濆瓨宸ヤ綔绨垮埌文件
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
  
  ; 8. xl/styles.xml
  Define stylesXml.i = PbXls_WriteStylesXML(*wb)
  If stylesXml
    PbXls_AddXMLToZIP(packId, stylesXml, #PbXls_ARCStyles$)
  EndIf
  
  ; 8.5. xl/theme/theme1.xml
  Define themeXmlStr.s = PbXls_GetThemeXML()
  If themeXmlStr <> ""
    Define themeBufferSize.i = Len(themeXmlStr) * 4 + 10
    Define *themeBuffer = AllocateMemory(themeBufferSize)
    If *themeBuffer
      PokeA(*themeBuffer, $EF)
      PokeA(*themeBuffer + 1, $BB)
      PokeA(*themeBuffer + 2, $BF)
      Define themeStrLen.i = PokeS(*themeBuffer + 3, themeXmlStr, themeBufferSize - 3, #PB_UTF8)
      Define themeTotalSize.i = 3 + themeStrLen
      PbXls_ZIPAddMemory(packId, *themeBuffer, themeTotalSize, #PbXls_ARCTheme$)
      FreeMemory(*themeBuffer)
    EndIf
  EndIf
  
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
; 分区11: XML读取器（解析Excel文件各部分XML）
;    参考openpyxl: reader/excel.py, reader/strings.py, reader/workbook.py
;    使用PureBASIC内置XML库解析Excel xlsx文件中的XML内容
; ***************************************************************************************

; PbXls_ExtractToTempDir - 将ZIP文件解压到临时目录
; 参考openpyxl: 使用ZipFile直接读取ZIP内文件
Procedure.s PbXls_ExtractToTempDir(filename.s)
  Define tempDir.s = GetTemporaryDirectory() + "PbXls_read_" + Str(Random(999999)) + "\"
  CreateDirectory(tempDir)
  
  Define packId.i = PbXls_ZIPOpen(filename)
  If packId = 0
    Debug "[Extract] 无法打开ZIP文件: " + filename
    ProcedureReturn ""
  EndIf
  
  If PbXls_ZIPExamine(packId) = #False
    ClosePack(packId)
    Debug "[Extract] 无法解析ZIP文件"
    ProcedureReturn ""
  EndIf
  
  Define entryCount.i = 0
  While PbXls_ZIPNextEntry(packId)
    Define entryName.s = PbXls_ZIPEntryName(packId)
    If entryName <> ""
      entryCount + 1
      ; 替换正斜杠为反斜杠
      entryName = ReplaceString(entryName, "/", "\")
      Define fullPath.s = tempDir + entryName
      
      ; 查找最后一个反斜杠位置
      Define lastSlash.i = 0
      Define searchPos.i = 1
      While searchPos <= Len(fullPath)
        If Mid(fullPath, searchPos, 1) = "\"
          lastSlash = searchPos
        EndIf
        searchPos + 1
      Wend
      
      If lastSlash > 0
        Define dirPath.s = Left(fullPath, lastSlash - 1)
        CreateDirectory(dirPath)
        
        ; 解压文件
        Define entrySize.i = PbXls_ZIPEntrySize(packId)
        If entrySize > 0
          Define *buffer = AllocateMemory(entrySize)
          If *buffer
            PbXls_ZIPExtractToMemory(packId, *buffer, entrySize)
            Define fOut.i = CreateFile(#PB_Any, fullPath)
            If fOut
              WriteData(fOut, *buffer, entrySize)
              CloseFile(fOut)
            EndIf
            FreeMemory(*buffer)
          EndIf
        Else
          ; 空文件也创建一下
          Define fOut2.i = CreateFile(#PB_Any, fullPath)
          If fOut2
            CloseFile(fOut2)
          EndIf
        EndIf
      EndIf
    EndIf
  Wend
  
  Debug "[Extract] 共解压 " + Str(entryCount) + " 个条目到 " + tempDir
  ClosePack(packId)
  ProcedureReturn tempDir
EndProcedure

; PbXls_CleanupTempDir - 清理临时目录
Procedure.b PbXls_CleanupTempDir(tempDir.s)
  ; 简单清理: 删除临时目录下的文件
  Define searchId.i = ExamineDirectory(#PB_Any, tempDir, "*.*")
  If searchId
    While NextDirectoryEntry(searchId)
      If DirectoryEntryType(searchId) = #PB_DirectoryEntry_File
        Define fileName.s = DirectoryEntryName(searchId)
        DeleteFile(tempDir + fileName)
      ElseIf DirectoryEntryType(searchId) = #PB_DirectoryEntry_Directory
        If DirectoryEntryName(searchId) <> "." And DirectoryEntryName(searchId) <> ".."
          PbXls_CleanupTempDir(tempDir + DirectoryEntryName(searchId) + "\")
        EndIf
      EndIf
    Wend
    FinishDirectory(searchId)
    ; PureBuilt-in没有RemoveDirectory，使用API
    RemoveDirectory_(tempDir)
  EndIf
  ProcedureReturn #True
EndProcedure

; PbXls_GetNodeNameWithoutNS - 获取节点名(去除命名空间前缀)
Procedure.s PbXls_GetNodeNameWithoutNS(node.i)
  Define fullName.s = GetXMLNodeName(node)
  ; 查找冒号位置，返回冒号后面的部分
  Define colonPos.i = FindString(fullName, ":", 1)
  If colonPos > 0
    ProcedureReturn Mid(fullName, colonPos + 1)
  EndIf
  ProcedureReturn fullName
EndProcedure

; PbXls_ReadSharedStrings - 读取共享字符串表 (xl/sharedStrings.xml)
; 参考openpyxl: reader/strings.py
; 输入: XML文档ID
; 输出: 字符串List (按索引顺序)
Procedure.b PbXls_ReadSharedStrings(xmlId.i, List stringList.s())
  If xmlId = 0
    ProcedureReturn #False
  EndIf
  
  ClearList(stringList())
  
  ; 递归遍历XML树查找si节点
  Define xmlStr.s = PbXls_XMLSaveToString(xmlId)
  If xmlStr = ""
    ProcedureReturn #False
  EndIf
  
  ; 简单方式: 按文本解析
  Define pos.i = 1
  While pos <= Len(xmlStr)
    ; 查找 <t> 标记
    Define tStart.i = FindString(xmlStr, "<t>", pos)
    If tStart = 0
      Break
    EndIf
    Define tEnd.i = FindString(xmlStr, "</t>", tStart + 3)
    If tEnd = 0
      Break
    EndIf
    Define tContent.s = Mid(xmlStr, tStart + 3, tEnd - tStart - 3)
    tContent = PbXls_UnescapeXML(tContent)
    AddElement(stringList())
    stringList() = tContent
    pos = tEnd + 4
  Wend
  
  ProcedureReturn #True
EndProcedure

; PbXls_ReadWorkbookInfo - 读取工作簿信息 (xl/workbook.xml)
; 参考openpyxl: reader/workbook.py
; 输出: 工作表名称List 和 activeSheetIndex
Procedure.b PbXls_ReadWorkbookInfo(xmlId.i, List sheetNames.s(), *activeSheetIndex.Integer)
  If xmlId = 0
    ProcedureReturn #False
  EndIf
  
  ClearList(sheetNames())
  
  Define rootNode.i = RootXMLNode(xmlId)
  Define workbookNode.i = ChildXMLNode(rootNode)
  While workbookNode And PbXls_GetNodeNameWithoutNS(workbookNode) <> "workbook"
    workbookNode = NextXMLNode(workbookNode)
  Wend
  
  If workbookNode = 0
    ProcedureReturn #False
  EndIf
  
  ; 读取activeTab (从bookViews中)
  Define bookViewsNode.i = ChildXMLNode(workbookNode)
  While bookViewsNode
    If PbXls_GetNodeNameWithoutNS(bookViewsNode) = "bookViews"
      Define workbookViewNode.i = ChildXMLNode(bookViewsNode)
      While workbookViewNode
        If PbXls_GetNodeNameWithoutNS(workbookViewNode) = "workbookView"
          Define activeTab.s = GetXMLAttribute(workbookViewNode, "activeTab")
          If activeTab <> ""
            *activeSheetIndex\i = Val(activeTab)
          EndIf
          Break
        EndIf
        workbookViewNode = NextXMLNode(workbookViewNode)
      Wend
      Break
    EndIf
    bookViewsNode = NextXMLNode(bookViewsNode)
  Wend
  
  ; 读取sheets节点
  Define sheetsNode.i = ChildXMLNode(workbookNode)
  While sheetsNode
    If PbXls_GetNodeNameWithoutNS(sheetsNode) = "sheets"
      Define sheetNode.i = ChildXMLNode(sheetsNode)
      While sheetNode
        If PbXls_GetNodeNameWithoutNS(sheetNode) = "sheet"
          Define sheetName.s = GetXMLAttribute(sheetNode, "name")
          If sheetName <> ""
            AddElement(sheetNames())
            sheetNames() = sheetName
          EndIf
        EndIf
        sheetNode = NextXMLNode(sheetNode)
      Wend
      Break
    EndIf
    sheetsNode = NextXMLNode(sheetsNode)
  Wend
  
  ProcedureReturn #True
EndProcedure

; PbXls_ReadWorksheetData - 读取工作表单元格数据 (xl/worksheets/sheetN.xml)
; 参考openpyxl: worksheet/_reader.py
; 输入: XML文档ID, 工作表ID, 共享字符串List
; 输出: 填充PbXls_AllCells Map
Procedure.b PbXls_ReadWorksheetData(xmlId.i, wsId.i, List sharedStrings.s())
  If xmlId = 0
    ProcedureReturn #False
  EndIf
  
  Define rootNode.i = RootXMLNode(xmlId)
  Define worksheetNode.i = ChildXMLNode(rootNode)
  While worksheetNode And PbXls_GetNodeNameWithoutNS(worksheetNode) <> "worksheet"
    worksheetNode = NextXMLNode(worksheetNode)
  Wend
  
  If worksheetNode = 0
    ProcedureReturn #False
  EndIf
  
  ; 查找sheetData节点
  Define sheetDataNode.i = ChildXMLNode(worksheetNode)
  While sheetDataNode
    If PbXls_GetNodeNameWithoutNS(sheetDataNode) = "sheetData"
      Break
    EndIf
    sheetDataNode = NextXMLNode(sheetDataNode)
  Wend
  
  If sheetDataNode = 0
    ProcedureReturn #True ; 空工作表
  EndIf
  
  ; 遍历所有行
  Define rowNode.i = ChildXMLNode(sheetDataNode)
  While rowNode
    If PbXls_GetNodeNameWithoutNS(rowNode) = "row"
      Define rowAttr.s = GetXMLAttribute(rowNode, "r")
      Define rowNum.i = 0
      If rowAttr <> ""
        rowNum = Val(rowAttr)
      EndIf
      
      ; 遍历行中的单元格
      Define cellNode.i = ChildXMLNode(rowNode)
      While cellNode
        If PbXls_GetNodeNameWithoutNS(cellNode) = "c"
          Define cellRef.s = GetXMLAttribute(cellNode, "r")
          Define cellType.s = GetXMLAttribute(cellNode, "t")
          Define cellStyle.s = GetXMLAttribute(cellNode, "s")
          
          Define row.i = 0, col.i = 0
          If cellRef <> ""
            PbXls_CoordinateToTuple(cellRef, @row, @col)
          EndIf
          
          If row > 0 And col > 0
            Define key.s = Str(wsId) + "_" + Str(row) + "," + Str(col)
            Define *cell.PbXls_Cell = @PbXls_AllCells(key)
            *cell\row = row
            *cell\column = col
            
            If cellStyle <> ""
              *cell\styleId = Val(cellStyle)
            EndIf
            
            ; 读取单元格值
            Define valueNode.i = ChildXMLNode(cellNode)
            Define cellValue.s = ""
            Define formulaValue.s = ""
            While valueNode
              Define vnLocal.s = PbXls_GetNodeNameWithoutNS(valueNode)
              Select vnLocal
                Case "v"
                  cellValue = GetXMLNodeText(valueNode)
                Case "f"
                  formulaValue = GetXMLNodeText(valueNode)
                Case "is"
                  ; 内联字符串
                  Define tNode2.i = ChildXMLNode(valueNode)
                  While tNode2
                    If PbXls_GetNodeNameWithoutNS(tNode2) = "t"
                      cellValue = GetXMLNodeText(tNode2)
                      Break
                    EndIf
                    tNode2 = NextXMLNode(tNode2)
                  Wend
              EndSelect
              valueNode = NextXMLNode(valueNode)
            Wend
            
            ; 根据类型设置值
            Select cellType
              Case "s"
                ; 共享字符串
                Define strIdx.i = Val(cellValue)
                If SelectElement(sharedStrings(), strIdx)
                  *cell\value = sharedStrings()
                  *cell\dataType = #PbXls_DataTypeString
                EndIf
              Case "str"
                ; 公式结果字符串
                If formulaValue <> ""
                  *cell\formula = formulaValue
                  *cell\value = "=" + formulaValue
                  *cell\dataType = #PbXls_DataTypeFormula
                Else
                  *cell\value = cellValue
                  *cell\dataType = #PbXls_DataTypeString
                EndIf
              Case "b"
                ; 布尔值
                *cell\value = cellValue
                *cell\dataType = #PbXls_DataTypeBoolean
              Case "inlineStr"
                ; 内联字符串
                *cell\value = PbXls_UnescapeXML(cellValue)
                *cell\dataType = #PbXls_DataTypeInline
              Case "e"
                ; 错误值
                *cell\value = cellValue
                *cell\dataType = #PbXls_DataTypeError
              Case ""
                ; 默认: 数值或通用
                If formulaValue <> ""
                  *cell\formula = formulaValue
                  *cell\value = "=" + formulaValue
                  *cell\dataType = #PbXls_DataTypeFormula
                Else
                  *cell\value = cellValue
                  *cell\dataType = #PbXls_DataTypeNumeric
                EndIf
              Default
                *cell\value = cellValue
                *cell\dataType = #PbXls_DataTypeString
            EndSelect
          EndIf
        EndIf
        cellNode = NextXMLNode(cellNode)
      Wend
    EndIf
    rowNode = NextXMLNode(rowNode)
  Wend
  
  ProcedureReturn #True
EndProcedure

; PbXls_ReadMergedCellsFromString - 读取合并单元格信息 (从XML字符串)
Procedure.b PbXls_ReadMergedCellsFromString(xmlStr.s, wsId.i)
  If xmlStr = ""
    Debug "[ReadMergedCells] xmlStr为空"
    ProcedureReturn #False
  EndIf
  
  Debug "[ReadMergedCells] wsId=" + Str(wsId) + " xmlStr len=" + Str(Len(xmlStr))
  
  ; 查找所有 mergeCells 节点内的 mergeCell ref="..." 属性
  Define pos.i = 1
  Define foundCount.i = 0
  While pos <= Len(xmlStr)
    Define mcTag.i = FindString(xmlStr, "<mergeCell ", pos)
    If mcTag = 0
      Break
    EndIf
    ; 在该标签内查找ref属性
    Define tagEnd.i = FindString(xmlStr, ">", mcTag)
    If tagEnd > 0
      Define tagContent.s = Mid(xmlStr, mcTag, tagEnd - mcTag + 1)
      Define refStart.i = FindString(tagContent, ~"ref=\"", 1)
      If refStart > 0
        ; ref=" 长度为5，值从+5开始
        Define refEnd.i = FindString(tagContent, ~"\"", refStart + 5)
        If refEnd > 0
          Define ref.s = Mid(tagContent, refStart + 5, refEnd - refStart - 5)
          ref = UCase(ref)
          Define mcKey.s = Str(wsId) + "_" + ref
          PbXls_MergedCells(mcKey) = ref
          foundCount + 1
          Debug "[ReadMergedCells] 找到合并单元格: ref=" + ref + " key=" + mcKey
        EndIf
      EndIf
      pos = tagEnd + 1
    Else
      pos = mcTag + 1
    EndIf
  Wend
  
  Debug "[ReadMergedCells] 共找到 " + Str(foundCount) + " 个合并单元格"
  ProcedureReturn #True
EndProcedure

; PbXls_ReadMergedCells - 读取合并单元格信息 (从xmlId，已废弃，使用FromString版本)
Procedure.b PbXls_ReadMergedCells(xmlId.i, wsId.i)
  ProcedureReturn PbXls_ReadMergedCellsFromString("", wsId)
EndProcedure

; PbXls_ReadColumnWidthsFromString - 读取列宽设置 (从XML字符串)
; col节点格式: <col min="1" max="1" width="15.0" .../>
Procedure.b PbXls_ReadColumnWidthsFromString(xmlStr.s, wsId.i)
  If xmlStr = ""
    Debug "[ReadColumnWidths] xmlStr为空"
    ProcedureReturn #False
  EndIf
  
  Debug "[ReadColumnWidths] wsId=" + Str(wsId) + " xmlStr len=" + Str(Len(xmlStr))
  
  ; 查找所有 <col 标签，提取其中的min和width属性
  Define pos.i = 1
  Define foundCount.i = 0
  While pos <= Len(xmlStr)
    Define colTag.i = FindString(xmlStr, "<col ", pos)
    If colTag = 0
      Break
    EndIf
    ; 确保不是 <cols> 标签
    Define afterCol.s = Mid(xmlStr, colTag + 4, 2)
    If afterCol <> "s " And afterCol <> "s>"
      ; 找到标签结束
      Define tagEnd.i = FindString(xmlStr, ">", colTag)
      If tagEnd > 0
        Define tagContent.s = Mid(xmlStr, colTag, tagEnd - colTag + 1)
        
        ; 提取min属性
        Define minPos.i = FindString(tagContent, ~"min=\"", 1)
        Define widthPos.i = FindString(tagContent, ~"width=\"", 1)
        
        If minPos > 0 And widthPos > 0
          ; 提取min值
          Define minEnd.i = FindString(tagContent, ~"\"", minPos + 5)
          Define minVal.s = ""
          If minEnd > 0
            minVal = Mid(tagContent, minPos + 5, minEnd - minPos - 5)
          EndIf
          
          ; 提取width值
          Define widthEnd.i = FindString(tagContent, ~"\"", widthPos + 7)
          Define widthVal.s = ""
          If widthEnd > 0
            widthVal = Mid(tagContent, widthPos + 7, widthEnd - widthPos - 7)
          EndIf
          
          If minVal <> "" And widthVal <> ""
            Define colIdx.i = Val(minVal)
            Define cwKey.s = Str(wsId) + "_" + Str(colIdx)
            PbXls_ColumnWidths(cwKey) = ValF(widthVal)
            foundCount + 1
            Debug "[ReadColumnWidths] 找到列宽: col=" + minVal + " width=" + widthVal + " key=" + cwKey
          EndIf
        EndIf
        
        pos = tagEnd + 1
      Else
        pos = colTag + 1
      EndIf
    Else
      pos = colTag + 5
    EndIf
  Wend
  
  Debug "[ReadColumnWidths] 共找到 " + Str(foundCount) + " 个列宽"
  ProcedureReturn #True
EndProcedure

; PbXls_ReadColumnWidths - 读取列宽设置 (从xmlId，已废弃，使用FromString版本)
Procedure.b PbXls_ReadColumnWidths(xmlId.i, wsId.i)
  ProcedureReturn PbXls_ReadColumnWidthsFromString("", wsId)
EndProcedure

; PbXls_ReadRowHeightsFromString - 读取行高设置 (从XML字符串)
; row节点格式: <row r="1" ht="25.0" ...>
Procedure.b PbXls_ReadRowHeightsFromString(xmlStr.s, wsId.i)
  If xmlStr = ""
    Debug "[ReadRowHeights] xmlStr为空"
    ProcedureReturn #False
  EndIf
  
  Debug "[ReadRowHeights] wsId=" + Str(wsId) + " xmlStr len=" + Str(Len(xmlStr))
  
  ; 查找所有 <row 标签 (在sheetData内)，提取其中的r和ht属性
  Define pos.i = 1
  Define foundCount.i = 0
  While pos <= Len(xmlStr)
    Define rowTag.i = FindString(xmlStr, "<row ", pos)
    If rowTag = 0
      Break
    EndIf
    
    ; 找到标签结束
    Define tagEnd.i = FindString(xmlStr, ">", rowTag)
    If tagEnd > 0
      Define tagContent.s = Mid(xmlStr, rowTag, tagEnd - rowTag + 1)
      
      ; 提取r属性(行号)
      Define rPos.i = FindString(tagContent, ~"r=\"", 1)
      ; 提取ht属性(行高)
      Define htPos.i = FindString(tagContent, ~"ht=\"", 1)
      
      If rPos > 0 And htPos > 0
        ; 提取r值
        Define rEnd.i = FindString(tagContent, ~"\"", rPos + 3)
        Define rowVal.s = ""
        If rEnd > 0
          rowVal = Mid(tagContent, rPos + 3, rEnd - rPos - 3)
        EndIf
        
        ; 提取ht值
        Define htEnd.i = FindString(tagContent, ~"\"", htPos + 4)
        Define htVal.s = ""
        If htEnd > 0
          htVal = Mid(tagContent, htPos + 4, htEnd - htPos - 4)
        EndIf
        
        If rowVal <> "" And htVal <> ""
          Define rowNum.i = Val(rowVal)
          Define rhKey.s = Str(wsId) + "_" + Str(rowNum)
          PbXls_RowHeights(rhKey) = ValF(htVal)
          foundCount + 1
          Debug "[ReadRowHeights] 找到行高: row=" + rowVal + " ht=" + htVal + " key=" + rhKey
        EndIf
      EndIf
      
      pos = tagEnd + 1
    Else
      pos = rowTag + 1
    EndIf
  Wend
  
  Debug "[ReadRowHeights] 共找到 " + Str(foundCount) + " 个行高"
  ProcedureReturn #True
EndProcedure

; PbXls_ReadRowHeights - 读取行高设置 (从xmlId，已废弃，使用FromString版本)
Procedure.b PbXls_ReadRowHeights(xmlId.i, wsId.i)
  ProcedureReturn PbXls_ReadRowHeightsFromString("", wsId)
EndProcedure

; PbXls_ReadSheetSettings - 读取工作表设置(冻结窗格、自动筛选、页面设置等)
Procedure.b PbXls_ReadSheetSettings(xmlId.i, *ws.PbXls_Worksheet)
  If xmlId = 0
    ProcedureReturn #False
  EndIf
  
  Define rootNode.i = RootXMLNode(xmlId)
  Define worksheetNode.i = ChildXMLNode(rootNode)
  While worksheetNode And PbXls_GetNodeNameWithoutNS(worksheetNode) <> "worksheet"
    worksheetNode = NextXMLNode(worksheetNode)
  Wend
  
  If worksheetNode = 0
    ProcedureReturn #False
  EndIf
  
  ; 读取冻结窗格 (从sheetViews -> sheetView -> pane)
  Define sheetViewsNode.i = ChildXMLNode(worksheetNode)
  While sheetViewsNode
    If PbXls_GetNodeNameWithoutNS(sheetViewsNode) = "sheetViews"
      Define sheetViewNode.i = ChildXMLNode(sheetViewsNode)
      While sheetViewNode
        If PbXls_GetNodeNameWithoutNS(sheetViewNode) = "sheetView"
          Define paneNode.i = ChildXMLNode(sheetViewNode)
          While paneNode
            If PbXls_GetNodeNameWithoutNS(paneNode) = "pane"
              Define topLeftCell.s = GetXMLAttribute(paneNode, "topLeftCell")
              If topLeftCell <> ""
                *ws\freezePanes = topLeftCell
              EndIf
              Break
            EndIf
            paneNode = NextXMLNode(paneNode)
          Wend
          Break
        EndIf
        sheetViewNode = NextXMLNode(sheetViewNode)
      Wend
      Break
    EndIf
    sheetViewsNode = NextXMLNode(sheetViewsNode)
  Wend
  
  ; 读取自动筛选
  Define autoFilterNode.i = ChildXMLNode(worksheetNode)
  While autoFilterNode
    If PbXls_GetNodeNameWithoutNS(autoFilterNode) = "autoFilter"
      Define ref.s = GetXMLAttribute(autoFilterNode, "ref")
      If ref <> ""
        *ws\autoFilter = ref
      EndIf
      Break
    EndIf
    autoFilterNode = NextXMLNode(autoFilterNode)
  Wend
  
  ; 读取页面设置
  Define pageSetupNode.i = ChildXMLNode(worksheetNode)
  While pageSetupNode
    If PbXls_GetNodeNameWithoutNS(pageSetupNode) = "pageSetup"
      Define orientation.s = GetXMLAttribute(pageSetupNode, "orientation")
      If orientation <> ""
        *ws\orientation = orientation
      EndIf
      Define paperSize.s = GetXMLAttribute(pageSetupNode, "paperSize")
      If paperSize <> ""
        *ws\paperSize = Val(paperSize)
      EndIf
      Break
    EndIf
    pageSetupNode = NextXMLNode(pageSetupNode)
  Wend
  
  ProcedureReturn #True
EndProcedure

; ***************************************************************************************
; 分区13: 公共API
; ***************************************************************************************

; PbXls_LoadWorkbook - 从文件加载Excel工作簿
; 参考openpyxl: reader/excel.py ExcelReader.read方法
; 读取流程: 1.解压 2.读共享字符串 3.读工作簿 4.读工作表 5.清理
Procedure.i PbXls_LoadWorkbook(filename.s)
  ; 检查文件是否存在
  If FileSize(filename) = -1
    Debug "[PbXls_LoadWorkbook] 文件不存在: " + filename
    ProcedureReturn -1
  EndIf
  
  ; 1. 解压xlsx到临时目录
  Define tempDir.s = PbXls_ExtractToTempDir(filename)
  If tempDir = ""
    Debug "[PbXls_LoadWorkbook] 解压失败"
    ProcedureReturn -1
  EndIf
  Debug "[PbXls_LoadWorkbook] 解压到: " + tempDir
  
  ; 2. 读取共享字符串表 (xl/sharedStrings.xml)
  Define NewList sharedStrings.s()
  Define sharedStringsFile.s = tempDir + "xl\sharedStrings.xml"
  If FileSize(sharedStringsFile) > 0
    Define ssXmlId.i = PbXls_XMLParseFile(sharedStringsFile)
    If ssXmlId
      PbXls_ReadSharedStrings(ssXmlId, sharedStrings())
      Debug "[PbXls_LoadWorkbook] 共享字符串数量: " + Str(ListSize(sharedStrings()))
      FreeXML(ssXmlId)
    EndIf
  Else
    Debug "[PbXls_LoadWorkbook] sharedStrings.xml 不存在"
  EndIf
  
  ; 3. 读取工作簿信息 (xl/workbook.xml)
  Define workbookFile.s = tempDir + "xl\workbook.xml"
  If FileSize(workbookFile) = -1
    Debug "[PbXls_LoadWorkbook] workbook.xml 不存在: " + workbookFile
    PbXls_CleanupTempDir(tempDir)
    ProcedureReturn -1
  EndIf
  
  Define wbXmlId.i = PbXls_XMLParseFile(workbookFile)
  If wbXmlId = 0
    PbXls_CleanupTempDir(tempDir)
    ProcedureReturn -1
  EndIf
  
  ; 使用临时数组存储工作表名称 (最多100个)
  Dim sheetNameArray.s(100)
  Define sheetNameCount.i = 0
  Define activeSheetVal.i = 0
  
  ; 解析workbook.xml获取工作表名称
  Define wbXmlStr.s = PbXls_XMLSaveToString(wbXmlId)
  If wbXmlStr <> ""
    ; 查找activeTab
    Define atPos.i = FindString(wbXmlStr, "activeTab=", 1)
    If atPos > 0
      Define atEnd.i = FindString(wbXmlStr, " ", atPos)
      If atEnd = 0 : atEnd = FindString(wbXmlStr, ">", atPos) : EndIf
      If atEnd > 0
        Define atVal.s = Mid(wbXmlStr, atPos + 11, atEnd - atPos - 12)
        atVal = ReplaceString(atVal, ~"\"", "")
        activeSheetVal = Val(atVal)
      EndIf
    EndIf
    
    ; 查找所有sheet name属性
    Define snPos.i = 1
    While snPos <= Len(wbXmlStr)
      Define nameAttr.s = "name=" + ~"\""
      Define naPos.i = FindString(wbXmlStr, nameAttr, snPos)
      If naPos = 0 : Break : EndIf
      Define naEnd.i = FindString(wbXmlStr, ~"\"", naPos + 6)
      If naEnd = 0 : Break : EndIf
      Define sn.s = Mid(wbXmlStr, naPos + 6, naEnd - naPos - 6)
      If sn <> "" And sheetNameCount < 100
        sheetNameArray(sheetNameCount) = sn
        sheetNameCount + 1
      EndIf
      snPos = naEnd + 1
    Wend
  EndIf
  
  FreeXML(wbXmlId)
  
  ; 4. 创建新的工作簿
  Define wbId.i = PbXls_CreateWorkbook()
  If wbId = -1
    PbXls_CleanupTempDir(tempDir)
    ProcedureReturn -1
  EndIf
  
  ; 5. 获取工作簿指针并设置路径
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(wbId)
  If *wb = 0
    PbXls_CleanupTempDir(tempDir)
    ProcedureReturn -1
  EndIf
  *wb\path = filename
  
  ; 6. 删除默认创建的Sheet并重置计数器
  Define wbKeyReset.s = Str(wbId)
  PbXls_WorkbookSheetCount(wbKeyReset) = 0
  Define defaultSheetCount.i = 1
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets()) And defaultSheetCount > 0
    If PbXls_AllWorksheets()\parent = wbId
      DeleteElement(PbXls_AllWorksheets())
      defaultSheetCount - 1
      ResetList(PbXls_AllWorksheets())
    EndIf
  Wend
  
  ; 7. 创建从workbook.xml读取到的工作表
  Define wsIdx.i = 0
  While wsIdx < sheetNameCount
    PbXls_CreateSheet(wbId, sheetNameArray(wsIdx))
    wsIdx + 1
  Wend
  
  ; 8. 读取每个工作表的XML
  wsIdx = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = wbId
      Define sheetFile.s = tempDir + "xl\worksheets\sheet" + Str(wsIdx + 1) + ".xml"
      If FileSize(sheetFile) > 0
        Define wsXmlId.i = PbXls_XMLParseFile(sheetFile)
        If wsXmlId
          Define *ws.PbXls_Worksheet = @PbXls_AllWorksheets()
          
          ; 读取工作表XML内容（用于文本解析）
          Define sheetXmlContent.s = PbXls_ReadFileToString(sheetFile)
          
          ; 读取单元格数据
          PbXls_ReadWorksheetData(wsXmlId, *ws\id, sharedStrings())
          
          ; 读取合并单元格（使用文件内容而不是XML ID）
          PbXls_ReadMergedCellsFromString(sheetXmlContent, *ws\id)
          
          ; 读取列宽
          PbXls_ReadColumnWidthsFromString(sheetXmlContent, *ws\id)
          
          ; 读取行高
          PbXls_ReadRowHeightsFromString(sheetXmlContent, *ws\id)
          
          ; 读取工作表设置(冻结窗格、自动筛选、页面设置)
          PbXls_ReadSheetSettings(wsXmlId, *ws)
          
          FreeXML(wsXmlId)
        EndIf
      EndIf
      wsIdx + 1
    EndIf
  Wend
  
  ; 9. 设置活动工作表
  *wb\activeSheetIndex = activeSheetVal
  
  ; 10. 清理临时文件
  PbXls_CleanupTempDir(tempDir)
  
  ProcedureReturn wbId
EndProcedure

Procedure.b PbXls_SaveWorkbook(workbook.i, filename.s)
  ProcedureReturn PbXls_SaveWorkbookToFile(workbook, filename)
EndProcedure

Procedure.b PbXls_CloseWorkbook(workbook.i)
  Define *wb.PbXls_Workbook = PbXls_GetWorkbookPtr(workbook)
  If *wb = 0
    ProcedureReturn #False
  EndIf
  
  ; 1. 娓呯悊涓庤宸ヤ綔绨垮叧鑱旂殑鎵€鏈夊崟鍏冩牸鏁版嵁
  Define wsKey.s = Str(*wb\id) + "_"
  Define wsKeyLen.i = Len(wsKey)
  
  ForEach PbXls_AllCells()
    If Left(MapKey(PbXls_AllCells()), wsKeyLen) = wsKey
      DeleteMapElement(PbXls_AllCells(), MapKey(PbXls_AllCells()))
    EndIf
  Next
  
  ; 2. 娓呯悊涓庤宸ヤ綔绨垮叧鑱旂殑鎵€鏈夊垪瀹芥暟鎹?
  ForEach PbXls_ColumnWidths()
    If Left(MapKey(PbXls_ColumnWidths()), wsKeyLen) = wsKey
      DeleteMapElement(PbXls_ColumnWidths(), MapKey(PbXls_ColumnWidths()))
    EndIf
  Next
  
  ; 3. 娓呯悊涓庤宸ヤ綔绨垮叧鑱旂殑鎵€鏈夎楂樻暟鎹?
  ForEach PbXls_RowHeights()
    If Left(MapKey(PbXls_RowHeights()), wsKeyLen) = wsKey
      DeleteMapElement(PbXls_RowHeights(), MapKey(PbXls_RowHeights()))
    EndIf
  Next
  
  ; 4. 娓呯悊涓庤宸ヤ綔绨垮叧鑱旂殑鎵€鏈夊悎骞跺崟鍏冩牸鏁版嵁 (including margins, headers, print settings)
  ForEach PbXls_MergedCells()
    Define mcKey2.s = MapKey(PbXls_MergedCells())
    If Left(mcKey2, wsKeyLen) = wsKey
      DeleteMapElement(PbXls_MergedCells(), mcKey2)
    EndIf
  Next
  
  ; 5. 娓呯悊涓庤宸ヤ綔绨垮叧鑱旂殑鎵€鏈夎秴閾炬帴鍜屾敞閲婃暟鎹?宸插寘鍚湪鍗曞厓鏍间腑,无需单独清理)
  
  ; 7. 娓呯悊涓庤宸ヤ綔绨垮叧鑱旂殑鎵€鏈夊伐浣滆〃
  Define wsIdx.i = 0
  ResetList(PbXls_AllWorksheets())
  While NextElement(PbXls_AllWorksheets())
    If PbXls_AllWorksheets()\parent = *wb\id
      wsIdx + 1
    EndIf
  Wend
  
  If wsIdx > 0
    ResetList(PbXls_AllWorksheets())
    While NextElement(PbXls_AllWorksheets()) And wsIdx > 0
      If PbXls_AllWorksheets()\parent = *wb\id
        DeleteElement(PbXls_AllWorksheets())
        wsIdx - 1
        ResetList(PbXls_AllWorksheets())
      EndIf
    Wend
  EndIf
  
  ; 8. 娓呯悊宸ヤ綔绨胯鏁?
  DeleteMapElement(PbXls_WorkbookSheetCount(), Str(*wb\id))
  
  ; 9. 娓呯悊鍏变韩瀛楃涓茬储寮曟槧灏?
  DeleteMapElement(PbXls_WorkbookSharedStrings(), Str(*wb\id))
  
  ; 10. 浠庡伐浣滅翱鍒楄〃涓Щ闄?
  ForEach PbXls_Workbooks()
    If PbXls_Workbooks()\id = *wb\id
      DeleteElement(PbXls_Workbooks())
      Break
    EndIf
  Next
  
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

Procedure.i PbXls_GetCellStyleAPI(worksheet.i, row.i, col.i)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn -1
  EndIf
  Define *cell.PbXls_Cell = PbXls_GetCell(*ws, row, col)
  If *cell = 0
    ProcedureReturn -1
  EndIf
  ProcedureReturn PbXls_GetCellStyle(*cell)
EndProcedure

Procedure.b PbXls_SetCellHyperlinkAPI(worksheet.i, row.i, col.i, url.s, tooltip.s = "")
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  Define *cell.PbXls_Cell = PbXls_GetCell(*ws, row, col)
  If *cell = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetCellHyperlink(*cell, url, tooltip)
EndProcedure

Procedure.b PbXls_InsertRowsAPI(worksheet.i, row.i, count.i = 1)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_InsertRows(*ws, row, count)
EndProcedure

Procedure.b PbXls_DeleteRowsAPI(worksheet.i, row.i, count.i = 1)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_DeleteRows(*ws, row, count)
EndProcedure

; 列插入/删除 API
Procedure.b PbXls_InsertColumnsAPI(worksheet.i, col.i, count.i = 1)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_InsertColumns(*ws, col, count)
EndProcedure

Procedure.b PbXls_DeleteColumnsAPI(worksheet.i, col.i, count.i = 1)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_DeleteColumns(*ws, col, count)
EndProcedure

; 页边距/打印设置 API

; 页边距/打印设置 API
Procedure.b PbXls_SetPageMarginsAPI(worksheet.i, left.f = 0.7, right.f = 0.7, top.f = 0.75, bottom.f = 0.75, header.f = 0.3, footer.f = 0.3)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetPageMargins(*ws, left, right, top, bottom, header, footer)
EndProcedure

Procedure.b PbXls_SetHeaderFooterAPI(worksheet.i, header.s = "", footer.s = "")
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetHeaderFooter(*ws, header, footer)
EndProcedure

Procedure.b PbXls_SetPrintOptionsAPI(worksheet.i, gridLines.b = #True, headings.b = #False, horizontalCentered.b = #False, verticalCentered.b = #False)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetPrintOptions(*ws, gridLines, headings, horizontalCentered, verticalCentered)
EndProcedure

Procedure.b PbXls_SetOrientationAPI(worksheet.i, orientation.s)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetOrientation(*ws, orientation)
EndProcedure

Procedure.b PbXls_SetPaperSizeAPI(worksheet.i, paperSize.i)
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetPaperSize(*ws, paperSize)
EndProcedure

Procedure.b PbXls_SetCellCommentAPI(worksheet.i, row.i, col.i, content.s, author.s = "")
  Define *ws.PbXls_Worksheet = worksheet
  If *ws = 0
    ProcedureReturn #False
  EndIf
  ProcedureReturn PbXls_SetCellComment(*ws, row, col, content, author)
EndProcedure

; 样式 API

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
  ClearList(PbXls_DataValidations())
  ClearList(PbXls_ConditionalFormats())
  ClearList(PbXls_Charts())
  ClearMap(PbXls_AllCells())
  ClearMap(PbXls_ColumnWidths())
  ClearMap(PbXls_RowHeights())
  ClearMap(PbXls_MergedCells())
  ClearMap(PbXls_MergedCellCount())
  ClearMap(PbXls_WorkbookSheetCount())
  ClearMap(PbXls_WorkbookSharedStrings())
  ClearMap(PbXls_ChartSeriesName())
  ClearMap(PbXls_ChartSeriesValues())
  ClearMap(PbXls_ChartSeriesCategories())
  PbXls_DxfCounter = 0
  PbXls_ChartCounter = 0
  ProcedureReturn #True
EndProcedure

PbXls_Init()
