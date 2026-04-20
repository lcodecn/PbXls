; ***************************************************************************************
; PbXls 示例代码 - 读写Excel工作簿
; 说明: 演示如何使用PbXls库创建、写入和读取Excel文件
; 注意: 读取功能当前为占位实现，本示例主要演示写入功能
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 读写Excel工作簿示例 ==="
Debug ""

; ============================================================
; 第一部分：创建并写入Excel文件
; ============================================================
Debug "--- 第一部分：创建并写入Excel文件 ---"
Debug ""

; 1. 创建新工作簿
Debug "1. 创建工作簿..."
wbId.i = PbXls_CreateWorkbook()
If wbId = -1
  Debug "错误: 无法创建工作簿"
  End
EndIf
Debug "  工作簿创建成功"

; 2. 获取活动工作表
wsId.i = PbXls_ActiveSheet(wbId)

; 3. 写入表头
Debug "2. 写入表头..."
PbXls_SetCell(wsId, 1, 1, "姓名")
PbXls_SetCell(wsId, 1, 2, "年龄")
PbXls_SetCell(wsId, 1, 3, "城市")
Debug "  表头: 姓名 | 年龄 | 城市"

; 4. 写入数据行
Debug "3. 写入数据..."
PbXls_SetCell(wsId, 2, 1, "张三")
PbXls_SetCell(wsId, 2, 2, "25")
PbXls_SetCell(wsId, 2, 3, "北京")
Debug "  行2: 张三, 25, 北京"

PbXls_SetCell(wsId, 3, 1, "李四")
PbXls_SetCell(wsId, 3, 2, "30")
PbXls_SetCell(wsId, 3, 3, "上海")
Debug "  行3: 李四, 30, 上海"

PbXls_SetCell(wsId, 4, 1, "王五")
PbXls_SetCell(wsId, 4, 2, "28")
PbXls_SetCell(wsId, 4, 3, "广州")
Debug "  行4: 王五, 28, 广州"

; 5. 添加公式行
Debug "4. 添加公式..."
PbXls_SetCell(wsId, 5, 1, "平均年龄")
PbXls_SetCellFormulaWS(wsId, 5, 2, "=AVERAGE(B2:B4)")
Debug "  B5 = =AVERAGE(B2:B4)"

; 6. 设置列宽
PbXls_SetColumnWidth(wsId, 1, 12)
PbXls_SetColumnWidth(wsId, 2, 8)
PbXls_SetColumnWidth(wsId, 3, 10)

; 7. 保存文件
Debug "5. 保存文件..."
outputFile.s = "demo_readwrite.xlsx"
If PbXls_SaveWorkbook(wbId, outputFile)
  Debug "  保存成功: " + outputFile
Else
  Debug "  保存失败!"
EndIf

; 清理第一个工作簿
PbXls_CloseWorkbook(wbId)

; ============================================================
; 第二部分：坐标工具函数演示
; ============================================================
Debug ""
Debug "--- 第二部分：坐标工具函数演示 ---"
Debug ""

; 列号转列字母
Debug "列号转列字母:"
Debug "  1 -> " + PbXls_GetColumnLetter(1)
Debug "  26 -> " + PbXls_GetColumnLetter(26)
Debug "  27 -> " + PbXls_GetColumnLetter(27)
Debug "  52 -> " + PbXls_GetColumnLetter(52)
Debug "  702 -> " + PbXls_GetColumnLetter(702)

; 列字母转列号
Debug ""
Debug "列字母转列号:"
Debug "  A -> " + Str(PbXls_ColumnIndexFromString("A"))
Debug "  Z -> " + Str(PbXls_ColumnIndexFromString("Z"))
Debug "  AA -> " + Str(PbXls_ColumnIndexFromString("AA"))
Debug "  AZ -> " + Str(PbXls_ColumnIndexFromString("AZ"))
Debug "  ZZ -> " + Str(PbXls_ColumnIndexFromString("ZZ"))

; 坐标解析
Debug ""
Debug "坐标解析:"
Define row.i, col.i
If PbXls_CoordinateToTuple("B5", @row, @col)
  Debug "  B5 -> 行=" + Str(row) + ", 列=" + Str(col)
EndIf

If PbXls_CoordinateToTuple("AA10", @row, @col)
  Debug "  AA10 -> 行=" + Str(row) + ", 列=" + Str(col)
EndIf

; 范围解析
Debug ""
Debug "范围解析:"
Define minCol.i, minRow.i, maxCol.i, maxRow.i
If PbXls_RangeBoundaries("A1:C3", @minCol, @minRow, @maxCol, @maxRow)
  Debug "  A1:C3 -> (" + Str(minRow) + "," + Str(minCol) + ") 到 (" + Str(maxRow) + "," + Str(maxCol) + ")"
EndIf

; XML转义
Debug ""
Debug "XML转义:"
Debug "  'A < B & C > D' -> '" + PbXls_EscapeXML("A < B & C > D") + "'"
Debug "  'Hello " + Chr(34) + "World" + Chr(34) + "'" + " -> '" + PbXls_EscapeXML("Hello " + Chr(34) + "World" + Chr(34)) + "'"

; ============================================================
; 第三部分：数据类型检测
; ============================================================
Debug ""
Debug "--- 第三部分：数据类型检测 ---"
Debug ""

Debug "数值检测:"
Debug "  '123.45' -> " + Str(PbXls_IsNumeric("123.45"))
Debug "  '-100' -> " + Str(PbXls_IsNumeric("-100"))
Debug "  'abc' -> " + Str(PbXls_IsNumeric("abc"))

Debug ""
Debug "布尔值检测:"
Debug "  'TRUE' -> " + Str(PbXls_IsBoolean("TRUE"))
Debug "  'false' -> " + Str(PbXls_IsBoolean("false"))
Debug "  'yes' -> " + Str(PbXls_IsBoolean("yes"))

Debug ""
Debug "公式检测:"
Debug "  '=SUM(A1:A10)' -> " + Str(PbXls_IsFormula("=SUM(A1:A10)"))
Debug "  '=A1+B1' -> " + Str(PbXls_IsFormula("=A1+B1"))
Debug "  'SUM(A1:A10)' -> " + Str(PbXls_IsFormula("SUM(A1:A10)"))

Debug ""
Debug "日期检测:"
Debug "  '2026-04-20' -> " + Str(PbXls_IsDate("2026-04-20"))
Debug "  '2026/4/20' -> " + Str(PbXls_IsDate("2026/4/20"))
Debug "  'April 20, 2026' -> " + Str(PbXls_IsDate("April 20, 2026"))

; 第四部分：清理
Debug ""
Debug "=== 示例完成 ==="
Debug "生成的文件: " + outputFile
Debug ""

PbXls_Free()
