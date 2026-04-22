; ***************************************************************************************
; PbXls 示例代码 - 测试加载详情
; 说明: 测试LoadWorkbook的详细功能，验证加载后数据的一致性
; 版本: 2.4
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 日志文件路径
Global logPath.s = GetPathPart(ProgramFilename()) + "pbxls_test2.txt"
Global logFile.i = CreateFile(#PB_Any, logPath)

; 日志写入过程
Procedure D2(msg.s)
  WriteStringN(logFile, msg)
EndProcedure

D2("=== 测试2: LoadWorkbook 详情测试 ===")

; 第一部分: 首先创建一个测试文件
Global testFile.s = GetTemporaryDirectory() + "pbxls_test2.xlsx"

; 创建工作簿
Global wb1.i = PbXls_CreateWorkbook()
D2("创建 wb1=" + Str(wb1))
Global ws1.i = PbXls_GetSheetByIndex(wb1, 0)
D2("ws1=" + Str(ws1))

; 写入测试数据
PbXls_SetCell(ws1, 1, 1, "Hello")
PbXls_SetCell(ws1, 1, 2, "World")
PbXls_MergeCells(ws1, "A1:B1")
PbXls_SetColumnWidth(ws1, 1, 20.0)
PbXls_SetRowHeight(ws1, 1, 25.0)

; 保存文件
PbXls_SaveWorkbook(wb1, testFile)
D2("已保存到: " + testFile + ", 文件大小=" + Str(FileSize(testFile)))

; 打印保存前的所有映射数据
D2("关闭前的映射数据:")
D2("  AllCells 数量: " + Str(MapSize(PbXls_AllCells())))
D2("  MergedCells 数量: " + Str(MapSize(PbXls_MergedCells())))
D2("  ColumnWidths 数量: " + Str(MapSize(PbXls_ColumnWidths())))
D2("  RowHeights 数量: " + Str(MapSize(PbXls_RowHeights())))
ForEach PbXls_MergedCells()
  D2("    合并单元格: " + MapKey(PbXls_MergedCells()) + " = " + PbXls_MergedCells())
Next
ForEach PbXls_ColumnWidths()
  D2("    列宽: " + MapKey(PbXls_ColumnWidths()) + " = " + StrF(PbXls_ColumnWidths(), 1))
Next
ForEach PbXls_RowHeights()
  D2("    行高: " + MapKey(PbXls_RowHeights()) + " = " + StrF(PbXls_RowHeights(), 1))
Next

; 关闭第一个工作簿
PbXls_CloseWorkbook(wb1)
D2("已关闭 wb1")

; 第二部分: 加载文件
D2("开始加载...")
Define wb2.i = PbXls_LoadWorkbook(testFile)
D2("加载返回: " + Str(wb2))

; 检查加载后的所有映射和列表
D2("加载后的映射数据:")
D2("  AllCells 数量: " + Str(MapSize(PbXls_AllCells())))
D2("  MergedCells 数量: " + Str(MapSize(PbXls_MergedCells())))
D2("  ColumnWidths 数量: " + Str(MapSize(PbXls_ColumnWidths())))
D2("  RowHeights 数量: " + Str(MapSize(PbXls_RowHeights())))
D2("  Workbooks 数量: " + Str(ListSize(PbXls_Workbooks())))
D2("  Worksheets 数量: " + Str(ListSize(PbXls_AllWorksheets())))

ForEach PbXls_AllCells()
  D2("    单元格: " + MapKey(PbXls_AllCells()) + " 值=" + PbXls_AllCells()\value)
Next
ForEach PbXls_MergedCells()
  D2("    合并单元格: " + MapKey(PbXls_MergedCells()) + " = " + PbXls_MergedCells())
Next
ForEach PbXls_ColumnWidths()
  D2("    列宽: " + MapKey(PbXls_ColumnWidths()) + " = " + StrF(PbXls_ColumnWidths(), 1))
Next
ForEach PbXls_RowHeights()
  D2("    行高: " + MapKey(PbXls_RowHeights()) + " = " + StrF(PbXls_RowHeights(), 1))
Next

ForEach PbXls_Workbooks()
  D2("    工作簿: id=" + Str(PbXls_Workbooks()\id) + " 标题=" + PbXls_Workbooks()\title)
Next
ForEach PbXls_AllWorksheets()
  D2("    工作表: id=" + Str(PbXls_AllWorksheets()\id) + " 标题=" + PbXls_AllWorksheets()\title + " 父级=" + Str(PbXls_AllWorksheets()\parent))
Next

; 验证加载的数据
If wb2 >= 0
  D2("工作表数量: " + Str(PbXls_GetSheetCount(wb2)))
  Define ws2.i = PbXls_GetSheetByIndex(wb2, 0)
  D2("ws2=" + Str(ws2))
  
  If ws2 <> 0
    D2("A1='" + PbXls_GetCellString(ws2, 1, 1) + "'")
    D2("B1='" + PbXls_GetCellString(ws2, 1, 2) + "'")
  EndIf
  
  PbXls_CloseWorkbook(wb2)
EndIf

; 清理资源
PbXls_Free()
CloseFile(logFile)
MessageRequester("完成", "请检查: " + logPath, #MB_ICONINFORMATION)

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 17
; FirstLine = 72
; Folding = -
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory
