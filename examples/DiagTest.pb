; 简单诊断测试
Global logPath.s = GetPathPart(ProgramFilename()) + "pbxls_diag.txt"
; 删除旧文件
DeleteFile(logPath)
Global logFile.i = CreateFile(#PB_Any, logPath)

If logFile = 0
  MessageRequester("错误", "无法创建日志文件: " + logPath, #MB_ICONERROR)
  End
EndIf

Procedure D(msg.s)
  WriteStringN(logFile, msg)
EndProcedure

XIncludeFile "..\PbXls.pb"

D("=== PbXls 诊断测试 ===")

; 创建文件
Define testFile.s = GetTemporaryDirectory() + "pbxls_diag.xlsx"

Define wb1.i = PbXls_CreateWorkbook()
Define ws1.i = PbXls_GetSheetByIndex(wb1, 0)

D("ws1 id = " + Str(ws1))

PbXls_SetCell(ws1, 1, 1, "Test")
PbXls_MergeCells(ws1, "A1:B1")
PbXls_SetColumnWidth(ws1, 1, 20.0)
PbXls_SetRowHeight(ws1, 1, 25.0)

D("写入后Maps:")
ForEach PbXls_MergedCells()
  D("  MC: '" + MapKey(PbXls_MergedCells()) + "' = '" + PbXls_MergedCells() + "'")
Next
ForEach PbXls_ColumnWidths()
  D("  CW: '" + MapKey(PbXls_ColumnWidths()) + "' = " + StrF(PbXls_ColumnWidths(), 1))
Next
ForEach PbXls_RowHeights()
  D("  RH: '" + MapKey(PbXls_RowHeights()) + "' = " + StrF(PbXls_RowHeights(), 1))
Next

PbXls_SaveWorkbook(wb1, testFile)
D("已保存: " + testFile + " size=" + Str(FileSize(testFile)))
PbXls_CloseWorkbook(wb1)

; 现在加载回来
D("--- 开始加载 ---")
Define wb2.i = PbXls_LoadWorkbook(testFile)
D("Load result: " + Str(wb2))

If wb2 >= 0
  D("工作表数量: " + Str(PbXls_GetSheetCount(wb2)))
  Define ws2.i = PbXls_GetSheetByIndex(wb2, 0)
  D("ws2 = " + Str(ws2))
  
  If ws2 <> 0
    D("ws2 id = " + Str(PbXls_AllWorksheets()\id))
    D("ws2 title = '" + PbXls_AllWorksheets()\title + "'")
  EndIf
  
  D("加载后Maps:")
  D("  AllCells: " + Str(MapSize(PbXls_AllCells())))
  D("  MergedCells: " + Str(MapSize(PbXls_MergedCells())))
  D("  ColumnWidths: " + Str(MapSize(PbXls_ColumnWidths())))
  D("  RowHeights: " + Str(MapSize(PbXls_RowHeights())))
  
  ForEach PbXls_MergedCells()
    D("  MC: '" + MapKey(PbXls_MergedCells()) + "' = '" + PbXls_MergedCells() + "'")
  Next
  ForEach PbXls_ColumnWidths()
    D("  CW: '" + MapKey(PbXls_ColumnWidths()) + "' = " + StrF(PbXls_ColumnWidths(), 1))
  Next
  ForEach PbXls_RowHeights()
    D("  RH: '" + MapKey(PbXls_RowHeights()) + "' = " + StrF(PbXls_RowHeights(), 1))
  Next
  
  ForEach PbXls_AllCells()
    D("  Cell: '" + MapKey(PbXls_AllCells()) + "' val='" + PbXls_AllCells()\value + "'")
  Next
  
  PbXls_CloseWorkbook(wb2)
EndIf

PbXls_Free()
CloseFile(logFile)
MessageRequester("诊断完成", "查看: " + logPath, #MB_ICONINFORMATION)

; IDE Options = PureBasic 6.40 (Windows - x86)
; Folding = -
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory