﻿; ***************************************************************************************
; PbXls 示例代码 - 快速加载测试
; 说明: 测试创建->保存->读取的快速流程
; 版本: 2.4
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 简单测试: 创建->保存->读取
Global testFile.s = GetPathPart(ProgramFilename()) + "PbXls_quick_test.xlsx"
Global logPath.s = GetPathPart(ProgramFilename()) + "PbXls_quick_debug.txt"

; 创建日志文件
Global logFile.i = CreateFile(#PB_Any, logPath)
If logFile = 0
  MessageRequester("错误", "无法创建日志文件: " + logPath, #MB_ICONERROR)
  End
EndIf

; 日志写入过程
Procedure D(msg.s)
  WriteStringN(logFile, msg)
EndProcedure

D("=== PbXls 快速加载测试 ===")

; 第一部分: 创建并保存
; 创建工作簿
Global wb1.i = PbXls_CreateWorkbook()
Global ws1.i = PbXls_GetSheetByIndex(wb1, 0)
D("创建 wb1=" + Str(wb1) + ", ws1=" + Str(ws1))

; 写入测试数据
PbXls_SetCell(ws1, 1, 1, "Hello")
PbXls_SetCell(ws1, 2, 1, "World")
PbXls_SetCell(ws1, 1, 2, "123")
PbXls_MergeCells(ws1, "A1:B1")
PbXls_SetColumnWidth(ws1, 1, 20.0)
D("单元格数据已设置")

; 保存文件
If PbXls_SaveWorkbook(wb1, testFile)
  D("保存成功: " + testFile)
Else
  D("保存失败!")
EndIf
PbXls_CloseWorkbook(wb1)
D("已关闭 wb1")

; 第二部分: 读取文件
D("正在加载文件...")
Define wb2.i = PbXls_LoadWorkbook(testFile)
D("加载结果: " + Str(wb2))

If wb2 >= 0
  D("工作表数量: " + Str(PbXls_GetSheetCount(wb2)))
  
  Define ws2.i = PbXls_GetSheetByIndex(wb2, 0)
  If ws2 <> 0
    D("ws2 获取成功: " + Str(ws2))
    
    Define a1.s = PbXls_GetCellString(ws2, 1, 1)
    D("A1='" + a1 + "'")
    
    Define b1.s = PbXls_GetCellString(ws2, 1, 2)
    D("B1='" + b1 + "'")
    
    PbXls_CloseWorkbook(wb2)
  Else
    D("ws2 = 0, 获取失败")
  EndIf
Else
  D("加载失败!")
EndIf

; 清理资源
PbXls_Free()
CloseFile(logFile)
MessageRequester("完成", "请检查: " + logPath, #MB_ICONINFORMATION)

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 58
; FirstLine = 40
; Folding = -
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory
