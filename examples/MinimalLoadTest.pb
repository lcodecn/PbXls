﻿; ***************************************************************************************
; PbXls 示例代码 - 最小加载测试
; 说明: 绝对最小的工作簿创建和加载测试
; 版本: 2.4
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 日志文件路径
Global logPath.s = GetPathPart(ProgramFilename()) + "pbxls_minimal_test.txt"
Global logFile.i = CreateFile(#PB_Any, logPath)

WriteStringN(logFile, "测试开始")

; 创建工作簿
Global wb1.i = PbXls_CreateWorkbook()
WriteStringN(logFile, "工作簿创建成功: " + Str(wb1))

If wb1 >= 0
  Global ws1.i = PbXls_GetSheetByIndex(wb1, 0)
  WriteStringN(logFile, "工作表获取成功: " + Str(ws1))
  
  ; 写入测试数据
  PbXls_SetCell(ws1, 1, 1, "Hello")
  WriteStringN(logFile, "单元格数据已设置")
  
  ; 保存文件
  If PbXls_SaveWorkbook(wb1, GetTemporaryDirectory() + "pbxls_minimal.xlsx")
    WriteStringN(logFile, "保存成功")
  Else
    WriteStringN(logFile, "保存失败!")
  EndIf
  
  PbXls_CloseWorkbook(wb1)
  WriteStringN(logFile, "已关闭 wb1")
  
  ; 现在尝试加载回来
  WriteStringN(logFile, "即将加载...")
  Define wb2.i = PbXls_LoadWorkbook(GetTemporaryDirectory() + "pbxls_minimal.xlsx")
  WriteStringN(logFile, "加载返回: " + Str(wb2))
Else
  WriteStringN(logFile, "工作簿创建失败!")
EndIf

WriteStringN(logFile, "测试结束")
CloseFile(logFile)
MessageRequester("完成", "请检查: " + logPath, #MB_ICONINFORMATION)

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 38
; FirstLine = 15
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory
