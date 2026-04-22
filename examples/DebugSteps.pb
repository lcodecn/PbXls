; ***************************************************************************************
; PbXls 示例代码 - 调试步骤测试
; 说明: 逐步调试PbXls库的基本功能，定位问题
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 步骤1: 使用简单的文件写入测试系统
Define f.i = CreateFile(#PB_Any, "test1.txt")
If f
  WriteStringN(f, "Test1")
  CloseFile(f)
EndIf

; 步骤2: 创建工作簿
Define wb.i = PbXls_CreateWorkbook()
Define f2.i = CreateFile(#PB_Any, "test2.txt")
If f2
  WriteStringN(f2, "工作簿ID=" + Str(wb))
  CloseFile(f2)
EndIf

; 步骤3: 获取工作表
Define ws.i = PbXls_GetSheetByIndex(wb, 0)
Define f3.i = CreateFile(#PB_Any, "test3.txt")
If f3
  WriteStringN(f3, "工作表ID=" + Str(ws))
  CloseFile(f3)
EndIf

; 步骤4: 写入单元格
PbXls_SetCell(ws, 1, 1, "Hello World")
Define f4.i = CreateFile(#PB_Any, "test4.txt")
If f4
  WriteStringN(f4, "单元格写入成功")
  CloseFile(f4)
EndIf

; 步骤5: 保存文件
Define saved.i = PbXls_SaveWorkbook(wb, "test_output.xlsx")
Define f5.i = CreateFile(#PB_Any, "test5.txt")
If f5
  WriteStringN(f5, "保存结果=" + Str(saved))
  If saved
    WriteStringN(f5, "文件大小=" + Str(FileSize("test_output.xlsx")) + " 字节")
  EndIf
  CloseFile(f5)
EndIf

; 清理资源
PbXls_Free()
