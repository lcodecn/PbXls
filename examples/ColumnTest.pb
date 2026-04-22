; ***************************************************************************************
; PbXls 示例代码 - 列插入/删除功能测试
; 说明: 测试列插入、列删除功能，验证单元格移动、列宽更新、合并单元格更新
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Procedure PbXls_RunColumnTest()
  Define outputFile.s = "column_test.xlsx"
  DeleteFile(outputFile, #PB_FileSystem_Force)
  
  Debug ""
  Debug "=== PbXls 列插入/删除功能测试 ==="
  Debug ""
  
  ; ===== 第一部分：创建基础数据 =====
  Debug "1. 创建基础工作簿和数据..."
  
  Define wbId.i = PbXls_CreateWorkbook()
  Define wsId.i = PbXls_GetSheetByIndex(wbId, 0)
  
  ; 写入初始数据: A1=1, B1=2, C1=3, D1=4, E1=5
  Define col.i
  For col = 1 To 5
    PbXls_SetCell(wsId, 1, col, "原始" + Str(col))
  Next
  
  ; 设置列宽
  PbXls_SetColumnWidth(wsId, 1, 10.0)
  PbXls_SetColumnWidth(wsId, 2, 15.0)
  PbXls_SetColumnWidth(wsId, 3, 20.0)
  PbXls_SetColumnWidth(wsId, 4, 25.0)
  PbXls_SetColumnWidth(wsId, 5, 30.0)
  
  ; 合并单元格 B1:D1
  PbXls_MergeCells(wsId, "B1:D1")
  
  Debug "   初始数据: A1=原始1, B1=原始2, C1=原始3, D1=原始4, E1=原始5"
  Debug "   列宽: A=10, B=15, C=20, D=25, E=30"
  Debug "   合并单元格: B1:D1"
  
  ; ===== 第二部分：列插入测试 =====
  Debug ""
  Debug "2. 在C列前插入2列..."
  
  PbXls_InsertColumns(wsId, 3, 2)
  
  Debug "   插入后数据应该为: A1=原始1, B1=原始2, C1=(空), D1=(空), E1=原始3, F1=原始4, G1=原始5"
  Debug "   合并单元格应该更新为: B1:F1 (原B1:D1)"
  
  ; ===== 第三部分：写入新数据到插入的列 =====
  Debug ""
  Debug "3. 在插入的列中写入数据..."
  
  PbXls_SetCell(wsId, 1, 3, "新C")
  PbXls_SetCell(wsId, 1, 4, "新D")
  PbXls_SetColumnWidth(wsId, 3, 12.0)
  PbXls_SetColumnWidth(wsId, 4, 18.0)
  
  Debug "   C1=新C, D1=新D"
  
  ; ===== 第四部分：列删除测试 =====
  Debug ""
  Debug "4. 删除B列..."
  
  PbXls_DeleteColumns(wsId, 2, 1)
  
  Debug "   删除后数据应该为: A1=原始1, B1=(空), C1=新C, D1=新D, E1=原始3, F1=原始4"
  Debug "   合并单元格应该更新: A1:E1 (原B1:F1向前移动)"
  
  ; ===== 第五部分：保存文件 =====
  Debug ""
  Debug "5. 保存文件..."
  
  If PbXls_SaveWorkbook(wbId, outputFile)
    Define fullPath.s = GetCurrentDirectory() + outputFile
    Define fileSize.i = FileSize(fullPath)
    Debug "   保存成功: " + outputFile
    Debug "   文件大小: " + Str(fileSize) + " 字节"
    Debug ""
    Debug "=== 列插入/删除测试完成 ==="
  Else
    Debug "   保存失败!"
  EndIf
  
  PbXls_CloseWorkbook(wbId)
  PbXls_Free()
EndProcedure

PbXls_RunColumnTest()

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 92
; Folding = -
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory