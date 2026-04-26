﻿; PbXls MiniStyleTest - 测试样式功能
XIncludeFile "..\PbXls.pb"

Procedure MiniTest()
  ; 步骤1: 创建基础对象
  Define wb.i = PbXls_CreateWorkbook()
  If wb < 0
    Debug "ERROR: CreateWorkbook failed"
    ProcedureReturn
  EndIf
  Debug "OK: CreateWorkbook=" + Str(wb)
  
  Define ws.i = PbXls_GetSheetByIndex(wb, 0)
  If ws = 0
    Debug "ERROR: GetSheetByIndex failed"
    ProcedureReturn
  EndIf
  Debug "OK: GetSheetByIndex=" + Str(ws)
  
  ; 步骤2: 创建字体
  Define font1.i = PbXls_CreateFont()
  Debug "OK: CreateFont=" + Str(font1)
  
  Define setFontResult.b = PbXls_SetFont(font1, "Arial", 14, #True, #False, "FF0000")
  Debug "OK: SetFont=" + Str(setFontResult)
  
  ; 步骤3: 创建填充
  Define fill1.i = PbXls_CreateFill()
  Debug "OK: CreateFill=" + Str(fill1)
  
  Define setFillResult.b = PbXls_SetFill(fill1, "solid", "00FF00", "")
  Debug "OK: SetFill=" + Str(setFillResult)
  
  ; 步骤4: 创建边框
  Define border1.i = PbXls_CreateBorder()
  Debug "OK: CreateBorder=" + Str(border1)
  
  Define setBorderResult.b = PbXls_SetBorder(border1, "left", "thin", "0000FF")
  PbXls_SetBorder(border1, "right", "thin", "0000FF")
  PbXls_SetBorder(border1, "top", "thin", "0000FF")
  PbXls_SetBorder(border1, "bottom", "thin", "0000FF")
  Debug "OK: SetBorder=" + Str(setBorderResult)
  
  ; 步骤5: 创建样式
  Define style1.i = PbXls_CreateAlignment()
  Debug "OK: CreateAlignment=" + Str(style1)
  
  SelectElement(PbXls_CellStyles(), style1)
  PbXls_CellStyles()\fontId = font1
  PbXls_CellStyles()\fillId = fill1
  PbXls_CellStyles()\borderId = border1
  Debug "OK: Style configured"
  
  ; 步骤6: 写入单元格并应用样式
  Define setCellResult.b = PbXls_SetCell(ws, 1, 1, "Styled Cell")
  Debug "OK: SetCell=" + Str(setCellResult)
  
  Define setStyleResult.b = PbXls_SetCellStyleWS(ws, 1, 1, style1)
  Debug "OK: SetCellStyleWS=" + Str(setStyleResult)
  
  ; 步骤7: 保存
  Define saveResult.b = PbXls_SaveWorkbook(wb, "minitest_style.xlsx")
  Debug "OK: SaveWorkbook=" + Str(saveResult)
  
  If saveResult
    Define fullPath.s = GetCurrentDirectory() + "minitest_style.xlsx"
    Define fileSize.i = FileSize(fullPath)
    Debug "OK: FileSize=" + Str(fileSize)
  EndIf
  
  PbXls_CloseWorkbook(wb)
  PbXls_Free()
  Debug "OK: Cleanup done"
EndProcedure

Debug "=== MiniStyleTest Start ==="
MiniTest()
Debug "=== MiniStyleTest End ==="
