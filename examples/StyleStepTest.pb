; ***************************************************************************************
; PbXls 示例代码 - 样式逐步测试
; 说明: 逐步测试样式功能，定位崩溃点
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 步骤1: 创建工作簿和工作表
Define wb.i = PbXls_CreateWorkbook()
Define ws.i = PbXls_GetSheetByIndex(wb, 0)

; 步骤2: 写入单元格
Debug "1. 写入单元格..."
PbXls_SetCell(ws, 1, 1, "Test1")
Debug "1. 完成"

; 步骤3: 创建字体
Debug "2. 创建字体..."
Define f1.i = PbXls_CreateFont()
PbXls_SetFont(f1, "", -1, #True, #False, "FF0000")
Debug "2. 完成, 字体ID=" + Str(f1)

; 步骤4: 创建样式
Debug "3. 创建样式..."
Define s1.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s1)
PbXls_CellStyles()\fontId = f1
PbXls_CellStyles()\fillId = 0
PbXls_CellStyles()\borderId = 0
Debug "3. 完成, 样式ID=" + Str(s1)

; 步骤5: 设置单元格样式
Debug "4. 设置单元格样式..."
Define r.b = PbXls_SetCellStyleWS(ws, 1, 1, s1)
Debug "4. 完成, 结果=" + Str(r)

; 步骤6: 保存文件
Debug "5. 保存文件..."
Define saved.b = PbXls_SaveWorkbook(wb, "style_step_test.xlsx")
Debug "5. 完成, 保存成功=" + Str(saved)
If saved
  Debug "文件大小=" + Str(FileSize("style_step_test.xlsx")) + " 字节"
EndIf

; 清理资源
PbXls_Free()
Debug "测试完成"
