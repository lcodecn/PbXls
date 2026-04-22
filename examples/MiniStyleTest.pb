; ***************************************************************************************
; PbXls 示例代码 - 迷你样式测试
; 说明: 测试基本的样式创建和应用功能
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 创建日志文件
Define log.i = CreateFile(#PB_Any, "minitest_log.txt")
WriteStringN(log, "测试开始")

; 步骤1: 创建工作簿
Define wb.i = PbXls_CreateWorkbook()
WriteStringN(log, "工作簿创建成功, WB=" + Str(wb))

; 步骤2: 获取工作表
Define ws.i = PbXls_GetSheetByIndex(wb, 0)
WriteStringN(log, "工作表获取成功, WS=" + Str(ws))

; 步骤3: 创建字体
Define f1.i = PbXls_CreateFont()
WriteStringN(log, "字体创建成功, Font=" + Str(f1))

; 步骤4: 设置字体样式
Define r1.b = PbXls_SetFont(f1, "Arial", 14, #True, #False, "FF0000")
WriteStringN(log, "字体设置完成, SetFont=" + Str(r1))

; 步骤5: 创建填充
Define fl1.i = PbXls_CreateFill()
WriteStringN(log, "填充创建成功, Fill=" + Str(fl1))

; 步骤6: 设置填充样式
Define r2.b = PbXls_SetFill(fl1, "solid", "00FF00", "")
WriteStringN(log, "填充设置完成, SetFill=" + Str(r2))

; 步骤7: 创建边框
Define b1.i = PbXls_CreateBorder()
WriteStringN(log, "边框创建成功, Border=" + Str(b1))

; 步骤8: 设置边框样式
Define r3.b = PbXls_SetBorder(b1, "left", "thin", "0000FF")
WriteStringN(log, "边框设置完成, SetBorder=" + Str(r3))

; 步骤9: 创建对齐/样式对象
Define s1.i = PbXls_CreateAlignment()
WriteStringN(log, "样式创建成功, Style=" + Str(s1))

; 步骤10: 配置样式
SelectElement(PbXls_CellStyles(), s1)
PbXls_CellStyles()\fontId = f1
PbXls_CellStyles()\fillId = fl1
PbXls_CellStyles()\borderId = b1
WriteStringN(log, "样式配置完成")

; 步骤11: 写入单元格并应用样式
PbXls_SetCell(ws, 1, 1, "Hello")
PbXls_SetCellStyleWS(ws, 1, 1, s1)
WriteStringN(log, "单元格写入完成")

; 步骤12: 保存文件
Define saved.b = PbXls_SaveWorkbook(wb, "minitest.xlsx")
WriteStringN(log, "保存结果=" + Str(saved))

If saved
  WriteStringN(log, "文件大小=" + Str(FileSize("minitest.xlsx")) + " 字节")
EndIf

; 清理资源
PbXls_CloseWorkbook(wb)
PbXls_Free()
WriteStringN(log, "测试完成")

CloseFile(log)
