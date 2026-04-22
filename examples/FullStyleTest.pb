; ***************************************************************************************
; PbXls 示例代码 - 完整样式功能测试
; 说明: 逐步测试找到崩溃点，测试所有样式功能
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 步骤1: 创建工作簿和工作表
Define wb.i = PbXls_CreateWorkbook()
Define ws.i = PbXls_GetSheetByIndex(wb, 0)

; 步骤2: 创建字体
; 字体1: 粗体红色
Define f1.i = PbXls_CreateFont()
PbXls_SetFont(f1, "", -1, #True, #False, "FF0000")
; 字体2: Arial 20号黑色
Define f2.i = PbXls_CreateFont()
PbXls_SetFont(f2, "Arial", 20, #False, #False, "000000")
; 字体3: 斜体紫色
Define f3.i = PbXls_CreateFont()
PbXls_SetFont(f3, "", -1, #False, #True, "800080")
; 字体4: 白色字体(用于深色背景)
Define f4.i = PbXls_CreateFont()
PbXls_SetFont(f4, "", -1, #False, #False, "FFFFFF")

; 步骤3: 创建填充
; 填充1: 绿色
Define fl1.i = PbXls_CreateFill()
PbXls_SetFill(fl1, "solid", "00FF00", "")
; 填充2: 黄色
Define fl2.i = PbXls_CreateFill()
PbXls_SetFill(fl2, "solid", "FFFF00", "")
; 填充3: 红色
Define fl3.i = PbXls_CreateFill()
PbXls_SetFill(fl3, "solid", "FF0000", "")
; 填充4: 蓝色
Define fl4.i = PbXls_CreateFill()
PbXls_SetFill(fl4, "solid", "0000FF", "")
; 填充5: 橙色
Define fl5.i = PbXls_CreateFill()
PbXls_SetFill(fl5, "solid", "FFA500", "")

; 步骤4: 创建边框
; 边框1: 细蓝色边框
Define b1.i = PbXls_CreateBorder()
PbXls_SetBorder(b1, "left", "thin", "0000FF")
PbXls_SetBorder(b1, "right", "thin", "0000FF")
PbXls_SetBorder(b1, "top", "thin", "0000FF")
PbXls_SetBorder(b1, "bottom", "thin", "0000FF")
; 边框2: 粗黑色边框
Define b2.i = PbXls_CreateBorder()
PbXls_SetBorder(b2, "left", "thick", "000000")
PbXls_SetBorder(b2, "right", "thick", "000000")
PbXls_SetBorder(b2, "top", "thick", "000000")
PbXls_SetBorder(b2, "bottom", "thick", "000000")
; 边框3: 双线橙色边框
Define b3.i = PbXls_CreateBorder()
PbXls_SetBorder(b3, "left", "double", "FF8C00")
PbXls_SetBorder(b3, "right", "double", "FF8C00")
PbXls_SetBorder(b3, "top", "double", "FF8C00")
PbXls_SetBorder(b3, "bottom", "double", "FF8C00")

; 步骤5: 创建组合样式
; 样式1: 粗体红色字体
Define s1.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s1)
PbXls_CellStyles()\fontId = f1
PbXls_CellStyles()\fillId = 0
PbXls_CellStyles()\borderId = 0

; 样式2: 绿色背景
Define s2.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s2)
PbXls_CellStyles()\fontId = 0
PbXls_CellStyles()\fillId = fl1
PbXls_CellStyles()\borderId = 0

; 样式3: 细蓝色边框
Define s3.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s3)
PbXls_CellStyles()\fontId = 0
PbXls_CellStyles()\fillId = 0
PbXls_CellStyles()\borderId = b1

; 样式4: Arial 20号大字体
Define s4.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s4)
PbXls_CellStyles()\fontId = f2
PbXls_CellStyles()\fillId = 0
PbXls_CellStyles()\borderId = 0

; 样式5: 斜体紫色字体
Define s5.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s5)
PbXls_CellStyles()\fontId = f3
PbXls_CellStyles()\fillId = 0
PbXls_CellStyles()\borderId = 0

; 样式6: 完整样式(粗体红字+黄底+橙双线边框)
Define s6.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s6)
PbXls_CellStyles()\fontId = f1
PbXls_CellStyles()\fillId = fl2
PbXls_CellStyles()\borderId = b3

; 样式7: 红色背景白字
Define s7.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s7)
PbXls_CellStyles()\fontId = f4
PbXls_CellStyles()\fillId = fl3
PbXls_CellStyles()\borderId = 0

; 样式8: 蓝色背景白字
Define s8.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s8)
PbXls_CellStyles()\fontId = f4
PbXls_CellStyles()\fillId = fl4
PbXls_CellStyles()\borderId = 0

; 样式9: 橙色背景白字
Define s9.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s9)
PbXls_CellStyles()\fontId = f4
PbXls_CellStyles()\fillId = fl5
PbXls_CellStyles()\borderId = 0

; 样式10: 粗黑色边框
Define s10.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s10)
PbXls_CellStyles()\fontId = 0
PbXls_CellStyles()\fillId = 0
PbXls_CellStyles()\borderId = b2

; 样式13: 居中对齐
Define s13.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), s13)
PbXls_CellStyles()\alignment\horizontal = "center"
PbXls_CellStyles()\alignment\vertical = "center"

; 步骤6: 写入单元格数据并应用样式
; 第一行: 基本字体样式
PbXls_SetCell(ws, 1, 1, "Test1")
PbXls_SetCellStyleWS(ws, 1, 1, s1)
PbXls_SetCell(ws, 1, 2, "Test2")
PbXls_SetCellStyleWS(ws, 1, 2, s2)
PbXls_SetCell(ws, 1, 3, "Test3")
PbXls_SetCellStyleWS(ws, 1, 3, s3)

; 第二行: 更多字体样式
PbXls_SetCell(ws, 2, 1, "Test4")
PbXls_SetCellStyleWS(ws, 2, 1, s4)
PbXls_SetCell(ws, 2, 2, "Test5")
PbXls_SetCellStyleWS(ws, 2, 2, s5)

; 第三行: 合并单元格+完整样式
PbXls_SetCell(ws, 3, 1, "Test6")
PbXls_SetCellStyleWS(ws, 3, 1, s6)
PbXls_MergeCells(ws, "A3:C3")

; 第四行: 背景色样式
PbXls_SetCell(ws, 4, 1, "Test7")
PbXls_SetCellStyleWS(ws, 4, 1, s7)
PbXls_SetCell(ws, 4, 2, "Test8")
PbXls_SetCellStyleWS(ws, 4, 2, s8)
PbXls_SetCell(ws, 4, 3, "Test9")
PbXls_SetCellStyleWS(ws, 4, 3, s9)

; 第五行: 边框样式
PbXls_SetCell(ws, 5, 1, "Test10")
PbXls_SetCellStyleWS(ws, 5, 1, s10)
PbXls_SetCell(ws, 5, 2, "Test11")
PbXls_SetCellStyleWS(ws, 5, 2, s3)

; 第六行: 居中对齐
PbXls_SetCell(ws, 5, 3, "Test12")
PbXls_SetCell(ws, 6, 1, "Test13")
PbXls_SetCellStyleWS(ws, 6, 1, s13)

; 设置列宽
PbXls_SetColumnWidth(ws, 1, 20.0)
PbXls_SetColumnWidth(ws, 2, 20.0)
PbXls_SetColumnWidth(ws, 3, 20.0)

; 步骤7: 保存文件
Define saved.b = PbXls_SaveWorkbook(wb, "full_style_test.xlsx")
If saved
  Define sz.i = FileSize("full_style_test.xlsx")
  Debug "保存成功, 文件大小=" + Str(sz) + " 字节"
Else
  Debug "保存失败!"
EndIf

; 清理资源
PbXls_CloseWorkbook(wb)
PbXls_Free()
