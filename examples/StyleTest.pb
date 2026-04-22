; ***************************************************************************************
; PbXls 示例代码 - 单元格样式功能测试
; 说明: 测试单元格字体(名称/字号/字色/粗体/斜体)、单元格底色、边框、对齐等样式功能
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Procedure PbXls_RunStyleTest()
  Define outputFile.s = "style_test.xlsx"
  DeleteFile(outputFile, #PB_FileSystem_Force)
  
  Debug ""
  Debug "=== PbXls 单元格样式功能测试 ==="
  Debug ""
  
  Define wbId.i = PbXls_CreateWorkbook()
  Define wsId.i = PbXls_GetSheetByIndex(wbId, 0)
  
  ; ===== 第1部分：创建字体 =====
  Debug "1. 创建字体..."
  
  ; 字体1：粗体红色 Calibri 14
  Define fontBoldRed.i = PbXls_CreateFont()
  PbXls_SetFont(fontBoldRed, "", 14, #True, #False, "FF0000")
  Debug "   fontBoldRed ID=" + Str(fontBoldRed)
  
  ; 字体2：大字号 Arial 20
  Define fontBig.i = PbXls_CreateFont()
  PbXls_SetFont(fontBig, "Arial", 20, #False, #False, "000000")
  Debug "   fontBig ID=" + Str(fontBig)
  
  ; 字体3：斜体紫色
  Define fontItalic.i = PbXls_CreateFont()
  PbXls_SetFont(fontItalic, "", -1, #False, #True, "800080")
  Debug "   fontItalic ID=" + Str(fontItalic)
  
  ; 字体4：白色字体（用于深色背景）
  Define fontWhite.i = PbXls_CreateFont()
  PbXls_SetFont(fontWhite, "", -1, #False, #False, "FFFFFF")
  Debug "   fontWhite ID=" + Str(fontWhite)
  
  ; 字体5：蓝色下划线
  Define fontBlueUnderline.i = PbXls_CreateFont()
  PbXls_SetFont(fontBlueUnderline, "", 12, #False, #False, "0000FF")
  PbXls_Fonts()\underline = #True
  Debug "   fontBlueUnderline ID=" + Str(fontBlueUnderline)
  
  ; ===== 第2部分：创建填充 =====
  Debug ""
  Debug "2. 创建填充..."
  
  Define fillGreen.i = PbXls_CreateFill()
  PbXls_SetFill(fillGreen, "solid", "00FF00", "")
  Debug "   fillGreen ID=" + Str(fillGreen)
  
  Define fillYellow.i = PbXls_CreateFill()
  PbXls_SetFill(fillYellow, "solid", "FFFF00", "")
  Debug "   fillYellow ID=" + Str(fillYellow)
  
  Define fillRed.i = PbXls_CreateFill()
  PbXls_SetFill(fillRed, "solid", "FF0000", "")
  Debug "   fillRed ID=" + Str(fillRed)
  
  Define fillBlue.i = PbXls_CreateFill()
  PbXls_SetFill(fillBlue, "solid", "0000FF", "")
  Debug "   fillBlue ID=" + Str(fillBlue)
  
  Define fillOrange.i = PbXls_CreateFill()
  PbXls_SetFill(fillOrange, "solid", "FFA500", "")
  Debug "   fillOrange ID=" + Str(fillOrange)
  
  Define fillGray.i = PbXls_CreateFill()
  PbXls_SetFill(fillGray, "solid", "CCCCCC", "")
  Debug "   fillGray ID=" + Str(fillGray)
  
  ; ===== 第3部分：创建边框 =====
  Debug ""
  Debug "3. 创建边框..."
  
  Define borderThin.i = PbXls_CreateBorder()
  PbXls_SetBorder(borderThin, "left", "thin", "000000")
  PbXls_SetBorder(borderThin, "right", "thin", "000000")
  PbXls_SetBorder(borderThin, "top", "thin", "000000")
  PbXls_SetBorder(borderThin, "bottom", "thin", "000000")
  Debug "   borderThin ID=" + Str(borderThin)
  
  Define borderThick.i = PbXls_CreateBorder()
  PbXls_SetBorder(borderThick, "left", "thick", "000000")
  PbXls_SetBorder(borderThick, "right", "thick", "000000")
  PbXls_SetBorder(borderThick, "top", "thick", "000000")
  PbXls_SetBorder(borderThick, "bottom", "thick", "000000")
  Debug "   borderThick ID=" + Str(borderThick)
  
  Define borderDouble.i = PbXls_CreateBorder()
  PbXls_SetBorder(borderDouble, "left", "double", "FF0000")
  PbXls_SetBorder(borderDouble, "right", "double", "FF0000")
  PbXls_SetBorder(borderDouble, "top", "double", "FF0000")
  PbXls_SetBorder(borderDouble, "bottom", "double", "FF0000")
  Debug "   borderDouble ID=" + Str(borderDouble)
  
  Define borderBlueDashed.i = PbXls_CreateBorder()
  PbXls_SetBorder(borderBlueDashed, "left", "dashed", "0000FF")
  PbXls_SetBorder(borderBlueDashed, "right", "dashed", "0000FF")
  PbXls_SetBorder(borderBlueDashed, "top", "dashed", "0000FF")
  PbXls_SetBorder(borderBlueDashed, "bottom", "dashed", "0000FF")
  Debug "   borderBlueDashed ID=" + Str(borderBlueDashed)
  
  ; ===== 第4部分：创建组合样式 =====
  Debug ""
  Debug "4. 创建组合样式..."
  
  ; 样式1：粗体红色字
  Define s1.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s1)
  PbXls_CellStyles()\fontId = fontBoldRed
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  Debug "   s1(粗体红色字) ID=" + Str(s1)
  
  ; 样式2：绿色背景
  Define s2.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s2)
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = fillGreen
  PbXls_CellStyles()\borderId = 0
  Debug "   s2(绿色背景) ID=" + Str(s2)
  
  ; 样式3：细边框
  Define s3.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s3)
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = borderThin
  Debug "   s3(细边框) ID=" + Str(s3)
  
  ; 样式4：大字体
  Define s4.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s4)
  PbXls_CellStyles()\fontId = fontBig
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  Debug "   s4(大字体) ID=" + Str(s4)
  
  ; 样式5：斜体紫色
  Define s5.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s5)
  PbXls_CellStyles()\fontId = fontItalic
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  Debug "   s5(斜体紫色) ID=" + Str(s5)
  
  ; 样式6：完整样式(粗体红字+黄底+红双边框)
  Define s6.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s6)
  PbXls_CellStyles()\fontId = fontBoldRed
  PbXls_CellStyles()\fillId = fillYellow
  PbXls_CellStyles()\borderId = borderDouble
  Debug "   s6(完整样式) ID=" + Str(s6)
  
  ; 样式7：红色背景白字
  Define s7.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s7)
  PbXls_CellStyles()\fontId = fontWhite
  PbXls_CellStyles()\fillId = fillRed
  PbXls_CellStyles()\borderId = 0
  Debug "   s7(红底白字) ID=" + Str(s7)
  
  ; 样式8：蓝色背景白字
  Define s8.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s8)
  PbXls_CellStyles()\fontId = fontWhite
  PbXls_CellStyles()\fillId = fillBlue
  PbXls_CellStyles()\borderId = 0
  Debug "   s8(蓝底白字) ID=" + Str(s8)
  
  ; 样式9：橙色背景白字
  Define s9.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s9)
  PbXls_CellStyles()\fontId = fontWhite
  PbXls_CellStyles()\fillId = fillOrange
  PbXls_CellStyles()\borderId = 0
  Debug "   s9(橙底白字) ID=" + Str(s9)
  
  ; 样式10：粗边框
  Define s10.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s10)
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = borderThick
  Debug "   s10(粗边框) ID=" + Str(s10)
  
  ; 样式11：蓝色虚线边框
  Define s11.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s11)
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = fillGray
  PbXls_CellStyles()\borderId = borderBlueDashed
  Debug "   s11(灰底蓝色虚线边框) ID=" + Str(s11)
  
  ; 样式12：居中对齐
  Define s12.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s12)
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  PbXls_CellStyles()\alignment\horizontal = "center"
  PbXls_CellStyles()\alignment\vertical = "center"
  Debug "   s12(居中) ID=" + Str(s12)
  
  ; 样式13：自动换行+缩进
  Define s13.i = PbXls_CreateAlignment()
  SelectElement(PbXls_CellStyles(), s13)
  PbXls_CellStyles()\fontId = 0
  PbXls_CellStyles()\fillId = 0
  PbXls_CellStyles()\borderId = 0
  PbXls_CellStyles()\alignment\horizontal = "left"
  PbXls_CellStyles()\alignment\vertical = "top"
  PbXls_CellStyles()\alignment\wrapText = #True
  PbXls_CellStyles()\alignment\indent = 2
  Debug "   s13(换行+缩进) ID=" + Str(s13)
  
  ; ===== 第5部分：写入数据并应用样式 =====
  Debug ""
  Debug "5. 写入数据并应用样式..."
  
  ; 第1行：基本字体样式
  PbXls_SetCell(wsId, 1, 1, "BoldRed")
  PbXls_SetCellStyleWS(wsId, 1, 1, s1)
  
  PbXls_SetCell(wsId, 1, 2, "BigFont")
  PbXls_SetCellStyleWS(wsId, 1, 2, s4)
  
  PbXls_SetCell(wsId, 1, 3, "Italic")
  PbXls_SetCellStyleWS(wsId, 1, 3, s5)
  
  PbXls_SetCell(wsId, 1, 4, "Underline")
  PbXls_SetCellStyleWS(wsId, 1, 4, s1)
  SelectElement(PbXls_CellStyles(), s1)
  PbXls_CellStyles()\fontId = fontBlueUnderline
  PbXls_SetCellStyleWS(wsId, 1, 4, s1)
  
  ; 第2行：背景色
  PbXls_SetCell(wsId, 2, 1, "GreenBG")
  PbXls_SetCellStyleWS(wsId, 2, 1, s2)
  
  PbXls_SetCell(wsId, 2, 2, "RedWhite")
  PbXls_SetCellStyleWS(wsId, 2, 2, s7)
  
  PbXls_SetCell(wsId, 2, 3, "BlueWhite")
  PbXls_SetCellStyleWS(wsId, 2, 3, s8)
  
  PbXls_SetCell(wsId, 2, 4, "OrangeWhite")
  PbXls_SetCellStyleWS(wsId, 2, 4, s9)
  
  ; 第3行：边框
  PbXls_SetCell(wsId, 3, 1, "ThinBorder")
  PbXls_SetCellStyleWS(wsId, 3, 1, s3)
  
  PbXls_SetCell(wsId, 3, 2, "ThickBorder")
  PbXls_SetCellStyleWS(wsId, 3, 2, s10)
  
  PbXls_SetCell(wsId, 3, 3, "DblBorder")
  PbXls_SetCellStyleWS(wsId, 3, 3, s6)
  
  PbXls_SetCell(wsId, 3, 4, "DashedBorder")
  PbXls_SetCellStyleWS(wsId, 3, 4, s11)
  
  ; 第4行：对齐
  PbXls_SetCell(wsId, 4, 1, "Center")
  PbXls_SetCellStyleWS(wsId, 4, 1, s12)
  
  PbXls_SetCell(wsId, 4, 2, "Wrap Indent Text Here For Testing")
  PbXls_SetCellStyleWS(wsId, 4, 2, s13)
  
  ; 第5行：合并单元格+完整样式
  PbXls_SetCell(wsId, 5, 1, "Full Style: Bold Red + Yellow BG + Red Double Border")
  PbXls_SetCellStyleWS(wsId, 5, 1, s6)
  PbXls_MergeCells(wsId, "A5:D5")
  
  ; 设置列宽
  PbXls_SetColumnWidth(wsId, 1, 18.0)
  PbXls_SetColumnWidth(wsId, 2, 18.0)
  PbXls_SetColumnWidth(wsId, 3, 18.0)
  PbXls_SetColumnWidth(wsId, 4, 18.0)
  
  ; 设置行高
  PbXls_SetRowHeight(wsId, 4, 40.0)
  PbXls_SetRowHeight(wsId, 5, 30.0)
  
  Debug "   写入完成: 5行 x 4列 + 1合并单元格"
  
  ; ===== 第6部分：保存文件 =====
  Debug ""
  Debug "6. 保存文件..."
  
  If PbXls_SaveWorkbook(wbId, outputFile)
    Define fullPath.s = GetCurrentDirectory() + outputFile
    Define fileSize.i = FileSize(fullPath)
    Debug "   保存成功: " + outputFile
    Debug "   文件大小: " + Str(fileSize) + " 字节"
    Debug ""
    Debug "=== 样式功能测试完成 ==="
  Else
    Debug "   保存失败!"
  EndIf
  
  PbXls_CloseWorkbook(wbId)
  PbXls_Free()
EndProcedure

PbXls_RunStyleTest()
