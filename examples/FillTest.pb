; ***************************************************************************************
; PbXls 示例代码 - 填充功能测试
; 说明: 测试单元格填充(背景色)功能的创建和应用
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "开始测试"

; 步骤1: 创建工作簿
Define wb.i = PbXls_CreateWorkbook()
Debug "工作簿创建成功, WB=" + Str(wb)

; 步骤2: 获取工作表
Define ws.i = PbXls_GetSheetByIndex(wb, 0)
Debug "工作表获取成功, WS=" + Str(ws)

; 步骤3: 创建填充前检查填充列表
Debug "创建填充前 Fills 列表大小: " + Str(ListSize(PbXls_Fills()))

; 步骤4: 创建填充
Define fill1.i = PbXls_CreateFill()
Debug "填充创建成功, Fill=" + Str(fill1)
Debug "创建填充后 Fills 列表大小: " + Str(ListSize(PbXls_Fills()))

; 步骤5: 设置填充样式(绿色实心填充)
Define result.b = PbXls_SetFill(fill1, "solid", "00FF00", "")
Debug "填充设置完成, SetFill=" + Str(result)

If result = #False
  Debug "填充设置失败! 请检查验证逻辑。"
EndIf

; 步骤6: 写入测试单元格
PbXls_SetCell(ws, 1, 1, "Test")

; 步骤7: 创建样式并应用填充
Define style1.i = PbXls_CreateAlignment()
Debug "样式创建成功, Style=" + Str(style1)
Debug "CellStyles 列表大小: " + Str(ListSize(PbXls_CellStyles()))

; 步骤8: 配置样式
SelectElement(PbXls_CellStyles(), style1)
PbXls_CellStyles()\fontId = 0
PbXls_CellStyles()\fillId = fill1
PbXls_CellStyles()\borderId = 0

; 步骤9: 保存文件
Define saved.b = PbXls_SaveWorkbook(wb, "filltest.xlsx")
Debug "保存结果=" + Str(saved)
If saved
  Debug "文件大小=" + Str(FileSize("filltest.xlsx")) + " 字节"
EndIf

; 清理资源
PbXls_Free()
Debug "测试结束"
