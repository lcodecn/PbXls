; ***************************************************************************************
; PbXls 示例代码 - 简单样式保存测试
; 说明: 测试样式创建和保存功能的基本流程
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 样式保存测试 ==="

; 步骤1: 创建工作簿和工作表
Define wbId.i = PbXls_CreateWorkbook()
Define wsId.i = PbXls_GetSheetByIndex(wbId, 0)

Debug "1. 创建样式..."

; 创建字体
Define font1.i = PbXls_CreateFont()
PbXls_SetFont(font1, "Arial", 14, #True, #False, "FF0000")
Debug "   字体 ID=" + Str(font1)

; 创建填充
Define fill1.i = PbXls_CreateFill()
PbXls_SetFill(fill1, "solid", "FFFF00", "")
Debug "   填充 ID=" + Str(fill1)

; 创建组合样式
Define style1.i = PbXls_CreateAlignment()
SelectElement(PbXls_CellStyles(), style1)
PbXls_CellStyles()\fontId = font1
PbXls_CellStyles()\fillId = fill1
PbXls_CellStyles()\borderId = 0
Debug "   样式 ID=" + Str(style1)

Debug "2. 写入数据..."
; 写入带样式的单元格
PbXls_SetCell(wsId, 1, 1, "Hello Style")
PbXls_SetCellStyleWS(wsId, 1, 1, style1)

; 写入无样式的单元格
PbXls_SetCell(wsId, 2, 1, "No Style")

Debug "3. 保存文件..."
Define outputFile.s = "simple_style_test.xlsx"
Define result.i = PbXls_SaveWorkbook(wbId, outputFile)
Debug "保存结果: " + Str(result)

If result
  Define fullPath.s = GetCurrentDirectory() + outputFile
  Define fileSize.i = FileSize(fullPath)
  Debug "文件大小: " + Str(fileSize) + " 字节"
EndIf

; 清理资源
PbXls_CloseWorkbook(wbId)
PbXls_Free()
Debug "=== 测试完成 ==="
