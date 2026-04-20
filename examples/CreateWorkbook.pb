; ***************************************************************************************
; PbXls 示例代码 - 创建Excel工作簿
; 说明: 演示如何使用PbXls库创建和保存Excel文件
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 创建工作簿
Debug "=== PbXls 创建Excel工作簿示例 ==="
Debug ""

; 1. 创建新工作簿
Debug "1. 创建新工作簿..."
wbId.i = PbXls_CreateWorkbook()
If wbId = -1
  Debug "错误: 无法创建工作簿"
  End
EndIf
Debug "  工作簿创建成功, ID: " + Str(wbId)

; 2. 获取活动工作表
Debug "2. 获取活动工作表..."
wsId.i = PbXls_ActiveSheet(wbId)
If wsId = 0
  Debug "错误: 无法获取活动工作表"
  End
EndIf
Debug "  活动工作表获取成功, 指针: " + Str(wsId)

; 3. 设置单元格值
Debug "3. 设置单元格值..."

; A1: 字符串
PbXls_SetCell(wsId, 1, 1, "Hello World")
Debug "  A1 = 'Hello World'"

; B1: 数字
PbXls_SetCell(wsId, 1, 2, "123.45")
Debug "  B1 = 123.45"

; C1: 布尔值
PbXls_SetCell(wsId, 1, 3, "TRUE")
Debug "  C1 = TRUE"

; D1: 公式
PbXls_SetCellFormulaWS(wsId, 1, 4, "B1*2")
Debug "  D1 = =B1*2"

; A2: 中文
PbXls_SetCell(wsId, 2, 1, "你好世界")
Debug "  A2 = '你好世界'"

; 4. 追加行数据
Debug "4. 追加行数据..."
NewList values.s()
AddElement(values())
values() = "张三"
AddElement(values())
values() = "25"
AddElement(values())
values() = "北京"
PbXls_AppendRow(wsId, values())
Debug "  第3行: 张三, 25, 北京"

ClearList(values())
AddElement(values())
values() = "李四"
AddElement(values())
values() = "30"
AddElement(values())
values() = "上海"
PbXls_AppendRow(wsId, values())
Debug "  第4行: 李四, 30, 上海"

; 5. 设置列宽
Debug "5. 设置列宽..."
PbXls_SetColumnWidth(wsId, 1, 15)  ; A列宽15
PbXls_SetColumnWidth(wsId, 2, 10)  ; B列宽10
PbXls_SetColumnWidth(wsId, 3, 15)  ; C列宽15
Debug "  A列宽=15, B列宽=10, C列宽=15"

; 6. 合并单元格
Debug "6. 合并单元格..."
PbXls_MergeCells(wsId, "A5:C5")
PbXls_SetCell(wsId, 5, 1, "这是合并的单元格 A5:C5")
Debug "  合并 A5:C5"

; 7. 创建第二个工作表
Debug "7. 创建第二个工作表..."
sheet2Idx.i = PbXls_CreateSheet(wbId, "数据表")
Debug "  第二个工作表创建成功, 索引: " + Str(sheet2Idx)

ws2Id.i = PbXls_GetSheetByIndex(wbId, sheet2Idx)
If ws2Id <> 0
  PbXls_SetCell(ws2Id, 1, 1, "数据表的第一行第一列")
  Debug "  设置数据表 A1 = '数据表的第一行第一列'"
EndIf

; 8. 保存工作簿
Debug "8. 保存工作簿..."
outputFile.s = "demo_create.xlsx"
If PbXls_SaveWorkbook(wbId, outputFile)
  Debug "  保存成功: " + outputFile
Else
  Debug "  保存失败!"
EndIf

; 9. 获取工作簿信息
Debug ""
Debug "=== 工作簿信息 ==="
Debug "  工作表数量: " + Str(PbXls_GetSheetCount(wbId))

For i.i = 0 To PbXls_GetSheetCount(wbId) - 1
  Debug "  工作表" + Str(i) + ": " + PbXls_GetSheetName(wbId, i)
Next

; 10. 清理
Debug ""
Debug "10. 清理资源..."
PbXls_CloseWorkbook(wbId)
PbXls_Free()

Debug ""
Debug "=== 示例完成 ==="
Debug "请用Excel打开 " + outputFile + " 查看结果"
