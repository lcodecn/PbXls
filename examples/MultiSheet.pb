; ***************************************************************************************
; PbXls 示例代码 - 多工作表操作
; 说明: 演示如何在一个工作簿中创建和管理多个工作表
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 多工作表演示 ==="
Debug ""

; 1. 创建新工作簿
Debug "1. 创建工作簿..."
wbId.i = PbXls_CreateWorkbook()
If wbId = -1
  Debug "  错误: 无法创建工作簿"
  End
EndIf
Debug "  工作簿创建成功"

; 2. 重命名默认工作表
Debug "2. 重命名默认工作表..."
ws1Id.i = PbXls_ActiveSheet(wbId)
Debug "  第一个工作表已就绪"

; 写入第一个工作表数据 - 员工信息
Debug "3. 写入第一个工作表（员工信息）..."
PbXls_SetCell(ws1Id, 1, 1, "员工信息表")
PbXls_MergeCells(ws1Id, "A1:C1")
PbXls_SetColumnWidth(ws1Id, 1, 15)
PbXls_SetColumnWidth(ws1Id, 2, 10)
PbXls_SetColumnWidth(ws1Id, 3, 12)

; 写入表头
PbXls_SetCell(ws1Id, 2, 1, "姓名")
PbXls_SetCell(ws1Id, 2, 2, "部门")
PbXls_SetCell(ws1Id, 2, 3, "职位")

; 写入员工数据
NewList empData.s()
AddElement(empData()): empData() = "张三"
AddElement(empData()): empData() = "技术部"
AddElement(empData()): empData() = "工程师"
PbXls_AppendRow(ws1Id, empData())

ClearList(empData())
AddElement(empData()): empData() = "李四"
AddElement(empData()): empData() = "市场部"
AddElement(empData()): empData() = "经理"
PbXls_AppendRow(ws1Id, empData())

ClearList(empData())
AddElement(empData()): empData() = "王五"
AddElement(empData()): empData() = "人事部"
AddElement(empData()): empData() = "主管"
PbXls_AppendRow(ws1Id, empData())
Debug "  员工信息写入完成"

; 3. 创建第二个工作表 - 部门统计
Debug "4. 创建第二个工作表（部门统计）..."
ws2Idx.i = PbXls_CreateSheet(wbId, "部门统计")
ws2Id.i = PbXls_GetSheetByIndex(wbId, ws2Idx)
If ws2Id <> 0
  PbXls_SetCell(ws2Id, 1, 1, "部门统计表")
  PbXls_MergeCells(ws2Id, "A1:B1")
  PbXls_SetColumnWidth(ws2Id, 1, 15)
  PbXls_SetColumnWidth(ws2Id, 2, 10)
  
  PbXls_SetCell(ws2Id, 2, 1, "部门")
  PbXls_SetCell(ws2Id, 2, 2, "人数")
  
  PbXls_SetCell(ws2Id, 3, 1, "技术部")
  PbXls_SetCell(ws2Id, 3, 2, "50")
  
  PbXls_SetCell(ws2Id, 4, 1, "市场部")
  PbXls_SetCell(ws2Id, 4, 2, "30")
  
  PbXls_SetCell(ws2Id, 5, 1, "人事部")
  PbXls_SetCell(ws2Id, 5, 2, "20")
  
  PbXls_SetCell(ws2Id, 6, 1, "总计")
  PbXls_SetCellFormulaWS(ws2Id, 6, 2, "=SUM(B3:B5)")
  Debug "  部门统计写入完成"
EndIf

; 4. 创建第三个工作表 - 备注
Debug "5. 创建第三个工作表（备注）..."
ws3Idx.i = PbXls_CreateSheet(wbId, "备注")
ws3Id.i = PbXls_GetSheetByIndex(wbId, ws3Idx)
If ws3Id <> 0
  PbXls_SetCell(ws3Id, 1, 1, "备注信息")
  PbXls_SetCell(ws3Id, 2, 1, "此报表由PbXls库自动生成")
  PbXls_SetCell(ws3Id, 3, 1, "日期: 2026-04-20")
  Debug "  备注写入完成"
EndIf

; 5. 显示工作簿信息
Debug ""
Debug "=== 工作簿信息 ==="
Debug "  工作表数量: " + Str(PbXls_GetSheetCount(wbId))

For i.i = 0 To PbXls_GetSheetCount(wbId) - 1
  Debug "  工作表" + Str(i) + ": " + PbXls_GetSheetName(wbId, i)
Next

; 6. 删除第三个工作表
Debug ""
Debug "6. 删除第三个工作表（备注）..."
If PbXls_RemoveSheet(wbId, 2)
  Debug "  删除成功"
Else
  Debug "  删除失败"
EndIf

Debug "  剩余工作表数量: " + Str(PbXls_GetSheetCount(wbId))

; 7. 保存工作簿
Debug ""
Debug "7. 保存工作簿..."
outputFile.s = "demo_multisheet.xlsx"
If PbXls_SaveWorkbook(wbId, outputFile)
  Debug "  保存成功: " + outputFile
Else
  Debug "  保存失败!"
EndIf

; 8. 清理
PbXls_CloseWorkbook(wbId)
PbXls_Free()

Debug ""
Debug "=== 示例完成 ==="
Debug "请用Excel打开 " + outputFile + " 查看结果"
