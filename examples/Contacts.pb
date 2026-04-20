; ***************************************************************************************
; PbXls 示例代码 - 联系人管理
; 说明: 演示如何使用PbXls库创建联系人管理表格，包含数据分类和筛选
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 联系人管理示例 ==="
Debug ""

; 1. 创建新工作簿
Debug "1. 创建工作簿..."
wbId.i = PbXls_CreateWorkbook()
If wbId = -1
  Debug "  错误: 无法创建工作簿"
  End
EndIf

wsId.i = PbXls_ActiveSheet(wbId)

; 2. 设置标题和列宽
Debug "2. 设置表格格式..."
PbXls_SetCell(wsId, 1, 1, "联系人列表")
PbXls_MergeCells(wsId, "A1:F1")

; 设置列宽
PbXls_SetColumnWidth(wsId, 1, 12)  ; 姓名
PbXls_SetColumnWidth(wsId, 2, 10)  ; 性别
PbXls_SetColumnWidth(wsId, 3, 15)  ; 电话
PbXls_SetColumnWidth(wsId, 4, 25)  ; 邮箱
PbXls_SetColumnWidth(wsId, 5, 15)  ; 部门
PbXls_SetColumnWidth(wsId, 6, 15)  ; 职位

; 3. 写入表头
Debug "3. 写入表头..."
PbXls_SetCell(wsId, 2, 1, "姓名")
PbXls_SetCell(wsId, 2, 2, "性别")
PbXls_SetCell(wsId, 2, 3, "电话")
PbXls_SetCell(wsId, 2, 4, "邮箱")
PbXls_SetCell(wsId, 2, 5, "部门")
PbXls_SetCell(wsId, 2, 6, "职位")

; 设置自动筛选
PbXls_SetAutoFilter(wsId, "A2:F2")

; 4. 写入联系人数据
Debug "4. 写入联系人数据..."

NewList contact.s()

; 联系人1
ClearList(contact())
AddElement(contact()): contact() = "张三"
AddElement(contact()): contact() = "男"
AddElement(contact()): contact() = "13800138001"
AddElement(contact()): contact() = "zhangsan@company.com"
AddElement(contact()): contact() = "技术部"
AddElement(contact()): contact() = "高级工程师"
PbXls_AppendRow(wsId, contact())

; 联系人2
ClearList(contact())
AddElement(contact()): contact() = "李四"
AddElement(contact()): contact() = "女"
AddElement(contact()): contact() = "13800138002"
AddElement(contact()): contact() = "lisi@company.com"
AddElement(contact()): contact() = "市场部"
AddElement(contact()): contact() = "市场经理"
PbXls_AppendRow(wsId, contact())

; 联系人3
ClearList(contact())
AddElement(contact()): contact() = "王五"
AddElement(contact()): contact() = "男"
AddElement(contact()): contact() = "13800138003"
AddElement(contact()): contact() = "wangwu@company.com"
AddElement(contact()): contact() = "技术部"
AddElement(contact()): contact() = "项目经理"
PbXls_AppendRow(wsId, contact())

; 联系人4
ClearList(contact())
AddElement(contact()): contact() = "赵六"
AddElement(contact()): contact() = "女"
AddElement(contact()): contact() = "13800138004"
AddElement(contact()): contact() = "zhaoliu@company.com"
AddElement(contact()): contact() = "人事部"
AddElement(contact()): contact() = "人事主管"
PbXls_AppendRow(wsId, contact())

; 联系人5
ClearList(contact())
AddElement(contact()): contact() = "钱七"
AddElement(contact()): contact() = "男"
AddElement(contact()): contact() = "13800138005"
AddElement(contact()): contact() = "qianqi@company.com"
AddElement(contact()): contact() = "财务部"
AddElement(contact()): contact() = "财务总监"
PbXls_AppendRow(wsId, contact())

; 联系人6
ClearList(contact())
AddElement(contact()): contact() = "孙八"
AddElement(contact()): contact() = "女"
AddElement(contact()): contact() = "13800138006"
AddElement(contact()): contact() = "sunba@company.com"
AddElement(contact()): contact() = "技术部"
AddElement(contact()): contact() = "测试工程师"
PbXls_AppendRow(wsId, contact())

; 联系人7
ClearList(contact())
AddElement(contact()): contact() = "周九"
AddElement(contact()): contact() = "男"
AddElement(contact()): contact() = "13800138007"
AddElement(contact()): contact() = "zhoujiu@company.com"
AddElement(contact()): contact() = "市场部"
AddElement(contact()): contact() = "销售代表"
PbXls_AppendRow(wsId, contact())

; 联系人8
ClearList(contact())
AddElement(contact()): contact() = "吴十"
AddElement(contact()): contact() = "男"
AddElement(contact()): contact() = "13800138008"
AddElement(contact()): contact() = "wushi@company.com"
AddElement(contact()): contact() = "技术部"
AddElement(contact()): contact() = "架构师"
PbXls_AppendRow(wsId, contact())

Debug "  共写入8条联系人数据"

; 5. 添加统计信息
Debug "5. 添加统计信息..."
Define lastRow.i = 10
PbXls_SetCell(wsId, lastRow, 1, "统计")
PbXls_MergeCells(wsId, "A10:F10")

Define countRow.i = 11
PbXls_SetCell(wsId, countRow, 1, "技术部人数")
PbXls_SetCellFormulaWS(wsId, countRow, 2, "=COUNTIF(E3:E9," + Chr(34) + "技术部" + Chr(34) + ")")

Define nextRow.i = 12
PbXls_SetCell(wsId, nextRow, 1, "市场部人数")
PbXls_SetCellFormulaWS(wsId, nextRow, 2, "=COUNTIF(E3:E9," + Chr(34) + "市场部" + Chr(34) + ")")

Define nextRow2.i = 13
PbXls_SetCell(wsId, nextRow2, 1, "总人数")
PbXls_SetCellFormulaWS(wsId, nextRow2, 2, "=COUNTA(A3:A9)")

; 6. 创建第二个的工作表 - 按部门分组
Debug "6. 创建按部门分组的工作表..."
ws2Idx.i = PbXls_CreateSheet(wbId, "部门分组")
ws2Id.i = PbXls_GetSheetByIndex(wbId, ws2Idx)
If ws2Id <> 0
  PbXls_SetCell(ws2Id, 1, 1, "按部门分组统计")
  PbXls_MergeCells(wsId, "A1:C1")
  
  PbXls_SetCell(ws2Id, 2, 1, "部门")
  PbXls_SetCell(ws2Id, 2, 2, "人数")
  PbXls_SetCell(ws2Id, 2, 3, "占比")
  
  PbXls_SetCell(ws2Id, 3, 1, "技术部")
  PbXls_SetCell(ws2Id, 3, 2, "4")
  PbXls_SetCellFormulaWS(ws2Id, 3, 3, "=B3/B7")
  
  PbXls_SetCell(ws2Id, 4, 1, "市场部")
  PbXls_SetCell(ws2Id, 4, 2, "2")
  PbXls_SetCellFormulaWS(ws2Id, 4, 3, "=B4/B7")
  
  PbXls_SetCell(ws2Id, 5, 1, "人事部")
  PbXls_SetCell(ws2Id, 5, 2, "1")
  PbXls_SetCellFormulaWS(ws2Id, 5, 3, "=B5/B7")
  
  PbXls_SetCell(ws2Id, 6, 1, "财务部")
  PbXls_SetCell(ws2Id, 6, 2, "1")
  PbXls_SetCellFormulaWS(ws2Id, 6, 3, "=B6/B7")
  
  PbXls_SetCell(ws2Id, 7, 1, "合计")
  PbXls_SetCellFormulaWS(ws2Id, 7, 2, "=SUM(B3:B6)")
  
  PbXls_SetColumnWidth(ws2Id, 1, 15)
  PbXls_SetColumnWidth(ws2Id, 2, 10)
  PbXls_SetColumnWidth(ws2Id, 3, 10)
EndIf

; 7. 冻结首行（第一张工作表）
PbXls_SetFreezePanes(wsId, "A3")

; 8. 保存工作簿
Debug "7. 保存工作簿..."
outputFile.s = "demo_contacts.xlsx"
If PbXls_SaveWorkbook(wbId, outputFile)
  Debug "  保存成功: " + outputFile
Else
  Debug "  保存失败!"
EndIf

; 9. 显示工作簿信息
Debug ""
Debug "=== 工作簿信息 ==="
Debug "  工作表数量: " + Str(PbXls_GetSheetCount(wbId))
For i.i = 0 To PbXls_GetSheetCount(wbId) - 1
  Debug "  工作表" + Str(i) + ": " + PbXls_GetSheetName(wbId, i)
Next

; 10. 清理
PbXls_CloseWorkbook(wbId)
PbXls_Free()

Debug ""
Debug "=== 示例完成 ==="
Debug "联系人列表已生成，请用Excel打开 " + outputFile + " 查看结果"
