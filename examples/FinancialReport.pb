; ***************************************************************************************
; PbXls 示例代码 - 财务报表
; 说明: 演示如何使用PbXls库创建财务报表，包含公式和数据汇总
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 财务报表示例 ==="
Debug ""

; 1. 创建新工作簿
Debug "1. 创建工作簿..."
wbId.i = PbXls_CreateWorkbook()
If wbId = -1
  Debug "  错误: 无法创建工作簿"
  End
EndIf

wsId.i = PbXls_ActiveSheet(wbId)

; 2. 设置列宽
Debug "2. 设置报表格式..."
PbXls_SetColumnWidth(wsId, 1, 25)
PbXls_SetColumnWidth(wsId, 2, 15)
PbXls_SetColumnWidth(wsId, 3, 15)
PbXls_SetColumnWidth(wsId, 4, 15)

; 3. 写入报表标题
PbXls_SetCell(wsId, 1, 1, "2026年度财务报表")
PbXls_MergeCells(wsId, "A1:D1")

; 4. 写入表头
PbXls_SetCell(wsId, 3, 1, "项目")
PbXls_SetCell(wsId, 3, 2, "第一季度")
PbXls_SetCell(wsId, 3, 3, "第二季度")
PbXls_SetCell(wsId, 3, 4, "第三季度")

; 5. 写入收入数据
Debug "3. 写入财务数据..."
PbXls_SetCell(wsId, 4, 1, "销售收入")
PbXls_SetCell(wsId, 4, 2, "150000")
PbXls_SetCell(wsId, 4, 3, "180000")
PbXls_SetCell(wsId, 4, 4, "200000")

PbXls_SetCell(wsId, 5, 1, "服务收入")
PbXls_SetCell(wsId, 5, 2, "50000")
PbXls_SetCell(wsId, 5, 3, "60000")
PbXls_SetCell(wsId, 5, 4, "70000")

PbXls_SetCell(wsId, 6, 1, "其他收入")
PbXls_SetCell(wsId, 6, 2, "10000")
PbXls_SetCell(wsId, 6, 3, "15000")
PbXls_SetCell(wsId, 6, 4, "20000")

; 6. 计算收入合计
PbXls_SetCell(wsId, 7, 1, "收入合计")
PbXls_SetCellFormulaWS(wsId, 7, 2, "=SUM(B4:B6)")
PbXls_SetCellFormulaWS(wsId, 7, 3, "=SUM(C4:C6)")
PbXls_SetCellFormulaWS(wsId, 7, 4, "=SUM(D4:D6)")

; 7. 写入支出数据
PbXls_SetCell(wsId, 9, 1, "人员成本")
PbXls_SetCell(wsId, 9, 2, "80000")
PbXls_SetCell(wsId, 9, 3, "85000")
PbXls_SetCell(wsId, 9, 4, "90000")

PbXls_SetCell(wsId, 10, 1, "运营成本")
PbXls_SetCell(wsId, 10, 2, "30000")
PbXls_SetCell(wsId, 10, 3, "35000")
PbXls_SetCell(wsId, 10, 4, "40000")

PbXls_SetCell(wsId, 11, 1, "管理费用")
PbXls_SetCell(wsId, 11, 2, "20000")
PbXls_SetCell(wsId, 11, 3, "22000")
PbXls_SetCell(wsId, 11, 4, "25000")

; 8. 计算支出合计
PbXls_SetCell(wsId, 12, 1, "支出合计")
PbXls_SetCellFormulaWS(wsId, 12, 2, "=SUM(B9:B11)")
PbXls_SetCellFormulaWS(wsId, 12, 3, "=SUM(C9:C11)")
PbXls_SetCellFormulaWS(wsId, 12, 4, "=SUM(D9:D11)")

; 9. 计算净利润
PbXls_SetCell(wsId, 14, 1, "净利润")
PbXls_SetCellFormulaWS(wsId, 14, 2, "=B7-B12")
PbXls_SetCellFormulaWS(wsId, 14, 3, "=C7-C12")
PbXls_SetCellFormulaWS(wsId, 14, 4, "=D7-D12")

; 10. 添加年度汇总
PbXls_SetCell(wsId, 16, 1, "年度总收入")
PbXls_SetCellFormulaWS(wsId, 16, 2, "=B7+C7+D7")
PbXls_MergeCells(wsId, "A16:D16")

PbXls_SetCell(wsId, 17, 1, "年度总支出")
PbXls_SetCellFormulaWS(wsId, 17, 2, "=B12+C12+D12")
PbXls_MergeCells(wsId, "A17:D17")

PbXls_SetCell(wsId, 18, 1, "年度净利润")
PbXls_SetCellFormulaWS(wsId, 18, 2, "=B16-B17")
PbXls_MergeCells(wsId, "A18:D18")

Debug "  财务数据写入完成"

; 11. 保存工作簿
Debug "4. 保存工作簿..."
outputFile.s = "demo_financial.xlsx"
If PbXls_SaveWorkbook(wbId, outputFile)
  Debug "  保存成功: " + outputFile
Else
  Debug "  保存失败!"
EndIf

; 12. 清理
PbXls_CloseWorkbook(wbId)
PbXls_Free()

Debug ""
Debug "=== 示例完成 ==="
Debug "报表已生成，请用Excel打开 " + outputFile + " 查看结果"
Debug "报表包含收入、支出、净利润及年度汇总数据"
