; ***************************************************************************************
; PbXls 示例代码 - 批量数据导出
; 说明: 演示如何使用PbXls库批量导出大量数据到Excel
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 批量数据导出示例 ==="
Debug ""

; 1. 创建新工作簿
Debug "1. 创建工作簿..."
wbId.i = PbXls_CreateWorkbook()
If wbId = -1
  Debug "  错误: 无法创建工作簿"
  End
EndIf

wsId.i = PbXls_ActiveSheet(wbId)

; 2. 写入表头
Debug "2. 写入表头..."
PbXls_SetCell(wsId, 1, 1, "序号")
PbXls_SetCell(wsId, 1, 2, "产品名称")
PbXls_SetCell(wsId, 1, 3, "单价")
PbXls_SetCell(wsId, 1, 4, "数量")
PbXls_SetCell(wsId, 1, 5, "金额")

; 设置列宽
PbXls_SetColumnWidth(wsId, 1, 8)
PbXls_SetColumnWidth(wsId, 2, 20)
PbXls_SetColumnWidth(wsId, 3, 12)
PbXls_SetColumnWidth(wsId, 4, 10)
PbXls_SetColumnWidth(wsId, 5, 12)

; 3. 批量写入数据
Debug "3. 批量写入100行产品数据..."
Define i.i, row.i = 2
Define productName.s, price.s, quantity.s, amount.s

; 定义产品名称数组（用于模拟数据）
Define NewList names.s()
AddElement(names()): names() = "笔记本电脑"
AddElement(names()): names() = "无线鼠标"
AddElement(names()): names() = "机械键盘"
AddElement(names()): names() = "显示器"
AddElement(names()): names() = "U盘"
AddElement(names()): names() = "移动硬盘"
AddElement(names()): names() = "网线"
AddElement(names()): names() = "路由器"
AddElement(names()): names() = "交换机"
AddElement(names()): names() = "电源插座"

Define nameCount.i = 10

For i = 1 To 100
  row = i + 1
  
  ; 生成序号
  PbXls_SetCell(wsId, row, 1, Str(i))
  
  ; 循环使用产品名称
  Define nameIdx.i = (i - 1) % nameCount
  ResetList(names())
  NextElement(names())
  Define j.i
  For j = 1 To nameIdx
    NextElement(names())
  Next
  PbXls_SetCell(wsId, row, 2, names())
  
  ; 生成随机价格 (10-999)
  price = Str(10 + (i * 13) % 990)
  PbXls_SetCell(wsId, row, 3, price)
  
  ; 生成随机数量 (1-100)
  quantity = Str(1 + (i * 7) % 100)
  PbXls_SetCell(wsId, row, 4, quantity)
  
  ; 计算金额 = 单价 * 数量 (使用Excel公式)
  Define colLetter.s = PbXls_GetColumnLetter(row)
  amountFormula.s = "=C" + Str(row) + "*D" + Str(row)
  PbXls_SetCellFormulaWS(wsId, row, 5, amountFormula)
  
  ; 每20行输出一次进度
  If i % 20 = 0
    Debug "  已写入 " + Str(i) + " 行数据..."
  EndIf
Next

Debug "  数据写入完成"

; 4. 添加汇总行
Debug "4. 添加汇总行..."
Define lastRow.i = 102
PbXls_SetCell(wsId, lastRow, 2, "总计")
PbXls_SetCellFormulaWS(wsId, lastRow, 3, "=SUM(C2:C101)")
PbXls_SetCellFormulaWS(wsId, lastRow, 4, "=SUM(D2:D101)")
PbXls_SetCellFormulaWS(wsId, lastRow, 5, "=SUM(E2:E101)")

; 5. 保存工作簿
Debug "5. 保存工作簿..."
outputFile.s = "demo_batch_export.xlsx"
If PbXls_SaveWorkbook(wbId, outputFile)
  Debug "  保存成功: " + outputFile
  
  ; 获取文件大小
  Define fileSize.i = FileSize(outputFile)
  If fileSize > 0
    Debug "  文件大小: " + Str(fileSize) + " 字节"
  EndIf
Else
  Debug "  保存失败!"
EndIf

; 6. 清理
PbXls_CloseWorkbook(wbId)
PbXls_Free()

Debug ""
Debug "=== 示例完成 ==="
Debug "已导出100行产品数据到 " + outputFile
