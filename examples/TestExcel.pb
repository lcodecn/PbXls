; ***************************************************************************************
; PbXls 示例代码 - 功能测试
; 说明: 对PbXls库的各项功能进行全面测试
; 版本: 2.3
; 作者: lcode.cn
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

; 测试函数：验证基本功能
Procedure.b PbXls_RunTests()
  Define testResult.b = #True
  Define testCount.i = 0
  Define passCount.i = 0
  
  Debug "========== PbXls Library 测试开始 =========="
  
  ; 测试1: 创建工作簿
  testCount + 1
  Debug "[测试 1] 创建工作簿..."
  Define wbId.i = PbXls_CreateWorkbook()
  If wbId >= 0
    Debug "  通过: 工作簿创建成功, ID=" + Str(wbId)
    passCount + 1
  Else
    Debug "  失败: 工作簿创建失败"
    testResult = #False
  EndIf
  
  ; 测试2: 获取当前工作表
  testCount + 1
  Debug "[测试 2] 获取当前工作表..."
  Define wsId.i = PbXls_GetSheetByIndex(wbId, 0)
  If wsId <> 0
    Debug "  通过: 工作表获取成功, 指针=" + Str(wsId)
    passCount + 1
  Else
    Debug "  失败: 工作表获取失败"
    testResult = #False
  EndIf
  
  ; 测试3: 写入字符串单元格
  testCount + 1
  Debug "[测试 3] 写入字符串单元格..."
  If PbXls_SetCell(wsId, 1, 1, "Hello PbXls!")
    Debug "  通过: 单元格写入成功 A1='Hello PbXls!'"
    passCount + 1
  Else
    Debug "  失败: 单元格写入失败"
    testResult = #False
  EndIf
  
  ; 测试4: 写入数值
  testCount + 1
  Debug "[测试 4] 写入数值..."
  If PbXls_SetCell(wsId, 1, 2, "123.45")
    Debug "  通过: 数值写入成功 B1='123.45'"
    passCount + 1
  Else
    Debug "  失败: 数值写入失败"
    testResult = #False
  EndIf
  
  ; 测试5: 坐标转换测试
  testCount + 1
  Debug "[测试 5] 坐标转换..."
  Define colStr.s = PbXls_GetColumnLetter(1)
  Define colNum.i = PbXls_ColumnIndexFromString("A")
  If colStr = "A" And colNum = 1
    Debug "  通过: 坐标转换正确 (1->'" + colStr + "', 'A'->" + Str(colNum) + ")"
    passCount + 1
  Else
    Debug "  失败: 坐标转换错误 (1->'" + colStr + "', 'A'->" + Str(colNum) + ")"
    testResult = #False
  EndIf
  
  ; 测试6: 写入多行数据
  testCount + 1
  Debug "[测试 6] 写入多行数据..."
  Define i.i
  For i = 2 To 10
    PbXls_SetCell(wsId, i, 1, "行" + Str(i))
    PbXls_SetCell(wsId, i, 2, Str(i * 10))
    PbXls_SetCell(wsId, i, 3, Str(i * 100))
  Next
  Debug "  通过: 写入10行数据成功"
  passCount + 1
  
  ; 测试7: 设置列宽
  testCount + 1
  Debug "[测试 7] 设置列宽..."
  If PbXls_SetColumnWidth(wsId, 1, 20.0)
    Debug "  通过: 列宽设置成功"
    passCount + 1
  Else
    Debug "  失败: 列宽设置失败"
    testResult = #False
  EndIf
  
  ; 测试8: 合并单元格
  testCount + 1
  Debug "[测试 8] 合并单元格..."
  If PbXls_MergeCells(wsId, "A1:C1")
    Debug "  通过: 单元格合并成功 A1:C1"
    passCount + 1
  Else
    Debug "  失败: 单元格合并失败"
    testResult = #False
  EndIf
  
  ; 测试9: 获取工作表数量
  testCount + 1
  Debug "[测试 9] 获取工作表数量..."
  Define sheetCount.i = PbXls_GetSheetCount(wbId)
  If sheetCount >= 1
    Debug "  通过: 工作表数量=" + Str(sheetCount)
    passCount + 1
  Else
    Debug "  失败: 工作表数量错误"
    testResult = #False
  EndIf
  
  ; 测试10: 保存工作簿
  testCount + 1
  Debug "[测试 10] 保存工作簿..."
  Define testFile.s = "test_output.xlsx"
  If PbXls_SaveWorkbook(wbId, testFile)
    Debug "  通过: 工作簿保存成功 '" + testFile + "'"
    passCount + 1
  Else
    Debug "  失败: 工作簿保存失败"
    testResult = #False
  EndIf
  
  ; 输出测试结果
  Debug ""
  Debug "========== 测试结果 =========="
  Debug "总测试数: " + Str(testCount)
  Debug "通过: " + Str(passCount)
  Debug "失败: " + Str(testCount - passCount)
  If testResult
    Debug "状态: 全部通过"
  Else
    Debug "状态: 部分失败"
  EndIf
  Debug "================================"
  
  PbXls_CloseWorkbook(wbId)
  ProcedureReturn testResult
EndProcedure

; 执行测试（仅在直接运行此文件时执行）
Debug ""
Debug "PbXls Library v2.3 - 测试模式"
Debug "================================"
Define testPassed.b = PbXls_RunTests()
Debug ""

; 清理资源
PbXls_Free()
