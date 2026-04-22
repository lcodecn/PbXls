; ***************************************************************************************
; PbXls 示例代码 - 读写循环测试 (RoundTrip Test)
; 说明: 创建Excel文件，然后读取回来验证数据一致性
; 版本: 2.4
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Procedure PbXls_RunRoundTripTest()
  Define testFile.s = "roundtrip_test.xlsx"
  Define testResult.b = #True
  
  Debug "========== PbXls RoundTrip 测试开始 =========="
  
  ; ====== 第一部分：创建Excel文件 ======
  Debug "--- 第1阶段: 创建Excel文件 ---"
  
  Define wb1.i = PbXls_CreateWorkbook()
  Define ws1.i = PbXls_GetSheetByIndex(wb1, 0)
  
  ; 写入各种类型数据
  PbXls_SetCell(ws1, 1, 1, "姓名")
  PbXls_SetCell(ws1, 1, 2, "年龄")
  PbXls_SetCell(ws1, 1, 3, "公式")
  
  PbXls_SetCell(ws1, 2, 1, "张三")
  PbXls_SetCell(ws1, 2, 2, "25")
  PbXls_SetCellFormulaWS(ws1, 2, 3, "=B2*2")
  
  PbXls_SetCell(ws1, 3, 1, "李四")
  PbXls_SetCell(ws1, 3, 2, "30")
  PbXls_SetCellFormulaWS(ws1, 3, 3, "=B3*2")
  
  ; 设置列宽
  PbXls_SetColumnWidth(ws1, 1, 15.0)
  PbXls_SetColumnWidth(ws1, 2, 10.0)
  
  ; 合并单元格
  PbXls_MergeCells(ws1, "A1:C1")
  
  ; 设置行高
  PbXls_SetRowHeight(ws1, 1, 25.0)
  
  ; 设置页边距
  PbXls_SetPageMargins(ws1, 0.8, 0.8, 0.9, 0.9, 0.4, 0.4)
  
  ; 设置页眉页脚
  PbXls_SetHeaderFooter(ws1, "&C测试报表", "&P")
  
  ; 设置打印选项
  PbXls_SetPrintOptions(ws1, #True, #True, #True, #True)
  
  ; 保存文件
  If PbXls_SaveWorkbook(wb1, testFile)
    Debug "  [OK] 文件创建成功: " + testFile
  Else
    Debug "  [FAIL] 文件创建失败"
    testResult = #False
  EndIf
  
  ; 关闭第一个工作簿
  PbXls_CloseWorkbook(wb1)
  
  ; ====== 第二部分：读取Excel文件 ======
  If testResult
    Debug "--- 第2阶段: 读取Excel文件 ---"
    
    Define wb2.i = PbXls_LoadWorkbook(testFile)
    If wb2 >= 0
      Debug "  [OK] 工作簿读取成功, ID=" + Str(wb2)
    Else
      Debug "  [FAIL] 工作簿读取失败"
      testResult = #False
    EndIf
    
    If testResult
      ; 验证工作表数量
      Define sheetCount.i = PbXls_GetSheetCount(wb2)
      If sheetCount >= 1
        Debug "  [OK] 工作表数量: " + Str(sheetCount)
      Else
        Debug "  [FAIL] 工作表数量错误: " + Str(sheetCount)
        testResult = #False
      EndIf
      
      ; 获取工作表
      Define ws2.i = PbXls_GetSheetByIndex(wb2, 0)
      If ws2 <> 0
        Debug "  [OK] 工作表获取成功"
        
        ; 验证单元格数据
        Define a1.s = PbXls_GetCellString(ws2, 1, 1)
        Debug "  A1='" + a1 + "'"
        If a1 = "姓名"
          Debug "  [OK] A1 数据正确"
        Else
          Debug "  [FAIL] A1 数据错误，期望='姓名' 实际='" + a1 + "'"
          testResult = #False
        EndIf
        
        Define b2.s = PbXls_GetCellString(ws2, 2, 2)
        Debug "  B2='" + b2 + "'"
        If b2 = "25"
          Debug "  [OK] B2 数据正确"
        Else
          Debug "  [FAIL] B2 数据错误，期望='25' 实际='" + b2 + "'"
          testResult = #False
        EndIf
        
        Define a2.s = PbXls_GetCellString(ws2, 2, 1)
        Debug "  A2='" + a2 + "'"
        If a2 = "张三"
          Debug "  [OK] A2 数据正确"
        Else
          Debug "  [FAIL] A2 数据错误，期望='张三' 实际='" + a2 + "'"
          testResult = #False
        EndIf
        
        ; 验证合并单元格 - 使用动态wsId
        Define ws2Id.s = "0_"  ; 默认第一个Sheet的ID为0
        Define mcKey.s = ws2Id + "A1:C1"
        Debug "  查找合并单元格 key='" + mcKey + "'"
        If FindMapElement(PbXls_MergedCells(), mcKey)
          Debug "  [OK] 合并单元格读取正确: A1:C1"
        Else
          Debug "  [FAIL] 合并单元格读取失败 (key='" + mcKey + "')"
          Debug "  MergedCells Map 内容:"
          ForEach PbXls_MergedCells()
            Debug "    - '" + MapKey(PbXls_MergedCells()) + "' = '" + PbXls_MergedCells() + "'"
          Next
          testResult = #False
        EndIf
        
        ; 验证列宽
        Define cwKey.s = ws2Id + "1"
        Debug "  查找列宽 key='" + cwKey + "'"
        If FindMapElement(PbXls_ColumnWidths(), cwKey)
          Define cw.f = PbXls_ColumnWidths(cwKey)
          If cw >= 14.9 And cw <= 15.1
            Debug "  [OK] 列宽读取正确: " + StrF(cw, 1)
          Else
            Debug "  [FAIL] 列宽读取错误: " + StrF(cw, 1) + " (期望15.0)"
            testResult = #False
          EndIf
        Else
          Debug "  [FAIL] 列宽数据未找到 (key='" + cwKey + "')"
          Debug "  ColumnWidths Map 内容:"
          ForEach PbXls_ColumnWidths()
            Debug "    - '" + MapKey(PbXls_ColumnWidths()) + "' = " + StrF(PbXls_ColumnWidths(), 1)
          Next
          testResult = #False
        EndIf
        
        ; 验证行高
        Define rhKey.s = ws2Id + "1"
        Debug "  查找行高 key='" + rhKey + "'"
        If FindMapElement(PbXls_RowHeights(), rhKey)
          Define rh.f = PbXls_RowHeights(rhKey)
          If rh >= 24.9 And rh <= 25.1
            Debug "  [OK] 行高读取正确: " + StrF(rh, 1)
          Else
            Debug "  [FAIL] 行高读取错误: " + StrF(rh, 1) + " (期望25.0)"
            testResult = #False
          EndIf
        Else
          Debug "  [FAIL] 行高数据未找到 (key='" + rhKey + "')"
          Debug "  RowHeights Map 内容:"
          ForEach PbXls_RowHeights()
            Debug "    - '" + MapKey(PbXls_RowHeights()) + "' = " + StrF(PbXls_RowHeights(), 1)
          Next
          testResult = #False
        EndIf
      Else
        Debug "  [FAIL] 工作表获取失败"
        testResult = #False
      EndIf
      
      ; 关闭第二个工作簿
      PbXls_CloseWorkbook(wb2)
    EndIf
  EndIf
  
  ; 输出测试结果
  Debug ""
  Debug "========== RoundTrip 测试结果 =========="
  If testResult
    Debug "状态: 全部通过"
  Else
    Debug "状态: 部分失败"
  EndIf
  Debug "========================================="
  
  ; 清理资源
  PbXls_Free()
  
  ProcedureReturn testResult
EndProcedure

Debug ""
Debug "PbXls Library v2.4 - RoundTrip 测试模式"
Debug "================================"
Define passed.b = PbXls_RunRoundTripTest()
Debug ""

; IDE Options = PureBasic 6.40 (Windows - x86)
; CursorPosition = 5
; Folding = -
; EnableThread
; EnableXP
; DPIAware
; CompileSourceDirectory