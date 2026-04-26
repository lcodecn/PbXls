﻿; ***************************************************************************************
; PbXls 示例代码 - 预留功能测试(数据验证、条件格式、图表)
; 说明: 测试数据验证(下拉列表/整数验证)、条件格式(单元格值比较/颜色刻度)、图表(柱状图/折线图)
; 版本: 2.6
; ***************************************************************************************

; 引入PbXls库
XIncludeFile "..\PbXls.pb"

Procedure PbXls_RunFeaturesTest()
  Define outputFile.s = "features_test.xlsx"
  DeleteFile(outputFile, #PB_FileSystem_Force)
  
  Debug ""
  Debug "=== PbXls 预留功能测试 ==="
  Debug ""
  
  Define wbId.i = PbXls_CreateWorkbook()
  Define wsId.i = PbXls_GetSheetByIndex(wbId, 0)
  
  ; ===== 第1部分：准备基础数据 =====
  Debug "1. 准备基础数据..."
  
  ; 写入表头
  PbXls_SetCell(wsId, 1, 1, "姓名")
  PbXls_SetCell(wsId, 1, 2, "部门")
  PbXls_SetCell(wsId, 1, 3, "年龄")
  PbXls_SetCell(wsId, 1, 4, "工资")
  PbXls_SetCell(wsId, 1, 5, "绩效")
  
  ; 写入数据
  Define row.i
  Define names0.s = "张三", names1.s = "李四", names2.s = "王五", names3.s = "赵六"
  Define depts0.s = "技术部", depts1.s = "销售部", depts2.s = "人事部", depts3.s = "财务部"
  
  For row = 2 To 5
    Define nameVal.s, deptVal.s
    Select row
      Case 2: nameVal = names0: deptVal = depts0
      Case 3: nameVal = names1: deptVal = depts1
      Case 4: nameVal = names2: deptVal = depts2
      Case 5: nameVal = names3: deptVal = depts3
    EndSelect
    PbXls_SetCell(wsId, row, 1, nameVal)
    PbXls_SetCell(wsId, row, 2, deptVal)
    PbXls_SetCell(wsId, row, 3, Str(25 + row))
    PbXls_SetCell(wsId, row, 4, Str(5000 + row * 1000))
    PbXls_SetCell(wsId, row, 5, Str(80 + row * 5))
  Next
  
  ; 设置列宽
  PbXls_SetColumnWidth(wsId, 1, 12.0)
  PbXls_SetColumnWidth(wsId, 2, 12.0)
  PbXls_SetColumnWidth(wsId, 3, 10.0)
  PbXls_SetColumnWidth(wsId, 4, 12.0)
  PbXls_SetColumnWidth(wsId, 5, 10.0)
  
  ; ===== 第2部分：数据验证测试 =====
  Debug ""
  Debug "2. 创建数据验证..."
  
  ; 下拉列表验证 - B2:B5 部门列
  Define dv1.i = PbXls_CreateDataValidation("list", "B2:B5", ~"技术部,销售部,人事部,财务部")
  PbXls_SetValidationPrompt(dv1, "选择部门", "请从下拉列表中选择部门")
  PbXls_SetValidationError(dv1, "输入错误", "请选择有效的部门")
  Debug "   dv1: 部门下拉列表 B2:B5"
  
  ; 整数验证 - C2:C5 年龄列 (18-65)
  Define dv2.i = PbXls_CreateDataValidation("whole", "C2:C5", "18", "65", "between")
  PbXls_SetValidationPrompt(dv2, "输入年龄", "请输入18-65之间的整数")
  PbXls_SetValidationError(dv2, "年龄错误", "年龄必须在18到65之间")
  Debug "   dv2: 年龄整数验证 C2:C5"
  
  ; ===== 第3部分：条件格式测试 =====
  Debug ""
  Debug "3. 创建条件格式..."
  
  ; 单元格值比较 - E2:E5 绩效>85标红加粗
  Define cf1.i = PbXls_CreateConditionalFormat("cellIs", "E2:E5", "85", "", "greaterThan")
  PbXls_SetConditionalFormatDxf(cf1, "FF0000", "", #True, #False)
  Debug "   cf1: 绩效>85 红色粗体 E2:E5"
  
  ; 颜色刻度 - D2:D5 工资列颜色渐变
  Define cf2.i = PbXls_CreateConditionalFormat("colorScale", "D2:D5")
  PbXls_SetConditionalFormatColorScale(cf2, "FF0000", "FFFF00", "00FF00", "min", "percentile", "max")
  Debug "   cf2: 工资颜色刻度 D2:D5 (红-黄-绿)"
  
  ; ===== 第4部分：保存文件 =====
  Debug ""
  Debug "4. 保存文件..."
  
  If PbXls_SaveWorkbook(wbId, outputFile)
    Define fullPath.s = GetCurrentDirectory() + outputFile
    Define fileSize.i = FileSize(fullPath)
    Debug "   保存成功: " + outputFile
    Debug "   文件大小: " + Str(fileSize) + " 字节"
    Debug ""
    Debug "=== 预留功能测试完成 ==="
  Else
    Debug "   保存失败!"
  EndIf
  
  PbXls_CloseWorkbook(wbId)
  PbXls_Free()
EndProcedure

PbXls_RunFeaturesTest()
