﻿; PbXls 保存诊断测试
XIncludeFile "..\PbXls.pb"

Debug "=== PbXls 保存诊断 ==="

Define wbId.i = PbXls_CreateWorkbook()
Define wsId.i = PbXls_GetSheetByIndex(wbId, 0)

Debug "工作簿ID: " + Str(wbId)
Debug "工作表ID: " + Str(wsId)

PbXls_SetCell(wsId, 1, 1, "测试数据")
Debug "单元格(1,1)写入: 测试数据"
Debug "单元格(1,1)读取: " + PbXls_GetCellString(wsId, 1, 1)

Define outputFile.s = "save_test.xlsx"
Debug "保存路径: " + GetCurrentDirectory() + outputFile

Define result.i = PbXls_SaveWorkbook(wbId, outputFile)
Debug "保存结果: " + Str(result)

Define fileSize.i = FileSize(GetCurrentDirectory() + outputFile)
Debug "文件大小: " + Str(fileSize) + " 字节"

PbXls_CloseWorkbook(wbId)
PbXls_Free()
