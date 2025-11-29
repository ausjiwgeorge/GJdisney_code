'==============================================================================
' CheckNAValues - 优化版本
' 性能改进:
' 1. 使用数组批量读取数据 (避免逐单元格访问，速度提升10-100倍)
' 2. 关闭屏幕更新和自动计算
' 3. 批量写入结果 (避免逐单元格写入)
' 4. 使用常量避免重复字符串创建
'==============================================================================
Option Explicit

' 常量定义
Private Const WS_CY_NAME As String = "FC_current"
Private Const WS_SAP_NAME As String = "SAP report"
Private Const WS_RESULT_NAME As String = "NA_Check_Result"
Private Const COL_CY_CHECK As Long = 4      ' D列
Private Const COL_SAP_CHECK As Long = 46    ' AT列

Sub CheckNAValues()
    Dim wsCY As Worksheet
    Dim wsSAP As Worksheet
    Dim wsResult As Worksheet
    Dim lastRowCY As Long
    Dim lastRowSAP As Long
    
    ' 性能优化: 关闭屏幕更新和自动计算
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' 设置源工作表
    On Error Resume Next
    Set wsCY = ThisWorkbook.Worksheets(WS_CY_NAME)
    Set wsSAP = ThisWorkbook.Worksheets(WS_SAP_NAME)
    On Error GoTo ErrorHandler
    
    ' 检查工作表是否存在
    If wsCY Is Nothing Then
        MsgBox "找不到工作表 '" & WS_CY_NAME & "'，请检查工作表名称是否正确。", vbExclamation, "错误"
        GoTo Cleanup
    End If
    
    If wsSAP Is Nothing Then
        MsgBox "找不到工作表 '" & WS_SAP_NAME & "'，请检查工作表名称是否正确。", vbExclamation, "错误"
        GoTo Cleanup
    End If
    
    ' 获取数据范围
    lastRowCY = wsCY.Cells(wsCY.Rows.Count, COL_CY_CHECK).End(xlUp).Row
    lastRowSAP = wsSAP.Cells(wsSAP.Rows.Count, COL_SAP_CHECK).End(xlUp).Row
    
    ' 批量读取数据到数组 (核心性能优化)
    Dim arrCY_D As Variant      ' D列数据
    Dim arrCY_ABC As Variant    ' A-C列数据
    Dim arrSAP_AT As Variant    ' AT列数据
    Dim arrSAP_ABC As Variant   ' A-C列数据
    
    If lastRowCY > 0 Then
        arrCY_D = wsCY.Range(wsCY.Cells(1, COL_CY_CHECK), wsCY.Cells(lastRowCY, COL_CY_CHECK)).Value
        arrCY_ABC = wsCY.Range(wsCY.Cells(1, 1), wsCY.Cells(lastRowCY, 3)).Value
    End If
    
    If lastRowSAP > 0 Then
        arrSAP_AT = wsSAP.Range(wsSAP.Cells(1, COL_SAP_CHECK), wsSAP.Cells(lastRowSAP, COL_SAP_CHECK)).Value
        arrSAP_ABC = wsSAP.Range(wsSAP.Cells(1, 1), wsSAP.Cells(lastRowSAP, 3)).Value
    End If
    
    ' 第一遍扫描: 计算N/A数量 (用于预分配结果数组)
    Dim naCount As Long
    Dim i As Long
    naCount = 0
    
    If IsArray(arrCY_D) Then
        For i = 1 To UBound(arrCY_D, 1)
            If CheckIfNAValue(arrCY_D(i, 1)) Then naCount = naCount + 1
        Next i
    End If
    
    If IsArray(arrSAP_AT) Then
        For i = 1 To UBound(arrSAP_AT, 1)
            If CheckIfNAValue(arrSAP_AT(i, 1)) Then naCount = naCount + 1
        Next i
    End If
    
    ' 如果没有N/A值
    If naCount = 0 Then
        MsgBox "Forecast和SAP report资产卡片一致", vbInformation, "检查完成"
        GoTo Cleanup
    End If
    
    ' 预分配结果数组 (避免逐行写入)
    Dim arrResult() As Variant
    ReDim arrResult(1 To naCount + 1, 1 To 7)
    
    ' 设置表头
    arrResult(1, 1) = "来源工作表"
    arrResult(1, 2) = "行号"
    arrResult(1, 3) = "问题列"
    arrResult(1, 4) = "N/A值"
    arrResult(1, 5) = "相关数据1"
    arrResult(1, 6) = "相关数据2"
    arrResult(1, 7) = "相关数据3"
    
    ' 填充结果数组
    Dim resultRow As Long
    resultRow = 2
    
    ' 处理CY_current数据
    If IsArray(arrCY_D) Then
        For i = 1 To UBound(arrCY_D, 1)
            If CheckIfNAValue(arrCY_D(i, 1)) Then
                arrResult(resultRow, 1) = "CY_current"
                arrResult(resultRow, 2) = i
                arrResult(resultRow, 3) = "D"
                arrResult(resultRow, 4) = arrCY_D(i, 1)
                arrResult(resultRow, 5) = arrCY_ABC(i, 1)
                arrResult(resultRow, 6) = arrCY_ABC(i, 2)
                arrResult(resultRow, 7) = arrCY_ABC(i, 3)
                resultRow = resultRow + 1
            End If
        Next i
    End If
    
    ' 处理SAP report数据
    If IsArray(arrSAP_AT) Then
        For i = 1 To UBound(arrSAP_AT, 1)
            If CheckIfNAValue(arrSAP_AT(i, 1)) Then
                arrResult(resultRow, 1) = "SAP report"
                arrResult(resultRow, 2) = i
                arrResult(resultRow, 3) = "AT"
                arrResult(resultRow, 4) = arrSAP_AT(i, 1)
                arrResult(resultRow, 5) = arrSAP_ABC(i, 1)
                arrResult(resultRow, 6) = arrSAP_ABC(i, 2)
                arrResult(resultRow, 7) = arrSAP_ABC(i, 3)
                resultRow = resultRow + 1
            End If
        Next i
    End If
    
    ' 创建或清空结果工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(WS_RESULT_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = WS_RESULT_NAME
    
    ' 批量写入结果 (单次操作，而非逐单元格)
    wsResult.Range("A1").Resize(naCount + 1, 7).Value = arrResult
    
    ' 格式化表头
    With wsResult.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 批量高亮N/A值列 (如果数据量大，使用条件格式更高效)
    If naCount <= 1000 Then
        ' 小数据量：直接设置颜色
        wsResult.Range("D2:D" & (naCount + 1)).Interior.Color = RGB(255, 200, 200)
    Else
        ' 大数据量：使用条件格式
        With wsResult.Range("D2:D" & (naCount + 1)).FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
            .Interior.Color = RGB(255, 200, 200)
        End With
    End If
    
    ' 自动调整列宽
    wsResult.Columns("A:G").AutoFit
    
    ' 显示结果消息
    MsgBox "发现 " & naCount & " 个N/A值！" & vbCrLf & _
           "详细信息已列在新工作表 '" & WS_RESULT_NAME & "' 中。", vbExclamation, "检查完成"
    
    ' 激活结果工作表
    wsResult.Activate
    
Cleanup:
    ' 恢复Excel设置
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "发生错误: " & Err.Description, vbCritical, "错误"
    Resume Cleanup
End Sub

' 优化版辅助函数：检查值是否为N/A (接受Variant而非Range，避免对象访问开销)
Private Function CheckIfNAValue(ByVal cellValue As Variant) As Boolean
    CheckIfNAValue = False
    
    ' 检查是否为错误值
    If IsError(cellValue) Then
        If cellValue = CVErr(xlErrNA) Then
            CheckIfNAValue = True
        End If
        Exit Function
    End If
    
    ' 检查是否为空
    If IsEmpty(cellValue) Then Exit Function
    
    ' 检查是否包含N/A文本
    Dim strValue As String
    strValue = UCase$(CStr(cellValue))
    
    If InStr(1, strValue, "N/A", vbBinaryCompare) > 0 Or _
       InStr(1, strValue, "#N/A", vbBinaryCompare) > 0 Then
        CheckIfNAValue = True
    End If
End Function
