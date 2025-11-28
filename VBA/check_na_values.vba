Sub CheckNAValues()
    Dim wsCY As Worksheet
    Dim wsSAP As Worksheet
    Dim wsResult As Worksheet
    Dim lastRowCY As Long
    Dim lastRowSAP As Long
    Dim i As Long
    Dim naCount As Long
    Dim resultRow As Long
    Dim cell As Range
    
    ' 设置源工作表
    On Error Resume Next
    Set wsCY = ThisWorkbook.Worksheets("FC_current")
    Set wsSAP = ThisWorkbook.Worksheets("SAP report")
    On Error GoTo 0
    
    ' 检查工作表是否存在
    If wsCY Is Nothing Then
        MsgBox "找不到工作表 'CY_current'，请检查工作表名称是否正确。", vbExclamation, "错误"
        Exit Sub
    End If
    
    If wsSAP Is Nothing Then
        MsgBox "找不到工作表 'SAP report'，请检查工作表名称是否正确。", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 获取CY_current D列最后一行
    lastRowCY = wsCY.Cells(wsCY.Rows.Count, "D").End(xlUp).Row
    
    ' 获取SAP report AT列最后一行
    lastRowSAP = wsSAP.Cells(wsSAP.Rows.Count, "AT").End(xlUp).Row
    
    ' 计数N/A值
    naCount = 0
    
    ' 检查CY_current的D列
    For i = 1 To lastRowCY
        Set cell = wsCY.Cells(i, "D")
        If CheckIfNA(cell) Then
            naCount = naCount + 1
        End If
    Next i
    
    ' 检查SAP report的AT列
    For i = 1 To lastRowSAP
        Set cell = wsSAP.Cells(i, "AT")
        If CheckIfNA(cell) Then
            naCount = naCount + 1
        End If
    Next i
    
    ' 如果没有N/A值
    If naCount = 0 Then
        MsgBox "Forecast和SAP report资产卡片一致", vbInformation, "检查完成"
        Exit Sub
    End If
    
    ' 如果有N/A值，创建新工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("NA_Check_Result").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "NA_Check_Result"
    
    ' 设置结果表头
    wsResult.Cells(1, 1).Value = "来源工作表"
    wsResult.Cells(1, 2).Value = "行号"
    wsResult.Cells(1, 3).Value = "问题列"
    wsResult.Cells(1, 4).Value = "N/A值"
    wsResult.Cells(1, 5).Value = "相关数据1"
    wsResult.Cells(1, 6).Value = "相关数据2"
    wsResult.Cells(1, 7).Value = "相关数据3"
    
    ' 格式化表头
    With wsResult.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 填充CY_current的N/A数据
    resultRow = 2
    For i = 1 To lastRowCY
        Set cell = wsCY.Cells(i, "D")
        If CheckIfNA(cell) Then
            wsResult.Cells(resultRow, 1).Value = "CY_current"
            wsResult.Cells(resultRow, 2).Value = i
            wsResult.Cells(resultRow, 3).Value = "D"
            wsResult.Cells(resultRow, 4).Value = cell.Value
            wsResult.Cells(resultRow, 5).Value = wsCY.Cells(i, 1).Value ' A列
            wsResult.Cells(resultRow, 6).Value = wsCY.Cells(i, 2).Value ' B列
            wsResult.Cells(resultRow, 7).Value = wsCY.Cells(i, 3).Value ' C列
            
            ' 高亮N/A值
            wsResult.Cells(resultRow, 4).Interior.Color = RGB(255, 200, 200)
            
            resultRow = resultRow + 1
        End If
    Next i
    
    ' 填充SAP report的N/A数据
    For i = 1 To lastRowSAP
        Set cell = wsSAP.Cells(i, "AT")
        If CheckIfNA(cell) Then
            wsResult.Cells(resultRow, 1).Value = "SAP report"
            wsResult.Cells(resultRow, 2).Value = i
            wsResult.Cells(resultRow, 3).Value = "AT"
            wsResult.Cells(resultRow, 4).Value = cell.Value
            wsResult.Cells(resultRow, 5).Value = wsSAP.Cells(i, 1).Value ' A列
            wsResult.Cells(resultRow, 6).Value = wsSAP.Cells(i, 2).Value ' B列
            wsResult.Cells(resultRow, 7).Value = wsSAP.Cells(i, 3).Value ' C列
            
            ' 高亮N/A值
            wsResult.Cells(resultRow, 4).Interior.Color = RGB(255, 200, 200)
            
            resultRow = resultRow + 1
        End If
    Next i
    
    ' 自动调整列宽
    wsResult.Columns("A:G").AutoFit
    
    ' 显示结果消息
    MsgBox "发现 " & naCount & " 个N/A值！" & vbCrLf & _
           "详细信息已列在新工作表 'NA_Check_Result' 中。", vbExclamation, "检查完成"
    
    ' 激活结果工作表
    wsResult.Activate
    
End Sub

' 辅助函数：检查单元格是否为N/A
Function CheckIfNA(cell As Range) As Boolean
    CheckIfNA = False
    
    ' 检查单元格是否包含N/A（包括#N/A错误、文本"N/A"等）
    If IsError(cell.Value) Then
        If CVErr(cell.Value) = CVErr(xlErrNA) Then
            CheckIfNA = True
        End If
    ElseIf Not IsEmpty(cell.Value) Then
        If InStr(1, CStr(cell.Value), "N/A", vbTextCompare) > 0 Or _
           InStr(1, CStr(cell.Value), "#N/A", vbTextCompare) > 0 Then
            CheckIfNA = True
        End If
    End If
End Function
