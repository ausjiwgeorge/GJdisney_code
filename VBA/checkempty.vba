Function CheckEmptyColumns(rng As Range) As String
    '检查P列至DR列中的空值
    '如果有空值，返回相应的列名称
    '如果没有空值，返回"数据完整无空值"
    
    Dim cell As Range
    Dim emptyColumns As String
    Dim colLetter As String
    Dim hasEmpty As Boolean
    
    emptyColumns = ""
    hasEmpty = False
    
    '遍历范围内的每个单元格
    For Each cell In rng
        If IsEmpty(cell.Value) Or Trim(cell.Value) = "" Then
            '获取列名
            colLetter = Split(cell.Address, "$")(1)
            
            '检查是否已经记录过该列
            If InStr(emptyColumns, colLetter) = 0 Then
                If emptyColumns = "" Then
                    emptyColumns = colLetter
                Else
                    emptyColumns = emptyColumns & ", " & colLetter
                End If
                hasEmpty = True
            End If
        End If
    Next cell
    
    '返回结果
    If hasEmpty Then
        CheckEmptyColumns = emptyColumns
    Else
        CheckEmptyColumns = "数据完整无空值"
    End If
End Function

Sub CheckAllRowsInSheet()
    '检查工作表CY26-34中P列到EE列是否有空值
    '将有空值的行对应的A列和C列值添加到新的sheet中
    
    '性能优化：关闭屏幕更新和自动计算
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim outputRow As Long
    
    Set ws = ThisWorkbook.Worksheets("CY26-34")
    
    '创建或清空输出工作表
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets("空值检查结果")
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        '如果工作表不存在，创建新工作表
        Set wsOutput = ThisWorkbook.Worksheets.Add
        wsOutput.Name = "空值检查结果"
    Else
        '如果工作表存在，清空内容
        wsOutput.Cells.Clear
    End If
    
    '添加标题行
    wsOutput.Cells(1, 1).Value = "A列值"
    wsOutput.Cells(1, 2).Value = "C列值"
    
    '找到数据最后一行
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    startCol = 16  'P列（第16列）
    endCol = 109   'EE列（第109列）
    outputRow = 2  '从第2行开始写入（第1行是标题）
    
    '一次性读取所有需要的数据到数组（提升性能）
    Dim dataArr As Variant
    Dim colAArr As Variant
    Dim colCArr As Variant
    
    'P到EE列的数据
    dataArr = ws.Range(ws.Cells(2, startCol), ws.Cells(lastRow, endCol)).Value
    'A列的数据
    colAArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).Value
    'C列的数据
    colCArr = ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 3)).Value
    
    '遍历数组检查空值
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim hasEmptyInRow As Boolean
    Dim cellValue As Variant
    
    For rowIdx = 1 To UBound(dataArr, 1)
        hasEmptyInRow = False
        
        '检查当前行的P到EE列是否有空值
        For colIdx = 1 To UBound(dataArr, 2)
            cellValue = dataArr(rowIdx, colIdx)
            
            '检查是否为空
            If IsEmpty(cellValue) Or cellValue = "" Then
                hasEmptyInRow = True
                Exit For '找到空值就退出内层循环
            End If
        Next colIdx
        
        '如果该行P到EE列有空值，将A列和C列的值写入新sheet
        If hasEmptyInRow Then
            wsOutput.Cells(outputRow, 1).Value = colAArr(rowIdx, 1) 'A列值
            wsOutput.Cells(outputRow, 2).Value = colCArr(rowIdx, 1) 'C列值
            outputRow = outputRow + 1
        End If
    Next rowIdx
    
    '自动调整列宽
    wsOutput.Columns("A:B").AutoFit
    
    '恢复Excel设置
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub



