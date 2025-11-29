'==============================================================================
' CheckEmptyColumns - 优化版本
' 性能改进:
' 1. 使用 Option Explicit 强制变量声明
' 2. 使用常量提高可维护性
' 3. CheckEmptyColumns 函数改用数组处理
' 4. CheckAllRowsInSheet 使用批量写入结果
' 5. 添加错误处理
'==============================================================================
Option Explicit

' 常量定义
Private Const WS_SOURCE_NAME As String = "CY26-34"
Private Const WS_OUTPUT_NAME As String = "空值检查结果"
Private Const START_COL As Long = 16    ' P列
Private Const END_COL As Long = 109     ' EE列
Private Const DATA_START_ROW As Long = 2

Function CheckEmptyColumns(rng As Range) As String
    '检查范围中的空值列
    '如果有空值，返回相应的列名称
    '如果没有空值，返回"数据完整无空值"
    '
    ' 优化: 使用数组处理避免逐单元格访问
    
    Dim dataArr As Variant
    Dim emptyColDict As Object
    Dim rowIdx As Long, colIdx As Long
    Dim cellValue As Variant
    Dim colLetter As String
    Dim result As String
    
    ' 单个单元格处理
    If rng.Cells.Count = 1 Then
        If IsEmpty(rng.Value) Or Trim(CStr(rng.Value)) = "" Then
            CheckEmptyColumns = Split(rng.Address, "$")(1)
        Else
            CheckEmptyColumns = "数据完整无空值"
        End If
        Exit Function
    End If
    
    ' 读取数据到数组
    dataArr = rng.Value
    
    ' 使用 Dictionary 去重 (比 InStr 检查更快)
    Set emptyColDict = CreateObject("Scripting.Dictionary")
    
    ' 遍历数组
    For rowIdx = 1 To UBound(dataArr, 1)
        For colIdx = 1 To UBound(dataArr, 2)
            cellValue = dataArr(rowIdx, colIdx)
            
            ' 检查是否为空
            If IsEmpty(cellValue) Then
                If Not emptyColDict.Exists(colIdx) Then
                    emptyColDict.Add colIdx, GetColumnLetter(rng.Columns(colIdx).Column)
                End If
            ElseIf VarType(cellValue) = vbString Then
                If Len(Trim$(cellValue)) = 0 Then
                    If Not emptyColDict.Exists(colIdx) Then
                        emptyColDict.Add colIdx, GetColumnLetter(rng.Columns(colIdx).Column)
                    End If
                End If
            End If
        Next colIdx
    Next rowIdx
    
    ' 构建结果字符串
    If emptyColDict.Count = 0 Then
        CheckEmptyColumns = "数据完整无空值"
    Else
        result = Join(emptyColDict.Items, ", ")
        CheckEmptyColumns = result
    End If
    
    Set emptyColDict = Nothing
End Function

' 辅助函数: 将列号转换为列字母
Private Function GetColumnLetter(ByVal colNum As Long) As String
    Dim colLetter As String
    
    Do While colNum > 0
        colLetter = Chr(((colNum - 1) Mod 26) + 65) & colLetter
        colNum = (colNum - 1) \ 26
    Loop
    
    GetColumnLetter = colLetter
End Function

Sub CheckAllRowsInSheet()
    '检查工作表CY26-34中P列到EE列是否有空值
    '将有空值的行对应的A列和C列值添加到新的sheet中
    '
    ' 优化: 使用批量写入结果而非逐行写入
    
    Dim ws As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    
    ' 性能优化：关闭屏幕更新和自动计算
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' 获取源工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_SOURCE_NAME)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "找不到工作表 '" & WS_SOURCE_NAME & "'", vbExclamation, "错误"
        GoTo Cleanup
    End If
    
    ' 创建或清空输出工作表
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets(WS_OUTPUT_NAME)
    On Error GoTo ErrorHandler
    
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Worksheets.Add
        wsOutput.Name = WS_OUTPUT_NAME
    Else
        wsOutput.Cells.Clear
    End If
    
    ' 找到数据最后一行
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < DATA_START_ROW Then
        MsgBox "源工作表没有数据", vbInformation, "提示"
        GoTo Cleanup
    End If
    
    ' 一次性读取所有需要的数据到数组
    Dim dataArr As Variant   ' P到EE列
    Dim colAArr As Variant   ' A列
    Dim colCArr As Variant   ' C列
    Dim rowCount As Long
    
    rowCount = lastRow - DATA_START_ROW + 1
    dataArr = ws.Range(ws.Cells(DATA_START_ROW, START_COL), ws.Cells(lastRow, END_COL)).Value
    colAArr = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, 1)).Value
    colCArr = ws.Range(ws.Cells(DATA_START_ROW, 3), ws.Cells(lastRow, 3)).Value
    
    ' 第一遍扫描: 计算有空值的行数 (用于预分配结果数组)
    Dim emptyRowCount As Long
    Dim rowIdx As Long, colIdx As Long
    Dim hasEmptyInRow As Boolean
    Dim cellValue As Variant
    
    emptyRowCount = 0
    For rowIdx = 1 To UBound(dataArr, 1)
        hasEmptyInRow = False
        For colIdx = 1 To UBound(dataArr, 2)
            cellValue = dataArr(rowIdx, colIdx)
            If IsEmpty(cellValue) Or cellValue = "" Then
                hasEmptyInRow = True
                Exit For
            End If
        Next colIdx
        If hasEmptyInRow Then emptyRowCount = emptyRowCount + 1
    Next rowIdx
    
    ' 预分配结果数组 (包含标题行)
    Dim resultArr() As Variant
    ReDim resultArr(1 To emptyRowCount + 1, 1 To 2)
    
    ' 设置标题
    resultArr(1, 1) = "A列值"
    resultArr(1, 2) = "C列值"
    
    ' 第二遍扫描: 填充结果数组
    outputRow = 2
    For rowIdx = 1 To UBound(dataArr, 1)
        hasEmptyInRow = False
        For colIdx = 1 To UBound(dataArr, 2)
            cellValue = dataArr(rowIdx, colIdx)
            If IsEmpty(cellValue) Or cellValue = "" Then
                hasEmptyInRow = True
                Exit For
            End If
        Next colIdx
        
        If hasEmptyInRow Then
            resultArr(outputRow, 1) = colAArr(rowIdx, 1)
            resultArr(outputRow, 2) = colCArr(rowIdx, 1)
            outputRow = outputRow + 1
        End If
    Next rowIdx
    
    ' 批量写入结果 (单次操作)
    If emptyRowCount > 0 Then
        wsOutput.Range("A1").Resize(emptyRowCount + 1, 2).Value = resultArr
    Else
        wsOutput.Range("A1:B1").Value = Array("A列值", "C列值")
    End If
    
    ' 格式化标题行
    With wsOutput.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' 自动调整列宽
    wsOutput.Columns("A:B").AutoFit
    
    ' 显示结果
    If emptyRowCount = 0 Then
        MsgBox "未发现空值", vbInformation, "检查完成"
    Else
        MsgBox "发现 " & emptyRowCount & " 行包含空值", vbInformation, "检查完成"
        wsOutput.Activate
    End If
    
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

