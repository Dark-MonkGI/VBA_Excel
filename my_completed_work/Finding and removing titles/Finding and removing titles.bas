Sub DeliteTitle()

    
    'Последняя строка
    Dim lastRow As Integer
    lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
    
    
    Dim LookingValue As String
    LookingValue = "****"
    
    'Поиск оглавлений и подсчет
    Dim CountLine, i, index As Integer
    Dim ArrStrings() As Integer
    Dim CellValue As String
    
    CountLine = 0
    
    'Поиск количесва вхождений заголовков в отчет
    For i = 1 To lastRow
        CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 1).Value
        If CellValue Like LookingValue Then
            CountLine = CountLine + 1
        End If
    Next i
    
    If CountLine = 1 Then
        Exit Sub
    End If
    
    If CountLine = 0 Then
        MsgBox ("В файле:" & Chr(10) & ActiveWorkbook.Name & Chr(10) & "не найден заголовок:" _
                                                        + LookingValue & Chr(10) & "Необходимо проверить фаил!")
        Exit Sub
    End If
    
    
    'Создаем массив из строк со вхождением заголовков
    ReDim ArrStrings(1 To CountLine) As Integer
    index = 1
    
    For i = 1 To lastRow
        CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 1).Value
        If CellValue Like LookingValue Then
            ArrStrings(index) = i
            index = index + 1
        End If
    Next i
    
    
    
    
    
    'Удаление значений
    Dim LineOne As Integer, LineTwo As Integer
    
    'Если всего два заголовка -удаляем второй
    If CountLine = 2 Then
        LineOne = ArrStrings(2)
        LineTwo = LineOne + 2
        Rows(CStr(LineOne) + ":" + CStr(LineTwo)).Delete Shift:=xlUp
        Exit Sub
    End If
    
    
    
    
    Dim tempRange As Range
    Set tempRange = Range(CStr(ArrStrings(2)) + ":" + CStr(ArrStrings(2)))
    
    For i = 2 To UBound(ArrStrings)
        Set tempRange = Union(Range(CStr(ArrStrings(i)) + ":" + CStr(ArrStrings(i))), tempRange)
        Set tempRange = Union(Range(CStr(ArrStrings(i) + 1) + ":" + CStr(ArrStrings(i) + 1)), tempRange)
        Set tempRange = Union(Range(CStr(ArrStrings(i) + 2) + ":" + CStr(ArrStrings(i) + 2)), tempRange)
    Next i
    
    tempRange.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub

