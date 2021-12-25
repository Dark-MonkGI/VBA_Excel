Attribute VB_Name = "Module1"
Sub K_T()
Attribute K_T.VB_ProcData.VB_Invoke_Func = "й\n14"
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '------------------------------------------------
    'Необходимые даты
    Dim NewData As Date
    Dim FistData As Date
    Dim LastData As Date
    
    NewData = Date
    FistData = NewData - 7
    LastData = NewData - 1
    'Проверяем дни недели
    If Weekday(NewData, vbMonday) <> 1 Then MsgBox ("Сегодня не понедельник!" & Chr(10) & "Даты формирования отчета будут сформированы неверно!" & Chr(10) & "Измени в ручную!")
    '------------------------------------------------
    
    'открываем отчет который необходдимо сформировать и создаем новый лист
    
    
    Workbooks.Open Filename:="U:\****.xlsx"
    'Workbooks.Open Filename:="C:\****.xlsx"
    'Запоминаем имя главной самой книги
    Dim HeadWorkbook As String
    HeadWorkbook = ActiveWorkbook.Name
    
    If ActiveSheet.Name <> Worksheets(1).Name Then Worksheets(1).Activate
    Worksheets(1).Copy Before:=Worksheets(1)
    Worksheets(1).Name = NewData
    Worksheets(1).Activate
    
    'Меняем заголовок
    Dim Heading As String, FirstValueDate As String, LastValueDate As String
    Heading = Range("A1")
    FirstValueDate = Mid(Heading, 67, 10)
    LastValueDate = Right(Heading, 10)
    Heading = Replace(Heading, FirstValueDate, FistData)
    Heading = Replace(Heading, LastValueDate, LastData)
    Range("A1") = Heading
    
    'Pain форматирования
    Dim NameSecondList As String
    NameSecondList = Worksheets(2).Name
    
    Range("C6:O22").FormatConditions.Delete
    'Строки с правилами 6 8 10 12 14 16 17 19 21
    
    '1-----------------------6row----------------------------------
    'Green
    Range("C6:O6").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$6>'" + NameSecondList + "'!C$6"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$6<'" + NameSecondList + "'!C$6"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '2-------------------------8row-----------------------------
    'Green
    Range("C8:O8").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$8<'" + NameSecondList + "'!C$8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$8>'" + NameSecondList + "'!C$8"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '3-------------------------10row----------------------------
    'Green
    Range("C10:O10").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$10>'" + NameSecondList + "'!C$10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$10<'" + NameSecondList + "'!C$10"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '4-------------------------12row----------------------------
    'Green
    Range("C12:O12").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$12<'" + NameSecondList + "'!C$12"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$12>'" + NameSecondList + "'!C$12"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '5-------------------------14row----------------------------
    'Green
    Range("C14:O14").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$14>'" + NameSecondList + "'!C$14"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$14<'" + NameSecondList + "'!C$14"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '6-------------------------16row----------------------------
    'Green
    Range("C16:O16").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$16<'" + NameSecondList + "'!C$16"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$16>'" + NameSecondList + "'!C$16"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '7-------------------------17row----------------------------
    'Green
    Range("C17:O17").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$17<'" + NameSecondList + "'!C$17"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$17>'" + NameSecondList + "'!C$17"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '8-------------------------19row----------------------------
    'Green
    Range("C19:O19").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$19>'" + NameSecondList + "'!C$19"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$19<'" + NameSecondList + "'!C$19"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '9-------------------------21row----------------------------
    'Green
    Range("C21:O21").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$21<'" + NameSecondList + "'!C$21"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    'Red
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=C$21>'" + NameSecondList + "'!C$21"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10461183
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '---------------------------------------------------------
    Range("A2").Select
    
    
    'Заполнение столбцов
    
    Dim ArrColumn() As Variant
    ArrColumn = Array("C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O")
    
    'Функция работает с выгруженными файлами по очериди для каждого филиала
    For i = LBound(ArrColumn) To UBound(ArrColumn)
        Call FillColumn(ArrColumn(i), HeadWorkbook)
    Next i
    
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    'ActiveWorkbook.Save
    
End Sub

Sub FillColumn(NameColumn As Variant, HeadWorkbook As String)
    
    Dim NameRegion As String
    Dim Directory As String
    Dim NameRegionFile As String
    Dim xlsx, xls As String
    
    'Открыть фаил с выгрузкой
    
    xlsx = ".xlsx"
    xls = ".xls"
    NameRegion = Range(NameColumn + "3").Value
    
    Directory = "U:\****\"
    'Directory = "C:\****\"
    
    '-----------------------------------
    If (Dir(Directory + NameRegion + xlsx) = "") And (Dir(Directory + NameRegion + xls) = "") Then
        MsgBox "Файл " + NameRegion + " не существует"
        'MsgBox ("Необходимо проверить наличие файла, затем запустить все с самого начала.")
        Exit Sub
    ElseIf (Dir(Directory + NameRegion + xlsx) = "") Then
        NameRegionFile = NameRegion + xls
        Workbooks.Open Filename:=Directory + NameRegionFile
    ElseIf (Dir(Directory + NameRegion + xls) = "") Then
        NameRegionFile = NameRegion + xlsx
        Workbooks.Open Filename:=Directory + NameRegionFile
    End If
    If ActiveWorkbook.Name <> NameRegionFile Then Windows(NameRegionFile).Activate
    'Книга открыта и активирована
    '-----------------------------------
    
    
    
    
    
    '-----------------------------------
    '---После измениеия удалить---функция удаления шляп--
    Dim ArrSectionTitle(1 To 4) As String
    ArrSectionTitle(1) = "I Основные потери превышают верхнюю границу коридора нулевого баланса"
    ArrSectionTitle(2) = "II Основные потери находятся в границах коридора нулевого баланса"
    ArrSectionTitle(3) = "III Основные потери меньше нижней отрицательной границы коридора нулевого баланса"
    ArrSectionTitle(4) = "IV Показания отсутствуют"
    
    For i = LBound(ArrSectionTitle) To UBound(ArrSectionTitle)
        Call DeliteTitle(ArrSectionTitle(i))
    Next i
    '-----------------------------------
    
    
    
    
    
    
    '-----------------------------------
    '--вставляем столбец и протягиваем в нем формулы---
    Sheets(1).Columns(9).Insert Shift:=xlToRightt
    
    Call InsertFormula(ArrSectionTitle(), NameColumn, HeadWorkbook, NameRegionFile)
    
    
    
    
    
    
    
    
    
    
    '-----------------------------------
    
    
    
    
    
    
    
    'Скопировать поочередно данные в главный лист
    'Сохранить и закрвыть фаил
    ActiveWorkbook.Close SaveChanges:=True
    
End Sub
Sub DeliteTitle(LookingValue As String)
    
    'Последняя строка
    Dim lastRow As Integer
    lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
    
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
    
    If CountLine = 2 Then
        LineOne = ArrStrings(2)
        LineTwo = LineOne + 4
        Rows(CStr(LineOne) + ":" + CStr(LineTwo)).Delete Shift:=xlUp
        Exit Sub
    End If
    
    
    
    
    Dim tempRange As Range
    Set tempRange = Range(CStr(ArrStrings(2)) + ":" + CStr(ArrStrings(2)))
    
    For i = 2 To UBound(ArrStrings)
        Set tempRange = Union(Range(CStr(ArrStrings(i)) + ":" + CStr(ArrStrings(i))), tempRange)
        Set tempRange = Union(Range(CStr(ArrStrings(i) + 1) + ":" + CStr(ArrStrings(i) + 1)), tempRange)
        Set tempRange = Union(Range(CStr(ArrStrings(i) + 2) + ":" + CStr(ArrStrings(i) + 2)), tempRange)
        Set tempRange = Union(Range(CStr(ArrStrings(i) + 3) + ":" + CStr(ArrStrings(i) + 3)), tempRange)
        Set tempRange = Union(Range(CStr(ArrStrings(i) + 4) + ":" + CStr(ArrStrings(i) + 4)), tempRange)
    Next i
    
    tempRange.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
End Sub

Sub InsertFormula(ByRef ArrSectionTitle() As String, ByRef NameColumn As Variant, ByRef HeadWorkbook As String, ByRef NameRegionFile As String)

    
    Dim i, j, index As Integer
    Dim CellValue, CellTwoValue, Total As String
    Dim ArrStrings(1 To 4) As Integer
    Dim ArrStringsTotal(1 To 3) As Integer
    Dim lastRow As Integer
    
    lastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
    
    
    index = 1
    
    For i = LBound(ArrSectionTitle) To UBound(ArrSectionTitle)
    
        For j = 1 To lastRow
            CellValue = ActiveWorkbook.Worksheets(1).Cells(j, 1).Value
            If CellValue Like ArrSectionTitle(i) Then
                ArrStrings(index) = j
                index = index + 1
            End If
        
        Next j
      
    Next i
    
    '-------Ищем "Всего"---------------
    Total = "Всего:"
    index = 1
    For i = 1 To lastRow
        CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 3).Value
            If CellValue Like Total Then
                ArrStringsTotal(index) = i
                index = index + 1
            End If
    Next i
    
    
    
    'For i = 1 To UBound(ArrStrings)
        'MsgBox (ArrStrings(i))
    'Next i
    
    
    '---------------------------------------
    '-Топорно вставляем формулы-
    
    For i = ArrStrings(1) + 5 To ArrStrings(2) - 3
        Range("I" + CStr(i)).FormulaR1C1 = "=RC[-1]-RC[-2]"
    Next i
    
    For i = ArrStrings(2) + 5 To ArrStrings(3) - 3
        Range("I" + CStr(i)).FormulaR1C1 = "=RC[-1]-RC[-2]"
    Next i
    
    For i = ArrStrings(3) + 5 To ArrStrings(4) - 5
        Range("I" + CStr(i)).FormulaR1C1 = "=RC[-1]+RC[-2]"
    Next i
    
    '---------------------------------------
    
    ActiveWorkbook.Save
    
    '---------------------------------------
    '-Общее количесво обьектов в выгрузке-
    
    
    Dim Summ As Integer
    Dim Summ2 As Integer
    Dim Summ3 As Integer
    Dim Summ4 As Integer
    Dim Summ5 As Integer
    
    '---------------------------------------
    'Строка 3
    
    
    Summ = ((ArrStringsTotal(1) - 1) - (ArrStrings(1) + 4)) + ((ArrStringsTotal(2) - 1) - (ArrStrings(2) + 4)) + ((ArrStringsTotal(3) - 1) - (ArrStrings(3) + 4)) + (lastRow - (ArrStrings(4) + 4))
    
    'MsgBox ("Всего **** " + CStr(Summ))
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "6").Value = Summ

    
    '---------------------------------------
    'Строка 5
    Summ2 = (lastRow - (ArrStrings(4) + 4))
    'MsgBox ("Кол-во  в 4 разделе" + CStr(Summ2))
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "8").Value = Summ2
    '---------------------------------------
    'Строка 7
    Summ3 = Summ - Summ2
    'MsgBox ("Кол-во в 1,2 и 3 разделах" + CStr(Summ3))
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "10").Value = Summ3
    
    '---------------------------------------
    'Строка 9
    Summ4 = ((ArrStrings(4) - 5) - (ArrStrings(3) + 4))
    'MsgBox ("Кол-во в 3 разделе" + CStr(Summ4))
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "12").Value = Summ4
    
    '---------------------------------------
    'Строка 11
    Summ5 = Summ3 - Summ4
    'MsgBox ("Кол-во  в 1 и 2 разделах" + CStr(Summ5))
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "14").Value = Summ5
    
    
    '---------------------------
    

    If ActiveWorkbook.Name <> NameRegionFile Then Windows(NameRegionFile).Activate
    
    
    
    '--------Пункт "Средне значние"-------------------
   
   
   
    Dim SummAllValue As Double
    Dim CauntValue As Integer
    Dim Result As Double
    
    CauntValue = 0
    SummAllValue = 0
    
   
    For i = 1 To lastRow
        If (i >= (ArrStrings(1) + 5) And i <= (ArrStringsTotal(1) - 1)) Or (i >= (ArrStrings(2) + 5) And i <= (ArrStringsTotal(2) - 1)) Or _
                                                                                                    (i >= (ArrStrings(3) + 5) And i <= (ArrStringsTotal(3) - 1)) Then
            
            CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 8).Value
                'If CellValue <> "" And CellValue < 100 Then
                If CellValue <> "" And CellValue < 80 And CellValue >= 0 Then
                    SummAllValue = SummAllValue + CellValue
                    
                    CauntValue = CauntValue + 1
                End If
        End If
    Next i
    
    Result = SummAllValue / CauntValue
    Result = Result / 100
    
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "16").Value = Result
    
    '-----Номера строк в переменные---------
    Dim FirstSectionFirstLine As Integer, FirstSectionLastLine As Integer, SecondSectionFirstLine As Integer, SecondSectionLastLine As Integer
    Dim ThirdSectionFirstLine As Integer, ThirdSectionLastLine As Integer
    
    FirstSectionFirstLine = ArrStrings(1) + 5
    FirstSectionLastLine = ArrStringsTotal(1) - 1
    
    SecondSectionFirstLine = ArrStrings(2) + 5
    SecondSectionLastLine = ArrStringsTotal(2) - 1
    
    ThirdSectionFirstLine = ArrStrings(3) + 5
    ThirdSectionLastLine = ArrStringsTotal(3) - 1
    '---------------------------
    
    
    '--------Пункт 14---"Кол-во  в 1 разделе "-----
    If ActiveWorkbook.Name <> NameRegionFile Then Windows(NameRegionFile).Activate
    
    Dim CountNambersValeu As Integer
    CountNambersValeu = 0
    
    For i = FirstSectionFirstLine To FirstSectionLastLine
        CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 8).Value
        CellTwoValue = ActiveWorkbook.Worksheets(1).Cells(i, 9).Value
        If CellValue <> "" Then
            If CellTwoValue <= (-2) Then
                CountNambersValeu = CountNambersValeu + 1
            End If
        End If
    Next i
    
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "17").Value = CountNambersValeu
    '---------------------------
    
    
    '--------Пункт 16---"Кол-во  в 2 разделе "-----
    If ActiveWorkbook.Name <> NameRegionFile Then Windows(NameRegionFile).Activate
    
    CountNambersValeu = 0
    For i = SecondSectionFirstLine To SecondSectionLastLine
        CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 8).Value
        CellTwoValue = ActiveWorkbook.Worksheets(1).Cells(i, 9).Value
        If CellValue <> "" Then
            If CellTwoValue >= (-2) And CellTwoValue <= 2.01 Then
                CountNambersValeu = CountNambersValeu + 1
            End If
        End If
    Next i
    
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "19").Value = CountNambersValeu
    '---------------------------
    
    
        '--------Пункт 18---"Кол-во  в 3 разделе "-----
    If ActiveWorkbook.Name <> NameRegionFile Then Windows(NameRegionFile).Activate
    
    CountNambersValeu = 0
    
    For i = ThirdSectionFirstLine To ThirdSectionLastLine
        CellValue = ActiveWorkbook.Worksheets(1).Cells(i, 8).Value
        CellTwoValue = ActiveWorkbook.Worksheets(1).Cells(i, 9).Value
        If CellValue <> "" Then
            If CellTwoValue <= (-2) Then
                CountNambersValeu = CountNambersValeu + 1
            End If
        End If
    Next i
    
    Workbooks(HeadWorkbook).Sheets(1).Range(NameColumn + "21").Value = CountNambersValeu
    '---------------------------
    If ActiveWorkbook.Name <> NameRegionFile Then Windows(NameRegionFile).Activate
    
    
    
    
End Sub
