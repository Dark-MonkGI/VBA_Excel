Sub ReportReceivingDevice()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '------------------------------------------------
    'Работа с Датами
    Dim ToDayData As Date, WeekdayToday As Integer, TodayYear As Integer, TodayIsDay As Integer, TodayIsMonth As Integer, DayOfWeek As String, FullNewDate As String
    
    ToDayData = Date
    WeekdayToday = Weekday(ToDayData, vbMonday)
    
    TodayIsDay = DatePart("d", ToDayData)
    TodayIsMonth = DatePart("m", ToDayData)
    TodayYear = DatePart("yyyy", ToDayData)
    FullNewDate = " " & TodayYear & "-" & TodayIsMonth & "-" & TodayIsDay
    
    
    Select Case WeekdayToday
        Case Is = 1
            DayOfWeek = "Пн"
        Case Is = 2
            DayOfWeek = "Вт"
        Case Is = 3
            DayOfWeek = "Ср"
        Case Is = 4
            DayOfWeek = "Чт"
        Case Is = 5
            DayOfWeek = "Пт"
        Case Is = 6
            DayOfWeek = "Сб"
        Case Else
            DayOfWeek = "Вс"
    End Select
    '------------------------------------------------
    
    
    
    '------------------------------------------------
    'Перемещаем, переименовываем, открываем файлы
    
    Dim DirectoryDownloaded As String, DestinationPathLoaded As String, MainFilePath As String, DestinationMainFilePath As String, DestinationMainFilePathDesktop As String
    Dim NameDownloadedReport As String, NameDestinationDownloadedReport As String, NameMainFiles As String, NameDestinationMainFile As String, DateCreateFile As String
    
    '-------
    
    DestinationMainFilePathDesktop = "C:\Users\Desktop\"
    DirectoryDownloaded = "C:\Users\Downloads\"
    '-------
    DestinationPathLoaded = "U:\"
    MainFilePath = "U:\"
    DestinationMainFilePath = "U:\"
    
    
    NameDownloadedReport = "Отчет по актуальности данных"
    NameMainFiles = "сводный с динамикой.xlsx"
    
    Dim xlsx, xls, ReportFormat As String
    xlsx = ".xlsx"
    xls = ".xls"
    
    If (Dir(DirectoryDownloaded + NameDownloadedReport + xlsx) = "") And (Dir(DirectoryDownloaded + NameDownloadedReport + xls) = "") Then
        MsgBox "Файл: " + NameDownloadedReport + " -не существует", vbCritical
        Exit Sub
    ElseIf (Dir(DirectoryDownloaded + NameDownloadedReport + xlsx) = "") Then
        ReportFormat = xls
        DateCreateFile = Left(CStr(FileDateTime(DirectoryDownloaded + NameDownloadedReport + xls)), 10)
        If CStr(ToDayData) Like DateCreateFile Then
            Name (DirectoryDownloaded + NameDownloadedReport + xls) As (DestinationPathLoaded + NameDownloadedReport + FullNewDate + xls)
            NameDownloadedReport = NameDownloadedReport + FullNewDate
            Workbooks.Open Filename:=DestinationPathLoaded + NameDownloadedReport + xls
        Else
            MsgBox "Файл: " + NameDownloadedReport + " -не верной даты выгрузки.", vbCritical
            Exit Sub
        End If
        
    ElseIf (Dir(DirectoryDownloaded + NameDownloadedReport + xls) = "") Then
        ReportFormat = xlsx
        DateCreateFile = Left(CStr(FileDateTime(DirectoryDownloaded + NameDownloadedReport + xlsx)), 10)
        
        If CStr(ToDayData) Like DateCreateFile Then
            Name (DirectoryDownloaded + NameDownloadedReport + xlsx) As (DestinationPathLoaded + NameDownloadedReport + FullNewDate + xlsx)
            NameDownloadedReport = NameDownloadedReport + FullNewDate
            Workbooks.Open Filename:=DestinationPathLoaded + NameDownloadedReport + xlsx
        Else
            MsgBox "Файл: " + NameDownloadedReport + " -не верной даты выгрузки.", vbCritical
            Exit Sub
        End If
    End If
    
    Dim MainWk As Workbook, Wk As Workbook
    
    Set MainWk = Workbooks.Open(MainFilePath + NameMainFiles)
    Set Wk = Workbooks(NameDownloadedReport + ReportFormat)
    If ActiveWorkbook.Name <> NameMainFiles Then MainWk.Activate
    
    
    '------------------------------------------------
    
    
    
    
    '------------------------------------------------
    'Работа с отчетом
    Columns(2).Insert Shift:=xlToRight
    Columns(2).Insert Shift:=xlToRight
    Columns(2).Insert Shift:=xlToRight
    Range("E2:G61").Copy Range("B2")
    Range("B:G").ColumnWidth = 16
    Range("E2:G2").Interior.Color = 49407
    Range("B2") = ToDayData
    
    '--ВПРим---
    Dim i As Integer, j As Integer, index As Integer
    Dim ArrValue(0 To 2) As Variant
    Dim LookingValue As String, CellValue As String
    Dim Optionals As Integer
    
    Dim LookingPenzaBranch As String
    Dim LookingIvanovoBranch As String
    
    LookingPenzaBranch = "Итого ***"
    LookingIvanovoBranch = "Итого ***"
    
    Optionals = 0
    
    For i = 7 To 60
        LookingValue = MainWk.Sheets(1).Cells(i, 1)
        '-------в следующей строке меняеться  цифра (Втрая - последняя) в зависимотсти от кол-ва строк в выгрузке----
        For j = 6 To 70
            CellValue = Wk.Worksheets(1).Cells(j, 1).Value
                
            If CellValue Like LookingValue Then
                ArrValue(0) = Wk.Worksheets(1).Cells(j, 2).Value2
                ArrValue(1) = Wk.Worksheets(1).Cells(j, 3).Value2
                ArrValue(2) = Wk.Worksheets(1).Cells(j, 5).Value
                Optionals = 1
            End If
        Next j
        
        If Optionals = 1 Then
            MainWk.Sheets(1).Cells(i, 2) = ArrValue(0)
            MainWk.Sheets(1).Cells(i, 3) = ArrValue(1)
            MainWk.Sheets(1).Cells(i, 4).Value = ArrValue(2)
            Optionals = 0
        ElseIf LookingValue Like LookingPenzaBranch Or LookingValue Like LookingIvanovoBranch Then
            MainWk.Sheets(1).Cells(i, 2) = Null
            MainWk.Sheets(1).Cells(i, 3) = Null
            MainWk.Sheets(1).Cells(i, 4).Value = Null
            Optionals = 0
        Else
            MainWk.Sheets(1).Cells(i, 2) = Null
            MainWk.Sheets(1).Cells(i, 2).Interior.Color = 255
            MainWk.Sheets(1).Cells(i, 3) = Null
            MainWk.Sheets(1).Cells(i, 3).Interior.Color = 255
            MainWk.Sheets(1).Cells(i, 4).Value = Null
            MainWk.Sheets(1).Cells(i, 4).Interior.Color = 255
            Optionals = 0
        End If
        
    Next i
    Wk.Close False
    If ActiveWorkbook.Name <> NameMainFiles Then MainWk.Activate
    
    Range("B31:D31").Value = Range("B30:D30").Value
    Range("B33:D33").Value = Range("B32:D32").Value
    Range("B41:D41").Value = Range("B40:D40").Value
    
    '--Владимирский филиал--
    Range("B39").Value = (Range("B37").Value + Range("B38").Value)
    Range("C39").Value = (Range("C37").Value + Range("C38").Value)
    Range("D39").Value = (Range("C39").Value / Range("B39").Value)
    
    
    '--Пермский филиал--
    Range("B57").Value = (Range("B53").Value + Range("B54").Value + Range("B55").Value + Range("B56").Value)
    Range("C57").Value = (Range("C53").Value + Range("C54").Value + Range("C55").Value + Range("C56").Value)
    Range("D57").Value = (Range("C57").Value / Range("B57").Value)
    
    '--Нижегородский филиал--
    Range("B60").Value = (Range("B58").Value + Range("B59").Value)
    Range("C60").Value = (Range("C58").Value + Range("C59").Value)
    Range("D60").Value = (Range("C60").Value / Range("B60").Value)
    
    
    
    
    
    '--Итоги по столбцам--
    
    Range("B61").Value = ((WorksheetFunction.Sum(Range("B7:B36"))) + (WorksheetFunction.Sum(Range("B37:B60")))) / 2
    Range("C61").Value = ((WorksheetFunction.Sum(Range("C7:C36"))) + (WorksheetFunction.Sum(Range("C37:C60")))) / 2
    Range("D61").Value = (Range("C61").Value / Range("B61").Value)
    
    '------------------------------------------------

    
    
    '------------------------------------------------
    '--Удаление лишних столбцов и группировка----
    Dim DinamicStatus As Integer
    
    DinamicStatus = 0
    
    If DayOfWeek Like "Вт" And Not Range("E3").Value Like "пн" Then
        MsgBox ("Внимание!!!" & Chr(13) & "Динамика будет построенна за неделю с сегодняшнего дня!")
        GoTo M:
    End If
    
    If DayOfWeek Like "Ср" And Not Range("E3").Value Like "вт" Then
        MsgBox ("Внимание!!!" & Chr(13) & "Динамика будет построенна за неделю с сегодняшнего дня!")
        GoTo M:
    End If
    
    
    
    If DayOfWeek Like "Пн" Then
        MsgBox ("Сегодня понедельник" & Chr(13) & "Формируется недельная динамика")
M:
        Dim SearchSector As String
        Dim ArrColumn(1 To 9) As String
        ArrColumn(1) = "E3"
        ArrColumn(2) = "H3"
        ArrColumn(3) = "K3"
        ArrColumn(4) = "N3"
        ArrColumn(5) = "Q3"
        ArrColumn(6) = "T3"
        ArrColumn(7) = "W3"
        ArrColumn(8) = "Z3"
        ArrColumn(9) = "AC3"
        
        DinamicStatus = 1
        
        For i = 2 To 9
            If Range(ArrColumn(i)).Value Like "пн" Then
                SearchSector = (Range(ArrColumn(i - 1)).Column) + 2
                GoTo N:
            End If
        Next i
N:
        If SearchSector = "" Then
            MsgBox ("Прошлый понедельник не найден." & Chr(13) & " динамика будет сформирована некорректно!")
            GoTo S:
        End If
        
        
        SearchSector = ((SearchSector - 1) / 3) - 1
        
        For i = 1 To SearchSector
            Columns("E:G").Delete Shift:=xlToLeft
        Next i
         
    End If
    
    '------------------------------------------------
    
    
    
    
    '------------------------------------------------
    '--Динамика-----
S:
    
    Dim LastColumn As Integer, DynamicsColumn As Integer, Dynamics2Column As Integer
    

    LastColumn = Cells(7, Columns.Count).End(xlToLeft).Column
    DynamicsColumn = LastColumn - 2
    Dynamics2Column = LastColumn - 1
    
    
   
    For i = 7 To 61
        MainWk.Sheets(1).Cells(i, DynamicsColumn) = (Range("B" & CStr(i)).Value - Range("E" & CStr(i)).Value)
        MainWk.Sheets(1).Cells(i, Dynamics2Column) = (Range("C" & CStr(i)).Value - Range("F" & CStr(i)).Value)
        MainWk.Sheets(1).Cells(i, LastColumn).Value = (Range("D" & CStr(i)).Value - Range("G" & CStr(i)).Value)
    Next i
    
    
    If DinamicStatus = 0 Then
        MainWk.Sheets(1).Cells(3, DynamicsColumn) = "За сутки"
        Cells(3, DynamicsColumn).Interior.Color = 65535
    Else
        MainWk.Sheets(1).Cells(3, DynamicsColumn) = "За неделю"
        Cells(3, DynamicsColumn).Interior.Color = 255
    End If
    '------------------------------------------------
    
    
    
    
    '------------------------------------------------
    '--Закрыть и переместить------
    Application.Calculation = xlCalculationAutomatic
    MainWk.Close True
    
    Dim NewMainFiles As String
    
    NewMainFiles = Left(NameMainFiles, 24) + FullNewDate + xlsx
    
    FileCopy (MainFilePath + NameMainFiles), (DestinationMainFilePath + NameMainFiles)
    Name (DestinationMainFilePath + NameMainFiles) As (DestinationMainFilePath + NewMainFiles)
    FileCopy (DestinationMainFilePath + NewMainFiles), (DestinationMainFilePathDesktop + NewMainFiles)
    
    '------------------------------------------------
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
