Sub MainCode()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
    'Создаем переменные с именами ***
    
Dim data As New Collection ' Массив


Dim folderName, variable, stations As String
folderName = "folderName"
variable = "variable"
stations = "stations"
    

Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("*** ***", "***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "Объекты с нарушением ***", _
                    Array("***", "***") _
        )



    
        
    'Присваиваем временной переменной имя открытой книги
    Dim TempVariable As String
    TempVariable = ActiveWorkbook.Name

    Dim xlsx, xls, branchFormat As String
    xlsx = ".xlsx"
    xls = ".xls"

    If Right(TempVariable, 5) = xlsx Then
        branchFormat = Right(TempVariable, 5)
        TempVariable = Left(TempVariable, Len(TempVariable) - 5)
    ElseIf Right(TempVariable, 4) = xls Then
        branchFormat = Right(TempVariable, 4)
        TempVariable = Left(TempVariable, Len(TempVariable) - 4)
    End If

    Dim isFindBrach As Boolean
    isFindBrach = False
   
   
    For Each branch In data
        If TempVariable = branch.Item(variable) Then
            Call WorkingPartOfTheCode(branch.Item(variable), branch.Item(folderName), branch.Item(stations), branchFormat)
            isFindBrach = True
        End If
    Next branch

    If Not isFindBrach Then MsgBox ("Неизвестная книга.(Возможно не верное имя файла или расширение)")
        
  
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ActiveWorkbook.Save
    
End Sub

Sub addPropInObject(data As Collection, folderName As String, variable As String, arr As Variant)

        Dim branch As Object
        Dim branchStations As New Collection
        
        
        Set branch = CreateObject("Scripting.Dictionary") 'создать обьект
        
        For Each station In arr
            branchStations.Add station
        Next station
        
        branch.Add "folderName", folderName
        branch.Add "variable", variable
        branch.Add "stations", branchStations
        data.Add branch
End Sub
  
Sub WorkingPartOfTheCode(TempVariable As String, folderName As String, NameStations As Variant, branchFormat As String)
            '3 Функция подготавливает фаил (Удаляя старую дату) и возвращает переменную - дата(Новая дата формирования отчета)
        Dim NewData As Date

        NewData = GenerateDate()
   
        '4 Программа копирования из выгрузки в отчет
        ' Данная программа принимает 1- имя отчета,2-название папки,3-дату, 4-имя ***-меняеться только послежняя переменая
        'Программа вызываеться столько раз, сколько *** в *** каждый раз передаеться в нее новая ***
        
        For Each NameStation In NameStations
            Dim TempNameStation As String
            TempNameStation = NameStation ' Обход ошибки с ByRef
            Call CopyPastXD(TempVariable, folderName, NewData, TempNameStation, branchFormat) 'Вот тут вызываем каждую ***!
        Next NameStation

        
        '5 Скрипт сортировки файла
        'применияем форматирование
        'Если Имя активного книги не равно TempVariable, то активировать
        If ActiveWorkbook.Name <> TempVariable Then Windows(TempVariable + branchFormat).Activate
        Dim TemplastRow As Integer
        TemplastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        Range("A7:M" + CStr(TemplastRow)).Select
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
        
        'Рамки
        Range("A7:M" + CStr(TemplastRow)).Borders.LineStyle = True
        Range("C7:M" + CStr(TemplastRow)).Select
        'Центруем
        Selection.HorizontalAlignment = xlCenter
        
        'Сортируем
        ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort.SortFields. _
                                    Add Key:=Range("J7:J" + CStr(TemplastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
                                                    DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort.SortFields. _
                                    Add Key:=Range("B7:B" + CStr(TemplastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
                                                    DataOption:=xlSortNormal
        
        With ActiveWorkbook.Worksheets("объекты нарушения").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        Range("A1").Select
        
        
                'ВОТ ТУТ КОД ПО ПЕРЕНОСУ НА ПЕРВУЮ СТРАНИЦУ
        
        Worksheets("Объекты с нарушениями").Activate
        Dim NewlastRow As Integer
        NewlastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row 'последняя строка
        
        Range("A8:B" + CStr(NewlastRow)).Select
        Selection.ClearContents
        
        Range("F8:F" + CStr(NewlastRow)).Select
        Selection.ClearContents
        
        Worksheets("объекты нарушения Т2").Activate
        Range("A7:B" + CStr(TemplastRow)).Select
        Selection.Copy
        
        Worksheets("Объекты с нарушениями").Activate
        Range("A8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
        
        Worksheets("объекты нарушения Т2").Activate
        Range("H7:H" + CStr(TemplastRow)).Select
        Selection.Copy
        
        Worksheets("Объекты с нарушениями").Activate
        Range("F8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                     :=False, Transpose:=False
        
        'Удаляем дубликаты
        Range("A8").Select
        ActiveSheet.Range("$A$7:$AI$" + CStr(TemplastRow + 1)).RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlNo 'Тут использовал переменную с прошлого листа просто добавив одну строку(Разница между листами)
        
        NewlastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row 'последняя строка
        
        ' Тут надо открывать книжки из МИ
        'Проверяем наличие книги если есть открываем ее
        
        Dim directoryV As String
        Dim directoryMI As String
        
        Dim FileSvodTI As String
        Dim FileSlovar As String
        Dim NameFileMI As String
        
        Dim CountNuber As Integer
        
        directoryV = "U:\***\"
        directoryMI = "U:\***\2021\Актуальная (после макроса)\"
        
        FileSvodTI = "Сводная с ТИ"
        FileSlovar = "Словарь полный"
        NameFileMI = Left(folderName, Len(folderName) - 1) 'Удаляем / с конца в названии папки
                
        CountNuber = 0
        
        Dim xlsx, xls, FormatFileSvodTI, FormatFileSlovar, FormatNameFileMI As String
        xlsx = ".xlsx"
        xls = ".xls"
        
        If (Dir(directoryV + FileSvodTI + xlsx) = "") And (Dir(directoryV + FileSvodTI + xls) = "") Then
            MsgBox "Файл " + FileSvodTI + " не существует"
        ElseIf (Dir(directoryV + FileSvodTI + xlsx) = "") Then
            FormatFileSvodTI = xls
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryV + FileSvodTI + xls
        ElseIf (Dir(directoryV + FileSvodTI + xls) = "") Then
            FormatFileSvodTI = xlsx
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryV + FileSvodTI + xlsx
        End If
        
        
        If (Dir(directoryV + FileSlovar + xlsx) = "") And (Dir(directoryV + FileSlovar + xls) = "") Then
            MsgBox "Файл " + FileSlovar + " не существует"
        ElseIf (Dir(directoryV + FileSlovar + xlsx) = "") Then
            FormatFileSlovar = xls
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryV + FileSlovar + xls
        ElseIf (Dir(directoryV + FileSlovar + xls) = "") Then
            FormatFileSlovar = xlsx
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryV + FileSlovar + xlsx
        End If
        
        
        ' Теперь проверка и открытие выгрузки из МИ
        
        
        If (Dir(directoryMI + NameFileMI + xlsx) = "") And (Dir(directoryMI + NameFileMI + xls) = "") Then
            MsgBox "Файл " + NameFileMI + " не существует"
        ElseIf (Dir(directoryMI + NameFileMI + xlsx) = "") Then
            FormatNameFileMI = xls
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryMI + NameFileMI + xls
        ElseIf (Dir(directoryMI + NameFileMI + xls) = "") Then
            FormatNameFileMI = xlsx
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryMI + NameFileMI + xlsx
        End If
        
        
        If CountNuber <> 3 Then
            MsgBox ("Открыты не все три книги, необходимые для расчета!")
            Exit Sub
        End If
        MsgBox ("Все три книги, необходимые для расчета - открыты!")
        
        'Если Имя активногой книги не равно TempVariable то переключиться на нее. В ином случае ничего не произойдёт
        If ActiveWorkbook.Name <> TempVariable Then Windows(TempVariable + branchFormat).Activate
        
        
        
        
        ' Протягивать столбцы
        Dim ArrColumn() As Variant
        
        ArrColumn = Array("C", "D", "E", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", _
                "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI")
                
        For i = LBound(ArrColumn) To UBound(ArrColumn)
            StretchColumn (ArrColumn(i))
        Next i
        
        ' Вставлять формулы Суммssss
        
        
        Dim ArrColumn2() As Variant
        
        ArrColumn2 = Array("M", "N", "O", "P", "Q", "S", "T", "U", "V", "W", _
                "Y", "Z", "AA", "AB", "AC", "AE", "AF", "AG", "AH", "AI")
                
        For i = LBound(ArrColumn2) To UBound(ArrColumn2)
            SummColumn (ArrColumn2(i))
        Next i
        
        ' Вставлять формулы СчетЕсли
        Dim ArrColumn3() As Variant
     
        ArrColumn3 = Array("I", "J")
                
        For i = LBound(ArrColumn3) To UBound(ArrColumn3)
            CountIFColumn (ArrColumn3(i))
        Next i
        
        
        
        'Форматирование листа
        
        'MsgBox (NewlastRow) Можно использовать
        

        Range("A8:K" + CStr(NewlastRow)).Select
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
        
        'Рамки
        Range("A8:K" + CStr(NewlastRow)).Borders.LineStyle = True
        'Центруем
        Range("C8:K" + CStr(NewlastRow)).Select
        Selection.HorizontalAlignment = xlCenter
  
        
        'Считаем все формулы и закрываем книги
        MsgBox ("Далее перерасчет формул.")
        Application.Calculation = xlCalculationAutomatic
        
        
         'ТУТ НАДО ВСТАВИТЬ  СОРТИРОВКУ!
         
         
        ActiveWorkbook.Worksheets("Объекты с нарушениями").AutoFilter.Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Объекты с нарушениями").AutoFilter.Sort.SortFields. _
            Add Key:=Range("I8:I" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlDescending, _
            DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Объекты с нарушениями").AutoFilter.Sort.SortFields. _
            Add Key:=Range("J8:J" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Объекты с нарушениями").AutoFilter.Sort.SortFields. _
            Add Key:=Range("H8:H" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlDescending, _
            DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Объекты с нарушениями").AutoFilter.Sort.SortFields. _
            Add Key:=Range("C8:C" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlDescending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Объекты с нарушениями").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
         
        Calculate 'Пересчет
        MsgBox ("Расчет окончен.")
 
         
        Workbooks(FileSvodTI + FormatFileSvodTI).Close False
        Workbooks(FileSlovar + FormatFileSlovar).Close False
        Workbooks(NameFileMI + FormatNameFileMI).Close False
        
        'Если Имя активного книги не равно TempVariable, то активировать
        If ActiveWorkbook.Name <> TempVariable Then Windows(TempVariable + branchFormat).Activate
        Range("A8").Select
        
        
        
End Sub
Sub CountIFColumn(NameColumn As String)
    'Протягиваем СчетЕсли
    Range(NameColumn + "7").Select
    Dim Row As Integer
    
    Row = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    Row = Row - 7

    ActiveCell.FormulaR1C1 = "=COUNTIF(R[1]C:R[" + CStr(Row) + "]C,TRUE)"
    
End Sub
Sub SummColumn(NameColumn As String)
    
    'Протягиваем автосумму по всем столбцам
    Range(NameColumn + "6").Select
    Dim Row As Integer
    
    Row = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    Row = Row - 6
    'MsgBox TypeName(Row)
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[" + CStr(Row) + "]C)"
    

End Sub
Sub StretchColumn(NameColumn As String)
        Dim TemplastRow As Integer
        Dim TemplastRow2 As Integer
            
        'Протягиваем колонку
        TemplastRow = 8
        TemplastRow2 = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        
        Range(NameColumn + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range(NameColumn + CStr(TemplastRow) + ":" + NameColumn + CStr(TemplastRow2))
        
End Sub

Sub CopyPastXD(ReportName As String, folderName As String, NewData As Date, NameStation As String, branchFormat As String)
                ' Имя отчета               Название папки        новая дата          имя ***   расширение файла с нарушениями
        
        'Проверяем наличие книги если есть открываем ее
        Dim directory As String
        directory = "U:\***\Актуальная Выгрузка После Макроса\"
        
        Dim xlsx, xls, stationFormat As String
        xlsx = ".xlsx"
        xls = ".xls"
        
        If (Dir(directory + folderName + NameStation + xlsx) = "") And (Dir(directory + folderName + NameStation + xls) = "") Then
            MsgBox "Файл " + NameStation + " не существует"
            Exit Sub
        ElseIf (Dir(directory + folderName + NameStation + xlsx) = "") Then
            stationFormat = xls
            Workbooks.Open Filename:=directory + folderName + NameStation + xls
        ElseIf (Dir(directory + folderName + NameStation + xls) = "") Then
            stationFormat = xlsx
            Workbooks.Open Filename:=directory + folderName + NameStation + xlsx
        End If
        
        'ТУТ КОД ПО КОПИРОВАНИЮ ИЗ ФАЙЛА В ФАЙЛ
        Dim TemplastRow As Integer
        Dim TemplastRow2 As Integer
        
        If Range("A6").Value <> "" Then
            TemplastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
            Range("A6:F" + CStr(TemplastRow)).Select
            Selection.Copy
        
            Windows(ReportName + branchFormat).Activate
            TemplastRow = (ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row) + 1
            Range("A" + CStr(TemplastRow)).Select
            ActiveSheet.Paste
        Else
            MsgBox ("Фаил:" + NameStation + " - пуст. Необходимо проверить.")
            Exit Sub
        End If
        
        'Далее протягивать колонки справа
        
        'Протягиваем колонку G
        TemplastRow = ActiveSheet.Range("G" & ActiveSheet.Rows.count).End(xlUp).Row
        TemplastRow2 = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        
        Range("G" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("G" + CStr(TemplastRow) + ":G" + CStr(TemplastRow2))
        
        'Протягиваем колонку H
        Dim TempNameStation As String
        TempNameStation = NameStation 'это чтоб сохранить переменную пока что неизменной
        'TempNameStation = Left(TempNameStation, Len(TempNameStation) - 5) 'Удаляем расширение с конца в названии файла
        Range("H" + CStr(TemplastRow + 1)).Value = TempNameStation 'Вставили значение
        If (TemplastRow + 1) <> TemplastRow2 Then
            Range("H" + CStr(TemplastRow + 1)).Select
            Selection.AutoFill Destination:=Range("H" + CStr(TemplastRow + 1) + ":H" + CStr(TemplastRow2)), Type:=xlFillCopy
        End If
        Range("H" + CStr(TemplastRow + 1) + ":H" + CStr(TemplastRow2)).Select
        
        'Далее это сразу форматируем
        'Центруем и устанавливаем шрифт
        Selection.HorizontalAlignment = xlCenter
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
    
        'Рамки
        Range("H" + CStr(TemplastRow + 1) + ":H" + CStr(TemplastRow2)).Borders.LineStyle = True
        
        'Протягиваем колонку I
        
        Range("I" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("I" + CStr(TemplastRow) + ":I" + CStr(TemplastRow2))
        
        'Протягиваем колонку J
        
        Range("J" + CStr(TemplastRow + 1)).Value = NewData 'Вставили значение
        
        If (TemplastRow + 1) <> TemplastRow2 Then
            Range("J" + CStr(TemplastRow + 1)).Select
            Selection.AutoFill Destination:=Range("J" + CStr(TemplastRow + 1) + ":J" + CStr(TemplastRow2)), Type:=xlFillCopy
        End If
        
        Range("J" + CStr(TemplastRow + 1) + ":J" + CStr(TemplastRow2)).Select
        
        'Далее это сразу форматируем
        'Центруем и устанавливаем шрифт
        Selection.HorizontalAlignment = xlCenter
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
        
        'Рамки
        Range("J" + CStr(TemplastRow + 1) + ":J" + CStr(TemplastRow2)).Borders.LineStyle = True
        
        
        'Протягиваем колонку K
        
        Range("K" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("K" + CStr(TemplastRow) + ":K" + CStr(TemplastRow2))
        
        'Протягиваем колонку L
        
        Range("L" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("L" + CStr(TemplastRow) + ":L" + CStr(TemplastRow2))
        
        
        'Протягиваем колонку M
        
        Range("M" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("M" + CStr(TemplastRow) + ":M" + CStr(TemplastRow2))
        
        MsgBox ("Скопирован: " + NameStation)
        'И закрыаем без сохранения и лишних вопросов
        Workbooks(NameStation + stationFormat).Close False
        
End Sub


Function GenerateDate() As Date
    
    'Название нужного листа
    Dim sheetName As String
    sheetName = "объекты нарушения Т2"

    'Если Имя активного листа не равно sheetName, то переключиться на лист с названием sheetName. В ином случае ничего не произойдёт
    If ActiveSheet.Name <> sheetName Then Worksheets("объекты нарушения ").Activate

    'Если фильтры вообще не включены, то включить их на 6 строке. (6 строка - жесткая привязка)
    If Not ActiveSheet.AutoFilterMode Then
        Rows(6).Select
        Selection.AutoFilter
    End If

    'Если есть активные фильтры, то очистить их. Если нету, ничего не произойдёт
    If ActiveSheet.AutoFilter.FilterMode Then ActiveSheet.ShowAllData

    'Объявление данных начальной ячейки. J7 - жёсткая привязка. Лучше делать поиск
    Dim initialColumn As String
    Dim initialColumnB As String
    Dim initialRow As Integer
    Dim initialCell As String

    initialRow = 7
    initialColumn = "J"
    initialColumnB = "B"
    initialCell = initialColumn + CStr(initialRow)
    initialColumnNumber = Range(initialCell).Column

    'Временные значения строки и ячейки для циклов while и for
    Dim tempRow As Integer
    Dim tempCell As String

    tempRow = initialRow
    tempCell = initialCell

    'Цикл while ищет первую попавшуюся пустую ячейку в столбце
    While Range(tempCell).Value <> ""
        tempRow = tempRow + 1
        tempCell = initialColumn + CStr(tempRow)
    Wend

    Dim lastRow As Integer
    Dim lastCell As String

    'Получение последней строки и ячейки
    lastRow = tempRow - 1
    lastCell = initialColumn + CStr(lastRow)

    'Объявление минимальной даты
    Dim minDate As Date

    'Поиск и присвоение минимальной даты
    For Row = initialRow To lastRow
        tempCell = initialColumn + CStr(Row)
        If (Row = initialRow) Then minDate = Range(initialCell).Value
        If ((Row <> initialRow) And (Range(tempCell).Value < minDate)) Then minDate = Range(tempCell).Value
    Next Row

    'Объявление начального диапазона для дальнейшего объединения. Если не сделать присвоения начального диапазона, то
    ' при применении Union будет ругаться
    Dim tempRange As Range
    'Set tempRange = Range(CStr(initialRow) + ":" + CStr(initialRow))

    'Ищем первую строку с минимальной датой
    For Row = initialRow To lastRow
        tempCell = initialColumn + CStr(Row)
        If (Range(tempCell).Value = minDate) Then
            Set tempRange = Range(CStr(Row) + ":" + CStr(Row))
            Exit For
        End If
    Next Row

    'Поиск всех строк содержащих минимальную дату и объединение их в один диапазон

    For Row = initialRow To lastRow
        tempCell = initialColumn + CStr(Row)
        If (Range(tempCell).Value = minDate) Then
            Set tempRange = Union(Range(CStr(Row) + ":" + CStr(Row)), tempRange)
        End If
    Next Row

    tempRange.Select
    Selection.Delete Shift:=xlUp

    'Снова ищем последнюю строку для сортировки

    'Временные значения строки и ячейки для циклов while и for
    Dim tempRow2 As Integer
    Dim tempCell2 As String

    tempRow2 = initialRow '7
    tempCell2 = initialCell 'J7  'initialColumn = "J"
                               'initialColumnB = "B"
                               'initialCell = initialColumn + CStr(initialRow)

    'Цикл while ищет первую попавшуюся пустую ячейку в столбце
    While Range(tempCell2).Value <> ""
        tempRow2 = tempRow2 + 1
        tempCell2 = initialColumn + CStr(tempRow2) 'J + строчное значение переменной tempRow2
    Wend

    Dim lastRow2 As Integer
    Dim lastCellJ2 As String
    Dim lastCellB2 As String

    'Получение последней строки и ячейки
    lastRow2 = tempRow2 - 1
    lastCellJ2 = initialColumn + CStr(lastRow2)
    lastCellB2 = initialColumnB + CStr(lastRow2)

    'Сортируем

    ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort.SortFields. _
        Add Key:=Range("J7:" + lastCellJ2), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort.SortFields. _
        Add Key:=Range("B7:" + lastCellB2), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("объекты нарушения Т2").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A" + CStr(tempRow2)).Select
      
    GenerateDate = (Range("J" + CStr(tempRow2 - 1)).Value) + 7 'Вернуть новую дату
    
    
End Function




