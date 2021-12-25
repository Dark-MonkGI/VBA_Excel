Sub MainCode()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
    '������� ���������� � ������� ***
    
Dim data As New Collection ' ������


Dim folderName, variable, stations As String
folderName = "folderName"
variable = "variable"
stations = "stations"
    

Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("*** ***", "***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***", "***", "***", "***", "***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***") _
        )
        
Call addPropInObject( _
                    data, _
                    "***\", _
                    "������� � ���������� ***", _
                    Array("***", "***") _
        )



    
        
    '����������� ��������� ���������� ��� �������� �����
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

    If Not isFindBrach Then MsgBox ("����������� �����.(�������� �� ������ ��� ����� ��� ����������)")
        
  
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ActiveWorkbook.Save
    
End Sub

Sub addPropInObject(data As Collection, folderName As String, variable As String, arr As Variant)

        Dim branch As Object
        Dim branchStations As New Collection
        
        
        Set branch = CreateObject("Scripting.Dictionary") '������� ������
        
        For Each station In arr
            branchStations.Add station
        Next station
        
        branch.Add "folderName", folderName
        branch.Add "variable", variable
        branch.Add "stations", branchStations
        data.Add branch
End Sub
  
Sub WorkingPartOfTheCode(TempVariable As String, folderName As String, NameStations As Variant, branchFormat As String)
            '3 ������� �������������� ���� (������ ������ ����) � ���������� ���������� - ����(����� ���� ������������ ������)
        Dim NewData As Date

        NewData = GenerateDate()
   
        '4 ��������� ����������� �� �������� � �����
        ' ������ ��������� ��������� 1- ��� ������,2-�������� �����,3-����, 4-��� ***-��������� ������ ��������� ���������
        '��������� ����������� ������� ���, ������� *** � *** ������ ��� ����������� � ��� ����� ***
        
        For Each NameStation In NameStations
            Dim TempNameStation As String
            TempNameStation = NameStation ' ����� ������ � ByRef
            Call CopyPastXD(TempVariable, folderName, NewData, TempNameStation, branchFormat) '��� ��� �������� ������ ***!
        Next NameStation

        
        '5 ������ ���������� �����
        '���������� ��������������
        '���� ��� ��������� ����� �� ����� TempVariable, �� ������������
        If ActiveWorkbook.Name <> TempVariable Then Windows(TempVariable + branchFormat).Activate
        Dim TemplastRow As Integer
        TemplastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        Range("A7:M" + CStr(TemplastRow)).Select
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
        
        '�����
        Range("A7:M" + CStr(TemplastRow)).Borders.LineStyle = True
        Range("C7:M" + CStr(TemplastRow)).Select
        '��������
        Selection.HorizontalAlignment = xlCenter
        
        '���������
        ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort.SortFields. _
                                    Add Key:=Range("J7:J" + CStr(TemplastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
                                                    DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort.SortFields. _
                                    Add Key:=Range("B7:B" + CStr(TemplastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
                                                    DataOption:=xlSortNormal
        
        With ActiveWorkbook.Worksheets("������� ���������").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        Range("A1").Select
        
        
                '��� ��� ��� �� �������� �� ������ ��������
        
        Worksheets("������� � �����������").Activate
        Dim NewlastRow As Integer
        NewlastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row '��������� ������
        
        Range("A8:B" + CStr(NewlastRow)).Select
        Selection.ClearContents
        
        Range("F8:F" + CStr(NewlastRow)).Select
        Selection.ClearContents
        
        Worksheets("������� ��������� �2").Activate
        Range("A7:B" + CStr(TemplastRow)).Select
        Selection.Copy
        
        Worksheets("������� � �����������").Activate
        Range("A8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
        
        Worksheets("������� ��������� �2").Activate
        Range("H7:H" + CStr(TemplastRow)).Select
        Selection.Copy
        
        Worksheets("������� � �����������").Activate
        Range("F8").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                     :=False, Transpose:=False
        
        '������� ���������
        Range("A8").Select
        ActiveSheet.Range("$A$7:$AI$" + CStr(TemplastRow + 1)).RemoveDuplicates Columns:=Array(1, 2), _
        Header:=xlNo '��� ����������� ���������� � �������� ����� ������ ������� ���� ������(������� ����� �������)
        
        NewlastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row '��������� ������
        
        ' ��� ���� ��������� ������ �� ��
        '��������� ������� ����� ���� ���� ��������� ��
        
        Dim directoryV As String
        Dim directoryMI As String
        
        Dim FileSvodTI As String
        Dim FileSlovar As String
        Dim NameFileMI As String
        
        Dim CountNuber As Integer
        
        directoryV = "U:\***\"
        directoryMI = "U:\***\2021\���������� (����� �������)\"
        
        FileSvodTI = "������� � ��"
        FileSlovar = "������� ������"
        NameFileMI = Left(folderName, Len(folderName) - 1) '������� / � ����� � �������� �����
                
        CountNuber = 0
        
        Dim xlsx, xls, FormatFileSvodTI, FormatFileSlovar, FormatNameFileMI As String
        xlsx = ".xlsx"
        xls = ".xls"
        
        If (Dir(directoryV + FileSvodTI + xlsx) = "") And (Dir(directoryV + FileSvodTI + xls) = "") Then
            MsgBox "���� " + FileSvodTI + " �� ����������"
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
            MsgBox "���� " + FileSlovar + " �� ����������"
        ElseIf (Dir(directoryV + FileSlovar + xlsx) = "") Then
            FormatFileSlovar = xls
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryV + FileSlovar + xls
        ElseIf (Dir(directoryV + FileSlovar + xls) = "") Then
            FormatFileSlovar = xlsx
            CountNuber = CountNuber + 1
            Workbooks.Open Filename:=directoryV + FileSlovar + xlsx
        End If
        
        
        ' ������ �������� � �������� �������� �� ��
        
        
        If (Dir(directoryMI + NameFileMI + xlsx) = "") And (Dir(directoryMI + NameFileMI + xls) = "") Then
            MsgBox "���� " + NameFileMI + " �� ����������"
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
            MsgBox ("������� �� ��� ��� �����, ����������� ��� �������!")
            Exit Sub
        End If
        MsgBox ("��� ��� �����, ����������� ��� ������� - �������!")
        
        '���� ��� ���������� ����� �� ����� TempVariable �� ������������� �� ���. � ���� ������ ������ �� ���������
        If ActiveWorkbook.Name <> TempVariable Then Windows(TempVariable + branchFormat).Activate
        
        
        
        
        ' ����������� �������
        Dim ArrColumn() As Variant
        
        ArrColumn = Array("C", "D", "E", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", _
                "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI")
                
        For i = LBound(ArrColumn) To UBound(ArrColumn)
            StretchColumn (ArrColumn(i))
        Next i
        
        ' ��������� ������� ����ssss
        
        
        Dim ArrColumn2() As Variant
        
        ArrColumn2 = Array("M", "N", "O", "P", "Q", "S", "T", "U", "V", "W", _
                "Y", "Z", "AA", "AB", "AC", "AE", "AF", "AG", "AH", "AI")
                
        For i = LBound(ArrColumn2) To UBound(ArrColumn2)
            SummColumn (ArrColumn2(i))
        Next i
        
        ' ��������� ������� ��������
        Dim ArrColumn3() As Variant
     
        ArrColumn3 = Array("I", "J")
                
        For i = LBound(ArrColumn3) To UBound(ArrColumn3)
            CountIFColumn (ArrColumn3(i))
        Next i
        
        
        
        '�������������� �����
        
        'MsgBox (NewlastRow) ����� ������������
        

        Range("A8:K" + CStr(NewlastRow)).Select
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
        
        '�����
        Range("A8:K" + CStr(NewlastRow)).Borders.LineStyle = True
        '��������
        Range("C8:K" + CStr(NewlastRow)).Select
        Selection.HorizontalAlignment = xlCenter
  
        
        '������� ��� ������� � ��������� �����
        MsgBox ("����� ���������� ������.")
        Application.Calculation = xlCalculationAutomatic
        
        
         '��� ���� ��������  ����������!
         
         
        ActiveWorkbook.Worksheets("������� � �����������").AutoFilter.Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("������� � �����������").AutoFilter.Sort.SortFields. _
            Add Key:=Range("I8:I" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlDescending, _
            DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("������� � �����������").AutoFilter.Sort.SortFields. _
            Add Key:=Range("J8:J" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("������� � �����������").AutoFilter.Sort.SortFields. _
            Add Key:=Range("H8:H" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlDescending, _
            DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("������� � �����������").AutoFilter.Sort.SortFields. _
            Add Key:=Range("C8:C" + CStr(NewlastRow)), SortOn:=xlSortOnValues, Order:=xlDescending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("������� � �����������").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
         
        Calculate '��������
        MsgBox ("������ �������.")
 
         
        Workbooks(FileSvodTI + FormatFileSvodTI).Close False
        Workbooks(FileSlovar + FormatFileSlovar).Close False
        Workbooks(NameFileMI + FormatNameFileMI).Close False
        
        '���� ��� ��������� ����� �� ����� TempVariable, �� ������������
        If ActiveWorkbook.Name <> TempVariable Then Windows(TempVariable + branchFormat).Activate
        Range("A8").Select
        
        
        
End Sub
Sub CountIFColumn(NameColumn As String)
    '����������� ��������
    Range(NameColumn + "7").Select
    Dim Row As Integer
    
    Row = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
    Row = Row - 7

    ActiveCell.FormulaR1C1 = "=COUNTIF(R[1]C:R[" + CStr(Row) + "]C,TRUE)"
    
End Sub
Sub SummColumn(NameColumn As String)
    
    '����������� ��������� �� ���� ��������
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
            
        '����������� �������
        TemplastRow = 8
        TemplastRow2 = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        
        Range(NameColumn + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range(NameColumn + CStr(TemplastRow) + ":" + NameColumn + CStr(TemplastRow2))
        
End Sub

Sub CopyPastXD(ReportName As String, folderName As String, NewData As Date, NameStation As String, branchFormat As String)
                ' ��� ������               �������� �����        ����� ����          ��� ***   ���������� ����� � �����������
        
        '��������� ������� ����� ���� ���� ��������� ��
        Dim directory As String
        directory = "U:\***\���������� �������� ����� �������\"
        
        Dim xlsx, xls, stationFormat As String
        xlsx = ".xlsx"
        xls = ".xls"
        
        If (Dir(directory + folderName + NameStation + xlsx) = "") And (Dir(directory + folderName + NameStation + xls) = "") Then
            MsgBox "���� " + NameStation + " �� ����������"
            Exit Sub
        ElseIf (Dir(directory + folderName + NameStation + xlsx) = "") Then
            stationFormat = xls
            Workbooks.Open Filename:=directory + folderName + NameStation + xls
        ElseIf (Dir(directory + folderName + NameStation + xls) = "") Then
            stationFormat = xlsx
            Workbooks.Open Filename:=directory + folderName + NameStation + xlsx
        End If
        
        '��� ��� �� ����������� �� ����� � ����
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
            MsgBox ("����:" + NameStation + " - ����. ���������� ���������.")
            Exit Sub
        End If
        
        '����� ����������� ������� ������
        
        '����������� ������� G
        TemplastRow = ActiveSheet.Range("G" & ActiveSheet.Rows.count).End(xlUp).Row
        TemplastRow2 = ActiveSheet.Range("A" & ActiveSheet.Rows.count).End(xlUp).Row
        
        Range("G" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("G" + CStr(TemplastRow) + ":G" + CStr(TemplastRow2))
        
        '����������� ������� H
        Dim TempNameStation As String
        TempNameStation = NameStation '��� ���� ��������� ���������� ���� ��� ����������
        'TempNameStation = Left(TempNameStation, Len(TempNameStation) - 5) '������� ���������� � ����� � �������� �����
        Range("H" + CStr(TemplastRow + 1)).Value = TempNameStation '�������� ��������
        If (TemplastRow + 1) <> TemplastRow2 Then
            Range("H" + CStr(TemplastRow + 1)).Select
            Selection.AutoFill Destination:=Range("H" + CStr(TemplastRow + 1) + ":H" + CStr(TemplastRow2)), Type:=xlFillCopy
        End If
        Range("H" + CStr(TemplastRow + 1) + ":H" + CStr(TemplastRow2)).Select
        
        '����� ��� ����� �����������
        '�������� � ������������� �����
        Selection.HorizontalAlignment = xlCenter
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
    
        '�����
        Range("H" + CStr(TemplastRow + 1) + ":H" + CStr(TemplastRow2)).Borders.LineStyle = True
        
        '����������� ������� I
        
        Range("I" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("I" + CStr(TemplastRow) + ":I" + CStr(TemplastRow2))
        
        '����������� ������� J
        
        Range("J" + CStr(TemplastRow + 1)).Value = NewData '�������� ��������
        
        If (TemplastRow + 1) <> TemplastRow2 Then
            Range("J" + CStr(TemplastRow + 1)).Select
            Selection.AutoFill Destination:=Range("J" + CStr(TemplastRow + 1) + ":J" + CStr(TemplastRow2)), Type:=xlFillCopy
        End If
        
        Range("J" + CStr(TemplastRow + 1) + ":J" + CStr(TemplastRow2)).Select
        
        '����� ��� ����� �����������
        '�������� � ������������� �����
        Selection.HorizontalAlignment = xlCenter
        
        With Selection.Font
            .Name = "Tahoma"
            .Size = 11
        End With
        
        '�����
        Range("J" + CStr(TemplastRow + 1) + ":J" + CStr(TemplastRow2)).Borders.LineStyle = True
        
        
        '����������� ������� K
        
        Range("K" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("K" + CStr(TemplastRow) + ":K" + CStr(TemplastRow2))
        
        '����������� ������� L
        
        Range("L" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("L" + CStr(TemplastRow) + ":L" + CStr(TemplastRow2))
        
        
        '����������� ������� M
        
        Range("M" + CStr(TemplastRow)).Select
        Selection.AutoFill Destination:=Range("M" + CStr(TemplastRow) + ":M" + CStr(TemplastRow2))
        
        MsgBox ("����������: " + NameStation)
        '� �������� ��� ���������� � ������ ��������
        Workbooks(NameStation + stationFormat).Close False
        
End Sub


Function GenerateDate() As Date
    
    '�������� ������� �����
    Dim sheetName As String
    sheetName = "������� ��������� �2"

    '���� ��� ��������� ����� �� ����� sheetName, �� ������������� �� ���� � ��������� sheetName. � ���� ������ ������ �� ���������
    If ActiveSheet.Name <> sheetName Then Worksheets("������� ��������� ").Activate

    '���� ������� ������ �� ��������, �� �������� �� �� 6 ������. (6 ������ - ������� ��������)
    If Not ActiveSheet.AutoFilterMode Then
        Rows(6).Select
        Selection.AutoFilter
    End If

    '���� ���� �������� �������, �� �������� ��. ���� ����, ������ �� ���������
    If ActiveSheet.AutoFilter.FilterMode Then ActiveSheet.ShowAllData

    '���������� ������ ��������� ������. J7 - ������ ��������. ����� ������ �����
    Dim initialColumn As String
    Dim initialColumnB As String
    Dim initialRow As Integer
    Dim initialCell As String

    initialRow = 7
    initialColumn = "J"
    initialColumnB = "B"
    initialCell = initialColumn + CStr(initialRow)
    initialColumnNumber = Range(initialCell).Column

    '��������� �������� ������ � ������ ��� ������ while � for
    Dim tempRow As Integer
    Dim tempCell As String

    tempRow = initialRow
    tempCell = initialCell

    '���� while ���� ������ ���������� ������ ������ � �������
    While Range(tempCell).Value <> ""
        tempRow = tempRow + 1
        tempCell = initialColumn + CStr(tempRow)
    Wend

    Dim lastRow As Integer
    Dim lastCell As String

    '��������� ��������� ������ � ������
    lastRow = tempRow - 1
    lastCell = initialColumn + CStr(lastRow)

    '���������� ����������� ����
    Dim minDate As Date

    '����� � ���������� ����������� ����
    For Row = initialRow To lastRow
        tempCell = initialColumn + CStr(Row)
        If (Row = initialRow) Then minDate = Range(initialCell).Value
        If ((Row <> initialRow) And (Range(tempCell).Value < minDate)) Then minDate = Range(tempCell).Value
    Next Row

    '���������� ���������� ��������� ��� ����������� �����������. ���� �� ������� ���������� ���������� ���������, ��
    ' ��� ���������� Union ����� ��������
    Dim tempRange As Range
    'Set tempRange = Range(CStr(initialRow) + ":" + CStr(initialRow))

    '���� ������ ������ � ����������� �����
    For Row = initialRow To lastRow
        tempCell = initialColumn + CStr(Row)
        If (Range(tempCell).Value = minDate) Then
            Set tempRange = Range(CStr(Row) + ":" + CStr(Row))
            Exit For
        End If
    Next Row

    '����� ���� ����� ���������� ����������� ���� � ����������� �� � ���� ��������

    For Row = initialRow To lastRow
        tempCell = initialColumn + CStr(Row)
        If (Range(tempCell).Value = minDate) Then
            Set tempRange = Union(Range(CStr(Row) + ":" + CStr(Row)), tempRange)
        End If
    Next Row

    tempRange.Select
    Selection.Delete Shift:=xlUp

    '����� ���� ��������� ������ ��� ����������

    '��������� �������� ������ � ������ ��� ������ while � for
    Dim tempRow2 As Integer
    Dim tempCell2 As String

    tempRow2 = initialRow '7
    tempCell2 = initialCell 'J7  'initialColumn = "J"
                               'initialColumnB = "B"
                               'initialCell = initialColumn + CStr(initialRow)

    '���� while ���� ������ ���������� ������ ������ � �������
    While Range(tempCell2).Value <> ""
        tempRow2 = tempRow2 + 1
        tempCell2 = initialColumn + CStr(tempRow2) 'J + �������� �������� ���������� tempRow2
    Wend

    Dim lastRow2 As Integer
    Dim lastCellJ2 As String
    Dim lastCellB2 As String

    '��������� ��������� ������ � ������
    lastRow2 = tempRow2 - 1
    lastCellJ2 = initialColumn + CStr(lastRow2)
    lastCellB2 = initialColumnB + CStr(lastRow2)

    '���������

    ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort.SortFields. _
        Add Key:=Range("J7:" + lastCellJ2), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort.SortFields. _
        Add Key:=Range("B7:" + lastCellB2), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("������� ��������� �2").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A" + CStr(tempRow2)).Select
      
    GenerateDate = (Range("J" + CStr(tempRow2 - 1)).Value) + 7 '������� ����� ����
    
    
End Function




