Sub ???????????????????2()

'
' ?????????????????? ??????
'
'
    '??? ???????? ???????? ???? ????????
    Dim StartDate As Date
    Dim EndDate As Date
    Dim DownloadDate As Date
    Dim DifferenceIsDays As Long
    Dim DifferenceIsDays2 As Long
    
    
    StartDate = Range("E8").Value
    EndDate = Range("L8").Value
    DownloadDate = Range("E4").Value
    
    '????????? ??? ??????
    If Weekday(StartDate, vbMonday) <> 6 Then MsgBox ("?? ?????? ???? ?????? ????????!")
    If Weekday(EndDate, vbMonday) <> 5 Then MsgBox ("?? ?????? ???? ????????? ????????!")
    


    DifferenceIsDays = DateDiff("d", StartDate, DownloadDate)
    DifferenceIsDays2 = DateDiff("d", EndDate, DownloadDate)
    
    '????????? ???????? ????????
    If DifferenceIsDays > 8 Or DifferenceIsDays2 > 2 Then MsgBox ("?? ?????? ?????? ????????")
    


    Dim numberSting As Long    '??? ????? ????? ????? ??????????? ? ???????
    Dim rangeToSearch As Range '???????? ??? ??????
    Dim AddressCell As Range
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
   
    
    numberSting = Cells(Rows.Count, "A").End(xlUp).Row '??? ???????? ????? ?????? ????? ??????? ?

    Set rangeToSearch = Range("A1:G" & numberSting) '???????? ???????? ?????? ? ??????????
    
    
    Set AddressCell = rangeToSearch.Find("IV ???????????, ???????????? ????????")

    
    Rows("1:" & AddressCell.Row).Select
    Selection.Delete Shift:=xlUp
    
    
    
    
    Dim numberSting2 As Long    '??? ????? ????? ????? ??????????? ? ???????
    Dim rangeToSearch2 As Range '???????? ??? ??????
    Dim AddressCell2 As Range
   
    
    numberSting2 = Cells(Rows.Count, "A").End(xlUp).Row '??? ???????? ????? ?????? ????? ??????? ?

    Set rangeToSearch2 = Range("A1:L" & numberSting2) '???????? ???????? ?????? ? ??????????
    
    Set AddressCell2 = rangeToSearch2.Find("????? ?? ????????????, ???????????? ????????:")
    'MsgBox AddressCell2.Cells
    'MsgBox AddressCell2.Address
    'MsgBox AddressCell2.Row '????? ??????
    'entire.row.rows(i) ????????
    
    If AddressCell2 Is Nothing Then
        Rows(1 & ":" & numberSting2).Delete
        Range("G1") = "?????????? ?????? ? ???????? ?????!"
        GoTo TheEnd
    End If
  
    
    Rows(AddressCell2.Row & ":" & numberSting2).Select
    Selection.Delete Shift:=xlUp


' ????? ?????? ??????



    Columns("B:B").Select
    Range("B348").Activate
    Selection.UnMerge
    ActiveWindow.SmallScroll Down:=-129

    Columns("G:G").Select
    Range("G348").Activate
    Selection.UnMerge
    
    Columns("N:N").Select
    Range("N348").Activate
    Selection.UnMerge
    ActiveWindow.SmallScroll Down:=-369
    Columns("P:P").Select
    Range("P348").Activate
    Selection.UnMerge
    ActiveWindow.SmallScroll Down:=-390
    Columns("Y:Y").Select
    Range("Y348").Activate
    Selection.UnMerge
    ActiveWindow.SmallScroll Down:=-366
    Columns("AL:AL").Select
    Range("AL348").Activate
    Selection.UnMerge
  
    Columns("AZ:AZ").Select
    Range("AZ348").Activate
    Selection.UnMerge

    Columns("BH:BH").Select
    Range("BH348").Activate
    Selection.UnMerge

   
    Columns("BC:BS").Select
    Selection.Delete Shift:=xlToLeft
     
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:AI").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:B").Select
    
     
    Selection.Hyperlinks.Delete
    Columns("C:Q").Select
    Columns("C:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Columns("A:A").ColumnWidth = 34.86
    Columns("B:B").ColumnWidth = 35.71
    Columns("C:C").ColumnWidth = 12.14
    Columns("D:D").ColumnWidth = 11.71
    Columns("E:E").ColumnWidth = 12.71
    Columns("F:F").ColumnWidth = 13.43
    
    '??? ??? ??????????? ??????????!
    

    '?????????? ?????? ????????? ??????.
    Dim initialColumnF As String
    Dim initialRow6 As Integer
    Dim initialCellF6 As String

    initialRow6 = 6
    initialColumnF = "F"
    initialCellF6 = initialColumnF + CStr(initialRow6)
    
    
    '????????? ????????? ?????? ? ??????
    Dim lastRow As Integer
    Dim lastCellF As String
    '??????? ????????? ??????, ????? ? ????? ?????!
    lastRow = ActiveSheet.Range("a" & ActiveSheet.Rows.Count).End(xlUp).Row
    'MsgBox (lastRow)
    lastCellF = initialColumnF + CStr(lastRow)
    
    
    
    
    '????????? ???????? ?????? ? ?????? ??? ?????? while ? for
    Dim tempCell As String


    '?????????? ?????????? ????????? ??? ??????????? ???????????. ???? ?? ??????? ?????????? ?????????? ?????????, ??
    ' ??? ?????????? Union ????? ????????
    
    Dim tempRange As Range
    'Set tempRange = Range(CStr(initialRow) + ":" + CStr(initialRow))
    '???? ?????? ?????? ? ??????????? ?????
    For Row = initialRow6 To lastRow
        tempCell = initialColumnF + CStr(Row)
        If (Range(tempCell).Value < 50 Or Range(tempCell).Value = "") Then
            'MsgBox (Range(tempCell).Value)
            Set tempRange = Range(CStr(Row) + ":" + CStr(Row))
            Exit For
        End If
    Next Row

    If tempRange Is Nothing Then GoTo M:
    


    '????? ???? ????? ?????????? ??????????? ???? ? ??????????? ?? ? ???? ????????

    For Row = initialRow6 To lastRow
        tempCell = initialColumnF + CStr(Row)
        If (Range(tempCell).Value < 50 Or Range(tempCell).Value = "") Then
            Set tempRange = Union(Range(CStr(Row) + ":" + CStr(Row)), tempRange)
        End If
    Next Row
    

    
    tempRange.Select
    Selection.Delete Shift:=xlUp


M:
    
    '??????? ????????? ??????, ????? ? ????? ?????!
    lastRow = ActiveSheet.Range("a" & ActiveSheet.Rows.Count).End(xlUp).Row
   
'? ??? ??????????????
        
    Range("A6:F" + CStr(lastRow)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
   
    
    Range("A6:F" + CStr(lastRow)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
   
   
    Range("A6:F" + CStr(lastRow)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
TheEnd:
    
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Range("A1").Select
    ActiveWorkbook.Save
End Sub
