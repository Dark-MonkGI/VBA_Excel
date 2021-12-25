Sub Удаляем_столбцы()

'
' Удаляем_столбцы Макрос
'
' Сочетание клавиш: Ctrl+й
'
    Range("E:E,G:G,K:K,L:L,M:M,N:N,O:O,P:P,R:R,S:S,T:T,U:U,V:V,W:W,Z:Z,AA:AA,AF:AF,AG:AG,AH:AH"). _
        Select
    Range("AF1").Activate
    Selection.Delete Shift:=xlToLeft
    
' Обьеденяем заголовки
    Range("A1:A3").Select
    Selection.Merge
    Range("B1:B3").Select
    Selection.Merge
    Range("C1:C3").Select
    Selection.Merge
    Range("D1:D3").Select
    Selection.Merge
    Range("E1:E3").Select
    Selection.Merge
    Range("F2:F3").Select
    Selection.Merge
    Range("G2:G3").Select
    Selection.Merge
    Range("H2:H3").Select
    Selection.Merge
    Range("F1:H1").Select
    Selection.Merge
    Range("I2").Select
    Selection.ClearContents
    Range("I1:I3").Select
    Selection.Merge
    Range("J1:K1").Select
    Selection.Merge
    Range("J2:K2").Select
    Selection.Merge
    Range("L1:M1").Select
    Selection.Merge
    Range("L2:M2").Select
    Selection.Merge
    Range("N1:O1").Select
    Selection.Merge
    Range("N2:N3").Select
    Selection.Merge
    Range("O2:O3").Select
    Selection.Merge
    
'Записываем заголовки
    
    Range("J1:K1").Select
    ActiveCell.FormulaR1C1 = "Отключение"
    Range("L1:M1").Select
    ActiveCell.FormulaR1C1 = "Включение"
    Range("L2").Select
    
'Создаем еще столбец "Количество зданий"
    
    Range("E1:E3").Select
    Selection.Copy
    Range("P1:P3").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Range("P1:P3").Select
    ActiveCell.FormulaR1C1 = "Количество зданий"
    
' Ширина колонок

    Columns("A:A").ColumnWidth = 13
    Columns("B:B").ColumnWidth = 11.75
    Columns("C:C").ColumnWidth = 11.75
    Columns("D:D").ColumnWidth = 9.38
    Columns("E:E").ColumnWidth = 9.63
    Columns("I:I").ColumnWidth = 20.25
    
    Columns("J:K").Select
    Range("J3").Activate
    Selection.ColumnWidth = 9.5
    
    Columns("L:M").Select
    Range("L2").Activate
    Selection.ColumnWidth = 9.38
      
    Columns("N:O").Select
    Range("N4").Activate
    Selection.ColumnWidth = 16.75

    Columns("P:P").ColumnWidth = 17.38
    
' Выделяем заголовки

Range("A1:P3").Select
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
    

'Центруем заголовки


    Range("A1:P3").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With

'Вводим формулу счетесли для подсчета домов

    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-15],RC[-15])"
'Протягиваем формулу

    Range("P4").Select
    Selection.AutoFill Destination:=Range("P4:P1477") 'тут надо изменить последнее значение на количесво нужных строк
    Range("P4:P1477").Select
    
    
'Копируем и вставляем как значение


    Columns("P:P").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

'Удаляем дубли в первом столбце


    ActiveSheet.Range("$A$3:$P$1477").RemoveDuplicates Columns:=1, Header:= _
        xlYes                                                  'Необходимо заменить количество строк
        

        
'Форматируем всю таблицу

    Range("A4:P4").Select
    Range(Selection, Selection.End(xlDown)).Select
    
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
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    

    Range("A4").Select

'Сейвимся и все

    ActiveWorkbook.Save



End Sub
