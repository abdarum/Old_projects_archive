Attribute VB_Name = "Fiz_3_3_wykresy"
Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro1 Makro
'

'
End Sub

Sub Wypelnij_Do_Poczatku()
'
' Wypelnij_Do_Poczatku Makro
'

'
    ActiveCell.Select
    Selection.AutoFill Destination:=Range("C2:" & ActiveCell.Address), Type:=xlFillDefault
    
End Sub
Sub Nowy_Wiersz_Konduktancja()
'
' Nowy_Wiersz_Konduktancja Makro
'
' Klawisz skrótu: Ctrl+q
'

Dim LastRow As Integer
Dim r As Long


    '''     !!! Statyczne Dane !!!

    Rows("1:1").EntireRow.Insert
    Range("A1:B1") = Array("Napiêcie [mV]", "Natê¿enie [mA]")
    Range("J1:L1") = Array("B³¹d X", "B³¹d Y", "Kondunktancja Wyjœciowa")
    Columns("A:B").ColumnWidth = 14
    Columns("J").ColumnWidth = 10
    Columns("K").ColumnWidth = 10
    Columns("L").ColumnWidth = 23
    Range("N4") = "y = c5 * x ^ 5 + c4 * x ^ 4 + c3 * x ^ 3 + c2 * x ^ 2 + c1 * x + b"

    Range("M5:M10") = Application.Transpose(Array("c5:", "c4:", "c3:", "c2: ", "c1: ", "b: "))


   '''      !!! Dynamiczne dane !!!
   LastRow = ActiveSheet.UsedRange.Rows.Count
   
   ' Obliczanie masymalnej i minimalnej wartoœci komórek
   Dim MIN_X As Integer
   Dim MAX_X As Integer
   Dim MIN_Y As Integer
   Dim MAX_Y As Integer
   MIN_X = WorksheetFunction.Min(ActiveSheet.Range("A2:A" & LastRow))
   MAX_X = WorksheetFunction.Max(ActiveSheet.Range("A2:A" & LastRow))
   MIN_Y = WorksheetFunction.Min(ActiveSheet.Range("B2:B" & LastRow))
   MAX_Y = WorksheetFunction.Max(ActiveSheet.Range("B2:B" & LastRow))

   ' Przypisanie tych wartoœci
   Range("D1").Value = LastRow
   Range("E1").Value = MIN_X
   Range("E2").Value = MAX_X
   Range("F1").Value = MIN_Y
   Range("F2").Value = MAX_Y


   ' Adresy Potrzebne do przybli¿ania
   Dim rY As Range, rX As Range
     Set rX = Range("A2:A" & LastRow)
     Set rY = Range("B2:B" & LastRow)
     
   ' Przybli¿enie funkcji wielomianem 5 stopnia
   Range("N5").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1)"
   Range("N6").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;2)"
   Range("N7").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;3)"
   Range("N8").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;4)"
   Range("N9").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;5)"
   Range("N10").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;6)"

    ' Pochodna tego wielomianu
   Range("O5").FormulaLocal = "=5*N5"
   Range("O6").FormulaLocal = "=4*N6"
   Range("O7").FormulaLocal = "=3*N7"
   Range("O8").FormulaLocal = "=2*N8"
   Range("O9").FormulaLocal = "=N9"
   Range("O10").FormulaLocal = "=0"
   
   ' Przypisanie formu³y obliczania b³êdów
   Range("J2:J" & LastRow).FormulaLocal = "=0,05%*A2+3"
   Range("K2:K" & LastRow).FormulaLocal = "=0,5%*B2+0,03"
   Range("L2:L" & LastRow).FormulaLocal = "=$O$5*A2^4+$O$6*A2^3+$O$7*A2^2+$O$8*A2+$O$9"


End Sub

Sub Nowy_Wiersz_Transonduktancja()
'
' Nowy_Wiersz_Konduktancja Makro
'
' Klawisz skrótu: Ctrl+e
'

Dim LastRow As Integer
Dim r As Long


    '''     !!! Statyczne Dane !!!

    Rows("1:1").EntireRow.Insert
    Range("A1:B1") = Array("Napiêcie [mV]", "Natê¿enie [mA]")
    Range("J1:L1") = Array("B³¹d X", "B³¹d Y", "Transkondunktancja")
    Columns("A:B").ColumnWidth = 14
    Columns("J").ColumnWidth = 10
    Columns("K").ColumnWidth = 10
    Columns("L").ColumnWidth = 17
    Range("N4") = "y = c5 * x ^ 5 + c4 * x ^ 4 + c3 * x ^ 3 + c2 * x ^ 2 + c1 * x + b"

    Range("M5:M10") = Application.Transpose(Array("c5:", "c4:", "c3:", "c2: ", "c1: ", "b: "))


   '''      !!! Dynamiczne dane !!!
   LastRow = ActiveSheet.UsedRange.Rows.Count
   
   ' Obliczanie masymalnej i minimalnej wartoœci komórek
   Dim MIN_X As Integer
   Dim MAX_X As Integer
   Dim MIN_Y As Integer
   Dim MAX_Y As Integer
   MIN_X = WorksheetFunction.Min(ActiveSheet.Range("A2:A" & LastRow))
   MAX_X = WorksheetFunction.Max(ActiveSheet.Range("A2:A" & LastRow))
   MIN_Y = WorksheetFunction.Min(ActiveSheet.Range("B2:B" & LastRow))
   MAX_Y = WorksheetFunction.Max(ActiveSheet.Range("B2:B" & LastRow))

   ' Przypisanie tych wartoœci
   Range("D1").Value = LastRow
   Range("E1").Value = MIN_X
   Range("E2").Value = MAX_X
   Range("F1").Value = MIN_Y
   Range("F2").Value = MAX_Y


   ' Adresy Potrzebne do przybli¿ania
   Dim rY As Range, rX As Range
     Set rX = Range("A2:A" & LastRow)
     Set rY = Range("B2:B" & LastRow)
     
   ' Przybli¿enie funkcji wielomianem 5 stopnia
   Range("N5").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1)"
   Range("N6").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;2)"
   Range("N7").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;3)"
   Range("N8").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;4)"
   Range("N9").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;5)"
   Range("N10").FormulaLocal = "=INDEKS(REGLINP(" & rY.Address & "; " & rX.Address & "^{1;2;3;4;5});1;6)"

    ' Pochodna tego wielomianu
   Range("O5").FormulaLocal = "=5*N5"
   Range("O6").FormulaLocal = "=4*N6"
   Range("O7").FormulaLocal = "=3*N7"
   Range("O8").FormulaLocal = "=2*N8"
   Range("O9").FormulaLocal = "=N9"
   Range("O10").FormulaLocal = "=0"
   
   ' Przypisanie formu³y obliczania b³êdów
   Range("J2:J" & LastRow).FormulaLocal = "=0,05%*A2+3"
   Range("K2:K" & LastRow).FormulaLocal = "=0,5%*B2+0,03"
   Range("L2:L" & LastRow).FormulaLocal = "=$O$5*A2^4+$O$6*A2^3+$O$7*A2^2+$O$8*A2+$O$9"

    

End Sub
Sub Rysuj_Wykres_Tranzystor()
'
' Rysuj_Wykres_Tranzystor Makro
'
' Klawisz skrótu: Ctrl+w
'
 '   ActiveSheet.Shapes.AddChart.Select
 '   ActiveChart.SetSourceData Source:=Range("'2. -0.3V'!$J$1:$O$184")
 '   ActiveChart.ChartType = xlXYScatter
 '   ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
 '   ActiveSheet.ChartObjects("Wykres 1").Activate
 '   ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
 '   ActiveSheet.ChartObjects("Wykres 1").Activate
 '   ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "I [mA]"
 '   ActiveSheet.ChartObjects("Wykres 1").Activate
 '   Selection.Delete
 Dim CH_Name As String
 Dim LastRow As Integer
 Dim LastRow_Apr As Integer
 Dim LastElem As Integer
 Dim Last_AprElem As Integer
 Dim Apr_Forward As Integer
 With ActiveSheet
  LastRow = .UsedRange.Rows.Count
  LastRow_Apr = .Cells(.Rows.Count, "C").End(xlUp).Row
'  LastElem = .Range("A" & LastRow).Value2
'  Last_AprElem = .Range("C" & LastRow_Apr).Value
'Apr_Forward = .Range("A" & LastRow) - .Range("C" & LastRow_Apr)
 'Apr_Forward = .Range("C" & LastRow_Apr).Value - .Range("A" & LastRow).Value
'  Apr_Forward = LastElem / 2
Apr_Forward = 0

 End With
'LastRow_Apr = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row
' LastRow_Apr = ActiveSheet.Range("A1", sh.Range("A1").End(xlDown)).Rows.Count
' LastRow_Apr = ActiveSheet.Range("C" & Rows.Count).End(xlUp).Rows.Count
'  LastRow_Apr = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
' LastRow_Apr = ActiveSheet.Cells(Rows.Count, 3).End(xlDown).Rows

 CH_Name = ActiveSheet.Range("C1").Value

     ActiveSheet.Shapes.AddChart.Select
     ActiveChart.Parent.Name = CH_Name
     With ActiveChart
      '  .SetSourceData Source:=Range("A2:B" & LastRow)
        .HasTitle = True
        .ChartTitle.Characters.Text = CH_Name
'        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
'        .SetElement (msoElementPrimaryValueAxisTitleRotated)
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "U [mV]"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "I [mA]"
        .ChartType = xlXYScatter
        .HasLegend = False
       '        .SetSourceData Source:=Sheets(1).Range("A2")
         
        With .SeriesCollection.NewSeries
            .Name = "Aproksymacja_Lini¹"
            .XValues = Range("C2:C" & LastRow)
            .Values = Range("B2:B" & LastRow)
            .Border.LineStyle = xlNone
            .MarkerStyle = xlNone
            .MarkerBackgroundColorIndex = xlColorIndexNone
            .Format.Fill.Visible = msoFals
            .MarkerStyle = xlMarkerStyleNone
          ' .MarkerSize = 2
          
           With .Trendlines.Add
           .Type = xlLinear
'           .Forward = Apr_Forward
'           .Backward = 100
           .Format.Line.Weight = 3
           '.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
           .Border.ColorIndex = 1
            End With
        End With
  
'        .ChartObjects(CH_Name).SeriesCollection(1).Format.Line.Visible = msoFalse
    '  .SeriesCollection("Aproksymacja_Lini¹").Format.Fill.Visible = msoFalse
        With .SeriesCollection.NewSeries
            .Name = "Dane z b³êdami"
            .XValues = ActiveSheet.Range("A2:A" & LastRow)
            .Values = ActiveSheet.Range("B2:B" & LastRow)
            .MarkerStyle = xlMarkerStyleCircle
            .ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeBoth, _
                Type:=xlErrorBarTypeCustom, Amount:=ActiveSheet.Range("K2:K" & LastRow), MinusValues:=ActiveSheet.Range("K2:K" & LastRow)
            .ErrorBar Direction:=xlX, Include:=xlErrorBarIncludeBoth, _
                Type:=xlErrorBarTypeCustom, Amount:=ActiveSheet.Range("J2:J" & LastRow), MinusValues:=ActiveSheet.Range("J2:J" & LastRow)
          ' .MarkerBackgroundColorIndex = 3
          ' .MarkerForegroundColorIndex = 3
            .MarkerBackgroundColor = RGB(100, 200, 0)
            .MarkerForegroundColor = RGB(100, 200, 0)
        End With
        
   
     End With
 
     Dim n As Long
     With ActiveChart
        For n = .SeriesCollection.Count - 2 To 1 Step -1
          .SeriesCollection(n).Delete
        Next n
     End With
     ActiveSheet.ChartObjects(CH_Name).Activate

End Sub

