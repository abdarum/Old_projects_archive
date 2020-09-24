Attribute VB_Name = "Rysuj_Wykres"
Sub Rysuj_Wykres_Tranzystor()
Attribute Rysuj_Wykres_Tranzystor.VB_ProcData.VB_Invoke_Func = "w\n14"
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



