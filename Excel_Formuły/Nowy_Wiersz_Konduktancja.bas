Attribute VB_Name = "Nowy_Wiersz_Konduktancja"
Sub Nowy_Wiersz_Konduktancja()
Attribute Nowy_Wiersz_Konduktancja.VB_ProcData.VB_Invoke_Func = "q\n14"
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


  
