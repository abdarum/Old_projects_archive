Attribute VB_Name = "Module4"
Sub CreateLetter()
'
' CreateLetter Makro
'
    Dim ws As Worksheet
    Dim objWord As Object
    Dim objDoc
    Dim objRange
    Dim objTable
    Dim objRangeTable
    ' Table 1
    Dim intNoOfRows
    Dim intNoOfColumns
    intNoOfRows = 7
    intNoOfColumns = 6
    ' Table 2
    Dim intNoOfRows2
    Dim intNoOfColumns2
    intNoOfRows2 = 3
    intNoOfColumns2 = 6
    ' Table 3
    Dim intNoOfRows3
    Dim intNoOfColumns3
    intNoOfRows3 = 4
    intNoOfColumns3 = 6

    Set ws = ThisWorkbook.Sheets("Podsumowanie")
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    objWord.Activate
    ' Otworz dokument istniejacy na dysku w ustalonej sciezce
    '    objWord.Documents.Open "D:\\PhilipsTask\\test.docx"
    ' Utworz nowy dokument bez lokalizacji na dysku(nalezy go poniej zapisac)
    Set objDoc = objWord.Documents.Add
    objDoc.Paragraphs.Add
    Set objRangeTable = objDoc.Paragraphs.Last.Range
    objDoc.Tables.Add objRangeTable, intNoOfRows, intNoOfColumns
    Set objTable = objDoc.Tables(objDoc.Tables.Count)
    ' objTable.Cell(Row, Column) - objTable.Cell(5, 3) - it means C5 cell - index starts from 1
    objTable.Cell(1, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A2").Value
    objTable.Cell(2, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A3").Value
    objTable.Cell(3, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A4").Value
    objTable.Cell(4, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A5").Value
    objTable.Cell(5, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A6").Value
    
    objTable.Cell(1, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B2").Value
    objTable.Cell(2, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B3").Value
    objTable.Cell(3, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B4").Value
    objTable.Cell(4, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B5").Value
    objTable.Cell(5, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B6").Value
    
    objTable.Cell(1, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D2").Value
    objTable.Cell(2, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D3").Value
    objTable.Cell(3, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D4").Value
    objTable.Cell(4, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D5").Value
    objTable.Cell(5, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D6").Value
    objTable.Cell(6, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D7").Value
    objTable.Cell(7, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D8").Value
    
    objTable.Cell(1, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E2").Value
    objTable.Cell(2, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E3").Value
    objTable.Cell(3, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E4").Value
    objTable.Cell(4, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E5").Value
    objTable.Cell(5, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E6").Value
    objTable.Cell(6, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E7").Value
    objTable.Cell(7, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E8").Value
    

    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A10").Value & " "
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("B10").Value & vbNewLine & vbNewLine
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A12").Value & vbNewLine
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A13").Value & vbNewLine
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A15").Value & vbNewLine
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A25").Value & vbNewLine & vbNewLine
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A26").Value & vbNewLine
    objDoc.Range.InsertAfter ActiveWorkbook.Worksheets("Podsumowanie").Range("A27").Value & vbNewLine



    objDoc.Paragraphs.Add
    Set objRangeTable = objDoc.Paragraphs.Last.Range
    objDoc.Tables.Add objRangeTable, intNoOfRows2, intNoOfColumns2
    Set objTable = objDoc.Tables(objDoc.Tables.Count)
    objTable.Cell(1, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A28").Value & " "
    objTable.Cell(1, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B28").Value & vbNewLine
    objTable.Cell(2, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A29").Value & " "
    objTable.Cell(2, 2).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("B29").Value & vbNewLine
    objTable.Cell(3, 1).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("A30").Value & " "
    
    objTable.Cell(1, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D28").Value & " "
    objTable.Cell(1, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E28").Value & vbNewLine
    objTable.Cell(2, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D29").Value & " "
    objTable.Cell(2, 6).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("E29").Value & vbNewLine
    objTable.Cell(3, 5).Range.Text = ActiveWorkbook.Worksheets("Podsumowanie").Range("D30").Value & " "


    objDoc.Paragraphs.Add
    objDoc.Paragraphs.Last.Range.InsertBreak
    
    objDoc.Paragraphs.Add
    objDoc.Paragraphs.Last.Range.InsertAfter "lakjsdlkajsdlksa"
    
    objDoc.Paragraphs.Add
    Set objRangeTable = objDoc.Paragraphs.Last.Range
    objDoc.Tables.Add objRangeTable, intNoOfRows3, intNoOfColumns3
    Set objTable = objDoc.Tables(objDoc.Tables.Count)
    objTable.Cell(1, 1).Range.Text = "akuhsauhsd;kas"
    objTable.Cell(1, 2).Range.Text = "kljn  ;mmmmmmmmmaaaaaaaaaaa"

'
End Sub
