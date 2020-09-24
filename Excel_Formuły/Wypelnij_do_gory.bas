Attribute VB_Name = "Module1"
Sub Wypelnij_Do_Poczatku()
Attribute Wypelnij_Do_Poczatku.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Wypelnij_Do_Poczatku Makro
'

'
    ActiveCell.Select
    Selection.AutoFill Destination:=Range("C2:" & ActiveCell.Address), Type:=xlFillDefault
    
End Sub
Sub Makro3()
Attribute Makro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro3 Makro
'

'
    Range("D9:I20").Select
    Selection.Copy
    Range("D23").Select
    ActiveSheet.Paste
End Sub
