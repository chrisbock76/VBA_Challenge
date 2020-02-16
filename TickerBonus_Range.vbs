Attribute VB_Name = "Module3"
Sub Range():
Dim Max As Double
Dim Min As Double

Max = Application.WorksheetFunction.Max(Columns("K"))
Min = Application.WorksheetFunction.Min(Columns("K"))
MsgBox (Max & Min)

End Sub
