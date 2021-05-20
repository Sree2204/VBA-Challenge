VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ticker()

Dim Ticker As String

Dim Volume_total As Double
Volume_total = 0

Dim Summary_table_row As Integer
Summary_table_row = 2

For i = 2 To 800000

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Volume_total = Volume_total + Cells(i, 7).Value
        Range("I" & Summary_table_row).Value = Ticker
        Range("L" & Summary_table_row).Value = Volume_total
        Summary_table_row = Summary_table_row + 1
        Volume_total = 0
    Else
        Volume_total = Volume_total + Cells(i, 7).Value
    
    End If
    
Next i
     

End Sub
