VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub ticker()

Dim LR As LongLong
Dim LR2 As LongLong
Dim vol As LongLong
Dim row As LongLong
Dim SumLine As Integer
Dim RowCount As LongLong
Dim RsCount As LongLong
Dim grd As Single
Dim gri As Single
Dim grtv As LongLong


'counting how many rows of data in this worksheet
LR = Cells(Rows.Count, 1).End(xlUp).row

vol = 0
RowCount = 2
SumLine = 2

'Fill up ticker
For i = RowCount To LR
If Cells(RowCount + 1, 1).Value = Cells(RowCount, 1).Value Then
     RowCount = RowCount + 1
End If

Cells(SumLine, 10).Value = Cells(RowCount, 1).Value
     
If Cells(RowCount + 1, 1).Value <> Cells(RowCount, 1).Value Then
    RowCount = RowCount + 1
    SumLine = SumLine + 1
    Cells(SumLine, 10).Value = Cells(RowCount, 1).Value
End If
Next i

'fill up ticker data

RowCount = 2
SumLine = 2

For j = RowCount To LR
    vol = 0
    RsCount = RowCount
    While Cells(RowCount, 1).Value = Cells(SumLine, 10).Value
         vol = vol + Cells(RowCount, 7).Value
         RowCount = RowCount + 1
    Wend
    Cells(SumLine, 11).Value = Cells(RowCount - 1, 6).Value - Cells(RsCount, 3).Value
    If Cells(SumLine, 11).Value > 0 Then
        Cells(SumLine, 11).Interior.ColorIndex = 4 'green
        Else: Cells(SumLine, 11).Interior.ColorIndex = 3 'Red
    End If
    Cells(SumLine, 12).Value = Round(100 * ((Cells(RowCount - 1, 6).Value / Cells(RsCount, 3).Value) - 1), 2) & "%"
    Cells(SumLine, 13).Value = vol
    SumLine = SumLine + 1
Next j

End Sub



Sub summary()
'summary data fill

LR2 = Cells(Rows.Count, 12).End(xlUp).row

grd = WorksheetFunction.Min(Range("L" & 2 & ":" & "L" & LR2))
gri = WorksheetFunction.Max(Range("L" & 2 & ":" & "L" & LR2))
grtv = WorksheetFunction.Max(Range("M" & 2 & ":" & "M" & LR2))

Range("s2").Value = grd
Range("s3").Value = gri
Range("s4").Value = grtv

For k = 2 To LR2
 If Cells(k, 12).Value = grd Then Range("R2").Value = Cells(k, 10).Value
 If Cells(k, 12).Value = gri Then Range("R3").Value = Cells(k, 10).Value
 If Cells(k, 13).Value = grtv Then Range("R4").Value = Cells(k, 10).Value
Next k

End Sub



