Attribute VB_Name = "Module1"
Option Explicit
Sub StockAnalysis()
Dim ticker As Variant
Dim QuarterlyChange As Variant
Dim STockVolume As Variant
Dim PercentChange As Variant
Dim openingdate As String
Dim closingdate As String
Dim openingprice As Variant
Dim closingprice As Variant
Dim Worksheet As Worksheets
Dim i As Long
openingdate = Range("B2").Value
closingdate = Range("B2").End(xlDown)
ticker = Range("A2:A96001").Value
i = 1 '(counter of each row)
'incase you run the script need to start over it will clear all the content
sheets("Q1").Range("I2:L96001").Value = ""
 'Looping across to determine  opening date
For Each ticker In Range("A2:A96001").Value
 If Range("B" & i + 1).Value = openingdate Then
openingprice = Range("C" & i + 1).Value
sheets("Q1").Range("I" & i + 1).End(xlUp).Offset(1, 0) = Range("A" & i + 1).Value
STockVolume = Range("G" & i + 1).Value
'to get the total volume.intial value or opingprice plus on each cell until the closing date
ElseIf Range("A" & i + 1).Value = Range("A" & i + 1).Value Or Range("B" & i + 1).Value = closingdate Then
STockVolume = STockVolume + Range("G" & i + 1).Value
 'if condition met put the value according to the appropriat cell.
If Range("B" & i + 1).Value = closingdate Then
closingprice = Range("F" & i + 1).Value
sheets("Q1").Range("J" & i + 1).End(xlUp).Offset(1, 0) = closingprice - openingprice
QuarterlyChange = closingprice - openingprice
sheets("Q1").Range("K" & i + 1).End(xlUp).Offset(1, 0) = QuarterlyChange / openingprice
sheets("Q1").Range("L" & i + 1).End(xlUp).Offset(1, 0) = STockVolume
End If
End If
i = i + 1
Next ticker
sheets("Q1").Select
MsgBox "Completed"
End Sub
Sub HighestAndLowest()
Dim maxvalue As LongLong
Dim i As Long
Dim VolumeRange As Variant
Dim PercentRange As Variant
Dim PercentMaxvalue As Variant
Dim STockVolumeRange As Variant
Dim PercentMinValue As Variant
PercentRange = Range("K2:K1501").Value
'Greatest% Increase
PercentMaxvalue = Excel.WorksheetFunction.max(PercentRange)
'to get greatest %decrease
PercentMinValue = Excel.WorksheetFunction.Min(PercentRange)
'to get max total volume
VolumeRange = Range("I2:L1501").Value
STockVolumeRange = Range("L2:L1501").Value
maxvalue = Excel.WorksheetFunction.max(STockVolumeRange)
'looping
For i = 2 To 1601
If maxvalue = Range("L" & i + 1).Value Then
sheets("Q1").Range("U7").Value = maxvalue
sheets("Q1").Range("T7").Value = Range("I" & i + 1)
ElseIf PercentMinValue = Range("K" & i + 1).Value Then
sheets("Q1").Range("U5").Value = PercentMinValue
sheets("Q1").Range("T5").Value = Range("I" & i + 1).Value
ElseIf PercentMaxvalue = Range("K" & i + 1).Value Then
sheets("Q1").Range("U3").Value = PercentMaxvalue
sheets("Q1").Range("T3").Value = Range("I" & i + 1).Value
End If
Next i
sheets("Q1").Select
End Sub

    
