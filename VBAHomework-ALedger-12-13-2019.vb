Sub HW()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets

'Variables
ws.Activate

Dim StockVar As String

Dim I2 As String

Dim I As Double

Dim TickerName() As String

Dim StockCount() As Double

Dim VolCount As Double

Dim LastStock As String

Dim OpenStock As Double

Dim CloseStock As Double

Dim MaxInc As Double
Dim MaxDec As Double
Dim MaxVol As Double
Dim MaxIncName As String
Dim MaxDecName As String
Dim MaxVolName As String

Columns("A:B").Sort key1:=Range("A2"), order1:=xlAscending, key2:=Range("B2"), order2:=xlAscending, Header:=xlYes

TickerCount = 0

RowCount = ThisWorkbook.ActiveSheet.UsedRange.Rows.Count

VolCount = 0

StockVar = "Empty"

CloseStock = 0

OpenStock = 0

Cells(1, 8).Value = "Ticker"

Cells(1, 11) = "Tot_Vol"

Cells(1, 9) = "YearlyChange"

Cells(1, 10) = "PercentChange"

Cells(2, 13) = "Greatest % Increase"
Cells(3, 13) = "Greatest % Decrease"
Cells(4, 13) = "Greatest Total Volume"
Cells(1, 14) = "Ticker"
Cells(1, 15) = "Value"

LastStock = Range("A" + CStr(RowCount)).Value

'Count and Return of Ticker Names
For I = 2 To RowCount

I2 = CStr(I)

'Return unique ticker IDs and total volume
If StockVar = "Empty" Then
    StockVar = Range("A" + I2).Value
    VolCount = VolCount + Range("G" + I2).Value
    MaxVolName = StockVar
    MaxVol = VolCount


ElseIf (StockVar <> Range("A" + I2).Value) Or (StockVar = LastStock And I = RowCount) Then
    TickerCount = TickerCount + 1
    Range("H" + CStr(TickerCount + 1)).Value = StockVar
    Range("K" + CStr(TickerCount + 1)).Value = VolCount
    StockVar = Range("A" + I2).Value
'    If VolCount > MaxVol Then
'        MaxVol = VolCount
'        Range("O4").Value = VolCount
'        Range("N4").Value = Range("A" + I2)
'        End If
   
    VolCount = 0


ElseIf StockVar = Range("A" + I2).Value Then
    VolCount = VolCount + Range("G" + I2).Value
    End If

Next I

'Calculate percent change from beginning of year to end

TickerCount = 1
StockVar = Range("A2")
OpenStock = Range("C2")


For J = 2 To RowCount

    J2 = CStr(J)

    If (StockVar <> Range("A" + J2).Value) Then
        CloseStock = Range("F" + CStr(J - 1)).Value
        'Range("M" + CStr(TickerCount + 1)).Value = CloseStock
        Range("I" + CStr(TickerCount + 1)).Value = CloseStock - OpenStock
        'If (OpenStock > CloseStock) Then
        'Range("J" + CStr(TickerCount + 1)).Value = Format(Range("I" + CStr(TickerCount + 1)).Value / OpenStock, "Percent")
                'Else
        If OpenStock <> 0 Then
        Range("J" + CStr(TickerCount + 1)).Value = Format(Range("I" + CStr(TickerCount + 1)).Value / OpenStock, "Percent")
        Else: Range("J" + CStr(TickerCount + 1)).Value = Format(0, "Percent")
        End If
        'Range("J" + CStr(TickerCount + 1)).Value = "-" + CStr(CDbl((Range("I" + CStr(TickerCount + 1)).Value / OpenStock) * 100)) + "%"
        'End If
        OpenStock = Range("C" + J2)
        'Range("L" + CStr(TickerCount + 2)).Value = OpenStock
        TickerCount = TickerCount + 1
        StockVar = Range("A" + J2).Value
   
    ElseIf (StockVar = LastStock And J = RowCount) Then
        CloseStock = Range("F" + J2).Value
        Range("I" + CStr(TickerCount + 1)).Value = CloseStock - OpenStock
        'If (OpenStock > CloseStock) Then
        'Range("J" + CStr(TickerCount + 1)).Value = CStr(CDbl((Range("I" + CStr(TickerCount + 1)).Value / OpenStock) * 100)) + "%"
        Range("J" + CStr(TickerCount + 1)).Value = Format(Range("I" + CStr(TickerCount + 1)).Value / OpenStock, "Percent")
        'Else
        'Range("J" + CStr(TickerCount + 1)).Value = "-" + CStr(CDbl((Range("I" + CStr(TickerCount + 1)).Value / OpenStock) * 100)) + "%"
        'End If
        TickerCount = TickerCount + 1
        StockVar = Range("A" + J2).Value
        'Range("L" + CStr(TickerCount + 1)).Value = OpenStock
        'Range("M" + CStr(TickerCount + 1)).Value = CloseStock
         End If

Next J

StockVar = Range("H2").Value
MaxInc = Range("J2").Value
MaxDec = Range("J2").Value
MaxVolName = Range("H2").Value
MaxDecName = Range("H2").Value
MaxcName = Range("H2").Value

For Z = 3 To TickerCount
    Z2 = CStr(Z)
 If Range("J" + Z2) > MaxInc Then
    MaxInc = Range("J" + Z2)
    MaxIncName = Range("H" + Z2)
    End If
  If Range("J" + Z2) < MaxDec Then
    MaxDec = Range("J" + Z2)
    MaxDecName = Range("H" + Z2)
    End If
If Range("K" + Z2) > MaxVol Then
    MaxVol = Range("K" + Z2)
    MaxVolName = Range("H" + Z2)
    End If
If Range("I" + Z2) >= 0 Then
   Range("I" + Z2).Interior.ColorIndex = 4
   Else: Range("I" + Z2).Interior.ColorIndex = 3
   End If
   
Next Z

If Range("I2") >= 0 Then
 Range("I2").Interior.ColorIndex = 4
 Else: Range("I2").Interior.ColorIndex = 3
 End If

Range("O2") = Format(MaxInc, "Percent")
Range("N2") = MaxIncName
Range("O3") = Format(MaxDec, "Percent")
Range("N3") = MaxDecName
Range("O4") = MaxVol
Range("N4") = MaxVolName

Next ws

starting_ws.Activate


End Sub