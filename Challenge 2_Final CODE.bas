Attribute VB_Name = "Module11"
Sub StockData()

Dim ws_num As Integer

ws_num = Worksheets.Count
MsgBox (ws_num)

For Each ws In Worksheets
ws.Activate

Dim i As Double
Dim Ticker As String
Dim lastrow As Double
Dim Greatest_Decrease As Double
Dim Greatest_Increase As Double
Dim Greatest_Vol As Double

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
Range("S1").Value = "Value"
Range("R1").Value = "Ticker"
Range("Q2").Value = "Greatest % Increase"
Range("Q3").Value = "Greatest % Decrease"
Range("Q4").Value = "Greatest Total Volume"

Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Vol = 0

    For i = 2 To lastrow

        Ticker = Cells(i, 1).Value
        Cells(i, 10).Value = Ticker
        Cells(i, 11).Value = Cells(i, 6).Value - Cells(i, 3).Value
        Cells(i, 11).NumberFormat = "$0.00"
        Cells(i, 12).Value = Cells(i, 11).Value / Cells(i, 3).Value
        Cells(i, 12).NumberFormat = "0.00%"
        Cells(i, 13).Value = Cells(i, 7).Value * Cells(i, 6).Value
    
        If Cells(i, 12).Value > Greatest_Increase And Range("S2").Value < Cells(i, 12).Value Then
        Range("S2").Value = Cells(i, 12).Value
        Range("S2").NumberFormat = "0.00%"
        Range("R2").Value = Cells(i, 1).Value
        End If
    
        If Cells(i, 12).Value < Greatest_Decrease And Range("S3").Value > Cells(i, 12).Value Then
        Range("S3").Value = Cells(i, 12).Value
        Range("S3").NumberFormat = "0.00%"
        Range("R3").Value = Cells(i, 1).Value
        End If
    
        If Cells(i, 13).Value > Greatest_Vol And Range("S4").Value < Cells(i, 13).Value Then
        Range("S4").Value = Cells(i, 13).Value
        Range("R4").Value = Cells(i, 1).Value
        End If
    
        If Cells(i, 12).Value > 0 Then
        Cells(i, 12).Interior.ColorIndex = 4
        ElseIf Cells(i, 12).Value = 0 Then
        Cells(i, 12).Interior.ColorIndex = 6
        Else
        Cells(i, 12).Interior.ColorIndex = 3
        End If
    
        If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
        ElseIf Cells(i, 11).Value = 0 Then
        Cells(i, 11).Interior.ColorIndex = 6
        Else
        Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i
Next ws
End Sub


