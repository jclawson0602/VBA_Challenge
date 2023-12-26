Attribute VB_Name = "Module1"
Sub Stock_Data()

Dim a As LongLong
Dim b As Integer
Dim Ticker_Name As Integer
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Volume As LongLong
Dim Lastrow As LongLong
Dim Greatest_Percent As Double
Dim Smallest_Percent As Double
Dim Greatest_Volume As LongLong

For Each ws In Worksheets

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Ticker_Name = 2
Opening_Price = ws.Cells(2, 3).Value
Greatest_Percent = ws.Range("K2")
Smallest_Percent = ws.Range("K2")
Greatest_Volume = ws.Range("L2")

For a = 2 To Lastrow

    Volume = Volume + ws.Cells(a, 7).Value
    
    If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
        ws.Cells(Ticker_Name, 9) = ws.Cells(a, 1).Value
        Closing_Price = ws.Cells(a, 6).Value
        ws.Cells(Ticker_Name, 10) = Closing_Price - Opening_Price
        ws.Cells(Ticker_Name, 11) = (Closing_Price - Opening_Price) / (Opening_Price)
        Opening_Price = ws.Cells(a + 1, 3).Value
        ws.Cells(Ticker_Name, 12) = Volume
        Volume = 0
        Ticker_Name = Ticker_Name + 1
    Else
        Ticker_Name = Ticker_Name
    End If

Next a

For b = 2 To 3001

    If ws.Cells(b + 1, 11).Value > Greatest_Percent Then
        Greatest_Percent = ws.Cells(b + 1, 11).Value
    Else
        Greatest_Percent = Greatest_Percent
    End If
    
    If ws.Cells(b + 1, 11).Value < Smallest_Percent Then
        Smallest_Percent = ws.Cells(b + 1, 11).Value
    Else
        Smallest_Percent = Smallest_Percent
    End If
    
    If ws.Cells(b + 1, 12).Value > Greatest_Volume Then
        Greatest_Volume = ws.Cells(b + 1, 12).Value
    Else
        Greatest_Volume = Greatest_Volume
    End If

Next b
        
ws.Range("P2").Value = Greatest_Percent
ws.Range("P3").Value = Smallest_Percent
ws.Range("P4").Value = Greatest_Volume

Next ws

End Sub
