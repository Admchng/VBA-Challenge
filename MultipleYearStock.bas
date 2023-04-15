Attribute VB_Name = "Module1"
Sub alphabetical_testing()
Dim Ticker As String
Dim Volume_Total As Double
Dim ws As Worksheet
Dim Summary_Row As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Year_Open As Double
Dim Year_Close As Double
Dim Summary_Volume As Double
Dim Summary_Ticker As String

Volume_Total = 0
Summary_Volume = 0

For Each ws In Worksheets
Summary_Row = 2
Lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(2, 15) = "Greatest % Increase Value"
    ws.Cells(3, 15) = "Greatest % Decrease Value"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    

For I = 2 To Lastrow1
    If ws.Cells(I + 1, 1) <> ws.Cells(I, 1) Then
        Ticker = ws.Cells(I, 1)
        Volume_Total = Volume_Total + ws.Cells(I, 7)
        ws.Cells(Summary_Row, 9) = Ticker
        ws.Cells(Summary_Row, 12) = Volume_Total
        Volume_Total = 0
        Year_Close = ws.Cells(I, 6)
        Yearly_Change = Year_Close - Year_Open
            ws.Cells(Summary_Row, 10) = Yearly_Change
        Percent_Change = (Yearly_Change / Year_Open)
            ws.Cells(Summary_Row, 11) = Percent_Change
        Summary_Row = Summary_Row + 1
    Else
        If ws.Cells(I - 1, 1) <> ws.Cells(I, 1) Then
        Year_Open = ws.Cells(I, 3)
        End If
        Volume_Total = Volume_Total + ws.Cells(I, 7)
    End If
Next I
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For I = 2 To LastRow2
        If ws.Cells(I, 12) > Summary_Volume Then
            Summary_Volume = ws.Cells(I, 12)
            Summary_Ticker = ws.Cells(I, 9)
        End If
        If ws.Cells(I, 11) > Greatest_Increase Then
            Greatest_Increase = ws.Cells(I, 11)
            Greatest_Increase_Ticker = ws.Cells(I, 9)
        End If
        If ws.Cells(I, 11) < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(I, 11)
            Greatest_Decrease_Ticker = ws.Cells(I, 9)
        End If
Next I

    ws.Cells(4, 17) = Summary_Volume
    ws.Cells(4, 16) = Summary_Ticker
    ws.Cells(17, 2) = Greatest_Increase
    ws.Cells(2, 16) = Greatest_Increase_Ticker
    ws.Cells(3, 17) = Greatest_Decrease
    ws.Cells(3, 16) = Greatest_Decrease_Ticker
    
    Ticker_Row = 2
    Summary_Volume = 0
    Summary_Ticker = ""
    
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
        For I = 2 To LastRow2
            If ws.Cells(I, 10) < 0 Then
            ws.Cells(I, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(I, 10).Interior.ColorIndex = 4
            End If
        Next I
Next ws
End Sub
