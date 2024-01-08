Sub Multiple_Year_Stock_Analysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume_Total As Double
    Dim Summary_Table_Row As Integer

    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    Dim GreatestIncreaseValue As Double
    Dim GreatestDecreaseValue As Double
    Dim GreatestVolumeValue As Double

    For Each ws In ThisWorkbook.Sheets
    
        Summary_Table_Row = 2
        GreatestIncreaseValue = 0
        GreatestDecreaseValue = 0
        GreatestVolumeValue = 0

      
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row 'they all have different row counts

        Open_Price = ws.Cells(2, 3).Value

        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
               
                Ticker = ws.Cells(i, 1).Value
                Close_Price = ws.Cells(i, 6).Value
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

                Yearly_Change = Close_Price - Open_Price
                
                Percent_Change = Yearly_Change / Open_Price
            
                ws.Cells(Summary_Table_Row, 11).Value = Ticker
                ws.Cells(Summary_Table_Row, 12).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 12).NumberFormat = "0.00"
                ws.Cells(Summary_Table_Row, 13).Value = Percent_Change
                ws.Cells(Summary_Table_Row, 13).NumberFormat = "0.00%"
                ws.Cells(Summary_Table_Row, 14).Value = Volume_Total

                
                If Yearly_Change >= 0 Then
                    ws.Cells(Summary_Table_Row, 12).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, 12).Interior.ColorIndex = 3
                End If
                
                If Percent_Change >= 0 Then
                    ws.Cells(Summary_Table_Row, 13).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, 13).Interior.ColorIndex = 3
                End If
                    
               
                If Percent_Change > GreatestIncreaseValue Then
                    GreatestIncreaseValue = Percent_Change
                    GreatestIncreaseTicker = Ticker
                ElseIf Percent_Change < GreatestDecreaseValue Then
                    GreatestDecreaseValue = Percent_Change
                    GreatestDecreaseTicker = Ticker
                End If

                If Volume_Total > GreatestVolumeValue Then
                    GreatestVolumeValue = Volume_Total
                    GreatestVolumeTicker = Ticker
                End If

                
                Summary_Table_Row = Summary_Table_Row + 1

                
                Open_Price = ws.Cells(i + 1, 3).Value
                Volume_Total = 0
            Else
                
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            End If
        Next i

        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percent Change"
        ws.Range("N1").Value = "Total Stock Volume"
        
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("T2").NumberFormat = "0.00%"
        ws.Range("S2").Value = GreatestIncreaseTicker
        ws.Range("T2").Value = GreatestIncreaseValue

        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("T3").NumberFormat = "0.00%"
        ws.Range("S3").Value = GreatestDecreaseTicker
        ws.Range("T3").Value = GreatestDecreaseValue

        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S4").Value = GreatestVolumeTicker
        ws.Range("T4").Value = GreatestVolumeValue

    Next ws
End Sub