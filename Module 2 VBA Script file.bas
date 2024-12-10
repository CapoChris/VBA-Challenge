Attribute VB_Name = "Module1"
Sub Stock()

Dim ticker As Range
Dim tickerRange As Range
Dim lastRow As Long
Dim ws As Worksheet
Dim openPrice As Double
Dim closePrice As Double
Dim i As Long
Dim totalVol As Double
Dim qChange As Double
Dim pChange As Double
Dim tickSym As String
Dim lasTicktRow As Long
Dim uniqueTick As Long
Dim cRange As String
Dim pclastRow As Long
Dim j As Long
        
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'Set Tickers
        Set ticker = ws.Range("A2:A" & lastRow)
        
        'Set output for ticker range
        Set tickerRange = ws.Range("I2")
        
        'Get unique range of tickers
        ticker.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=tickerRange, Unique:=True
        
        'Dim the variables
        totalVol = 0
        
        'Find last row for unique tickers
        lastTickRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        'Loop through unique tickers
        For uniqueTick = 2 To lastTickRow
            tickSym = ws.Cells(uniqueTick, "I").Value
            totalVol = 0
            openPrice = 0
            closePrice = 0
            
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value = tickSym Then
                    'Calculate total volume
                    totalVol = totalVol + ws.Cells(i, 7).Value
                    
                    'Get the open price
                    If openPrice = 0 Then
                        openPrice = ws.Cells(i, 3).Value
                    End If
                    
                    'Update close price
                    closePrice = ws.Cells(i, 6).Value
                
                End If
            Next i
                
            'Calculate quarterly change
            qChange = closePrice - openPrice
            
            'Calculate percentage change
            If openPrice <> 0 Then
                pChange = (qChange / openPrice)
            Else
                pChange = 0
            End If
            
            'Wirte results for ticker
            ws.Cells(uniqueTick, 10).Value = qChange
            ws.Cells(uniqueTick, 11).Value = pChange
            ws.Cells(uniqueTick, 12).Value = totalVol
            
            cRange = "J"
            
            If ws.Cells(uniqueTick, cRange).Value > 0 Then
                ws.Cells(uniqueTick, cRange).Interior.ColorIndex = 4
            ElseIf ws.Cells(uniqueTick, cRange).Value < 0 Then
                ws.Cells(uniqueTick, cRange).Interior.ColorIndex = 3
            Else
                ws.Cells(uniqueTick, cRange).Interior.ColorIndex = 2
            End If
            
            pclastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            
            For j = 2 To pclastRow
                ws.Range("K2:K" & pclastRow).NumberFormat = "0.00%"
            Next j
            
        Next uniqueTick
        
    Next ws
    
End Sub


