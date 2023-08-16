Attribute VB_Name = "StockAnalysis"
Sub StockAnalysis()

    'Yearly breakdown of each stocks
    
    
    'Loop through 'A' Column to pull:
        'Closing - Opening values of the year
        '% Change from opening year price to closing year price
        'The total stock volume of the stock
        
    'Declare var
    Dim ticker As String    'Used to store the Ticker ID
    Dim opnPrice As Double  'Store Year opening price for Ticker
    Dim clsPrice As Double  'Store Year end price for Ticker
    Dim vol As LongLong     'Total stock volume
    Dim lr As Long          'lr = LastRow
    Dim ws As Worksheet     'Worksheet
    Dim i As Long           'i used to iterate through each row
    Dim outRow As Integer   'Output Row for summary of ticker
    Dim StockYear As String 'StockYear taken from the sheet name
    
    For Each ws In Worksheets
    
        ''Add Column Headings
         StockYear = Str(ws.Name)   'Comment out in alphabetical testing sheet
         ws.Cells(1, 10).Value = "Ticker"
         ws.Cells(1, 11).Value = StockYear & " Yearly Change"
         ws.Cells(1, 12).Value = StockYear & " Percent Change"
         ws.Cells(1, 13).Value = StockYear & " Total Stock Volume"
         With ws.Range("J1:M1")
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
         End With
         
         ''Add Headings for Greatest % inc/dec and Volume
         ws.Cells(1, 17).Value = "Ticker"
         ws.Cells(1, 18).Value = "Value"
         ws.Cells(2, 16).Value = "Greatest % Increase:"
         ws.Cells(3, 16).Value = "Greatest % Decrease:"
         ws.Cells(4, 16).Value = "Greatest Total Volume:"
         With ws.Range("Q1:R1,P2:P5")
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
         End With
    
    'Set initial values
    lr = ws.Cells(Rows.Count, "A").End(xlUp).Row
    opnPrice = 0
    clsPrice = 0
    vol = 0
    outRow = 1
    
            'For loop through Stock Data
            For i = 2 To lr
                
                ''If opnPrice = 0 then this is the first iteration of the ticker, so store the opening price
                If opnPrice = 0 Then
                    opnPrice = ws.Cells(i, 3).Value
                End If
        
        
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        ''Store Required information for outputs table
                        ticker = ws.Cells(i, 1).Value
                        vol = vol + ws.Cells(i, 7).Value
                        clsPrice = ws.Cells(i, 6).Value
            
                        ''Write to outputs table
                        ws.Cells(outRow + 1, 10).Value = ticker
                        ws.Cells(outRow + 1, 11).Value = clsPrice - opnPrice
                        ''Format Yearly Change Cell based on value (Red = Neg, Grn = Pos)
                        If clsPrice - opnPrice < 0 Then
                            ws.Cells(outRow + 1, 11).Interior.ColorIndex = 3
                        Else: ws.Cells(outRow + 1, 11).Interior.ColorIndex = 4
                    End If
                        ws.Cells(outRow + 1, 12).Value = (clsPrice - opnPrice) / opnPrice
                        ''Format Yearly Change Cell based on value (Red = Neg, Grn = Pos)
                    If (clsPrice - opnPrice) / opnPrice < 0 Then
                            ws.Cells(outRow + 1, 12).Interior.ColorIndex = 3
                        Else: ws.Cells(outRow + 1, 12).Interior.ColorIndex = 4
                    End If

                        ws.Cells(outRow + 1, 13).Value = vol
        
                    ''Reset var values
                    ticker = ""
                    opnPrice = 0
                    clsPrice = 0
                    vol = 0
                    outRow = outRow + 1
        
                Else: vol = vol + ws.Cells(i, 7).Value
        
                End If
        
            Next i
        
        
        ''Find Greatest inc/dec/vol
        Dim perIncTick As String    'Greatest % Inc Ticker
        Dim perDecTick As String    'Greatest % Dec Ticker
        Dim volTick As String       'Greatest Vol Ticker
        Dim perInc As Double        'Greatest % Inc
        Dim perDec As Double        'Greatest % Dec
        Dim TotVol As LongLong      'Greatest vol
                
        ''Set initial values
        perInc = 0
        perDec = 0
        TotVol = 0
        
        ''Iterate through the summary table to find Greatest % inc/dec and Total Vol
        For i = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            ''Check if % is greater than value already stored, if it is then store this value and retrieve the Ticker
            If ws.Cells(i, 12).Value > perInc Then
                perInc = ws.Cells(i, 12).Value
                perIncTick = ws.Cells(i, 10).Value
            ''If the % is not greater, then check if it's less than the greatest % decrease, if it is then store this value and retrieve the ticker
            ElseIf ws.Cells(i, 12).Value < perDec Then
                perDec = ws.Cells(i, 12).Value
                perDecTick = ws.Cells(i, 10).Value
            End If
            
            
            ''Check if vol is greater than the TotVol currently stored, if it is then retrieve this value and store the ticker
            If ws.Cells(i, 13).Value > TotVol Then
                TotVol = ws.Cells(i, 13).Value
                volTick = ws.Cells(i, 10).Value
            End If
            
        
        Next i
        
        ''Write Greatest inc/dec/vol
        ws.Range("Q2").Value = perIncTick
        ws.Range("Q3").Value = perDecTick
        ws.Range("Q4").Value = volTick
        With ws.Range("R2")
            .Value = perInc
            .NumberFormat = "0.00%"
        End With
        With ws.Range("R3")
            .Value = perDec
            .NumberFormat = "0.00%"
        End With
        With ws.Range("R4")
            .Value = TotVol
            .NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
        End With
        
        ''Format J:M Columns
        ws.Columns("J:M").AutoFit
        ws.Range("K2:K" & Cells(Rows.Count, 12).End(xlUp).Row).NumberFormat = "[$$-en-US]#,##0.00"
        ws.Range("L2:L" & Cells(Rows.Count, 12).End(xlUp).Row).NumberFormat = "0.00%"
        ws.Range("M2:M" & Cells(Rows.Count, 12).End(xlUp).Row).NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
        
        ''Format P:R Columns (Autofit)
        ws.Columns("P:R").AutoFit
        
    Next ws
    
    

End Sub
