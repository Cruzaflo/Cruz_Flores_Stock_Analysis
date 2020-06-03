Attribute VB_Name = "Module1"
Sub Stock_Analysis()

For Each ws In Worksheets

    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Table header names
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Voume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
    'Variable Declaration
        Dim tickername As String
        Dim tickercount As Integer
        Dim openp As Double
        Dim closep As Double
        Dim yrchange As Double
        Dim volumetotal As Double
        Dim rowcount As Double
        
    'Variable Assignment
        openp = 0
        closep = 0
        yrchange = 0
        tickercount = 2
        'WorksheetName = ws.Name

        
    'Row Index
        For ri = 2 To Lastrow
            volumetotal = volumetotal + ws.Cells(ri, 7).Value
        
    'ticker search
        'If cell is not equal to proceeding Cell and open value is not 0
        If ws.Cells(ri + 1, 1).Value <> ws.Cells(ri, 1).Value And openp <> 0 Then
            
            'set variables
            tickername = ws.Cells(ri, 1).Value
            closep = ws.Cells(ri, 6).Value
            yrchange = closep - openp
            
            'Perform
            ws.Range("I" & tickercount).Value = tickername
            ws.Range("J" & tickercount).Value = yrchange
            ws.Range("L" & tickercount).Value = volumetotal
            ws.Range("K" & tickercount).Value = FormatPercent(yrchange / openp)
                
            'reset variables
            tickercount = tickercount + 1
                    openp = 0
                    closep = 0
                    yrchange = 0
                    volumetotal = 0
                    
                ElseIf ws.Cells(ri + 1, 1).Value <> ws.Cells(ri, 1).Value And openp = 0 Then
            
                    'set variables
                    tickername = ws.Cells(ri, 1).Value
                    closep = ws.Cells(ri, 6).Value
                    yrchange = closep - openp
            
                    'Perform
                    ws.Range("I" & tickercount).Value = tickername
                    ws.Range("J" & tickercount).Value = yrchange
                    ws.Range("L" & tickercount).Value = volumetotal
                    ws.Range("K" & tickercount).Value = 0
                
                    'reset variables
                    tickercount = tickercount + 1
                        openp = 0
                        closep = 0
                        yrchange = 0
                        volumetotal = 0
            
            'If Cell is not equal to proceeding cell and open value is 0
                ElseIf ws.Cells(ri - 1, 1).Value <> ws.Cells(ri, 1).Value And ws.Cells(ri, 3).Value <> 0 Then

                        openp = ws.Cells(ri, 3).Value
                        
                ElseIf ws.Cells(ri - 1, 1).Value <> ws.Cells(ri, 1).Value And ws.Cells(ri, 3).Value = 0 Then

                        openp = 0

                                 
        End If
        
     Next ri
     
     'Return overall values
        
        'greates and lowest values
        Dim greatest_increase As Double

        Dim lowest_increase As Double

        Dim greatest_volume As Double

        greatest_increase = 0
        
        lowest_increase = 0
        
        greatest_volume = 0
        
        For ri = 2 To tickercount
        
            If ws.Cells(ri, 11).Value > greatest_increase Then
            
                greatest_increase = ws.Cells(ri, 11).Value
                ws.Range("Q2").Value = FormatPercent(ws.Cells(ri, 11).Value)
                ws.Range("P2").Value = ws.Cells(ri, 9).Value
            
                
            ElseIf ws.Cells(ri, 11).Value < lowest_increase Then
            
                lowest_increase = ws.Cells(ri, 11).Value
                ws.Range("Q3").Value = FormatPercent(ws.Cells(ri, 11).Value)
                ws.Range("P3").Value = ws.Cells(ri, 9).Value
                              
            End If
        
        Next ri
        
        For ri = 2 To tickercount
        
            If ws.Cells(ri, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(ri, 12).Value
                ws.Range("Q4").Value = ws.Cells(ri, 12).Value
                ws.Range("P4").Value = ws.Cells(ri, 9).Value
            
            End If
            
        Next ri
        
Next ws

        
End Sub

