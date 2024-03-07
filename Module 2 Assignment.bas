Attribute VB_Name = "Module3"
Sub alphatest2():

'Define variables

Dim i As Integer
Dim tickersymbol As String
Dim lastRow As Integer
Dim summary_table_row As Integer
Dim stocktotal As Double
Dim currentTicker As String
Dim initialOpen As Double
Dim finalClose As Double
Dim yearlychange As Double
Dim ws As Worksheet
Dim percentChange As Double



'starting value of summary table rows

summary_table_row = 2



'Loop through all sheets

    For Each ws In Worksheets
        stocktotal = 0
        summary_table_row = 2
        currentTicker = Cells(2, 1)
        initialOpen = Cells(2, 3).Value
        
        'define last row for each worksheet
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
            'loop through each row
            
            For i = 2 To lastRow
                
                    
                    'store ticker symbol
            
            tickersymbol = Cells(i, 1).Value
            'initialOpen = Cells(i, 3).Value
            
                    'If the ticker has changed
                
                    If tickersymbol <> currentTicker Then
                    
                        
                         finalClose = Cells((i - 1), 6).Value
                        yearlychange = initialOpen - finalClose
                        percentChange = yearlychange / initialOpen
                        

                        
                    'add up the stock total for when ticker values are different
                                        
                    ws.Range("L" & summary_table_row).Value = stocktotal
                    
                    'display the yearly change
                    
                    ws.Range("J" & summary_table_row).Value = yearlychange
                    ws.Range("K" & summary_table_row).Value = Format(percentChange, "0.00%")
                    
                    'add in conditional format
                    
                        If yearlychange < 0 Then
                        Range("J" & summary_table_row).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        
                        ElseIf yearlychange > 0 Then
                        Range("J" & summary_table_row).Select
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 5287936
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        End If
                        
                                        
                    ws.Range("I" & summary_table_row).Value = currentTicker
                    
                     'add one to the summary table row to allow for next change in ticker
                    summary_table_row = summary_table_row + 1
                    
                    'Reset the stock volume total
                    initialOpen = Cells(i, 3).Value
                    stocktotal = 0
                    currentTicker = Cells(i, 1).Value
                        
                   
                    'exit the loop for the last row
                    
                    ElseIf currentTicker = "" Then
                        Exit For
                      Else
                      
                    'add to the stock total
                    
                    stocktotal = stocktotal + Cells(i, 7).Value
                    
                    
                    
                    
                    
                    End If
                                                    
                  
                    
                      
            Next i
            
    Next ws
     
End Sub
