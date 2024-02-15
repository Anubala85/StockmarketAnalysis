Attribute VB_Name = "Module1"
Sub stockmarketanalysis():

    Dim i As Double
    
    Dim current_ticker As String
    Dim start_price As Double
    Dim end_price As Double
    Dim trade_volume As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim lastupdate_row As Integer
    
    Dim greatest_percent_inc As Double
    Dim greatest_percent_inc_ticker As String
    
    Dim greatest_percent_dec As Double
    Dim greatest_percent_dec_ticker As String
    
    Dim greatest_total_volume As Double
    Dim greatest_total_volume_ticker As String
    
    greatest_percent_inc = -999999
    greatest_percent_dec = 999999
    greatest_total_volume = -1
    
    Dim ws_Count As Integer
    Dim sheet_index As Integer
    
    'PART A
    
    ws_Count = ActiveWorkbook.Worksheets.Count

    For sheet_index = 1 To ws_Count:
    
        Worksheets(sheet_index).Activate
            
        'Generate column Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
     'PART B
     
               
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row:
    
           
            If i = 2 Then 'declare starting values
                
                current_ticker = Cells(i, 1).Value 'starting ticker
                start_price = Cells(i, 3).Value
                trade_volume = Cells(i, 7).Value
                lastupdate_row = 1
            
            Else
                    
                If Cells(i, 1).Value = current_ticker Then
                        
                    trade_volume = trade_volume + Cells(i, 7).Value
                    
                Else
                    
                    close_price = Cells(i - 1, 6).Value
                    yearly_change = close_price - start_price
                    percent_change = (yearly_change / start_price)
                    
                    
      'PART C
        
                    'GreatestPercentageIncrease
                    
                    If greatest_percent_inc < percent_change Then
                        greatest_percent_inc = percent_change
                        greatest_percent_inc_ticker = current_ticker
                    End If
                    
                         
                    'GreatestPercentageDecrease
                    
                    If greatest_percent_dec > percent_change Then
                        greatest_percent_dec = percent_change
                        greatest_percent_dec_ticker = current_ticker
                    End If
                    
                    'GreatestTotalVolume
                                
                    If greatest_total_volume < trade_volume Then
                       greatest_total_volume = trade_volume
                        greatest_total_volume_ticker = current_ticker
                    End If
                    
                    ' print output
                    Cells(lastupdate_row + 1, 9).Value = current_ticker
                    Cells(lastupdate_row + 1, 10).Value = yearly_change
                    Cells(lastupdate_row + 1, 11).Value = FormatPercent(percent_change, 2)
                    Cells(lastupdate_row + 1, 12).Value = trade_volume
                    
      'PART D
      
         
                    'Conditional formatting
                    
                    If yearly_change < 0 Then
                        Cells(lastupdate_row + 1, 10).Interior.ColorIndex = 3
                    Else
                        Cells(lastupdate_row + 1, 10).Interior.ColorIndex = 4
                    End If
                    
                    'MsgBox ("Ticker:" & current_ticker & " " & "Start Price:" & start_price & " " & "Close Price:" & close_price)
                    
                    ' Reset Variable
                    current_ticker = Cells(i, 1).Value
                    start_price = Cells(i, 3).Value
                    trade_volume = Cells(i, 7).Value
                    lastupdate_row = lastupdate_row + 1
                          
                End If
            
            End If
           
        Next i
    
    
        'Sumamary
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
      
        Cells(2, 16).Value = greatest_percent_inc_ticker
        Cells(2, 17).Value = FormatPercent(greatest_percent_inc, 2)
        
        Cells(3, 16).Value = greatest_percent_dec_ticker
        Cells(3, 17).Value = FormatPercent(greatest_percent_dec, 2)
        
        Cells(4, 16).Value = greatest_total_volume_ticker
        Cells(4, 17).Value = greatest_total_volume

    Next sheet_index
   
End Sub

