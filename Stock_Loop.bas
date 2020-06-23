Attribute VB_Name = "Module1"
Sub Stock_Loop()

    'Set variable types
    Dim openValue As Double
    Dim closeValue As Double
    Dim totalVolume As Variant

    Dim LastRow As Double
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set variable values for placing values
    x = 2
    firstRow = 2
    
    'Find total stock volume for year
    totalVolume = 0
    
    'Set headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Format "summation table"
     Range("I1:L1").Font.FontStyle = "Bold"
     Range("I1:L4").Columns.AutoFit
     
    'Loop through all tickers
    For i = 2 To LastRow
     
        totalVolume = totalVolume + Cells(i, 7).Value
        
        'Set initial opening value
        If i = 2 Then
        openValue = Cells(i, 3).Value
        ticker = Cells(i, 1).Value
        
       'Find when the ticker changes
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ticker = Cells(i, 1).Value
        
            LastRow = i
            'Find close value new ticker
            closeValue = Cells(i, 6).Value
        
            'Find yearly change value for new ticker
            Change = closeValue - openValue
            
            'Find yearly percentage change for new ticker
            PercentChange = (Change / openValue)
            
            'Reset open and ticker value
            openValue = Cells(i + 1, 3).Value
            
            If openValue = 0 Then
                For J = firstRow To LastRow
                    If Cells(J, 3).Value <> 0 Then
                    openValue = Cells(J, 3).Value
                    Exit For
                    End If
                Next J
            End If
            
            If totalVolume = 0 Then
                PercentChange = 0
                Change = 0
            End If
            
            'Deploy values in "summation table"
            Cells(x, 9).Value = ticker
            Cells(x, 10).Value = Change
            Cells(x, 11).Value = PercentChange
            Cells(x, 12).Value = totalVolume
            
            'Reset Total Volume
            totalVolume = 0
                
            'Cell formatting
            Cells(x, 11).NumberFormat = "0.00%"
    
                If Cells(x, 10).Value >= 0 Then
                    Cells(x, 10).Interior.ColorIndex = 4
                ElseIf Cells(x, 10).Value < 0 Then
                    Cells(x, 10).Interior.ColorIndex = 3
                End If
             
                'Move down one row to publish values
                x = x + 1
                firstRow = i + 1
          End If
    Next i
       
End Sub



