Sub Analyze()
'initialize variables
    Dim ticker As String
    Dim vol As Double
    Dim row As Long
    Dim open_track As Double
    Dim close_track As Double
    Dim leaderboard_row As Integer
    Dim lastRow As Long
    Dim ws As Integer
    Dim tws As Integer
    
    tws = ActiveWorkbook.Worksheets.Count
    
    'Move through each worksheet
    For ws = 1 To tws
    
            'assign initial values for some variables.
            leaderboard_row = 2
            vol = 0
            ticker = Worksheets(ws).Cells(2, 1).Value
            open_track = Worksheets(ws).Cells(2, 3).Value
            
            'set up the leaderboard
            Worksheets(ws).Cells(1, 10).Value = "Ticker"
            Worksheets(ws).Cells(1, 11).Value = "Quarterly Change"
            Worksheets(ws).Cells(1, 12).Value = "Percent Change"
            Worksheets(ws).Cells(1, 13).Value = "Volume"
            
            
            ' Find the last row with data in the first column
            lastRow = Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).row
        
        
            'go row by row checking if this is the last row and adding to the volume tracker
            For row = 2 To lastRow
                'If the next row ticker is the same, add to the volume and move on
                If (ticker = Worksheets(ws).Cells(row + 1, 1).Value) Then
                    vol = vol + Worksheets(ws).Cells(row, 7).Value
                    
                'if the next row isnt the same, add to volume and output to leaderboard. Then change the tracking variables
                Else
                    'add to the volume and update the close value
                    vol = vol + Worksheets(ws).Cells(row, 7).Value
                    close_track = Worksheets(ws).Cells(row, 6).Value
                    
                    
                    'output to leaderboard
                    Worksheets(ws).Cells(leaderboard_row, 10).Value = ticker
                    Worksheets(ws).Cells(leaderboard_row, 11).Value = close_track - open_track
                    Worksheets(ws).Cells(leaderboard_row, 12).Value = 100 * Worksheets(ws).Cells(leaderboard_row, 11).Value / open_track
                    Worksheets(ws).Cells(leaderboard_row, 13).Value = vol
                    If Worksheets(ws).Cells(leaderboard_row, 11).Value > 0 Then
                        Worksheets(ws).Cells(leaderboard_row, 11).Interior.ColorIndex = 4
                    ElseIf Worksheets(ws).Cells(leaderboard_row, 11).Value < 0 Then
                        Worksheets(ws).Cells(leaderboard_row, 11).Interior.ColorIndex = 3
                    End If
                    
                    leaderboard_row = leaderboard_row + 1
                    
                    'reset the variables for the next ticker
                    vol = 0
                    open_track = Worksheets(ws).Cells(row + 1, 3).Value
                    ticker = Worksheets(ws).Cells(row + 1, 1).Value
                End If
            Next row
        'Find the greatest increase/decrease and volume and output it to another leaderboard
        'First, set up the leaderboard
        Worksheets(ws).Cells(2, 15).Value = "Greatest Volume"
        Worksheets(ws).Cells(3, 15).Value = "Greatest % Increase"
        Worksheets(ws).Cells(4, 15).Value = "Greatest % Decrease"
        Worksheets(ws).Cells(1, 16).Value = "Ticker"
        Worksheets(ws).Cells(1, 17).Value = "Value"
        
        'then, go through the first leaderboard and find the max and min values. I know I could use a max or min function but this lets me get the row index for the Ticker symbol
        Dim minTrack As Double
        Dim maxTrack As Double
        Dim volTrack As Double
        
        Dim minRow As Integer
        Dim maxRow As Integer
        Dim volRow As Integer
        Dim i As Integer
        minTrack = 0
        maxTrack = 0
        volTrack = 0
        lastRow = Worksheets(ws).Cells(Rows.Count, 12).End(xlUp).row
        For i = 2 To lastRow
            If Worksheets(ws).Cells(i, 12).Value > maxTrack Then
                maxTrack = Worksheets(ws).Cells(i, 12).Value
                maxRow = i
            ElseIf Worksheets(ws).Cells(i, 12).Value < minTrack Then
                minTrack = Worksheets(ws).Cells(i, 12).Value
                minRow = i
            ElseIf Worksheets(ws).Cells(i, 13).Value > volTrack Then
                volTrack = Worksheets(ws).Cells(i, 13).Value
                volRow = i
            End If
        Next i
        
        'output the values to the leaderboard
        Worksheets(ws).Cells(2, 16).Value = Worksheets(ws).Cells(volRow, 10).Value
        Worksheets(ws).Cells(3, 16).Value = Worksheets(ws).Cells(maxRow, 10).Value
        Worksheets(ws).Cells(4, 16).Value = Worksheets(ws).Cells(minRow, 10).Value
        Worksheets(ws).Cells(2, 17).Value = volTrack
        Worksheets(ws).Cells(3, 17).Value = maxTrack
        Worksheets(ws).Cells(4, 17).Value = minTrack
        



        Next ws

End Sub

