Attribute VB_Name = "Module1"
Sub stock_data()

For Each ws in Worksheets

    ' Variables
    Dim lastrow As Long
    Dim x, y, I, maxline, minline, maxrow As Integer
    Dim closingnumber, openingnumber, maxnumber, minnumber, maxvolume As Double
    Dim total As Double

    ' Variable's Starting Values
    x = 2
    y = 9
    maxnumber = 0
    minnumber = 0
    maxvolume = 0
    openingnumber = ws.Cells(2, 3).Value
    lastrow = ws.Range("A999999").End(xlUp).Row
    total = ws.Cells(2, 7).Value

    ' Name Columns        
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(2, 15).Value = "Greatest Percent Increase"
    ws.Cells(3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Value"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' First Loop for Ticker, Yearly Change, Percentage Change, and Total Stock Value
    For I = 2 To lastrow
        If ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value Then
            total = total + ws.Cells(I + 1, 7).Value
        End If
        If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
            closingnumber = ws.Cells(I, 6).Value
            ws.Cells(x, y).Value = ws.Cells(I, 1).Value
            ws.Cells(x, y + 1).Value = closingnumber - openingnumber
            ws.Cells(x, y + 1).NumberFormat = "0.000000000"
                If ws.Cells(x, y + 1).Value >= 0 Then
                    ws.Cells(x, y + 1).Interior.ColorIndex = 4
                Else:
                    ws.Cells(x, y + 1).Interior.ColorIndex = 3
                End If
                If openingnumber <> 0 Then
                    ws.Cells(x, y + 2).Value = (closingnumber - openingnumber) / openingnumber
                Else:
                    ws.Cells(x, y + 2).Value = 0
                End If
            ws.Cells(x, y + 2).NumberFormat = "0.00%"
            ws.Cells(x, y + 3).Value = total
            openingnumber = ws.Cells(I + 1, 3).Value
                        
            If openingnumber = 0 Then
                For J = I + 1 To lastrow
                    If ws.Cells(J, 1).Value <> ws.Cells(J + 1, 1).Value Then Exit For
                    If ws.Cells(J, 3).Value > 0 Then
                        openingnumber = ws.Cells(J, 3).Value
                        Exit For
                    End If
                Next J
            End If
            x = x + 1
            total = ws.Cells(I + 1, 7).Value
        End If
    Next I
    
    ' Loop for Greatest Percent Increase
    For I = 2 To lastrow
        If ws.Cells(I, 11) > maxnumber Then
            maxnumber = ws.Cells(I, 11).Value
            maxline = I
        End If
    Next I
    ws.Cells(2, 16).Value = ws.Cells(maxline, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(maxline, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ' Loop for Greatest Percent % Decrease
    For I = 2 To lastrow
        If ws.Cells(I, 11) < minnumber Then
            minnumber = ws.Cells(I, 11).Value
            minline = I
        End If
    Next I
    ws.Cells(3, 16).Value = ws.Cells(minline, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(minline, 11).Value
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Loop for Greatest Total Value
    For I = 2 To lastrow
        If ws.Cells(I, 12) > maxvolume Then
            maxvolume = ws.Cells(I, 12).Value
            maxrow = I
        End If
    Next I
    ws.Cells(4, 16).Value = ws.Cells(maxrow, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(maxrow, 12).Value
       
Next ws
                        
End Sub