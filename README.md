# VBA-Challenge
Data Science Bootcamp Hw 2
Sub Stonks():

'How much did each Ticker Change per year?

Dim Ticker As String
Dim Summary As Integer
Dim YearlyChange As Double
Dim Initial As Double
Dim Final As Double
Dim Percent As Double
Dim Volume As Double


Summary = 2

'Row
For i = 2 To 20

Ticker = Cells(i, 1).Value



    'Column
    For j = 1 To 7
        
        'If Initial is 0 (and it will be because that's how numerical variables start)
        'then it takes the value of the open column and moves on.
        'If the Cell above in the ticker column isn't equal to the cell the loop is on
        'then initial will become 0 again.
        
        If Initial = 0 Then
        Initial = Cells(i, 3).Value
        ElseIf Cells(i - 1, 1).Value <> Ticker Then
        Initial = 0
        Volume = 0
    
        End If
    
        'Checks If Cell Below has the same ticker value.
        'If it does, it drops the year change in the summary table and adds 1 to the row
        
        If Cells(i + 1, 1).Value <> Ticker Then
        Final = Cells(i, 6).Value
        YearlyChange = Final - Initial
        Percent = ((Final - Initial) / Initial) * 100
        Range("J" & Summary).Value = YearlyChange
        Range("I" & Summary).Value = Ticker
        Range("K" & Summary).Value = Percent
        Range("L" & Summary).Value = Volume
        
        Summary = Summary + 1
        
        Else
        Volume = Volume + Cells(i, 7).Value
        
    
        End If
        

    
    Next j
    
Next i


End Sub
