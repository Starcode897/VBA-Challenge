# VBA-Challenge
Data Science Bootcamp Hw 2
Sub Stonks():

'How much did each Ticker Change per year?

Dim Ticker As String
Dim Summary As Long
Dim YearlyChange As Double
Dim Initial As Double
Dim Final As Double
Dim Percent As Double
Dim Volume As Double
Dim GreatestPercent As Double

'This is a variable to make the for loop head to the last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row



Summary = 2

For i = 2 To LastRow
        
    'If Initial is 0 (and it will be because that's how numerical variables start)
    'then it takes the value of the open column and moves on.
    'If the Cell above in the ticker column isn't equal to the cell the loop is on
    'then initial will become 0 again.
        
    Ticker = Cells(i, 1).Value
        
    If Initial = 0 Then
    Initial = Cells(i, 3).Value
        
        
    End If
    
    'Checks If Cell Below has the same ticker value.
    'If it does, it drops the year change, percent and volume in the summary table and adds 1 to the row
    'After dropping values in the summary table it turns volume and initial back to 0
        
    If Cells(i + 1, 1).Value <> Ticker Then
    Final = Cells(i, 6).Value
    YearlyChange = Final - Initial
    Percent = ((Final - Initial) / Initial) * 100
    Range("J" & Summary).Value = YearlyChange
    Range("I" & Summary).Value = Ticker
    Range("K" & Summary).Value = Percent
    Volume = Volume + Range("G" & i).Value 'That way it adds up the value for the row it's on instead of skipping.
    Range("L" & Summary).Value = Volume
    Summary = Summary + 1
    Volume = 0
    Initial = 0
        
        
    Else
    Volume = Volume + Range("G" & i).Value
        
    
    End If
    
Next i

'For k = 2 To 20

'Range("Q2").Value = Application.WorksheetFunction.Max(Range("K" & k).Value)
'Range("P2").Value = Cells(k, 1).Value

'Next k

End Sub
