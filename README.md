# VBA-Challenge

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
Dim SmallestPercent As Double
Dim GreatestVolume As Double

'This is a variable to make the for loop head to the last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Summary = 2
GreatestPercent = 0
SmallestPercent = 0
GreatestVolume = 0
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"






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
        
            If Range("J" & Summary).Value > 0 Then
                Range("J" & Summary).Interior.ColorIndex = 4
            ElseIf Range("J" & Summary).Value < 0 Then
                Range("J" & Summary).Interior.ColorIndex = 3
            End If
        
        Range("I" & Summary).Value = Ticker
        Range("K" & Summary).Value = Percent
    
            If Percent < 1 Then
                Percent = Range("K" & Summary).Value * 100
            End If
        
        Range("K" & Summary).NumberFormat = "0.00\%"
            
        'Takes the value of Percent and compares it to the value of GreatestPercent.
        'If the value is greater than it is documented. If it's less then it compares the value of
        'percent to SmallestPercent. If it is less, than it is documented.
            
            If Percent > GreatestPercent Then
                GreatestPercent = Percent
                Range("Q2").Value = GreatestPercent
                Range("P2").Value = Ticker
            ElseIf Percent < SmallestPercent Then
                SmallestPercent = Percent
                Range("Q3").Value = SmallestPercent
                Range("P3").Value = Ticker
            End If
        
        Range("Q2").NumberFormat = "0.00\%"
        Range("Q3").NumberFormat = "00.00\%"
        Volume = Volume + Range("G" & i).Value 'That way it adds up the value for the row it's on instead of skipping.
        Range("L" & Summary).Value = Volume
        
            If Volume > GreatestVolume Then
                GreatestVolume = Volume
                Range("Q4").Value = GreatestVolume
                Range("P4").Value = Ticker
            End If
        
        Summary = Summary + 1
        Volume = 0
        Initial = 0
        
        
    Else
        Volume = Volume + Range("G" & i).Value
        
    End If
        
Next i

End Sub


