Attribute VB_Name = "Module1"
Sub ForLoop()
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

' Add Heading for summary
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
        
' Define Variables
 Dim OpenPrice As Double
 Dim ClosePrice As Double
 Dim YearlyChange As Double
 Dim TickerName As String
 Dim PercentChange As Double
 Dim TickerVolume As Double
 TickerVolume = 0
 Dim Row As Double
 Row = 2
 Dim Column As Integer
 Column = 1
 Dim i As Long
        
For i = 2 To LastRow

' Determine the Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
' Determine Ticker name
    TickerName = WS.Cells(i, 1).Value
    Cells(Row, Column + 8).Value = TickerName

' Define Initial Prices
    OpenPrice = Cells(2, Column + 2).Value
    ClosePrice = Cells(i, Column + 5).Value
    YearlyChange = ClosePrice - OpenPrice
    Cells(Row, Column + 9).Value = YearlyChange
    
 ' Define Percent Change
    If YearlyOpen = 0 Then
    PercentChange = 0
    
    Else
    PercentChange = YearlyChange / YearlyOpen
    Cells(Row, Column + 10).Value = PercentChange
    Cells(Row, Column + 10).NumberFormat = "0.00%"
    
    End If
' Refefine Total Volumne
    TickerVolume = TickerVolume + Cells(i, Column + 6).Value
    Cells(Row, Column + 11).Value = TickerVolume
    Row = Row + 1
    OpenPrice = Cells(i + 1, Column + 2)
    TickerVolume = 0
            
    Else
    Volume = Volume + Cells(i, Column + 6).Value
    End If
    Next i
        
        ' Determine the Last Row of Yearly Change per WS
      YearlyChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' Set the Cell Colors
        For j = 2 To YearlyChangeLastRow
                       If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j

        
    Next WS
        
End Sub

