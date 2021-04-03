Attribute VB_Name = "Module1"
Sub stockchallenge()

Dim ws As Worksheet
Set ws = ActiveSheet

For Each ws In Worksheets

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim YearOpen As Double
Dim YearClose As Double
Dim TotalVolume As LongLong
TotalVolume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim LastRow As Long

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
YearOpen = ws.Cells(2, 3).Value

    For i = 2 To LastRow
          
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        YearClose = ws.Cells(i, 6).Value
        YearlyChange = YearClose - YearOpen
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        ws.Range("L" & Summary_Table_Row).Value = TotalVolume
            If YearlyChange <> 0 Then
                PercentChange = YearlyChange / YearOpen
            Else
                PercentChange = 0
            End If
        ws.Range("K" & Summary_Table_Row).Value = PercentChange
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            If (YearlyChange < 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf (YearlyChange > 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
        Summary_Table_Row = Summary_Table_Row + 1
        YearOpen = ws.Cells(i + 1, 3).Value
        TotalVolume = 0
    Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    End If
    Next i
Next ws
    
End Sub

