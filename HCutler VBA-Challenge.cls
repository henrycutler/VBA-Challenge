VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AlphabeticalTesting():

    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim Ticker As String
  
    Dim TickerOpen As Single
    PriceOpen = Cells(2, 3).Value
  
    Dim Vol As Double
    Vol = 0
  
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
     
     For i = 2 To LastRow
     
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Vol = Vol + Cells(i, 7).Value
            Range("I" & SummaryTableRow).Value = Ticker
            Range("L" & SummaryTableRow).Value = Vol
            
            Range("J" & SummaryTableRow).Value = Cells(i, 6).Value - PriceOpen
            
            Range("K" & SummaryTableRow).Value = (Cells(i, 6).Value - PriceOpen) / PriceOpen
            Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            
            SummaryTableRow = SummaryTableRow + 1
            Vol = 0
            PriceOpen = Cells(i + 1, 3).Value
        Else
            Vol = Vol + Cells(i, 7).Value
    
        End If
            
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    ElseIf Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    ElseIf Cells(i, 10).Value = 0 Then
        Cells(i, 10).Interior.ColorIndex = 2
    End If
    
    
    Next i

End Sub
