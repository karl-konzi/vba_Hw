VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 Sub stockMarket()
 For Each ws In Worksheets
 Dim summaryRow As Integer
 Dim totalValue, yearChange, percentChange As Double
 Dim ticker As String
 totalValue = 0
 summaryRow = 2
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Value"

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        totalValue = totalValue + ws.Cells(i, 7).Value
        yearChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
        
        If ws.Cells(i, 3).Value = 0 Then
            percentChange = 0
        Else
            percentChange = yearChange / ws.Cells(i, 3).Value
        End If
        If yearChange < 0 Then
        ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
        End If
        
        ws.Range("I" & summaryRow).Value = ticker
        ws.Range("J" & summaryRow).Value = yearChange
        ws.Range("K" & summaryRow).Value = percentChange
        ws.Range("K" & summaryRow).NumberFormat = "0.00%"
        ws.Range("L" & summaryRow).Value = totalValue
        totalValue = 0
        summaryRow = summaryRow + 1
   Else
   
    yearChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
    totalValue = totalValue + ws.Cells(i, 7).Value
    
    End If
  
  Next i
 
 Next ws

End Sub
