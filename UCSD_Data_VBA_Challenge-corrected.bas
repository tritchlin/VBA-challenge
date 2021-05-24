Attribute VB_Name = "Module1"
Option Explicit

Sub StockSummary()

    ' Declare ws Variable
    Dim ws As Worksheet
    For Each ws In Worksheets

    Dim i As Double
    
    Dim rowcounter As Integer
    rowcounter = 0
    
    ' Set variable for ticker names
    Dim ticker As String

    ' Set variable for total amount change from annual opening price.
    ' Format positive with green background, negative with red background.
    Dim yearlychange As Double
    yearlychange = 0

    ' Set variable for total percentage change from annual opening price.
    ' Set formatting to show as Percentage
    Dim percentchange As Double
    percentchange = 0
    
    ' Set Variable for Greatest % Increase
    Dim Maximum As Double
    Dim IncreaseName As String
    
    ' Set Variable for Greatest % Decrease
    Dim Minimum As Double
    Dim DecreaseName As String
    
    ' Set variable for Greatest Total Volume
    Dim greatestvol As Double
    Dim VolName As String

    ' Identify and set variable for Last Row
    Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Identify and set variable for last column
    Dim LastColumn As Double
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Identify starting row for summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2

    ' Set starting value for total stock volume
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    ' Set headers for summary tables
    ws.Cells(1, LastColumn + 2).Value = "Ticker"
    ws.Cells(1, LastColumn + 3).Value = "Yearly Change"
    ws.Cells(1, LastColumn + 4).Value = "Percent Change"
    ws.Cells(1, LastColumn + 5).Value = "Total Stock Volume"
    ws.Cells(2, LastColumn + 7).Value = "Greastest % Increase"
    ws.Cells(3, LastColumn + 7).Value = "Greastest % Decrease"
    ws.Cells(4, LastColumn + 7).Value = "Greatest Total Volume"
    ws.Cells(1, LastColumn + 8).Value = "Ticker"
    ws.Cells(1, LastColumn + 9).Value = "value"
       
       
    ' Start For Loop
    

    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Set ticker name
            ticker = ws.Cells(i, 1).Value
            
            ' Calculate Total Stock Volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)
            
            ' Print ticker name
            ws.Cells(SummaryTableRow, LastColumn + 2).Value = ticker
            
            ' Print Total Stock Volume for ticker value
            ws.Cells(SummaryTableRow, LastColumn + 5).Value = TotalStockVolume

            
            ' Calculate and set Yearly Change
            yearlychange = ws.Cells(i, 6).Value - ws.Cells(i - rowcounter, 3).Value
            
            ws.Cells(SummaryTableRow, LastColumn + 3).Value = yearlychange
                If yearlychange >= 0 Then
                ws.Cells(SummaryTableRow, LastColumn + 3).Interior.ColorIndex = 4
                Else
                ws.Cells(SummaryTableRow, LastColumn + 3).Interior.ColorIndex = 3
                End If
                                    
            ' Calculate and set Percentage Change
                If (yearlychange = "0") Or (ws.Cells(i - rowcounter, 3).Value = "0") Then
                percentchange = "0"
            
                Else
                percentchange = (yearlychange / ws.Cells(i - rowcounter, 3).Value)
            
                End If
            
            ws.Cells(SummaryTableRow, LastColumn + 4).NumberFormat = "0.00%"
            ws.Cells(SummaryTableRow, LastColumn + 4).Value = percentchange
         
            ' Identify and set Greatest % Increase
            
            Maximum = ws.Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
            ws.Range("P2").NumberFormat = "0.00%"
            ws.Range("P2").Value = Maximum
            
            IncreaseName = ws.Evaluate("Index(I:I,Match(max(k:k),k:k,0))")
            ws.Range("O2").Value = IncreaseName
            
            ' Identify and set Greatest % Decrease
            
            Minimum = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
            ws.Range("P3").NumberFormat = "0.00%"
            ws.Range("P3").Value = Minimum
            
            DecreaseName = ws.Evaluate("Index(I:I,Match(min(k:k),k:k,0))")
            ws.Range("O3").Value = DecreaseName
            
            ' Identify and set Greatest Total Volume
            greatestvol = ws.Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
            ws.Range("P4").Value = greatestvol
            
            VolName = ws.Evaluate("Index(I:I,Match(max(L:L),L:L,0))")
            ws.Range("O4").Value = VolName
                                                                      
            ' Add one row to Summary Table Row count for next entry
            SummaryTableRow = SummaryTableRow + 1
            
            ' Reset the Total Stock Volume
            TotalStockVolume = 0
            rowcounter = 0
            
        Else
        
            ' Add to the running Stock Volume total
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            rowcounter = rowcounter + 1
               
        End If
        
    Next i
    
        '  Autofit columns in all worksheets

        ws.Cells.EntireColumn.AutoFit
    
    
    Next ws
    

End Sub



