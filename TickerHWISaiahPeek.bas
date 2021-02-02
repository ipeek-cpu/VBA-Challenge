Attribute VB_Name = "Module1"
Sub TickerHW()

Dim WS As Worksheet
    
    For Each WS In ActiveWorkbook.Worksheets
    
    WS.Activate
        ' Declare Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Summary Headers
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock vol"
       
        Dim OPrice As Double
        Dim CPrice As Double
        Dim Yrly_Change As Double
        Dim Ticker As String
        Dim pcnt_change As Double
        Dim vol As Double
        Dim Row As Double
        Dim Column As Integer
        Dim i As Long
        vol = 0
        Row = 2
        Column = 1
       
       
        OPrice = Cells(2, Column + 2).Value
         ' Loop through  tickers
        
        For i = 2 To LastRow
         ' Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Set Ticker name
                Ticker = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker
                ' Set Close Price
                CPrice = Cells(i, Column + 5).Value
                ' Calculate Yearly Change
                Yrly_Change = CPrice - OPrice
                Cells(Row, Column + 9).Value = Yrly_Change
                
                ' Add Percent Change
                
                If (OPrice = 0 And CPrice = 0) Then
                    pcnt_change = 0
                
                ElseIf (OPrice = 0 And CPrice <> 0) Then
                    pcnt_change = 1
                
                Else
                    pcnt_change = Yrly_Change / OPrice
                    Cells(Row, Column + 10).Value = pcnt_change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Calculate Total Volume
                vol = vol + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = vol
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                OPrice = Cells(i + 1, Column + 2)
                ' reset the Volumn Total
                vol = 0
            'if tickers are thhe same
            Else
                vol = vol + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change for eachws
        
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Format Cell Colors
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        
            Next WS
        
End Sub
