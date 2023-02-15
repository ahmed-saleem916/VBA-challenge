VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Data()

    Dim TickerRow As String
    Dim NumberofRow As Integer
    Dim FinalRow As Long
    Dim s As Long
    Dim ws As Worksheet
    
        For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    Total_Volume = 0
    i_pointer = 2
    c_pointer = 2
    
    
    FinalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For s = 2 To FinalRow
        If ws.Cells(s, "A").Value = ws.Cells(s + 1, "A").Value Then
            Total_Volume = Total_Volume + ws.Cells(s, "G").Value
               
        Else
            Total_Volume = Total_Volume + ws.Cells(s, "G").Value
            
            openprice = ws.Cells(c_pointer, "C").Value
            closeprice = ws.Cells(s, "F").Value
            yearchange = closeprice - openprice
            PercentChange = yearchange / openprice * 100
            
            ws.Cells(i_pointer, "I").Value = ws.Cells(s, "A").Value
            ws.Cells(i_pointer, "J").Value = yearchange
            ws.Cells(i_pointer, "K").Value = "%" & PercentChange
            ws.Cells(i_pointer, "L").Value = Total_Volume
            
            If yearchange > 0 Then
                ws.Cells(i_pointer, "J").Interior.ColorIndex = 4
            Else
                ws.Cells(i_pointer, "J").Interior.ColorIndex = 3
        End If
        
    
            Total_Volume = 0
            c_pointer = s + 1
            i_pointer = i_pointer + 1
        End If

            
            
    Next s
    Next ws

End Sub
