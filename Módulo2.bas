Attribute VB_Name = "Módulo2"
Sub Stock3()
    Dim wb As Workbook
    Dim sheets As Variant
    Dim sheet As Variant
    
    Set wb = ActiveWorkbook
    sheets = Array("2018", "2019", "2020")
    
    For Each sheet In sheets
        wb.sheets(sheet).Activate
        primer_parte
        segunda_parte
        
    Next sheet
    
End Sub


Sub primer_parte()

    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim z As Long
    Dim ticker As String
    Dim open1 As Double
    Dim total_volume As Double
    Dim close1 As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    Set ws = ActiveSheet
    i = 2
    j = 2
    z = 1
    
    While ws.Cells(z, 1).Value <> ""
        z = z + 1
    Wend
    
    While i < z
        ticker = ws.Cells(i, 1).Value
        open1 = ws.Cells(i, 3).Value
        total_volume = 0
        
        While ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            i = i + 1
        Wend
        
        close1 = ws.Cells(i - 1, 6).Value
        yearly_change = close1 - open1
        percent_change = yearly_change / open1
        
        ws.Cells(j, 9).Value = ticker
        ws.Cells(j, 10).Value = yearly_change
        ws.Cells(j, 11).Value = percent_change
        ws.Cells(j, 12).Value = total_volume
        ws.Cells(j, 11).NumberFormat = "0.00%"
        
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        j = j + 1
    Wend
End Sub

Sub segunda_parte()
    Dim ws As Worksheet
    Dim count1 As Double
    Dim count2 As Double
    Dim count3 As Double
    Dim ticker1 As String
    Dim ticker2 As String
    Dim ticker3 As String
    Dim u As Long
    
    Set ws = ActiveSheet
    
    u = 2
    
    While ws.Cells(u, 11).Value <> ""
        If u = 2 Then
            count1 = ws.Cells(u, 11).Value
            ticker1 = ws.Cells(u, 9).Value
        Else
            If ws.Cells(u, 11).Value > count1 Then
                count1 = ws.Cells(u, 11).Value
                ticker1 = ws.Cells(u, 9).Value
            End If
        End If
        
        u = u + 1
    Wend
    
    ws.Cells(2, 17).Value = count1
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(2, 16).Value = ticker1
    
    u = 2
    
    While ws.Cells(u, 11).Value <> ""
        If u = 2 Then
            count2 = ws.Cells(u, 11).Value
            ticker2 = ws.Cells(u, 9).Value
        End If
        
        If count2 > ws.Cells(u, 11).Value Then
            count2 = ws.Cells(u, 11).Value
            ticker2 = ws.Cells(u, 9).Value
        End If
        
        u = u + 1
    Wend
    
    ws.Cells(3, 17).Value = count2
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = ticker2
    
    u = 2
    
    While ws.Cells(u, 12).Value <> ""
        If u = 2 Then
            count3 = ws.Cells(u, 12).Value
            ticker3 = ws.Cells(u, 9).Value
        End If
        
        If count3 < ws.Cells(u, 12).Value Then
            count3 = ws.Cells(u, 12).Value
            ticker3 = ws.Cells(u, 9).Value
        End If
        
        u = u + 1
    Wend
    
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 17).Value = count3
    ws.Cells(4, 17).NumberFormat = "##0.00E+0"
    ws.Cells(4, 16).Value = ticker3
End Sub

