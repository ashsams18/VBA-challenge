VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabeticaltesting()
  
  
    Dim ticker As String
    Dim Start As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim vol_Total As Double
    Dim Per_Change As Double
    Dim Year_Change As Double
    Dim start_ticker_year_row As Long
    Dim end_ticker_year_row As Long
    Dim i As Long
      
            

For Each ws In Worksheets
    
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 12).Value = "Total Stock volume"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
     
    vol_Total = 0
    Start = 2
   
          
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    start_ticker_year_row = 2
      
   For i = 2 To LastRow
   
   
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
    end_ticker_year_row = i
   
        Year_Change = ws.Cells(i, 6).Value - ws.Cells(start_ticker_year_row, 3).Value
        
        If ws.Cells(start_ticker_year_row, 3).Value = 0 Then
        Per_Change = 0
        Else
        Per_Change = Year_Change / ws.Cells(start_ticker_year_row, 3).Value
   
   
        End If
        
        ws.Range("K" & Start).NumberFormat = "0.00%"


        
        ws.Range("J" & Start).Value = Year_Change
        ws.Range("K" & Start).Value = Per_Change
        

        start_ticker_year_row = i + 1
    
        ticker = ws.Cells(i, 1).Value
        
        vol_Total = vol_Total + ws.Cells(i, 7).Value
        
        If Year_Change > 0 Then

        ws.Range("J" & Start).Interior.ColorIndex = 4

        ElseIf Year_Change < 0 Then

        ws.Range("J" & Start).Interior.ColorIndex = 3
        
        End If

        
        ws.Range("I" & Start).Value = ticker
        ws.Range("L" & Start).Value = vol_Total
        
        Start = Start + 1
        vol_Total = 0
        
    Else
    
        vol_Total = vol_Total + ws.Cells(i, 7).Value
        
    End If
          
   Next i
    
    
Next ws



End Sub
