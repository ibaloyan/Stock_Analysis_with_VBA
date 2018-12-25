VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Easy Solution VB code - HW2 VBA - Inna Baloyan May2018 Bootcamp
'Going through all worksheets
Sub WorksheetsLoop()

        ' Set CurrentWs as a worksheet object variable.
        Dim CurrentWs As Worksheet
        Dim Results_Sheet As Boolean
        Need_Summary_Table_Header = False       'Set Header flag
        
        ' Loop through all of the worksheets in the active workbook.
        For Each CurrentWs In Worksheets
        
            ' Set initial variable for holding the ticker name
            Dim Ticker_Name As String
            
            ' Set an initial variable for holding the total per ticker name
            Dim Total_Ticker_Volume As Double
            Total_Ticker_Volume = 0
            
            ' Keep track of the location for each ticker name
            ' in the summary table for the current worksheet
            Dim Summary_Table_Row As Long
            Summary_Table_Row = 2
            
            ' Set initial row count for the current worksheet
            Dim Lastrow As Long
            Dim i As Long
            
            Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

            ' For all worksheet except the first one, the Results
            If Need_Summary_Table_Header Then
                ' Set Titles for the Summary Table for current worksheet
                CurrentWs.Range("I1").Value = "Ticker"
                CurrentWs.Range("J1").Value = "Total Stock Volume"
            Else
                'This is the first, resulting worksheet, reset flag for the rest of worksheets
                Need_Summary_Table_Header = True
            End If
            
            ' Loop from the beginning of the current worksheet(Row2) till its last row
            For i = 2 To Lastrow
            
                ' Check if we are still within the same ticker name,
                ' if not - write results to summary table
                If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                
                    ' Set the ticker name, we are ready to insert this ticker name data
                    Ticker_Name = CurrentWs.Cells(i, 1).Value
                    
                    ' Add to the Ticker name total volume
                    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                    
                    ' Print the Ticker Name in the Summary Table, Column I
                    CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    
                    ' Print the Ticker Name in the Summary Table, Column J
                    CurrentWs.Range("J" & Summary_Table_Row).Value = Total_Ticker_Volume
                    
                    ' Add 1 to the summary table row count
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    Total_Ticker_Volume = 0
                    
                'Else - If the cell immediately following a row is still the same ticker name,
                'just add to Total Ticker Volume
                Else
                    ' Encrease the Total Ticker Volume
                    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                End If
                ' For debugging MsgBox (CurrentWs.Rows(i).Cells(2, 1))
          
            Next i
            
         Next CurrentWs
End Sub
