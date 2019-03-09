Attribute VB_Name = "Module2"

Sub tickerNEW()

'Identify loop element for worksheets
Dim ws As Worksheet

    'Set loop to run through each worksheet in the workbook
    For Each ws In Worksheets
    ws.Activate

    'Name the table that holds the ticker volume summary information
    Dim TickerName As String
    Dim TotalVolume As String

    Cells(1, 9).Value = "TickerName"
    Cells(1, 10).Value = "TotalVolume"

    'Assign a variable to hold the ticker names
    Dim ticker_name As String

    'Assign a variable to hold the volume totals and set it to 0
    Dim volume_total As Double
    volume_total = 0

    'Name the space where the first rows of data will drop in the new table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    'Assign a variable to the last row in the ws and find it
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
        'Loop through all ticker names in the ws
        For I = 2 To LastRow

            'Look to see if the ticker name changes
             If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
             ticker_name = Cells(I, 1).Value

             'If it does, add the volume totals together
             volume_total = volume_total + Cells(I, 7).Value

             'Drop the ticker name into the new Table
             Range("I" & Summary_Table_Row).Value = ticker_name

             'Drop the volume total into the new table
              Range("J" & Summary_Table_Row).Value = volume_total

             'Add another row into the new table
              Summary_Table_Row = Summary_Table_Row + 1
      
            'Reset the volume total
             volume_total = 0

            'If the ticker name doesn't change, then keep adding to the volume total
            Else: volume_total = volume_total + Cells(I, 7).Value

            End If

         Next I
  
  Next ws

End Sub
