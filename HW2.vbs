Sub Homework()

'Create titles for columns
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"

'Variables/Dims
  ' Set an initial variable for holding the Ticker Number
  Dim Ticker As String

  ' Set an initial variable for Open
  Dim Yearly_Open As LongLong

  ' Set a Yearly change
  Dim Tupac as Double
  Tupac = 0

  ' Set an initial variable for holding the Percentage Change
  Dim Percent_Change As Double
  Percent_Change = 0

  ' Set an initial variable for holding the total Volume
  Dim Total_Volume As LongLong
  Total_Volume = 0

  ' Keep track of the location for each ticker in a summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Find last row of data
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'HERE WE LOOOOP

  ' Loop through all Ticker Numbers
  For i = 2 To lastRow

  ' Find yearly open for each value
  Yearly_Open = Cells(Summary_Table_Row, 3).Value

    ' Check if we are still within the Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Name
      Ticker = Cells(i, 1).Value

      ' Add to the Stock Volume
      Total_Volume = Total_Volume + Cells(i, 7).Value

      ' Find Diff from open and close
      Tupac = (Cells(i, 6).Value - Yearly_Open)

      ' Find percent change
      Percent_Change = (Tupac / Yearly_Open)

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      'Print Yearly Change
      Range("J" & Summary_Table_Row).Value = Tupac

      'Print Percent Change and apply Percentage
      Range("K" & Summary_Table_Row).Value = Percent_Change
      Range("K" & Summary_Table_Row).Style = "Percent"

      ' Print the Total_Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' If the cell immediately following a row is the same Ticker...
      Else

        ' Add to the Total Volume
        Total_Volume = Total_Volume + Cells(i, 7).Value

    End If


    Next i

'Declare Cells for formatting
Change_end = Cells(Rows.Count, 11).End(xlUp).Row

'Add looooop
  for i = 2 to Change_end

    If Cells(i, 11).Value >= 0 Then
      Cells(i, 11).Interior.ColorIndex = 4

    Else
      Cells(i, 11).Interior.ColorIndex = 3

  End If

Next i

End Sub
