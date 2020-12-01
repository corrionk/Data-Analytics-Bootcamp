VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ticker()

  ' Set an initial variable
  Dim Ticker As String

  ' Set an initial variable for holding the total per stock
  Dim Volume_Total As Double
  Volume_Total = 0
  
  Dim Year_Open As Double
  Year_Open = 0
  
  Dim Year_Close As Double
  Year_Close = 0
  
  Dim Year_Change As Double
  Year_Change = 0
  
  Dim Percent_Change As Double
  Percent_Change = 0

  ' Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
    Cells(1, "I").Value = "Stock Symbol"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"

  ' Loop through all stocks
  For i = 2 To 705714

    ' Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Year_Open = Cells(i, 3).Value
      
    End If
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker = Cells(i, 1).Value
      Range("I" & Summary_Table_Row).Value = Ticker
      
      Volume_Total = Volume_Total + Cells(i, 7).Value
      Range("L" & Summary_Table_Row).Value = Volume_Total
      Range("L" & Summary_Table_Row).NumberFormat = "1,000"
      
      Year_Close = Cells(i, 6).Value
      Year_Change = Year_Open - Year_Close
      Range("J" & Summary_Table_Row).Value = Year_Change
      Range("J" & Summary_Table_Row).NumberFormat = "$0.00"
      
        If Year_Change >= 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
    
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        If Year_Open = 0 And Year_Close = 0 Then
                    'Starting at zero and ending at zero will be a zero increase.  Cannot use a formula because
                    'it would be dividing by zero.
                    Percent_Change = 0
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ElseIf Year_Open = 0 Then
                    'If a stock starts at zero and increases, it grows by infinite percent.
                    'Because of this, we only need to evaluate actual price increase by dollar amount and therefore put
                    '"New Stock" as percent change.
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                Else
                    Percent_Change = Year_Change / Year_Open
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                End If
    
    
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Totals
      Volume_Total = 0
      Year_Change = 0
      Year_Open = 0
      Year_Close = 0
      Percent_Change = 0
      
   End If
   
Next i
   
 

End Sub

