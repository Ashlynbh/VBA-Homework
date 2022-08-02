VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub testing()

'variables

  
  Dim Ticker As String
  Dim Change As Double
  Dim Stdate As Double
  Dim Endate As Double
  Dim yearly_percentage As Double
  Dim Volume As Double
  Dim Max_Name As String
  Dim Min_Name As String
  Dim Max_Percent As Double
  Dim Min_Percent As Double
  Dim Max_Volume As Double
  Dim Volume_Name As String
  
  Endate = 0
  Stdate = 0

  


 ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
  ' Set beginning price value
  Stdate = Cells(2, 3).Value
  
  
' loop through column for ticker
 For i = 2 To 705714

' Check if we are still within the same ticker, if it is not...
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

 ' Find the ticker name
 Ticker = Cells(i, 1).Value
    
' Calculate yearly price change
 Endate = Cells(i, 6).Value
 Change = Endate - Stdate


'Calculations
If Stdate <> 0 Then
yearly_percentage = (Change / Stdate) * 100

End If


'Add the total volume
Volume = Volume + Cells(i, 7).Value

  
  ' Print the Ticker in the Summary Table
    Range("I" & Summary_Table_Row).Value = Ticker

  ' Print the yearly change in the Summary Table
    Range("J" & Summary_Table_Row).Value = Change
    
     'Print yearly price change as %
    Range("K" & Summary_Table_Row).Value = (CStr(yearly_percentage) & "%")
    
    'Print the total stock volume in the Summary table
   Range("L" & Summary_Table_Row).Value = Volume
    
    ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Set the next start price
Stdate = Cells(i + 1, 3).Value

 'Colour fill yearly price change
 If (Change > 0) Then
  Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4
  
 ElseIf (Change <= 0) Then
   Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
   
 End If

 'Reset Values
Volume = 0
Change = 0

Else
Volume = Volume + Cells(i, 7).Value
 
End If
Next i


'Challenge
For k = 2 To 2836

'Max increase
Max_Percent = WorksheetFunction.Max(Range("K2:K2836"))
   Cells(2, 17).Value = Max_Percent
   Cells(2, 17).NumberFormat = "00.0%"
    
'Print ticker
   Max_Name = WorksheetFunction.Match(Max_Percent, Range("K2:K2836"), 0)
   Cells(2, 16).Value = Cells(Max_Name + 1, 9)

'Min increase

 Min_Percent = WorksheetFunction.Min(Range("K2:K2836"))
    Cells(3, 17).Value = Min_Percent
    Cells(3, 17).NumberFormat = "00.0%"

'Print ticker
  Min_Name = WorksheetFunction.Match(Min_Percent, Range("K2:K2836"), 0)
  Cells(3, 16).Value = Cells(Min_Name + 1, 9)
    
 'Max Volume

 Max_Volume = WorksheetFunction.Max(Range("L2:L2836"))
 Cells(4, 17).Value = Max_Volume

    
 Volume_Name = WorksheetFunction.Match(Max_Volume, Range("L2:L2836"), 0)
 Cells(4, 16).Value = Cells(Volume_Name + 1, 9)
   


Next k



'Print
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'Headers

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Volume"

 

End Sub
