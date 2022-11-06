VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker():
Dim ticker As String
Dim tickerCurrent As String
Dim tickerPrevious As String
Dim tickerNext As String
Dim firstPrice As Double
Dim lastPrice As Double
Dim Change As Double
Dim percentChange As Double
Dim Z As Integer
Dim volumeCounterA As Double
Dim volumeTotal As Long
Dim volumeFirst As Long
Dim Sheet1 As Worksheet
Dim Sheet2 As Worksheet
Dim Sheet3 As Worksheet
Dim tickerHeader, changeHeader, percentChangeHeader, volumeHeader As String
'__________________________________________________________________________
'setting values
'creating the headers for the additional columns
tickerHeader = "Ticker"
changeHeader = "Change"
percentChangeHeader = "Percent Change"
volumeHeader = "Total Volume"

'set the worksheets as sheet 1-3
Set Sheet1 = Worksheets("2018")
Set Sheet2 = Worksheets("2019")
Set Sheet3 = Worksheets("2020")

'set z value for the data the rows will be populated with
Z = 2
'__________________________________________________________________________
'find last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'__________________________________________________________________________
'if statements
'start loops for values
For i = 2 To lastrow

    'setting the ticker to check values
    tickerPrevious = Range("A" & i - 1).Value
    tickerCurrent = Range("A" & i).Value
    tickerNext = Range("A" & i + 1).Value
    
        'configuring the ticker
        If tickerCurrent <> tickerPrevious Then
        
            'reseting the volume count
            volumeCounterA = 0
            'set tickerCurrent as ticker
            ticker = tickerCurrent
            'grab ticker value
            Range("I" & Z).Value = ticker
            'set first price
            firstPrice = Range("C" & i).Value
            'set initial volumeCounterA value
            volumeCounterA = Range("G" & i).Value
        
        'If the tickerCurrent is equal to tickerPrevious then add the row's value to volumeCounterA
        Else
        
            volumeCounterA = volumeCounterA + Range("G" & i).Value
        
        End If
        
        'find and set last price, and values for the new columns
        If tickerNext <> tickerCurrent Then
            
            lastPrice = Range("F" & i).Value
            Change = (lastPrice - firstPrice)
            percentChange = (lastPrice - firstPrice) / firstPrice
            Range("J" & Z).Value = Change
            Range("K" & Z).Value = percentChange
            Range("K" & Z).NumberFormat = "0.00%"
            Range("L" & Z) = volumeCounterA
            Z = Z + 1
            volumeCounterA = 0

        End If
'__________________________________________________________________________
'formatting
    'Format "Change" column by color red, green, and black
    If Change = 0 Then
        'black
    If Not (Range("J" & 1).Value) And (Change > 0) Then
        'green
        Range("J" & Z - 1).Interior.ColorIndex = 4
        
    End If
        'red
    If Not (Range("J" & 1).Value) And (Change < 0) Then
        Range("J" & Z - 1).Interior.ColorIndex = 3
        
    End If
             
    'set column headers, select next sheet, reset i and z
    If i = lastrow Then
            
        Range("I" & 1).Value = tickerHeader
        Range("J" & 1).Value = changeHeader
        Range("K" & 1).Value = percentChangeHeader
        Range("L" & 1).Value = volumeHeader
        ActiveSheet.Next.Select
        i = 1
        Z = 2
        
    End If
      
Next i

End Sub
