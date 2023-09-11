Sub MultiYrStock()
'print all similar items from Col A to a consolidated list in col I

Dim Ticker As String
Dim Cell As Long
Dim PrintCtr As Long 'Counter for printing sequentially one below the other
Dim YrOpen As Double
Dim YrClose As Double
Dim YrlyDiff As Double 'Variable to hold yearly change in Open and Close prices of a stock
Dim PercentChg As Double 'Variable to old the Percent Change
Dim TotStkVol As Double ' Total Volume of the stock
Dim MaxPInc As Double   'Greatest Percent Increase Var
Dim MaxPDec As Double   'Greatest Percent Decrease Var
Dim MaxTotVol As Double     'Greatest Total Volume Var
Dim Tempvar As Double
Dim TicMaxInc As Long 'Var to hold position for Ticker value of Greatest Perc. Increase
Dim TicMaxDec As String 'Var to hold position for Ticker value of Greatest Perc. Decrease
Dim TicMaxVol As String 'Var to hold position for Ticker value of Greatest Max Vol
Dim WS_Count As Integer
Dim I As Integer

'Setting WS_Count equal ot the number of
'worksheets in the active workbook
WS_Count = ActiveWorkbook.Worksheets.Count

'Beginning looping through worksheets
For I = 1 To WS_Count

    Ticker = Range("A2").Value
    Cell = 1
    PrintCtr = 1
    YrClose = Range("F2").Value
    YrlyDiff = 0
    PercentChg = 0
    TotStkVol = 0
    MaxPInc = 0
    MaxPDec = 0
    MaxTotVol = 0
    
    'Intermediate code check
    MsgBox "Chk 1"
    
    'Parse through the first column to check for unique ticker symbols
    'and compile them in Column I
Compare_Continue:
        YrOpen = Worksheets(I).Cells(Cell + 1, 3).Value
        'YrOpen = Cells(Cell + 1, 3).Value
        Do While Ticker = Worksheets(I).Range("A2").Offset(Cell - 1).Value
            TotStkVol = TotStkVol + Worksheets(I).Cells(Cell + 1, 7).Value
            Cell = Cell + 1
        Loop
        
             
        YrClose = Worksheets(I).Cells(Cell, 6).Value
        YrlyDiff = YrClose - YrOpen
        PercentChg = YrlyDiff / Abs(YrOpen)
        'Post Yearly Change in Column J for each row
        Worksheets(I).Range("J2").Offset(PrintCtr - 1).Value = YrlyDiff
        'Color Yearly Change cell as per value
        If YrlyDiff < 0 Then
            Worksheets(I).Range("J2").Offset(PrintCtr - 1).Interior.Color = RGB(255, 0, 0)
        ElseIf YrlyDiff > 0 Then
            Worksheets(I).Range("J2").Offset(PrintCtr - 1).Interior.Color = RGB(0, 255, 0)
        End If
        'Post Percent Change in Column K for each row
        Worksheets(I).Range("K2").Offset(PrintCtr - 1).Value = PercentChg
        'Post Total Volume per Ticker in Column L
        Worksheets(I).Range("L2").Offset(PrintCtr - 1).Value = TotStkVol
        'Print current Ticker in Column I
        Worksheets(I).Cells(PrintCtr + 1, 9).Value = Ticker
        PrintCtr = PrintCtr + 1
        'Set Ticker to the new value from Row below
        Ticker = Worksheets(I).Range("A2").Offset(Cell).Value
        'Set Total Stock Volume value to the first one of the next Ticker
        TotStkVol = Worksheets(I).Range("L2").Offset(Cell).Value
        'Check if next cell down is empty Row is over.
        'If not, go back and repeat for next Ticker.
        'Else, skip program control to finish
        If (Not IsEmpty(Worksheets(I).Range("A2").Offset(Cell - 1).Value)) Then
            GoTo Compare_Continue
        Else: GoTo ExitCompare
        End If
ExitCompare:
    'Reset Cell variable value to 1
    Cell = 1
    Do While Not (IsEmpty(Worksheets(I).Range("K2").Offset(Cell - 1).Value))
    'Tempvar = Range("K2").Offset(Cell - 1).Value
            'Parsing for Greatest Percent Increase
            If MaxPInc < Worksheets(I).Range("K2").Offset(Cell - 1).Value Then
                MaxPInc = Worksheets(I).Range("K2").Offset(Cell - 1).Value
                TicMaxInc = Cell
            End If
            'Parsing for Greatest Percent Decrease
            If MaxPDec > Worksheets(I).Range("K2").Offset(Cell - 1).Value Then
                MaxPDec = Worksheets(I).Range("K2").Offset(Cell - 1).Value
                TicMaxDec = Cell
            End If
            'Parsing for Greatest Total Volume
            If MaxTotVol < Worksheets(I).Range("L2").Offset(Cell - 1).Value Then
                MaxTotVol = Worksheets(I).Range("L2").Offset(Cell - 1).Value
                TicMaxVol = Cell
            End If
            
            Cell = Cell + 1
    Loop
    
    
     'Intermediate sheet check
    MsgBox "Chk 2"
    
    
    'Print Ticker and corresponding Greatest Percent Increase
    Worksheets(I).Cells(2, 15).Value = Worksheets(I).Range("I2").Offset(TicMaxInc - 1).Value
    Worksheets(I).Cells(2, 16).Value = MaxPInc
    
    'Print Ticker and corresponding Greatest Percent Decrease
    Worksheets(I).Cells(3, 15).Value = Worksheets(I).Range("I2").Offset(TicMaxDec - 1).Value
    Worksheets(I).Cells(3, 16).Value = MaxPDec
    
    'Print Ticker and corresponding Greatest Total Volume
    Worksheets(I).Cells(4, 15).Value = Worksheets(I).Range("I2").Offset(TicMaxVol - 1).Value
    Worksheets(I).Cells(4, 16).Value = MaxTotVol

MsgBox ActiveWorkbook.Worksheets(I).Name

Next I

End Sub






