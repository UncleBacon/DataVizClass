Option Explicit
Sub DoStuff() 'Excel VBA to extract the unique items.
'turn off background excel functions to speed up code
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Dim ws As Worksheet
    Dim UItem As Collection
    Dim rng, crit, sumcrit, datecrit, opnrng, clsrng, VarC, VarF, VarI,VarJ, VarK, VarL As Range
    Dim i, mxday, mnday, mxvol, mxchng, mnchng, lrow As Long
    Dim opn, cls As Double
    Dim StartTime As Double
    Dim SecondsElapsed As Double

    'Remember time when macro starts
    StartTime = Timer
    For Each ws In Sheets
        'set ranges to variables
        Set crit = ws.Range("A2", ws.Range("A" & Rows.Count).End(xlUp))
        Set sumcrit = ws.Range("G2", ws.Range("G" & Rows.Count).End(xlUp))
        Set datecrit = ws.Range("B2", ws.Range("B" & Rows.Count).End(xlUp))
        Set opnrng = ws.Range("C2", ws.Range("C" & Rows.Count).End(xlUp))
        Set clsrng = ws.Range("F2", ws.Range("F" & Rows.Count).End(xlUp))
        Set UItem = New Collection
        
        'write data titles
        ws.Range("I1,O1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("N2") = "Greatest Increase"
        ws.Range("N3") = "Greatest Decrease"
        ws.Range("N4") = "Greatest Volume"
        ws.Range("P1") = "Values"
        
        'put unique tickers into array
        On Error Resume Next 'in case of empty cells
        For Each rng In ws.Range("A2", ws.Range("A" & Rows.Count).End(xlUp))
        UItem.Add CStr(rng), CStr(rng)
        Next

        'iterate through for as many times as there are unique tickers
        For i = 1 To UItem.Count
        set VarC = ws.range("C"&i+1) 
        set VarF = ws.range("F"&i+1)
        Set VarI = ws.Range("I" & i + 1)
        set VarJ = ws.range("J"&i+1)
        set VarK = ws.range("K"&i+1)
        set VarL = ws.range("L"&i+1)

        'write unique tickers to I column and sum of volume in L column 
        VarI = UItem(i)
        VarL = Application.WorksheetFunction.SumIf(crit, UItem(i), sumcrit)
       
       ' find first and last day and put in variables
        mxday = Application.WorksheetFunction.MaxIfs(datecrit, crit, VarI)
        mnday = Application.WorksheetFunction.MinIfs(datecrit, crit, VarI)
       
       'put closing and opening prices into variables
        If crit = VarI And datecrit = mxday Then
            cls = VarC.Value
        
        If crit = VarI And datecrit = mnday Then
            opn = VarF.Value
        End If
        End If
       'calculate difference and % change for each and place in columns J and K
        VarJ = cls - opn
        VarK = (cls - opn) / opn
        
        
        'conditional formating for positive or negative difference
        If VarJ > 0 Then
            VarJ.Interior.Color = vbGreen
        ElseIf VarJ < 0 Then
            VarJ.Interior.Color = vbRed
        End If
        
        'get Max Volume
        ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L2", ws.Range("L" & Rows.Count).End(xlUp)))
        
        'get greatest increase
        ws.Range("p2") = Application.WorksheetFunction.Max(ws.Range("k2", ws.Range("k" & Rows.Count).End(xlUp)))
        
        'get greatest decrease
        ws.Range("p3") = Application.WorksheetFunction.Min(ws.Range("k2", ws.Range("k" & Rows.Count).End(xlUp)))
        
        ' put ticker symbols next to summary data
        If VarL = ws.Range("P4") Then
            ws.Range("O4") = VarI
        ElseIf VarK = ws.Range("P3") Then
            ws.Range("O3") = VarI
        ElseIf VarK = ws.Range("P2") Then
            ws.Range("O2") = VarI
        End If
        Next
        
        ' format columns/cells
        lrow = Cells(Rows.Count, 11).End(xlUp).Row
        ws.Columns("J:P").NumberFormat = "#,#0.0#" 'set numbers to include comma if over 1,000 and at min show value as 0.0
        ws.Range("P2:P3,K2:K" & lrow).NumberFormat = "0.00%" ' set select cells/columns to percent
        ws.Columns("i:p").AutoFit ' autofit columns
    Next ws
'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user how long it took code to run in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
' turn on background excel functions
ActiveSheet.DisplayPageBreaks = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
End Sub



