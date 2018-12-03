Option Explicit

Sub aatally()

Dim ticker As String 'Ticker variabe
Dim total As Double 'Volume adding variable
Dim i As Long 'Cell looping variable
Dim count As Integer 'number of worksheets
Dim output As Integer 'Values for the summary table
Dim LastRow As Long 'Last row of data variable
Dim sht As Integer 'Worksheet counting variable
Dim first As Double 'first opening value for each ticker
Dim last As Double 'last closing value for each ticker
Dim chg As Double 'percent change

count = ActiveWorkbook.Worksheets.count 'Get the worksheet count so we can loop through them

For sht = 1 To count
    Worksheets(sht).Activate 'focus on the sheet
        With ActiveSheet
        
        'Set up the headings for the summary data
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"

'Populate the starting variables

            LastRow = Range("A1").CurrentRegion.Rows.count 'Find the last row of data on the sheet
            ticker = Cells(2, 1).Value 'Set the default value to the first cell in column A
            total = 0 'Set the default tally to zero
            first = Cells(2, 3).Value
            output = 2 'reset the value for where output goes

'Loop through the ticker volumn, grabbing opening/closing values to compare

        For i = 2 To LastRow + 2  'Making sure to capture the last row
            
             If Cells(i, 1).Value <> ticker Then 'Have we gotten to the next ticker? If so, get the closing value of the previous
               
                last = Cells(i - 1, 6).Value 'The closing value
                Cells(output, 9).Value = ticker 'Put the ticker name in the summary columns
                Cells(output, 12).Value = total   'Put the total for the ticker in the summary columns
                'Calculate the percent change
                
                Cells(output, 10).Value = last - first
                If first <> 0 Then
                chg = (last - first) / (first) 'percent change
                'Put the percent change in the Summary column, change to percent type
                Cells(output, 11).Value = chg
                Cells(output, 11).NumberFormat = "0.00%"


                End If
                Select Case True
                Case last - first > 0:
                        Cells(output, 10).Interior.Color = 65285
                Case last - first < 0:
                        Cells(output, 10).Interior.Color = 255
                Case last - first = 0:
                        Cells(output, 10).Interior.TintAndShade = 0
                End Select
                
                ticker = Cells(i, 1).Value  'Now reset the ticker variable to the new value
                output = output + 1 'increment the output row
                total = Cells(i, 7).Value 'Now start totalling volumn with the first new ticker
                first = Cells(i, 3).Value 'The opening value
             
             Else 'just keep adding the volume until the ticker changes.
               
                    total = total + Cells(i, 7).Value
          End If
             
        Next i
    
    '_________________________________________________________________
    'Enter the ticker w/ maximum increase and value in the summary
    'Declare the variables
    Dim max_inc As Double
    Dim inc_ticker As Integer
    Dim max_dec As Double
    Dim dec_ticker As Integer
    Dim max_vol As Double
    Dim vol_ticker As Integer
    LastRow = Range("I1").CurrentRegion.Rows.count 'Find the last row of data on the sheet
    '_________________________________________________________________
    max_inc = WorksheetFunction.Max(Range("K2:K" & LastRow))
    Range("Q2").Value = max_inc
    Range("Q2").NumberFormat = "0.00%"
    inc_ticker = WorksheetFunction.Match(max_inc, Range("K2:K" & LastRow), 0)
    Range("P2").Value = Cells(inc_ticker + 1, 9)
    
    max_dec = WorksheetFunction.Min(Range("K2:K" & LastRow))
    Range("Q3").Value = max_dec
    Range("Q3").NumberFormat = "0.00%"
    dec_ticker = WorksheetFunction.Match(max_dec, Range("K2:K" & LastRow), 0)
    Range("P3").Value = Cells(dec_ticker + 1, 9)
    
    max_vol = WorksheetFunction.Max(Range("L2:L" & LastRow))
    Range("Q4").Value = max_vol
    vol_ticker = WorksheetFunction.Match(max_vol, Range("L2:L" & LastRow), 0)
    Range("P4").Value = Cells(vol_ticker + 1, 9)
    
    'Autofit the new columns
    Worksheets(sht).Columns("I:Q").AutoFit
    
    End With
Next sht


End Sub
