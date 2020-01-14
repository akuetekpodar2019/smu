Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock():

'Set the variables

Dim x, y As Long
Dim total_price As Double
Dim ticker As String
Dim Row As Long
Dim Year_change As Double
Dim Percent_change As Double
Dim Open_price As Double
Dim Close_price As Double
Dim Upgrade_price As Long

'Loop through All worksheets

For Each Sh In Worksheets

'Add header

Sh.Range("I1").Value = "Ticker"
Sh.Range("J1").Value = "Yearly Change"
Sh.Range("K1").Value = "Percent Change"
Sh.Range("L1").Value = "Total Stock Value"

'Set total
total_price = 0
y = 2
Upgrade_price = 2

'Set last row
Row = Sh.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through stock data
For x = 2 To Row

        If Sh.Range("A" & x + 1).Value = Sh.Range("A" & x).Value Then
        
            total_price = total_price + Sh.Range("G" & x).Value
            
        Else
            
            ticker = Sh.Range("A" & x).Value
            
            'Yearly change and Percent change
            
            Open_price = Sh.Range("C" & Upgrade_price)
            Close_price = Sh.Range("F" & x)
            Year_change = Close_price - Open_price
            
            'Percent change
            If Open_price = 0 Then
                Percent_change = 0
            Else
                Percent_change = Year_change / Open_price
            End If
            
            
            'Display cells
            Sh.Range("I" & y).Value = ticker
            Sh.Range("L" & y).Value = total_price + Sh.Range("G" & x).Value
            Sh.Range("J" & y).Value = Year_change
            Sh.Range("K" & y).Value = Percent_change
            Sh.Range("K" & y).NumberFormat = "0.00%"
            
            y = y + 1
            total_price = 0
            Upgrade_price = x + 1
            
        End If
Next x
Next Sh
End Sub
