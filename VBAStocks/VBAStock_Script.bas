Attribute VB_Name = "Module1"
Sub StockScript()

'ThisWorkbook.Active.Worksheet


'#*****************************
'# Step 1: Dim Them Variables
'#*****************************
Dim total_stock As Double
Dim output_row As Integer
Dim start_price As Double
Dim close_price As Double
Dim sheet As Worksheet
Dim numsheets As Integer

numsheets = Worksheets.Count
MsgBox (numsheets)



For Each sheet In Worksheets
    sheet.Activate

'#*****************************
'# Step 2: Initialize It
'#*****************************
    total_stock = 0
    output_row = 2 'for output tables
    
'#*****************************
'# Step 3: Sort, or the real way
'# to make the code really work
'#*****************************
'Sort Columns A thru G by Column 1, in asending order. and yes, there are headers
    Range("A:G").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

'#*****************************
'# Step 4: Make Output Table 1
'#*****************************
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

'#*****************************
'# Step 5: The Meat
'#*****************************


    For i = 2 To 70926
    '# Starting on row 1 of data (row 1), add stock volume from date
    '# to variable "total_stock", so long as "Ticker Value" is the same
        If Cells(i, 2) <> Cells(i + 1, 2) Then
            total_stock = total_stock + Cells(i, 7)
        Else:
            Range("A" & i & ":G" & i).Interior.Color = RGB(255, 200, 200) 'format them dupes
            Range("A" & i & ":G" & i).Font.Color = RGB(255, 0, 0)
        End If
    '#If date is less than the next row's date, save over the start_price field
    '# with this row's start price; if the date number is greater than the one
    '# before it, then set the closing price
        If Cells(i, 2) < Cells(i + 1, 2) And Cells(i, 1) = Cells(i + 1, 1) Then
            start_price = Cells(i, 3)
        Else
            If Cells(i, 2) > Cells(i - 1, 2) And i > 2 Then
                close_price = Cells(i, 6)
            End If
        End If
    '#If the next row has a different ticker, print current value
    '# for Ticker value, take saved close_price and subtract start_price for
    '# year change and print into output table; use those to print percent
    '# increase, but check first to make sure starting price wasn't zero.
    '# if 0, print message and ignore.
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            Cells(output_row, 9) = Cells(i, 1)
            Cells(output_row, 10) = (close_price - start_price)
            If start_price = 0 Then
                Cells(output_row, 11) = "Cannot Calculate" 'future: could take earliest non-zero price
            Else: Cells(output_row, 11) = (close_price - start_price) / start_price
            End If
        '# format cells in percent and your total stock amount for the variable you are working on
        Cells(output_row, 11).NumberFormat = "0.00%"
        Cells(output_row, 12) = total_stock
        
        '# reset and increment the things
        output_row = output_row + 1
        total_stock = 0
        start_price = 0
        end_price = 0
        
    End If
    
    Next i
    Range("I1:L" & i).Borders.Color = RGB(59, 56, 85)
    Range("I1:L1").Font.Bold = True
    
    ' # now make the fancy summary of the summary
    Dim greatup As Double
    Dim greatdown As Double
    Dim greatvol As Double
    Dim greatupname As String
    Dim greatdownname As String
    Dim greatvolname As String
    
    Cells(1, 15) = "Ticker"
    Cells(1, 16) = "Value"
    Cells(2, 14) = "Greatest % ^"
    Cells(3, 14) = "Greatest % v"
    Cells(4, 14) = "Greatest vol"

    greatup = 0
    greatdown = 0
    greatvol = 0
    greatupname = ""
    greatdownname = ""
    greatvolname = ""
    
    For i = 2 To output_row
        If Cells(i, 11) > greatup Then
            greatup = Cells(i, 11)
            greatupname = Cells(i, 9)
        End If
        If Cells(i, 11) < greatdown Then
            greatdown = Cells(i, 11)
            greatdownname = Cells(i, 9)

        End If
        If Cells(i, 12) > greatvol Then
            greatvol = Cells(i, 12)
            greatvolname = Cells(i, 9)
            
            
        End If
                    
        Cells(2, 16) = greatup '%^
        Cells(2, 15) = greatupname
        Cells(3, 16) = greatdown '%v; negatives are > 0/blank
        Cells(3, 15) = greatdownname
        Cells(4, 16) = greatvol
        Cells(4, 15) = greatvolname
            
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.Color = RGB(222, 102, 102)
        End If
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.Color = RGB(101, 201, 152)
        End If
    Next i
'# format the percents
    Range("P2:P3").NumberFormat = "0.00%"
    Range("P4").NumberFormat = "###,###"
    Range("N1:P4").Font.Color = RGB(28, 72, 144)
    Range("N1:P4").Interior.Color = RGB(180, 196, 218)
    Range("N1:P4").Borders.Color = RGB(59, 56, 85)
    Range("N1:P1").Font.Bold = True
    Columns("J:L").AutoFit
    
Next sheet
MsgBox ("Just VBA'ed " + Str(numsheets) + " Worksheets.")
End Sub




