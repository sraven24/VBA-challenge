Attribute VB_Name = "Module1"
Sub Wrkbookupdate()
    
    'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    'set the code to run on each worksheet of the workbook
    'set the variable for the active worksheet
    
    Dim actwrksht As Worksheet
    Application.ScreenUpdating = False
    
    'loop to activate code for each worksheet
    
    For Each actwrksht In Worksheets
        actwrksht.Select
        'call the subroutine below for each sheet
        
        Call stockreport
    Next
    Application.ScreenUpdating = True


End Sub


Sub stockreport()
'sub to alter stock worksheet

    'Set upa variable to hold the stock name
    Dim stock_name As String
       
    'Set an initial variable for holding the total per stock
    Dim stock_total As Double
    
    ' Set intitial variables for open and close stock price comparisons
        
    Dim open_price As Double
    
    Dim close_price As Double
    
    Dim diff As Double
    
    Dim percent_change As Double
    
    'set the headers for the new columns
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

'bonus
    Dim G_inc As Double
    Dim G_dec As Double
    Dim G_total As Double
    Dim count As Integer
   
'setting headers for bonus chart
    Cells(2, 15) = "Greatest % increase:"
    Cells(3, 15) = "Greatest % decrease:"
    Cells(4, 15) = "Greatest Total Volume:"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"

    'set count to the second row
    count = 2
    
    
        
    'Set the stock counter to zero
    stock_total = 0
    
    'Track the location of each stock in the summary table
    Dim summary_stock_row As Integer
    summary_stock_row = 2
    
    I = 2
    
    open_price = Cells(I, 3).Value
        
    'Loop through all the stock list
    
      ' Select cell A2, *first line of data*.
      'https://learn.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro
      Range("A2").Select
      ' Set Do loop to stop when an empty cell is reached.
      Do Until IsEmpty(ActiveCell)
    
          ' Select cell A2, *first line of data*.

              ' Check to see if the next stock name different than current
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                           
              'Set the stock name
              stock_name = Cells(I, 1).Value
              
              'Add to the stock total
              stock_total = stock_total + Cells(I, 7).Value
                           
              ' Print the Stock symbol in the Summary Table
              Range("I" & summary_stock_row).Value = stock_name
              
              ' Print the stock total to the summary stock row
              Range("L" & summary_stock_row).Value = stock_total
              
              'Calculate out opening price vs ending price
              close_price = Cells(I, 6).Value
              
              diff = close_price - open_price
              
              Range("J" & summary_stock_row).Value = diff
              
              'Set the color compared to change
              
                    If diff > 0 Then
                        Range("J" & summary_stock_row).Interior.ColorIndex = 4
                
                    ElseIf diff = 0 Then
                        Range("J" & summary_stock_row).Interior.ColorIndex = 6
                        
                    ElseIf diff < 0 Then
                        Range("J" & summary_stock_row).Interior.ColorIndex = 3
                    
                    End If
                
              
              
              'Calculate the percentage change
              'https://www.skillsyouneed.com/num/percent-change.html#:~:text=First%3A%20work%20out%20the%20difference,two%20numbers%20you%20are%20comparing.&text=Then%3A%20divide%20the%20increase%20by,this%20is%20a%20percentage%20decrease.
              percent_change = diff / open_price
              
              'Print out the percent change
              
              Range("K" & summary_stock_row).Value = percent_change
              
              'Format the percentage
              
                    If percent_change >= 0 Then
                        Range("K" & summary_stock_row).NumberFormat = "0.00%"
              
                    ElseIf percent_change < 0 Then
                        Range("K" & summary_stock_row).NumberFormat = "0.00%;[red]-0.00%"
              
                    End If
              
                                                    
              ' Add one to the summary table row
              summary_stock_row = summary_stock_row + 1
              
              ' Reset the stock total
              stock_total = 0
              
              ' If cells are the same - Add to the total
              
              I = I + 1
              
              'Reset the open price to the next stock
              open_price = Cells(I, 3).Value
              
            Else
            
            'If the column stock name didn't change, add to the totals and update I
              
              stock_total = stock_total + Cells(I, 7).Value
              
              I = I + 1
              
              
            End If
            
              ' Step down 1 row from present location.
                ActiveCell.Offset(1, 0).Select
      
   Loop
  
   
   'Format the output columns to fit the new widths
   
   Columns(9).AutoFit
   Columns(10).AutoFit
   Columns(11).AutoFit
   Columns(12).AutoFit
   Columns(15).AutoFit
   
   'bonus - setup variables
   
    Dim maxnum As Double
    Dim minnum As Double
    Dim maxtot As Double
    Dim tickmax As String
    Dim tickmin As String
    Dim ticktot As String
    
    'bonus - set counters in loops to define row values
    Dim x As Integer
    Dim t As Integer
    x = 2
    t = 2
    
    'bonus - set values to zero so that they are ready for comparators
    maxtot = 0
    maxnum = 0
    minnum = 0
   
       'bonus - loop through created column to find range in Percent Change column
    Range("K2").Select
      ' Set Do loop to stop when an empty cell is reached.
    Do Until IsEmpty(ActiveCell)

        'compare each line with the next in order to find highest value
        If Cells(x, 11).Value > maxnum Then
        
        ' set highest value  and associated ticker name
        maxnum = Cells(x, 11).Value
        tickmax = Cells(x, 9).Value

        End If
        
        
        If Cells(x, 11).Value < minnum Then

        
        minnum = Cells(x, 11).Value
        tickmin = Cells(x, 9).Value

        End If
        
        'add to the counter (x) to move to next row in loop
        x = x + 1

   ' Step down 1 row from present location.
        ActiveCell.Offset(1, 0).Select
                
    Loop
    
           'loop through totals to find ranges
    Range("L2").Select
      ' Set Do loop to stop when an empty cell is reached.
    Do Until IsEmpty(ActiveCell)
        
        'compare each line in total stock volume with the next in order to find highest value
        If Cells(t, 12).Value > maxtot Then

        ' set highest value  and associated ticker name
        maxtot = Cells(t, 12).Value
        ticktot = Cells(t, 9).Value

        End If
        
        'add to the counter (t) to move to next row in loop
        t = t + 1

   ' Step down 1 row from present location.
        ActiveCell.Offset(1, 0).Select
                
    Loop
               
    'Display the ticker name and value for the max, min, and max total
    Cells(2, 17).Value = maxnum
    Cells(2, 16).Value = tickmax
    Cells(3, 17).Value = minnum
    Cells(3, 16).Value = tickmin
    Cells(4, 17).Value = maxtot
    Cells(4, 16).Value = ticktot
    
    'Format the output cells
    Cells(3, 17).NumberFormat = "0.00%;[red]-0.00%"
    Cells(2, 17).NumberFormat = "0.00%"
    'How I learned to turn off scientific notation -https://superuser.com/questions/452832/turn-off-scientific-notation-in-excel#:~:text=Unfortunately%20excel%20does%20not%20allow,your%20data%20to%20scientific%20notation.
    Cells(4, 17).NumberFormat = "0"
    
    
    
    'Format the output columns
    Columns(16).AutoFit
    Columns(17).AutoFit
    
End Sub

