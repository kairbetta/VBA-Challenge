VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_year_stock_analysis()


 
 
'***Loop through each Worksheet***
For Each ws In Worksheets

'***Declare diffenet variables for row, column, etc.***'
 Dim r As Long
 Dim c As Long
 Dim tikcer_name As String
 Dim ticker_row As Integer
 ticker_row = 2
' Dim m As Integer
 Dim n As Long
 Dim ticker_volume As Double
 ticker_volume = 0
 Dim opening_price As Double
 opening_price = Cells(2, 3).Value
 Dim closing_price As Double
 Dim yearly_change As Double
 Dim percent_change As Double

    '***Create new column headers***
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
   'Print labels "Greatest % increase", "Greatest % decrease", and "Greatest total volume", etc.
        
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"
   
 ' Determine number of rows by fetching the last row count
  
      last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
       For r = 2 To last_row
       
         
         If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            ticker_name = ws.Cells(r, 1).Value
         
            'ws.Cells(r, 9).Value = ticker_name
             ws.Range("I" & ticker_row).Value = ticker_name
            
             'Calculate trade volume
             ticker_volume = ticker_volume + ws.Cells(r, 7).Value
             
             'Generate values in ticker volume column
             ws.Range("L" & ticker_row).Value = ticker_volume

            ' Calculate and print Yearly change from the opening price at the beginning of a given year to the
            ' closing price at the end of that year.
              
            closing_price = ws.Cells(r, 6).Value
            
            yearly_change = (closing_price - opening_price)
            
                      
            ws.Range("J" & ticker_row).Value = yearly_change
            
            'Use Conditional formatting to highlight positive values, zero in green and negative values in red.
            If (yearly_change >= 0) Then
                ws.Range("J" & ticker_row).Interior.ColorIndex = 10
            Else
               ws.Range("J" & ticker_row).Interior.ColorIndex = 3
                
            End If
            
            
            If (opening_price = 0) Then
                percent_change = 0
            Else
                percent_change = yearly_change / opening_price
            End If
            
           
            ws.Range("K" & ticker_row).Value = percent_change
            ws.Range("K" & ticker_row).NumberFormat = "0.00%"
           
         
             ticker_row = ticker_row + 1
             ticker_volume = 0
             opening_price = ws.Cells(r + 1, 3)
           
           Else
            ticker_volume = ticker_volume + ws.Cells(r, 7).Value
        End If
     

    Next r
    
    'Below for loop is for the final part of the challenge - To calculate max %age, min %age and max vol of trade
       
       For n = 2 To last_row
      
        'Determine maximum percentage
        'Find if each row is equal to max value, if yes then assign that as max value in Cell (2, 17) and also find the ticker value asociated with this
        'max cell and display ticker in Cell (2,16)
         If ws.Cells(n, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row)) Then
                ws.Cells(2, 16).Value = ws.Cells(n, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(n, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
         'Determine minimum percentage
                ElseIf ws.Cells(n, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row)) Then
                ws.Cells(3, 16).Value = ws.Cells(n, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(n, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
         'Determine maximum volume of trade
                ElseIf ws.Cells(n, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row)) Then
                ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(n, 12).Value
        End If
   Next n
    
  Next ws
                       
End Sub



