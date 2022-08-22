# stock_analysis

### Incomplete

I am caught with the All Stocks Analysis Sheet. After trying for over 8 hours I decided to put it down and continue my lessons.

This is the code I have currently.

-----------------------------------------------------------

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    cells(3, 1).Value = "Year"
    cells(3, 2).Value = "Total Daily Volume"
    cells(3, 3).Value = "Return"


'-------------------------------------------

    Worksheets("2018").Activate
    
    totalvolume = 0
    
    Dim startingprice As Double
    Dim endingprice As Double
    
    
    
    'establish number of rows to loop over
    rowstart = 2
    
    rowend = cells(rows.Count, "a").End(xlUp).Row
    'code taken from https://tinyurl.com/mr46da4xz
    
    
    'loop over all rows
    For i = rowstart To rowend
    
        'increase totalVolume if ticker is "DQ"
        If cells(i, 1).Value = "DQ" Then
            totalvolume = totalvolume + cells(i, 8).Value
            
            
        End If
        
    'if IS DQ but prior is not then set starting price
     If cells(i, 1).Value = "DQ" And cells(i - 1, 1).Value <> "DQ" Then
 
 'set starting price
 
startingprice = cells(i, 6).Value

End If

Dim test_stocks As Integer

rowstart = 3
rowend = 3013



    'if cell is DQ and next one isn't then
    If cells(i, 1).Value = "DQ" And cells(i + 1, 1).Value <> "DQ" Then

endingprice = cells(i, 6).Value
 
 End If
 
 Next i
 

'-------------------------------------------
Worksheets("DQ Analysis").Activate
    cells(4, 1).Value = 2018
    cells(4, 2).Value = totalvolume
    cells(4, 3).Value = endingprice / startingprice - 1

End Sub
'-------------------------------------------
Sub AllStocksAnalysis()

Worksheets("All Stocks Analysis").Activate

cells(1, 1).Value = "All Stocks (2018)"
Range("A1:H1").Interior.ColorIndex = 38

cells(3, 1).Value = "Ticker"
cells(3, 2).Value = "Total Daily Volume"
cells(3, 3).Value = "Return"

End Sub


Sub StartArray()
Worksheets("Test Stocks 2018").Activate


'THIS IS THE ARRAY START
    Dim tickers(11) As String
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"




'THIS IS THE LOOP START

Worksheets("All Stocks Analysis").Activate
  
    Dim stocknames As String
    
    rowstart = 4
    rowend = 15
    
    For i = rowstart To rowend
    
        cells(4, 1).Value = "AY"
        cells(5, 1).Value = "CSIQ"
        cells(6, 1).Value = "DQ"
        cells(7, 1).Value = "ENPH"
        cells(8, 1).Value = "FSLR"
        cells(9, 1).Value = "HASI"
        cells(10, 1).Value = "JKS"
        cells(11, 1).Value = "RUN"
        cells(12, 1).Value = "SEDG"
        cells(13, 1).Value = "SPWR"
        cells(14, 1).Value = "TERP"
        cells(15, 1).Value = "VSLR"
        
   Next i
    


 Dim total_volume As Double
 
For j = rowstart To rowend


If cells(i - 1, 1).Value <> "AY" And cells(i, 1).Value = "AY" Then
    totalvolume = totalvolume + cells(i, 8).Value
    
End If

Dim cells As Double

If cells(i + 1, 1).Value <> "AY" And cells(i, 1).Value = "AY" Then
    
    
    


End If
Next j

End Sub
