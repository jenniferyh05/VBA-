Sub stock():

For each ws in Worksheets

    Dim stockcount as integer
    Dim totalvolumn as LongLong 
    Dim newstockcount as LongLong
    Dim increase as double 
    Dim decrease as double 
    Dim volume as LongLong 

    stockcount = 1
    totalvolumn = 0
    

    LastRow = ws.Cells(Rows.Count,1).End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("X1").Value = "Opening Price"
    ws.Range("Y1").Value = "Closing Price"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"


        for i = 2 to LastRow 
            If ws.Cells(i,1) = ws.Cells(i+1,1) Then 
            totalvolumn = totalvolumn + Cells(i,7).Value 


            Else
            stockcount = stockcount +1
            totalvolumn = totalvolumn + Cells(i,7).Value
            ws.Range("I"&stockcount).Value = ws.Cells(i,1).Value 
            ws.Range("L"&stockcount).value = totalvolumn 
            totalvolumn = 0 
            End if 

            If ws.Cells(i,2).Value = "20160101" Then 
            newstockcount = stockcount + 1
            ws.Range("X"&newstockcount).Value = ws.Cells(i,3).Value 
            Elseif ws.Cells(i,2).Value = "20161230" Then 
            ws.Range("Y"&newstockcount).Value = ws.Cells(i,6).Value 
            End if 

            ws.Range("J"&newstockcount).Value = ws.Range("Y"&newstockcount).Value - ws.Range("X"&newstockcount).Value 
            ws.Range("K"&newstockcount).Value= Round((ws.Range("J"&newstockcount).Value/ws.Range("X"&newstockcount).Value),2)
            ws.Range("K"&newstockcount).Style = "Percent"

                If ws.Range("J"&newstockcount).Value > 0 Then 
                ws.Range("J"&newstockcount).Interior.ColorIndex = 4
                Else
                ws.Range("J"&newstockcount).Interior.ColorIndex = 3
                End if 

        Next i 

        increase = WorksheetFunction.Max(ws.Range("K2:K"&newstockcount))
        decrease = WorksheetFunction.Min(ws.Range("K2:K"&newstockcount))
        volume = WorksheetFunction.Max(ws.Range("L2:L"&newstockcount))
        matchincrease = WorksheetFunction.match(increase,ws.Range("K1:K"&newstockcount),0)
        matchdecrease = WorksheetFunction.match(decrease,ws.Range("K1:K"&newstockcount),0)
        matchvolume = WorksheetFunction.match(volume,ws.Range("L1:L"&newstockcount),0)

         ws.Range("O2").Value = ws.Range("I"&matchincrease).Value 
         ws.Range("O3").Value = ws.Range("I"&matchdecrease).Value 
         ws.Range("O4").Value = ws.Range("I"&matchvolume).Value 
         ws.Range("P2").Value = increase
         ws.Range("P3").Value = decrease
         ws.Range("P4").Value = volume
         ws.Range("P2").Style = "Percent"
         ws.Range("P3").Style = "Percent"


Next ws

End Sub 


