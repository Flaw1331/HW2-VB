Sub StockAnalyze():
	
    ' Initializing loop to go through each sheet
    For Each ws in Worksheets	
	
		' First row in all worksheets
		Dim firstRow As Integer : firstRow = 2

		' Determine the Last Row
		LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

		' Declaring a counter
		dim counter as integer : counter = 1
		
		' Declaring results matrix categories variables
		Dim totalVol as Single : totalVol = 0
		Dim totalCha as Single : totalCha = 0
		Dim percentCha as Single : percentCha = 0
		Dim openPrice as Single : openPrice = 0
		Dim closePrice as Single : closePrice = 0
		Dim maxVol as Single : maxVol = 0
		Dim grtInc as Single : grtInc = 0
		Dim grtDec as Single : grtDec = 0
		
		' Create Results matrix on each sheet
		ws.Cells(1,9).Value = "Ticker"
		ws.Cells(1,10).Value = "Total Change"
		ws.Cells(1,11).Value = "% of Change"
		ws.Cells(1,12).Value = "Total Volume"
		ws.Cells(1,16).Value = "Ticker"
		ws.Cells(2,14).Value = "Total Shares"
		ws.Cells(4,14).Value = "Greatest % Increase"
		ws.Cells(6,14).Value = "Greatest % Decrease"
		
		' Formatting the matrix
		ws.Range("I1:P1").Font.Bold = True
		ws.Range("I1:P1").Font.Underline = True
		ws.Range("N1:N6").Font.Bold = True
		ws.Range("N1:N6").Font.Underline = True
		
		' Anchoring loop start for openPrice
		openPrice = ws.Cells(2,3).Value
	
		' Populating results matrix
		For i = 2 to LastRow
			
			' total volume runs until it detects a shift in ticker
			totalVol = totalVol + ws.Cells(i,7).Value
			
			' Checking if ticker is different in the next row and running end logic
			if (ws.Cells(i+1,1).Value <> ws.Cells(i,1).Value) then
				
				' Ticker shift
				ws.Cells(counter+1,9).Value = ws.Cells(i,1).Value
				counter = counter + 1
				
				' totalVol print for each ticker
				ws.Cells(counter,12).Value = totalVol
				
				' Checking if totalVol is new maxVol and printing it
				if totalVol > maxVol then
					maxVol = totalVol
					ws.Cells(2,15).Value = maxVol
					ws.Cells(2,16).Value = ws.Cells(i,1).Value
				end if
				
				' Resetting totalVol for next ticker
				totalVol = 0
				
				' Finding close price
				closePrice = ws.Cells(i,3).Value
				
				' Calculating yearly change in units and then print
				totalCha = (closePrice - openPrice)
				ws.Cells(counter,10).value	= totalCha
				
				' Formatting total change
				if totalCha > 0 then
					' Green for positive change
					ws.Cells(counter,10).interior.colorindex = 4
				elseif totalCha < 0 then
					' Red for negative change
					ws.Cells(counter,10).interior.colorindex = 3
				else
					' Yellow for the ever-elusive no-change
					ws.Cells(counter,10).interior.colorindex = 6
				end if
				
				' Checking if % change is NULL, if yes, setting change to 0
				if (closePrice <> 0 & openPrice <> 0) then
					percentCha = (closePrice-openPrice)/openPrice
				else
					percentCha = 0
				end if
				
				' Printing % change and formatting it
				ws.Cells(counter,11).Value = percentCha
				ws.Cells(counter,11).style = "Percent"
				
				' Finding and printing + formatting greatest improved
				if percentCha > grtInc then
					grtInc = percentCha
					ws.Cells(4,15).Value = grtInc
					ws.Cells(4,15).Style = "Percent"
					ws.Cells(4,16).Value = ws.Cells(i,1).Value
				end if
				
				' Finding and printing + formatting least improved
				if percentCha < grtDec then
					grtDec = percentCha
					ws.Cells(6,15).Value = grtDec
					ws.Cells(6,15).Style = "Percent"
					ws.Cells(6,16).Value = ws.Cells(i,1).Value
				End If

				' Setting new open price for next ticker
				openPrice = ws.Cells(i+1,3).Value
				
			end if
		next i
	Next ws
End Sub