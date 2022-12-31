{\rtf1\ansi\ansicpg1252\cocoartf2639
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Stocks()\
' --------------------------------------------\
' LOOP THROUGH ALL SHEETS\
' --------------------------------------------\
For Each ws In Worksheets\
\
    'Add Headings across worksheets\
    ws.Range("I1").Value = "Ticker"\
    ws.Range("J1").Value = "Yearly Change"\
    ws.Range("K1").Value = "Percent Change"\
    ws.Range("L1").Value = "Total Stock Volume"\
    ws.Range("P1").Value = "Ticker"\
    ws.Range("Q1").Value = "Value"\
    ws.Range("O2").Value = "Greatest % Increase"\
    ws.Range("O3").Value = "Greatest % Decrease"\
    ws.Range("O4").Value = "Greatest Total Volume"\
    \
    ' Set an initial variable for holding the ticker\
    Dim ticker As String\
    \
    ' Determine the Last Row\
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
   \
    ' Set an initial variable for holding the total stock volume,year start price, year end price\
    Dim yearstart As Double\
    Dim yearend As Double\
    Dim yearchange As Double\
    Dim percentchange As Double\
    Dim voltotal As Double\
    Dim year As String\
    'For Alpha list VBA Code\
    'year = "2020"\
    'For Multi-Year List\
    year = ws.Name\
    \
    voltotal = 0\
    \
\
    ' Keep track of the location for each ticker in the summary table\
    Dim Summary_Table_Row As Integer\
    Summary_Table_Row = 2\
\
    ' Loop through all tickers\
    For i = 2 To LastRow\
\
        ' Check if we are still within the same ticker, if it is not...\
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
 \
            ' Set the ticker name\
            ticker = ws.Cells(i, 1).Value\
\
            ' Add to the Stock Volume\
            voltotal = voltotal + ws.Cells(i, 7).Value\
\
            ' Print the Ticker in the Summary Table\
            ws.Range("I" & Summary_Table_Row).Value = ticker\
\
            ' Print the Stock Volume to the Summary Table\
            ws.Range("L" & Summary_Table_Row).Value = voltotal\
            'Determine year open value\
            If ws.Cells(i, 2).Value = (year + "0102") Then\
                yearstart = ws.Cells(i, 3).Value\
            End If\
            'Determine year close value\
            If ws.Cells(i, 2).Value = (year + "1231") Then\
                yearend = ws.Cells(i, 6).Value\
            End If\
            'Caculate Year Change, and percent change\
            yearchange = yearend - yearstart\
            percentchange = (yearchange / yearstart)\
            ' Print the Year Change to the Summary Table\
            ws.Range("J" & Summary_Table_Row).Value = yearchange\
            ws.Range("K" & Summary_Table_Row).Value = percentchange\
            \
            ' Add one to the summary table row\
            Summary_Table_Row = Summary_Table_Row + 1\
      \
            ' Reset the Stock Vol Total\
            voltotal = 0\
\
        ' If the cell immediately following a row is the same ticker...\
        Else\
\
            ' Add to the Stock Volume Total\
            voltotal = voltotal + ws.Cells(i, 7).Value\
            'Determine year open value\
            If ws.Cells(i, 2).Value = (year + "0102") Then\
                yearstart = ws.Cells(i, 3).Value\
            End If\
            'Determine year close value\
            If ws.Cells(i, 2).Value = (year + "1231") Then\
                yearend = ws.Cells(i, 6).Value\
            End If\
\
        End If\
\
    Next i\
    \
    'Determine Last Row for Bonus Answer\
    LastRowK = ws.Range("K2").End(xlDown).Row\
          \
    'Declare Variables for Bonus Answer\
    Dim GreatestIncrease As Double\
    Dim GreatestDecrease As Double\
    Dim GreatestVol As Double\
    \
    'Determine Values for Bonus Answer\
    GreatestIncrease = WorksheetFunction.Max(ws.Range("K2:K" & LastRowK))\
    ws.Range("Q2").Value = GreatestIncrease\
    GreatestDecrease = WorksheetFunction.Min(ws.Range("K2:K" & LastRowK))\
    ws.Range("Q3").Value = GreatestDecrease\
    GreatestVol = WorksheetFunction.Max(ws.Range("L2:L" & LastRowK))\
    ws.Range("Q4").Value = GreatestVol\
    \
    'Determine Ticker for Bonus Values\
    Dim y As Integer\
    For y = 2 To LastRowK\
        If GreatestIncrease = ws.Cells(y, 11).Value Then\
            ws.Range("P2").Value = ws.Cells(y, 9).Value\
        End If\
        If GreatestDecrease = ws.Cells(y, 11).Value Then\
            ws.Range("P3").Value = ws.Cells(y, 9).Value\
        End If\
        If GreatestVol = ws.Cells(y, 12).Value Then\
            ws.Range("P4").Value = ws.Cells(y, 9).Value\
        End If\
    Next y\
    \
    \
    'Formatting Notes\
    ws.Columns("K").NumberFormat = "0.00%"\
    ws.Range("Q2:Q3").NumberFormat = "0.00%"\
    ws.Range("I:Q").Columns.AutoFit\
    Dim rng As Range\
    Dim condition1 As FormatCondition, condition2 As FormatCondition\
    Set rng = ws.Columns("J")\
    \
    rng.FormatConditions.Delete\
    \
    Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")\
    Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")\
    With condition1\
        .Interior.Color = vbGreen\
    End With\
    With condition2\
        .Interior.Color = vbRed\
    End With\
    ws.Range("J1").FormatConditions.Delete\
    \
\
Next ws\
\
End Sub}