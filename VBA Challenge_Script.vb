{\rtf1\ansi\ansicpg1252\cocoartf2709
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub StockDataAnalysis()\
\
'Apply to all worksheets in workbook\
Dim ws As Worksheet\
For Each ws In Worksheets\
\
'Setting the last row and other variables\
Dim Lastrow As Long\
Dim openprice As Double\
Dim closeprice As Double\
Dim yeardiff As Double\
Dim percentdiff As Double\
Dim totalstock As Double\
Dim tickervalue As String\
Dim summarytablerow As Integer\
\
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
totalstock = 0\
summarytablerow = 2\
openprice = ws.Cells(2, 3).Value\
\
' New column headers\
ws.Cells(1, 9).Value = "Ticker"\
ws.Cells(1, 10).Value = "Yearly Change"\
ws.Cells(1, 11).Value = "Percent Change"\
ws.Cells(1, 12).Value = "Total Stock Volume"\
\
\
For i = 2 To Lastrow\
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
        tickervalue = ws.Cells(i, 1).Value\
        closeprice = ws.Cells(i, 6).Value\
        yeardiff = closeprice - openprice\
        percentdiff = yeardiff / openprice\
        totalstock = totalstock + Cells(i, 7).Value\
        ws.Range("I" & summarytablerow).Value = tickervalue\
        ws.Range("J" & summarytablerow).Value = yeardiff\
        ws.Range("K" & summarytablerow).Value = percentdiff\
        ws.Range("L" & summarytablerow).Value = totalstock\
        summarytablerow = summarytablerow + 1\
        totalstock = 0\
        closeprice = 0\
        openprice = ws.Cells(i + 1, 3).Value\
    \
    Else\
    totalstock = totalstock + ws.Cells(i, 7).Value\
    \
    End If\
    \
Next i\
\
'General and conditional formatting\
ws.Columns(10).AutoFit\
ws.Columns(11).AutoFit\
ws.Columns(12).AutoFit\
\
Dim Lastrow2 As Integer\
Lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row\
\
' Set number format for column K percentage\
ws.Range("K2:K" & Lastrow2).NumberFormat = "0.00%"\
\
For i = 2 To Lastrow2\
    If IsNumeric(ws.Cells(i, 10).Value) And ws.Cells(i, 10).Value >= 0 Then\
        ws.Cells(i, 10).Interior.Color = vbGreen\
    Else\
    ws.Cells(i, 10).Interior.Color = vbRed\
    \
    End If\
Next i\
\
' Second summary table formatting\
ws.Range("O2").Value = "Greatest % Increase"\
ws.Range("O3").Value = "Greatest % Decrease"\
ws.Range("O4").Value = "Greatest Total Volume"\
ws.Range("P1").Value = "Ticker"\
ws.Range("Q1").Value = "Value"\
ws.Columns(15).AutoFit\
ws.Range("Q2:Q3").NumberFormat = "0.00%"\
\
'Declaring and initializing variables for second summary table\
Dim greatestinc As Double: greatestinc = ws.Cells(2, 11).Value\
Dim greatestdec As Double: greatestdec = ws.Cells(2, 11).Value\
Dim greatestvol As Double: greatestvol = ws.Cells(2, 12).Value\
Dim tickervalue2 As String\
\
'Greatest increase\
For i = 3 To Lastrow2\
    If ws.Cells(i, 11).Value > greatestinc Then\
        greatestinc = ws.Cells(i, 11).Value\
        tickervalue2 = ws.Cells(i, 9).Value\
    End If\
Next i\
ws.Range("P2").Value = tickervalue2\
ws.Range("Q2").Value = greatestinc\
\
'Greatest decrease\
For i = 3 To Lastrow2\
    If ws.Cells(i, 11).Value < greatestdec Then\
        greatestdec = ws.Cells(i, 11).Value\
        tickervalue2 = ws.Cells(i, 9).Value\
    End If\
Next i\
ws.Range("P3").Value = tickervalue2\
ws.Range("Q3").Value = greatestdec\
\
'Greatest volume\
For i = 3 To Lastrow2\
    If ws.Cells(i, 12).Value > greatestvol Then\
        greatestvol = ws.Cells(i, 12).Value\
        tickervalue2 = ws.Cells(i, 9).Value\
        End If\
Next i\
ws.Range("P4").Value = tickervalue2\
ws.Range("Q4").Value = greatestvol\
\
Next ws\
\
End Sub\
\
\
\
}