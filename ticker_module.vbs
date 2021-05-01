Attribute VB_Name = "Module1"
Sub Stock_Analyis()

'yearly open and close price, total volume, and
'year chance and percent change ( format increase/decrease)

For Each ws In Worksheets

'set sheet as active
'ActiveSheet.Select

'information needed

Dim sticker As String
Dim openp As Double
Dim openprize As Double
Dim closep As Double
Dim closeprize As Double
Dim stockvol As Double
Dim stockvoltotal As Double
Dim lrow As Double


'counters
Dim i, j, k As Integer
j = 1
k = 1


'headers

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 13).Value = "Number of trade days"

'find how big each set is
lrow = ws.Cells(Rows.Count, 1).End(xlUp).row
'msgbox (str(lrow))

'read data loop

For i = 2 To lrow

sticker = ws.Cells(i, 1).Value
openp = ws.Cells(i, 3).Value
closep = ws.Cells(i, 6).Value
stockvol = ws.Cells(i, 7).Value
stockvoltotal = (stockvoltotal + stockvol)

'keep opening prize
If k = 1 Then
openprize = openp
End If

'skip if 0 opening prize
If openprize > 0 Then

'check if next stock is the same
If ws.Cells((i + 1), 1).Value <> sticker Then
    closeprize = closep
'calculate stats
    ws.Cells(j + 1, 9).Value = sticker
    ws.Cells(j + 1, 10).Value = (closeprize - openprize)
    ws.Cells(j + 1, 11).Value = (((closeprize - openprize)) / openprize)
    ws.Cells(j + 1, 12) = stockvoltotal
    ws.Cells(j + 1, 13) = k
'format
If (closeprize - openprize) > 0 Then
    ws.Cells(j + 1, 10).Interior.ColorIndex = 4
Else
    ws.Cells(j + 1, 10).Interior.ColorIndex = 3
End If

    ws.Cells(j + 1, 11).NumberFormat = "#.##%"

' resets
    stockvoltotal = 0
    j = j + 1
    k = 1
Else
    k = k + 1
End If

Else
End If

Next i

'find max and min'
Dim maxp As Double
Dim minp As Double
Dim maxvol As Double
Dim x As Double
Dim t As Double

'analysis row count
x = ws.Cells(Rows.Count, 9).End(xlUp).row
'ws.Cells(2, 15).Value = x

'max increase
maxp = WorksheetFunction.Max(ws.Range("K" & "2" & ":" & "K" & x))
ws.Cells(2, 18).Value = maxp
ws.Cells(2, 18).NumberFormat = "#.##%"
ws.Cells(2, 16).Value = "Greatest % Increase"
'for ticker
For t = 2 To x
If ws.Cells(t, 11).Value = maxp Then
ws.Cells(2, 17).Value = ws.Cells(t, 9).Value
Else
End If
Next t

'max decrease
minp = WorksheetFunction.Min(ws.Range("K" & "2" & ":" & "K" & x))
ws.Cells(3, 18).Value = minp
ws.Cells(3, 18).NumberFormat = "#.##%"
ws.Cells(3, 16).Value = "Greatest % Decrease"
'for ticker
For t = 2 To x
If ws.Cells(t, 11).Value = minp Then
ws.Cells(3, 17).Value = ws.Cells(t, 9).Value
Else
End If
Next t

' max volume
maxvol = WorksheetFunction.Min(ws.Range("L" & "2" & ":" & "L" & x))
ws.Cells(4, 18).Value = maxvol
ws.Cells(4, 16).Value = "Greatest Total Volume"
'for ticker
For t = 2 To x
If ws.Cells(t, 12).Value = maxvol Then
ws.Cells(4, 17).Value = ws.Cells(t, 9).Value
Else
End If
Next t

'headers
ws.Cells(1, 18).Value = "Value"
ws.Cells(1, 17).Value = "Ticker"



Next ws

End Sub

