Attribute VB_Name = "Module1"
Sub ComparePrices_JulyOnly()
    Dim wsJuly As Worksheet, wsAug As Worksheet, wsSep As Worksheet
    Dim wsOut As Worksheet
    Dim dictAug As Object, dictSep As Object
    Dim lastRow As Long, i As Long
    Dim itsb As String, priceJuly As Variant, priceAug As Variant, priceSep As Variant
    Dim outputRow As Long

    Set dictAug = CreateObject("Scripting.Dictionary")
    Set dictSep = CreateObject("Scripting.Dictionary")

    Set wsJuly = ThisWorkbook.Sheets("JulyAB")
    Set wsAug = ThisWorkbook.Sheets("AugustAB")
    Set wsSep = ThisWorkbook.Sheets("SeptemberAB")

    ' Read August data into dictionary
    lastRow = wsAug.Cells(wsAug.Rows.Count, "H").End(xlUp).Row
    For i = 4 To lastRow
        itsb = wsAug.Range("H" & i).Value
        If Not dictAug.exists(itsb) Then
            dictAug(itsb) = wsAug.Range("K" & i).Value
        End If
    Next i

    ' Read September data into dictionary
    lastRow = wsSep.Cells(wsSep.Rows.Count, "H").End(xlUp).Row
    For i = 4 To lastRow
        itsb = wsSep.Range("H" & i).Value
        If Not dictSep.exists(itsb) Then
            dictSep(itsb) = wsSep.Range("K" & i).Value
        End If
    Next i

    ' Create output sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Comparison").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "Comparison"

    outputRow = 1

    ' Loop through July data and compare
    lastRow = wsJuly.Cells(wsJuly.Rows.Count, "H").End(xlUp).Row
    For i = 4 To lastRow
        itsb = wsJuly.Range("H" & i).Value
        priceJuly = wsJuly.Range("K" & i).Value

        ' Only process if this ITSB hasn't been written yet
        If Application.WorksheetFunction.CountIf(wsOut.Range("A:A"), itsb) = 0 Then
            wsOut.Cells(outputRow, "A").Value = itsb
            wsOut.Cells(outputRow, "B").Value = priceJuly

            If dictAug.exists(itsb) Then
                wsOut.Cells(outputRow, "C").Value = dictAug(itsb)
            End If
            If dictSep.exists(itsb) Then
                wsOut.Cells(outputRow, "D").Value = dictSep(itsb)
            End If

            outputRow = outputRow + 1
        End If
    Next i

    MsgBox "Comparison complete using only July ITSB numbers.", vbInformation
End Sub
