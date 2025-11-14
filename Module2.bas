Attribute VB_Name = "Module2"
Sub ExtractNegativePrices_ExcludeGST()
    Dim monthNames As Variant
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim i As Long
    Dim itsb As Variant, price As Variant, category As String
    Dim monthName As Variant

    monthNames = Array("JulyAB", "AugustAB", "SeptemberAB")

    For Each monthName In monthNames
        Set wsSource = ThisWorkbook.Sheets(monthName)

        ' Create or reset the target sheet
        On Error Resume Next
        Application.DisplayAlerts = False
        Worksheets("Negative_" & monthName).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Set wsTarget = ThisWorkbook.Sheets.Add
        wsTarget.Name = "Negative_" & monthName

        outputRow = 1
        lastRow = wsSource.Cells(wsSource.Rows.Count, "H").End(xlUp).Row

        ' Loop through rows in the source sheet
        For i = 4 To lastRow
            itsb = wsSource.Range("H" & i).Value
            price = wsSource.Range("K" & i).Value
            category = Trim(wsSource.Range("G" & i).Value)

            If IsNumeric(price) And price < 0 Then
                If UCase(category) <> "GST" Then
                    wsTarget.Cells(outputRow, "A").Value = itsb
                    wsTarget.Cells(outputRow, "B").Value = price
                    outputRow = outputRow + 1
                End If
            End If
        Next i
    Next monthName

    MsgBox "Negative price extraction complete (excluding GST rows).", vbInformation
End Sub
