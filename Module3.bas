Attribute VB_Name = "Module3"
Sub ExtractNegativeRows_ExcludeGST()
    Dim monthNames As Variant
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, lastCol As Long, outputRow As Long
    Dim i As Long
    Dim price As Variant, category As String
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
        lastCol = wsSource.Cells(3, wsSource.Columns.Count).End(xlToLeft).Column ' assumes headers are in row 3

        ' Loop through rows in the source sheet
        For i = 4 To lastRow
            price = wsSource.Cells(i, "K").Value
            category = Trim(wsSource.Cells(i, "G").Value)

            If IsNumeric(price) And price < 0 Then
                If UCase(category) <> "GST" Then
                    ' Copy the entire row
                    wsSource.Range(wsSource.Cells(i, 1), wsSource.Cells(i, lastCol)).Copy _
                        Destination:=wsTarget.Cells(outputRow, 1)
                    outputRow = outputRow + 1
                End If
            End If
        Next i
    Next monthName

    MsgBox "Negative row extraction complete (excluding GST rows).", vbInformation
End Sub

