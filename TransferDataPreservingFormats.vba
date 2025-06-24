
Sub TransferDataPreservingFormats_MatchedHeaders()
    Dim wsMaster As Worksheet, wsLog As Worksheet
    Dim wbOutput As Workbook, wsOutput As Worksheet
    Dim inputHeaders As Object, outputHeaders As Object
    Dim dictGroups As Object
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, key As Variant, rowNum As Variant
    Dim logRow As Long

    Set wsMaster = ThisWorkbook.Sheets("Master Data")
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set inputHeaders = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "Log"
        wsLog.Cells(1, 1).Value = "File Path"
        wsLog.Cells(1, 2).Value = "Error Details"
    End If
    logRow = wsLog.Cells(Rows.Count, 1).End(xlUp).Row + 1

    lastCol = wsMaster.Cells(4, Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        inputHeaders(wsMaster.Cells(4, i).Value) = i
    Next i

    lastRow = wsMaster.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 6 To lastRow
        Dim filePath As String
        filePath = wsMaster.Cells(i, 1).Value

        If Not dictGroups.exists(filePath) Then
            dictGroups.Add filePath, New Collection
        End If
        dictGroups(filePath).Add i
    Next i

    For Each key In dictGroups.keys
        filePath = key

        If Dir(filePath) = "" Then
            wsLog.Cells(logRow, 1).Value = filePath
            wsLog.Cells(logRow, 2).Value = "File not found"
            logRow = logRow + 1
            GoTo NextGroup
        End If

        Set wbOutput = Workbooks.Open(filePath)
        Set wsOutput = wbOutput.Sheets(1)

        Set outputHeaders = CreateObject("Scripting.Dictionary")
        lastCol = wsOutput.Cells(4, Columns.Count).End(xlToLeft).Column
        For i = 1 To lastCol
            outputHeaders(wsOutput.Cells(4, i).Value) = i
        Next i

        Dim outputRow As Long
        outputRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1

        For Each rowNum In dictGroups(key)
            Dim header As Variant
            For Each header In inputHeaders.keys
                If outputHeaders.exists(header) Then
                    Dim masterCell As Range
                    Dim outputCell As Range

                    Set masterCell = wsMaster.Cells(rowNum, inputHeaders(header))
                    Set outputCell = wsOutput.Cells(outputRow, outputHeaders(header))

                    outputCell.Value = masterCell.Value
                    outputCell.NumberFormat = masterCell.NumberFormat
                End If
            Next header
            outputRow = outputRow + 1
        Next rowNum

        wbOutput.Save
        wbOutput.Close False

NextGroup:
    Next key

    MsgBox "Batch transfer complete. Check 'Log' for issues.", vbInformation
End Sub
