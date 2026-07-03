Attribute VB_Name = "PackingListModule"
'====================================================================
' CARTON DISTRIBUTION MACRO - "Input" sheet driven
'
' INPUT SHEET LAYOUT ("Input"):
'   B8  = Box Type (text: "01 - Large Box (60Lx37Wx30H)" or
'                          "02 - Small Box (37Lx30Wx20H)")
'   E8  = Capacity in pcs (24/21/18 for Large, 1-10 for Small)
'   B12:G12 = Size headers (XS,S,M,L,XL,XXL)  -- not used by macro
'   B13:G13 = Average qty per box per size    -- reference only
'   B17:H17 = Header row: "Style/Color/Fit", XS, S, M, L, XL, XXL
'   B18:H21 = Up to 4 data rows:
'             Col B = Style/Color/Fit label (e.g. "HM3010837/573/")
'             Col C:H = Order quantities per size (XS..XXL)
'
' OUTPUT:
'   For each non-empty data row (B18:B21), the macro creates/refreshes
'   a sheet named after the Style/Color/Fit label (sanitized) containing
'   a carton breakdown in the same J/K (Size/Qty) format as the
'   original packing list:
'       Row 1   = Title (Style/Color/Fit + Box Type + Capacity)
'       Rows 3+ = Carton rows: Col B = Carton No., Col J = Size,
'                 Col K = Qty in that carton
'       After last carton row: TOTAL CARTONS, TOTAL PIECES
'====================================================================

Sub GeneratePackingList()

    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Sheets("Input")

    Const HDR_ROW As Long = 17
    Const FIRST_DATA_ROW As Long = 18
    Const LAST_DATA_ROW As Long = 21
    Const MAX_CARTONS As Long = 100   ' safety cap on carton rows per style

    Dim sizes(1 To 6) As String
    sizes(1) = "XS": sizes(2) = "S": sizes(3) = "M"
    sizes(4) = "L": sizes(5) = "XL": sizes(6) = "XXL"

    '--- 1. Read Box Type and Capacity -----------------------------------
    Dim boxType As String
    Dim Capacity As Long
    boxType = Trim(wsIn.Range("B8").Value)
    Capacity = wsIn.Range("E8").Value

    If Capacity <= 0 Then
        MsgBox "Please enter a valid carton capacity in cell E8.", vbExclamation
        Exit Sub
    End If

    ' Validate capacity against box type
    If InStr(boxType, "01") = 1 Or InStr(boxType, "Large") > 0 Then
        If Capacity <> 18 And Capacity <> 21 And Capacity <> 24 Then
            MsgBox "For Box Type 01 (Large Box), capacity must be 18, 21, or 24.", vbExclamation
            Exit Sub
        End If
    ElseIf InStr(boxType, "02") = 1 Or InStr(boxType, "Small") > 0 Then
        If Capacity < 1 Or Capacity > 10 Then
            MsgBox "For Box Type 02 (Small Box), capacity must be between 1 and 10.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Please select a valid Box Type in cell B8.", vbExclamation
        Exit Sub
    End If

    '--- 2. Loop through each Style/Color/Fit data row -------------------
    Dim dataRow As Long
    Dim processedCount As Long
    processedCount = 0

    For dataRow = FIRST_DATA_ROW To LAST_DATA_ROW

        Dim styleLabel As String
        styleLabel = Trim(wsIn.Cells(dataRow, "B").Value)

        If styleLabel = "" Then GoTo NextRow

        ' Read order quantities for this row (C:H -> XS..XXL)
        Dim Qty(1 To 6) As Long
        Dim i As Long
        Dim totalPieces As Long
        totalPieces = 0
        For i = 1 To 6
            Qty(i) = wsIn.Cells(dataRow, 2 + i).Value   ' C=3..H=8
            totalPieces = totalPieces + Qty(i)
        Next i

        If totalPieces = 0 Then GoTo NextRow

        '--- 2a. Continuous fill algorithm --------------------------------
        Dim cartonsNeeded As Long
        cartonsNeeded = Application.WorksheetFunction.RoundUp(totalPieces / Capacity, 0)

        If cartonsNeeded > MAX_CARTONS Then
            MsgBox "Style " & styleLabel & " requires " & cartonsNeeded & _
                   " cartons, exceeding the safety limit of " & MAX_CARTONS & ".", vbCritical
            GoTo NextRow
        End If

        Dim remaining(1 To 6) As Long
        For i = 1 To 6
            remaining(i) = Qty(i)
        Next i

        ' rowSizeQty(cartonNo, sizeIndex)
        Dim rowSizeQty(1 To MAX_CARTONS, 1 To 6) As Long
        Dim r As Long, spaceLeft As Long, sizeIdx As Long, take As Long
        Dim remTotal As Long
        remTotal = totalPieces

        r = 1
        spaceLeft = Capacity
        sizeIdx = 1

        Do While remTotal > 0
            If remaining(sizeIdx) = 0 Then
                sizeIdx = sizeIdx + 1
                If sizeIdx > 6 Then Exit Do
                GoTo ContinueLoop
            End If

            take = remaining(sizeIdx)
            If take > spaceLeft Then take = spaceLeft

            rowSizeQty(r, sizeIdx) = rowSizeQty(r, sizeIdx) + take
            remaining(sizeIdx) = remaining(sizeIdx) - take
            spaceLeft = spaceLeft - take
            remTotal = remTotal - take

            If spaceLeft = 0 Then
                r = r + 1
                If r > MAX_CARTONS Then Exit Do
                spaceLeft = Capacity
            End If

            If remaining(sizeIdx) = 0 Then sizeIdx = sizeIdx + 1
            If sizeIdx > 6 Then Exit Do
ContinueLoop:
        Loop

        Dim totalCartons As Long
        totalCartons = 0
        For r = 1 To MAX_CARTONS
            Dim hasData As Boolean
            hasData = False
            For i = 1 To 6
                If rowSizeQty(r, i) > 0 Then hasData = True
            Next i
            If hasData Then totalCartons = totalCartons + 1
        Next r

        '--- 2b. Create / refresh output sheet ----------------------------
        Dim sheetName As String
        sheetName = SanitizeSheetName(styleLabel)

        Dim wsOut As Worksheet
        On Error Resume Next
        Set wsOut = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0

        If wsOut Is Nothing Then
            Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsOut.Name = sheetName
        Else
            wsOut.Cells.Clear
        End If

        ' Title
        wsOut.Range("B1").Value = "Style/Color/Fit: " & styleLabel
        wsOut.Range("B1").Font.Bold = True
        wsOut.Range("B2").Value = "Box Type: " & boxType
        wsOut.Range("B3").Value = "Carton Capacity: " & Capacity & " pcs"

        ' Header row for carton table
        wsOut.Range("B5").Value = "Carton No."
        wsOut.Range("J5").Value = "Size"
        wsOut.Range("K5").Value = "Qty"
        wsOut.Range("B5:K5").Font.Bold = True

        ' Write carton rows
        Dim outRow As Long
        Dim cartonCounter As Long
        cartonCounter = 0
        outRow = 6

        For r = 1 To totalCartons
            Dim firstSizeWritten As Boolean
            firstSizeWritten = False
            cartonCounter = cartonCounter + 1
            For i = 1 To 6
                If rowSizeQty(r, i) > 0 Then
                    If Not firstSizeWritten Then
                        wsOut.Cells(outRow, "B").Value = cartonCounter
                        firstSizeWritten = True
                    End If
                    wsOut.Cells(outRow, "J").Value = sizes(i)
                    wsOut.Cells(outRow, "K").Value = rowSizeQty(r, i)
                    outRow = outRow + 1
                End If
            Next i
        Next r

        ' Summary
        outRow = outRow + 1
        wsOut.Cells(outRow, "B").Value = "TOTAL CARTONS:"
        wsOut.Cells(outRow, "B").Font.Bold = True
        wsOut.Cells(outRow, "K").Value = totalCartons
        wsOut.Cells(outRow, "K").Font.Bold = True

        outRow = outRow + 1
        wsOut.Cells(outRow, "B").Value = "TOTAL PIECES:"
        wsOut.Cells(outRow, "B").Font.Bold = True
        wsOut.Cells(outRow, "K").Value = totalPieces
        wsOut.Cells(outRow, "K").Font.Bold = True

        ' Per-size totals (SHIP QTY summary)
        outRow = outRow + 2
        wsOut.Cells(outRow, "B").Value = "Size"
        wsOut.Cells(outRow, "B").Font.Bold = True
        For i = 1 To 6
            wsOut.Cells(outRow, 2 + i).Value = sizes(i)
            wsOut.Cells(outRow, 2 + i).Font.Bold = True
        Next i
        outRow = outRow + 1
        wsOut.Cells(outRow, "B").Value = "Order Qty"
        For i = 1 To 6
            wsOut.Cells(outRow, 2 + i).Value = Qty(i)
        Next i

        wsOut.Columns("B:K").AutoFit

        processedCount = processedCount + 1

NextRow:
    Next dataRow

    If processedCount = 0 Then
        MsgBox "No data rows found. Please enter quantities in B18:H21 on the Input sheet.", vbExclamation
    Else
        MsgBox "Packing list generated for " & processedCount & " style/color/fit row(s).", vbInformation
    End If

End Sub

'--- Helper: sanitize sheet names (Excel limits: 31 chars, no \ / ? * [ ] :) ---
Private Function SanitizeSheetName(ByVal s As String) As String
    Dim badChars As String
    badChars = "\/?*[]:"
    Dim i As Long
    For i = 1 To Len(badChars)
        s = Replace(s, Mid(badChars, i, 1), "-")
    Next i
    If Len(s) > 31 Then s = Left(s, 31)
    ' Trim trailing "-" left over from trailing "/"
    Do While Right(s, 1) = "-"
        s = Left(s, Len(s) - 1)
    Loop
    SanitizeSheetName = s
End Function
