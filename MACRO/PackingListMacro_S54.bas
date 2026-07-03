Option Explicit

'====================================================================
' HACKETT PACKING LIST - CARTON DISTRIBUTION MACRO
'
' INPUT CELLS (on the active/packing-list sheet):
'   S54      = Carton Capacity (18 / 21 / 24)
'   I55:N55  = Order Qty per size: XS, S, M, L, XL, XXL
'
' OUTPUT (starting at row 22):
'   Column B = Carton No.
'   Column J = Size
'   Column K = Units
'   Rows are inserted automatically if more carton rows are needed
'   than currently exist in the template.
'
'   Updates:
'     - Total Cartons (D22, carried down for each carton row)
'     - Total Pieces  (K45-equivalent, recalculated SUM formula)
'====================================================================

Sub GeneratePackingList()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Const FIRST_ROW As Long = 22

    Dim sizes(1 To 6) As String
    sizes(1) = "XS": sizes(2) = "S": sizes(3) = "M"
    sizes(4) = "L": sizes(5) = "XL": sizes(6) = "XXL"

    '--- 1. Read inputs ---------------------------------------------------
    Dim Capacity As Long
    Capacity = ws.Range("S54").Value

    If Capacity <= 0 Then
        MsgBox "Please enter a valid carton capacity (18, 21 or 24) in cell S54.", vbExclamation
        Exit Sub
    End If

    Dim Qty(1 To 6) As Long
    Dim i As Long
    Dim totalPieces As Long
    totalPieces = 0
    For i = 1 To 6
        Qty(i) = Val(ws.Cells(55, 8 + i).Value)   ' I55:N55 -> cols 9..14
        totalPieces = totalPieces + Qty(i)
    Next i

    If totalPieces = 0 Then
        MsgBox "Please enter order quantities in I55:N55.", vbExclamation
        Exit Sub
    End If

    '--- 2. Build carton list: full cartons first, then combined mixed ----
    ' cartonSize(n)  = size index (1-6) for carton-row n
    ' cartonQty(n)   = units for carton-row n
    ' cartonNo(n)    = carton number for carton-row n (mixed cartons can
    '                  span multiple rows sharing the same carton number)

    Dim maxRows As Long
    maxRows = (totalPieces \ 1) + 10   ' generous safety size for arrays
    Dim cartonSize() As Long
    Dim cartonQty() As Long
    Dim cartonNo() As Long
    ReDim cartonSize(1 To maxRows)
    ReDim cartonQty(1 To maxRows)
    ReDim cartonNo(1 To maxRows)

    Dim rowCount As Long
    rowCount = 0
    Dim cartonCounter As Long
    cartonCounter = 0

    Dim BalanceQty(1 To 6) As Long
    Dim FullCartons As Long
    Dim j As Long

    ' --- Full cartons first (per size, in XS->XXL order) ---
    For i = 1 To 6
        FullCartons = Qty(i) \ Capacity
        For j = 1 To FullCartons
            cartonCounter = cartonCounter + 1
            rowCount = rowCount + 1
            cartonSize(rowCount) = i
            cartonQty(rowCount) = Capacity
            cartonNo(rowCount) = cartonCounter
        Next j
        BalanceQty(i) = Qty(i) Mod Capacity
    Next i

    ' --- Combined mixed cartons from remaining balances ---
    Dim remaining(1 To 6) As Long
    Dim totalRemaining As Long
    totalRemaining = 0
    For i = 1 To 6
        remaining(i) = BalanceQty(i)
        totalRemaining = totalRemaining + remaining(i)
    Next i

    If totalRemaining > 0 Then
        cartonCounter = cartonCounter + 1   ' start first mixed carton

        Dim sizeIdx As Long
        Dim spaceLeft As Long
        Dim take As Long

        sizeIdx = 1
        spaceLeft = Capacity

        Do While totalRemaining > 0
            Do While sizeIdx <= 6 And remaining(sizeIdx) = 0
                sizeIdx = sizeIdx + 1
            Loop
            If sizeIdx > 6 Then Exit Do

            take = remaining(sizeIdx)
            If take > spaceLeft Then take = spaceLeft

            rowCount = rowCount + 1
            cartonSize(rowCount) = sizeIdx
            cartonQty(rowCount) = take
            cartonNo(rowCount) = cartonCounter

            remaining(sizeIdx) = remaining(sizeIdx) - take
            spaceLeft = spaceLeft - take
            totalRemaining = totalRemaining - take

            If spaceLeft = 0 And totalRemaining > 0 Then
                cartonCounter = cartonCounter + 1
                spaceLeft = Capacity
            End If
        Loop
    End If

    '--- 3. Insert rows if needed -------------------------------------------
    ' Determine how many rows are currently reserved for carton output
    ' by checking the existing block starting at FIRST_ROW: count rows
    ' until we find a row whose column B contains a non-numeric label
    ' (e.g. "TOTAL", "G.TOTAL") or an empty row marking the end.
    Dim existingRows As Long
    Dim r As Long
    existingRows = 0
    r = FIRST_ROW
    Do
        Dim bVal As Variant
        bVal = ws.Cells(r, "B").Value
        If IsNumeric(bVal) And bVal <> "" Then
            existingRows = existingRows + 1
            r = r + 1
        Else
            Exit Do
        End If
    Loop

    If existingRows = 0 Then existingRows = 1   ' at least one row assumed

    If rowCount > existingRows Then
        Dim rowsToInsert As Long
        rowsToInsert = rowCount - existingRows
        ws.Rows(FIRST_ROW + existingRows & ":" & (FIRST_ROW + existingRows + rowsToInsert - 1)).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

    '--- 4. Clear and write output -------------------------------------------
    Dim lastRow As Long
    lastRow = FIRST_ROW + Application.Max(rowCount, existingRows) - 1

    ws.Range("B" & FIRST_ROW & ":B" & lastRow).ClearContents
    ws.Range("J" & FIRST_ROW & ":K" & lastRow).ClearContents

    For r = 1 To rowCount
        ws.Cells(FIRST_ROW + r - 1, "B").Value = cartonNo(r)
        ws.Cells(FIRST_ROW + r - 1, "J").Value = sizes(cartonSize(r))
        ws.Cells(FIRST_ROW + r - 1, "K").Value = cartonQty(r)
    Next r

    '--- 5. Update Total Cartons and Total Pieces ----------------------------
    Dim totalCartons As Long
    totalCartons = cartonCounter

    ws.Range("D" & FIRST_ROW).Value = totalCartons
    For r = FIRST_ROW + 1 To FIRST_ROW + rowCount - 1
        ws.Cells(r, "D").Formula = "=D" & FIRST_ROW
    Next r

    ' Total Pieces - SUM of column K across the output rows
    Dim totalPiecesCell As Range
    Set totalPiecesCell = ws.Range("K" & (FIRST_ROW + rowCount + 1))
    totalPiecesCell.Formula = "=SUM(K" & FIRST_ROW & ":K" & (FIRST_ROW + rowCount - 1) & ")"
    ws.Cells(FIRST_ROW + rowCount + 1, "B").Value = "TOTAL PIECES"

    '--- 6. Done --------------------------------------------------------------
    MsgBox "Packing list generated:" & vbCrLf & _
           "Total Cartons: " & totalCartons & vbCrLf & _
           "Total Pieces: " & totalPieces & vbCrLf & _
           "Carton Capacity: " & Capacity, vbInformation

End Sub
