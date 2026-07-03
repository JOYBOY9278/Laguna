Option Explicit

'====================================================================
' HACKETT PACKING LIST - CARTON DISTRIBUTION MACRO (2-sheet version)
'
' SHEETS:
'   "Input"        - holds Carton Capacity (C8) and Order Qty (C12:H12,
'                     XS-XXL). This sheet is NEVER modified by the macro,
'                     so row-insertion on the output sheet cannot break it.
'   "Packing List" - output sheet. Carton breakdown written starting at
'                     row 22:
'                       Column B = Carton No.
'                       Column D = Total Cartons (carried down)
'                       Column J = Size
'                       Column K = Units
'                     Row (22 + rowCount + 1) = TOTAL PIECES (col B/K)
'====================================================================

Sub GeneratePackingList()

    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    Set wsIn = ThisWorkbook.Sheets("Input")
    Set wsOut = ThisWorkbook.Sheets("Packing List")

    Const FIRST_ROW As Long = 22

    Dim sizes(1 To 6) As String
    sizes(1) = "XS": sizes(2) = "S": sizes(3) = "M"
    sizes(4) = "L": sizes(5) = "XL": sizes(6) = "XXL"

    '--- 1. Read inputs from "Input" sheet --------------------------------
    Dim Capacity As Long
    Capacity = wsIn.Range("C8").Value

    If Capacity <= 0 Then
        MsgBox "Please enter a valid carton capacity (18, 21 or 24) in cell C8 on the Input sheet.", vbExclamation
        Exit Sub
    End If

    Dim Qty(1 To 6) As Long
    Dim i As Long
    Dim totalPieces As Long
    totalPieces = 0
    For i = 1 To 6
        Qty(i) = Val(wsIn.Cells(12, 2 + i).Value)   ' C12:H12 -> cols 3..8
        totalPieces = totalPieces + Qty(i)
    Next i

    If totalPieces = 0 Then
        MsgBox "Please enter order quantities in C12:H12 on the Input sheet.", vbExclamation
        Exit Sub
    End If

    '--- 2. Build carton list: full cartons first, then combined mixed ----
    Dim maxRows As Long
    maxRows = totalPieces + 10   ' generous safety size for arrays

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

    '--- 3. Determine existing output block size & insert rows if needed ---
    ' Count existing carton rows by checking column B for numeric values
    ' starting at FIRST_ROW.
    Dim existingRows As Long
    Dim r As Long
    existingRows = 0
    r = FIRST_ROW
    Do
        Dim bVal As Variant
        bVal = wsOut.Cells(r, "B").Value
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
        wsOut.Rows(FIRST_ROW + existingRows & ":" & (FIRST_ROW + existingRows + rowsToInsert - 1)).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

    '--- 4. Clear and write output ------------------------------------------
    Dim lastRow As Long
    lastRow = FIRST_ROW + Application.Max(rowCount, existingRows) - 1

    wsOut.Range("B" & FIRST_ROW & ":B" & lastRow).ClearContents
    wsOut.Range("D" & FIRST_ROW & ":D" & lastRow).ClearContents
    wsOut.Range("J" & FIRST_ROW & ":K" & lastRow).ClearContents

    For r = 1 To rowCount
        wsOut.Cells(FIRST_ROW + r - 1, "B").Value = cartonNo(r)
        wsOut.Cells(FIRST_ROW + r - 1, "J").Value = sizes(cartonSize(r))
        wsOut.Cells(FIRST_ROW + r - 1, "K").Value = cartonQty(r)
    Next r

    '--- 5. Update Total Cartons and Total Pieces ----------------------------
    Dim totalCartons As Long
    totalCartons = cartonCounter

    wsOut.Range("D" & FIRST_ROW).Value = totalCartons
    For r = FIRST_ROW + 1 To FIRST_ROW + rowCount - 1
        wsOut.Cells(r, "D").Formula = "=D" & FIRST_ROW
    Next r

    ' Clear any previous TOTAL PIECES row remnants in the now-unused tail
    Dim oldTotalRowSearch As Long
    For oldTotalRowSearch = FIRST_ROW To FIRST_ROW + Application.Max(rowCount, existingRows) + 5
        If wsOut.Cells(oldTotalRowSearch, "B").Value = "TOTAL PIECES" Then
            wsOut.Cells(oldTotalRowSearch, "B").ClearContents
            wsOut.Cells(oldTotalRowSearch, "K").ClearContents
        End If
    Next oldTotalRowSearch

    Dim totalRow As Long
    totalRow = FIRST_ROW + rowCount + 1
    wsOut.Cells(totalRow, "B").Value = "TOTAL PIECES"
    wsOut.Cells(totalRow, "B").Font.Bold = True
    wsOut.Cells(totalRow, "K").Formula = "=SUM(K" & FIRST_ROW & ":K" & (FIRST_ROW + rowCount - 1) & ")"
    wsOut.Cells(totalRow, "K").Font.Bold = True

    '--- 6. Done --------------------------------------------------------------
    MsgBox "Packing list generated:" & vbCrLf & _
           "Total Cartons: " & totalCartons & vbCrLf & _
           "Total Pieces: " & totalPieces & vbCrLf & _
           "Carton Capacity: " & Capacity, vbInformation

End Sub
