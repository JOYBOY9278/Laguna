Option Explicit

Sub GeneratePackingList()
    Dim ws As Worksheet
    Dim msgLog As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets
        If UCase(ws.Name) <> "SUMMARY" Then
            Dim result As String
            result = ProcessSheet(ws)
            If result <> "" Then msgLog = msgLog & result & vbCrLf
        End If
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If msgLog = "" Then
        MsgBox "Packing List Generated Successfully", vbInformation
    Else
        MsgBox "Packing List Generated with warnings:" & vbCrLf & vbCrLf & msgLog, vbExclamation
    End If
End Sub

'Returns a warning string if the sheet was skipped or truncated, else ""
Private Function ProcessSheet(ws As Worksheet) As String
    Const OUTPUT_START_ROW As Long = 22

    ProcessSheet = ""

    '=========================================
    ' READ CAPACITY
    '=========================================
    Dim Capacity As Long
    Capacity = ws.Range("P2").Value
    If Capacity <= 0 Then Exit Function

    '=========================================
    ' READ ORDER QUANTITIES (XS,S,M,L,XL,XXL)
    '=========================================
    Dim Qty(1 To 6) As Long
    Qty(1) = Val(ws.Range("I39").Value)
    Qty(2) = Val(ws.Range("J39").Value)
    Qty(3) = Val(ws.Range("K39").Value)
    Qty(4) = Val(ws.Range("L39").Value)
    Qty(5) = Val(ws.Range("M39").Value)
    Qty(6) = Val(ws.Range("N39").Value)

    Dim SizeName(1 To 6) As String
    SizeName(1) = "XS"
    SizeName(2) = "S"
    SizeName(3) = "M"
    SizeName(4) = "L"
    SizeName(5) = "XL"
    SizeName(6) = "XXL"

    '=========================================
    ' FIND BOUNDS (G.TOTAL row) - search column D
    ' for the label "G.TOTAL" starting below the
    ' output start row
    '=========================================
    Dim gTotalRow As Long
    Dim searchRow As Long
    gTotalRow = 0
    For searchRow = OUTPUT_START_ROW To OUTPUT_START_ROW + 200
        If UCase(Trim(ws.Cells(searchRow, "D").Value)) = "G.TOTAL" Then
            gTotalRow = searchRow
            Exit For
        End If
    Next searchRow

    Dim maxRow As Long
    If gTotalRow > 0 Then
        maxRow = gTotalRow - 1   ' last usable carton row
    Else
        maxRow = OUTPUT_START_ROW + 200   ' fallback: no G.TOTAL found
    End If

    '=========================================
    ' CLEAR OLD DATA (only within usable range)
    '=========================================
    ws.Range("J" & OUTPUT_START_ROW & ":K" & maxRow).ClearContents

    Dim r As Long
    Dim i As Long
    Dim j As Long
    r = OUTPUT_START_ROW

    '=========================================
    ' FULL CARTONS FIRST
    '=========================================
    Dim FullCartons As Long
    Dim BalanceQty(1 To 6) As Long

    For i = 1 To 6
        FullCartons = Qty(i) \ Capacity
        For j = 1 To FullCartons
            If r > maxRow Then GoTo Overflow
            ws.Cells(r, "J").Value = SizeName(i)
            ws.Cells(r, "K").Value = Capacity
            r = r + 1
        Next j
        BalanceQty(i) = Qty(i) Mod Capacity
    Next i

    '=========================================
    ' BALANCES - COMBINE INTO SHARED MIXED CARTONS
    ' (continuous fill across the remaining sizes
    '  in XS->XXL order)
    '=========================================
    Dim remaining(1 To 6) As Long
    For i = 1 To 6
        remaining(i) = BalanceQty(i)
    Next i

    Dim spaceLeft As Long
    Dim sizeIdx As Long
    Dim take As Long
    Dim totalRemaining As Long

    totalRemaining = 0
    For i = 1 To 6
        totalRemaining = totalRemaining + remaining(i)
    Next i

    sizeIdx = 1
    spaceLeft = Capacity
    Dim cartonOpen As Boolean
    cartonOpen = False

    Do While totalRemaining > 0
        ' advance to next size with remaining qty
        Do While sizeIdx <= 6 And remaining(sizeIdx) = 0
            sizeIdx = sizeIdx + 1
        Loop
        If sizeIdx > 6 Then Exit Do

        If r > maxRow Then GoTo Overflow

        take = remaining(sizeIdx)
        If take > spaceLeft Then take = spaceLeft

        ws.Cells(r, "J").Value = SizeName(sizeIdx)
        ws.Cells(r, "K").Value = take
        cartonOpen = True

        remaining(sizeIdx) = remaining(sizeIdx) - take
        spaceLeft = spaceLeft - take
        totalRemaining = totalRemaining - take

        r = r + 1

        If spaceLeft = 0 Then
            spaceLeft = Capacity
            cartonOpen = False
        End If
    Loop

    GoTo Finish

Overflow:
    ProcessSheet = "Sheet '" & ws.Name & "': output exceeds available rows " & _
                    "(G.TOTAL at row " & gTotalRow & "). Data truncated - increase " & _
                    "carton capacity or expand the carton block."

Finish:
    '=========================================
    ' UPDATE TOTAL PIECES
    '=========================================
    If r > OUTPUT_START_ROW Then
        ws.Range("K45").Formula = "=SUM(K" & OUTPUT_START_ROW & ":K" & r - 1 & ")"
    Else
        ws.Range("K45").Value = 0
    End If

    '=========================================
    ' UPDATE TOTAL CARTONS
    '=========================================
    ws.Range("D22").Value = r - OUTPUT_START_ROW

End Function
