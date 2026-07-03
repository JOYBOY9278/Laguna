
Option Explicit

Sub GeneratePackingList()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim Capacity As Long
    Capacity = CLng(ws.Range("S54").Value)

    If Capacity <= 0 Then
        MsgBox "Invalid capacity in S54", vbExclamation
        Exit Sub
    End If

    Dim Qty(1 To 6) As Long
    Dim Sizes(1 To 6) As String

    Sizes(1) = "XS"
    Sizes(2) = "S"
    Sizes(3) = "M"
    Sizes(4) = "L"
    Sizes(5) = "XL"
    Sizes(6) = "XXL"

    Qty(1) = Val(ws.Range("I55").Value)
    Qty(2) = Val(ws.Range("J55").Value)
    Qty(3) = Val(ws.Range("K55").Value)
    Qty(4) = Val(ws.Range("L55").Value)
    Qty(5) = Val(ws.Range("M55").Value)
    Qty(6) = Val(ws.Range("N55").Value)

    ws.Range("B22:B5000").ClearContents
    ws.Range("J22:K5000").ClearContents

    Dim r As Long, cartonNo As Long, i As Long
    Dim remQty(1 To 6) As Long, fullCartons As Long, j As Long

    r = 22
    cartonNo = 1

    ' Full cartons first
    For i = 1 To 6
        fullCartons = Qty(i) \ Capacity
        remQty(i) = Qty(i) Mod Capacity

        For j = 1 To fullCartons
            ws.Cells(r, "B").Value = cartonNo
            ws.Cells(r, "J").Value = Sizes(i)
            ws.Cells(r, "K").Value = Capacity
            cartonNo = cartonNo + 1
            r = r + 1
        Next j
    Next i

    ' Mixed cartons from remainders
    Dim idx As Long, takeQty As Long, spaceLeft As Long
    idx = 1

    Do While Remaining(remQty) > 0

        spaceLeft = Capacity

        Do While spaceLeft > 0 And idx <= 6

            If remQty(idx) = 0 Then
                idx = idx + 1
            Else
                takeQty = Application.Min(remQty(idx), spaceLeft)

                ws.Cells(r, "B").Value = cartonNo
                ws.Cells(r, "J").Value = Sizes(idx)
                ws.Cells(r, "K").Value = takeQty

                remQty(idx) = remQty(idx) - takeQty
                spaceLeft = spaceLeft - takeQty

                r = r + 1
            End If
        Loop

        cartonNo = cartonNo + 1
        idx = 1
    Loop

    MsgBox "Packing List Generated. Cartons: " & cartonNo - 1

End Sub

Private Function Remaining(arr() As Long) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Remaining = Remaining + arr(i)
    Next i
End Function
