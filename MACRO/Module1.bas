Attribute VB_Name = "Module1"
Attribute VB_Name = "HugoBossPackingMacro"
'==============================================================================
' HUGO BOSS PACKING LIST GENERATOR - FIXED VERSION 2.3
' Fully working with your workbook
'==============================================================================

Option Explicit

' ====================== CONSTANTS ======================
Private Const CARTON_CAPACITY As Integer = 14
Private Const MAX_SIZES_PER_CTN As Integer = 3
Private Const OUTPUT_DATA_START As Integer = 17
Private Const INPUT_SHEET_NAME As String = "input "
Private Const OUTPUT_SHEET_NAME As String = "output"

' Column indices
Private Const COL_CTN_FROM As Integer = 1
Private Const COL_DASH As Integer = 2
Private Const COL_CTN_TO As Integer = 3
Private Const COL_PO As Integer = 4
Private Const COL_STYLE As Integer = 5
Private Const COL_FORM_NO As Integer = 6
Private Const COL_FORM_NAME As Integer = 7
Private Const COL_QUALITY As Integer = 8
Private Const COL_COLOUR_NO As Integer = 9
Private Const COL_COLOUR As Integer = 10
Private Const COL_ITEM As Integer = 11
Private Const COL_SIZE_FIRST As Integer = 13
Private Const COL_NUM_CTNS As Integer = 26
Private Const COL_PER_CTN As Integer = 27
Private Const COL_TOTAL_PCS As Integer = 28
Private Const COL_GW_CTN As Integer = 29
Private Const COL_NW_CTN As Integer = 30
Private Const COL_GW As Integer = 31
Private Const COL_NW As Integer = 32
Private Const COL_CTN_SIZE As Integer = 33

Private Const NUM_SIZES As Integer = 13

' Module level variables
Private WeightCoeff(0 To 12) As Double
Private m_groupCount(0 To 9999) As Long

' ====================== TYPE DEFINITION ======================
Private Type CartonRecord
    qtys(0 To 12) As Long
    totalPcs As Long
    numSizes As Integer
    isMixed As Boolean
End Type

' ====================== MAIN MACRO ======================
Public Sub GeneratePackingList()
    Dim wsIn As Worksheet, wsOut As Worksheet
    On Error GoTo ErrHandler

    If Not SheetExists(INPUT_SHEET_NAME) Then
        MsgBox "Input sheet not found!", vbCritical
        Exit Sub
    End If
    If Not SheetExists(OUTPUT_SHEET_NAME) Then
        MsgBox "Output sheet not found!", vbCritical
        Exit Sub
    End If

    Set wsIn = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET_NAME)

    Call InitWeightCoefficients

    Dim orderQty(0 To NUM_SIZES - 1) As Long
    Dim poNo As String, styleNo As String, formNo As String, formName As String
    Dim quality As String, colourNo As String, colour As String, item As String

    Call ReadInputSheet(wsIn, orderQty, poNo, styleNo, formNo, formName, quality, colourNo, colour, item)

    Dim cartons() As CartonRecord
    Dim numCartons As Long
    numCartons = BuildCartonPlan(orderQty, cartons)

    If numCartons = 0 Then
        MsgBox "No quantities entered in Input sheet.", vbInformation
        Exit Sub
    End If

    Call WriteOutputSheet(wsOut, cartons, numCartons, poNo, styleNo, formNo, formName, quality, colourNo, colour, item)

    MsgBox "? Success! " & numCartons & " carton groups generated.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ====================== READ INPUT ======================
Private Sub ReadInputSheet(ws As Worksheet, orderQty() As Long, _
                           poNo As String, styleNo As String, formNo As String, formName As String, _
                           quality As String, colourNo As String, colour As String, item As String)
    
    Dim i As Integer
    For i = 0 To NUM_SIZES - 1
        Dim raw As Variant
        raw = ws.Cells(11, 3 + i).Value          ' Row 11, columns C to J
        orderQty(i) = IIf(IsNumeric(raw) And Not IsEmpty(raw), CLng(raw), 0)
    Next i

    poNo = Trim(CStr(ws.Cells(4, 6).Value))
    styleNo = Trim(CStr(ws.Cells(4, 3).Value))
    formNo = Trim(CStr(ws.Cells(7, 3).Value))
    formName = Trim(CStr(ws.Cells(6, 6).Value))
    quality = Trim(CStr(ws.Cells(7, 6).Value))
    colourNo = Trim(CStr(ws.Cells(6, 3).Value))
    colour = Trim(CStr(ws.Cells(5, 3).Value))
    item = Trim(CStr(ws.Cells(5, 6).Value))
End Sub

' ====================== BUILD CARTONS ======================
Private Function BuildCartonPlan(orderQty() As Long, cartons() As CartonRecord) As Long
    Dim remaining(0 To 12) As Long, i As Integer, ctnIdx As Long
    Dim groupCount(0 To 9999) As Long

    For i = 0 To NUM_SIZES - 1
        remaining(i) = orderQty(i)
    Next i

    ReDim cartons(0 To 9999)
    ctnIdx = 0

    ' PHASE 1: Full single-size cartons
    For i = 0 To NUM_SIZES - 1
        Dim fullCtns As Long
        fullCtns = remaining(i) \ CARTON_CAPACITY
        If fullCtns > 0 Then
            Dim cr As CartonRecord
            Dim j As Integer
            For j = 0 To NUM_SIZES - 1: cr.qtys(j) = 0: Next j
            cr.qtys(i) = CARTON_CAPACITY
            cr.totalPcs = CARTON_CAPACITY
            cr.numSizes = 1
            cr.isMixed = False

            cartons(ctnIdx) = cr
            groupCount(ctnIdx) = fullCtns
            ctnIdx = ctnIdx + 1
            remaining(i) = remaining(i) Mod CARTON_CAPACITY
        End If
    Next i

    ' PHASE 2: Mixed remainders
    Do While True
        Dim totalRem As Long: totalRem = 0
        For i = 0 To NUM_SIZES - 1: totalRem = totalRem + remaining(i): Next i
        If totalRem = 0 Then Exit Do

        Dim sortIdx(0 To 12) As Integer
        For i = 0 To NUM_SIZES - 1: sortIdx(i) = i: Next i
        Dim swapped As Boolean, k As Integer, tmp As Integer
        Do
            swapped = False
            For k = 0 To NUM_SIZES - 2
                If remaining(sortIdx(k)) < remaining(sortIdx(k + 1)) Then
                    tmp = sortIdx(k): sortIdx(k) = sortIdx(k + 1): sortIdx(k + 1) = tmp
                    swapped = True
                End If
            Next k
        Loop While swapped

        Dim mc As CartonRecord
        For j = 0 To NUM_SIZES - 1: mc.qtys(j) = 0: Next j
        mc.totalPcs = 0
        mc.numSizes = 0
        mc.isMixed = True

        Dim spaceFilled As Long: spaceFilled = 0
        Dim sizesUsed As Integer: sizesUsed = 0

        For k = 0 To NUM_SIZES - 1
            If sizesUsed >= MAX_SIZES_PER_CTN Then Exit For
            Dim sIdx As Integer: sIdx = sortIdx(k)
            If remaining(sIdx) <= 0 Then GoTo NextS

            Dim take As Long
            take = Application.Min(remaining(sIdx), CARTON_CAPACITY - spaceFilled)
            If take > 0 Then
                mc.qtys(sIdx) = take
                spaceFilled = spaceFilled + take
                sizesUsed = sizesUsed + 1
            End If
NextS:
        Next k

        If spaceFilled = 0 Then Exit Do

        mc.totalPcs = spaceFilled
        mc.numSizes = sizesUsed

        For i = 0 To NUM_SIZES - 1
            remaining(i) = remaining(i) - mc.qtys(i)
        Next i

        cartons(ctnIdx) = mc
        groupCount(ctnIdx) = 1
        ctnIdx = ctnIdx + 1
    Loop

    ' Save group counts
    For i = 0 To ctnIdx - 1
        m_groupCount(i) = groupCount(i)
    Next i

    BuildCartonPlan = ctnIdx
End Function

' ====================== WRITE OUTPUT ======================
Private Sub WriteOutputSheet(ws As Worksheet, cartons() As CartonRecord, numCartons As Long, _
                             poNo As String, styleNo As String, formNo As String, formName As String, _
                             quality As String, colourNo As String, colour As String, item As String)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= OUTPUT_DATA_START Then
        ws.Rows(OUTPUT_DATA_START & ":" & lastRow).ClearContents
    End If

    Dim rowIdx As Long: rowIdx = OUTPUT_DATA_START
    Dim ctnNumber As Long: ctnNumber = 1
    Dim c As Long

    For c = 0 To numCartons - 1
        Dim numInGroup As Long: numInGroup = m_groupCount(c)
        If numInGroup < 1 Then numInGroup = 1

        Dim cr As CartonRecord: cr = cartons(c)

        With ws
            .Cells(rowIdx, COL_CTN_FROM).Value = ctnNumber
            .Cells(rowIdx, COL_DASH).Value = "-"
            .Cells(rowIdx, COL_CTN_TO).Value = ctnNumber + numInGroup - 1

            .Cells(rowIdx, COL_PO).Value = poNo
            .Cells(rowIdx, COL_STYLE).Value = styleNo
            .Cells(rowIdx, COL_FORM_NO).Value = formNo
            .Cells(rowIdx, COL_FORM_NAME).Value = formName
            .Cells(rowIdx, COL_QUALITY).Value = quality
            .Cells(rowIdx, COL_COLOUR_NO).Value = colourNo
            .Cells(rowIdx, COL_COLOUR).Value = colour
            .Cells(rowIdx, COL_ITEM).Value = item

            Dim s As Integer
            For s = 0 To NUM_SIZES - 1
                If cr.qtys(s) > 0 Then .Cells(rowIdx, COL_SIZE_FIRST + s).Value = cr.qtys(s)
            Next s

            .Cells(rowIdx, COL_NUM_CTNS).Value = numInGroup
            .Cells(rowIdx, COL_PER_CTN).Value = cr.totalPcs
            .Cells(rowIdx, COL_TOTAL_PCS).Value = cr.totalPcs * numInGroup

            Dim nw As Double: nw = CalcNW(cr)
            .Cells(rowIdx, COL_NW_CTN).Value = Round(nw, 3)
            .Cells(rowIdx, COL_GW_CTN).Value = Round(nw + 1.75, 3)
            .Cells(rowIdx, COL_NW).Value = Round(nw * numInGroup, 3)
            .Cells(rowIdx, COL_GW).Value = Round((nw + 1.75) * numInGroup, 3)
            .Cells(rowIdx, COL_CTN_SIZE).Value = "584X382X312 MM"
        End With

        ctnNumber = ctnNumber + numInGroup
        rowIdx = rowIdx + 1
    Next c

    ' S-TOTAL
    Dim totalRow As Long: totalRow = rowIdx
    ws.Cells(totalRow, COL_CTN_FROM).Value = "S-Total"

    Dim s2 As Integer
    For s2 = 0 To NUM_SIZES - 1
        Dim colLetter As String
        colLetter = Split(ws.Cells(1, COL_SIZE_FIRST + s2).Address, "$")(1)
        ws.Cells(totalRow, COL_SIZE_FIRST + s2).Formula = "=SUM(" & colLetter & OUTPUT_DATA_START & ":" & colLetter & (totalRow - 1) & ")"
    Next s2

    ws.Cells(totalRow, COL_NUM_CTNS).Formula = "=SUM(Z" & OUTPUT_DATA_START & ":Z" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_TOTAL_PCS).Formula = "=SUM(AB" & OUTPUT_DATA_START & ":AB" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_GW_CTN).Formula = "=SUM(AC" & OUTPUT_DATA_START & ":AC" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_NW_CTN).Formula = "=SUM(AD" & OUTPUT_DATA_START & ":AD" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_GW).Formula = "=SUM(AE" & OUTPUT_DATA_START & ":AE" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_NW).Formula = "=SUM(AF" & OUTPUT_DATA_START & ":AF" & (totalRow - 1) & ")"

    Call CopyRowFormats(ws, OUTPUT_DATA_START, rowIdx - 1)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function CalcNW(cr As CartonRecord) As Double
    Dim nw As Double, i As Integer
    For i = 0 To NUM_SIZES - 1
        nw = nw + cr.qtys(i) * WeightCoeff(i)
    Next i
    CalcNW = nw
End Function

Private Sub InitWeightCoefficients()
    WeightCoeff(0) = 0.257: WeightCoeff(1) = 0.259: WeightCoeff(2) = 0.268
    WeightCoeff(3) = 0.27: WeightCoeff(4) = 0.278: WeightCoeff(5) = 0.281
    WeightCoeff(6) = 0.288: WeightCoeff(7) = 0.293: WeightCoeff(8) = 0.3
    WeightCoeff(9) = 0.307: WeightCoeff(10) = 0.314
    WeightCoeff(11) = 0.321: WeightCoeff(12) = 0.328
End Sub

Private Sub CopyRowFormats(ws As Worksheet, first As Long, last As Long)
    If last < first Then Exit Sub
    Dim r As Long
    For r = first + 1 To last
        ws.Rows(first).Copy
        ws.Rows(r).PasteSpecial Paste:=xlPasteFormats
    Next r
    Application.CutCopyMode = False
End Sub

Private Function SheetExists(name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

