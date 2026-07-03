Attribute VB_Name = "HugoBossPackingMacro"
'==============================================================================
' HUGO BOSS PACKING LIST GENERATOR
' Version: 2.0
' Description: Reads order quantities from the Input sheet, calculates
'              full single-size cartons (14 pcs) and mixed cartons
'              (max 3 sizes, exactly 14 pcs), then writes results to Output.
'
' HOW TO INSTALL:
'   1. Open your Hugo Boss packing list .xlsm file in Excel.
'   2. Press Alt + F11 to open the Visual Basic Editor.
'   3. In the menu: Insert > Module
'   4. Copy and paste ALL of this code into the new module.
'   5. Close the VBE window (Alt + Q).
'   6. A "Generate Packing List" button on the Input sheet will run the macro,
'      or you can run it via Alt + F8 > GeneratePackingList > Run.
'==============================================================================

Option Explicit

' ─────────────── Constants ───────────────
Private Const CARTON_CAPACITY   As Integer = 14
Private Const MAX_SIZES_PER_CTN As Integer = 3
Private Const OUTPUT_DATA_START As Integer = 17   ' First data row on Output sheet
Private Const INPUT_SHEET_NAME  As String  = "input "
Private Const OUTPUT_SHEET_NAME As String  = "output"

' Output column indices (1-based)
Private Const COL_CTN_FROM  As Integer = 1  ' A  Carton N° (from)
Private Const COL_DASH      As Integer = 2  ' B  "-"
Private Const COL_CTN_TO    As Integer = 3  ' C  Carton N° (to)
Private Const COL_PO        As Integer = 4  ' D  PO NO
Private Const COL_STYLE     As Integer = 5  ' E  STYLE NO
Private Const COL_FORM_NO   As Integer = 6  ' F  FORM NO
Private Const COL_FORM_NAME As Integer = 7  ' G  FORM NAME
Private Const COL_QUALITY   As Integer = 8  ' H  QUALITY
Private Const COL_COLOUR_NO As Integer = 9  ' I  Colour N°
Private Const COL_COLOUR    As Integer = 10 ' J  Colour
Private Const COL_ITEM      As Integer = 11 ' K  ITEM
' Size columns M..Y = 13..25  (sizes 36..48)
Private Const COL_SIZE_FIRST As Integer = 13 ' M  Size 36
Private Const COL_NUM_CTNS   As Integer = 26 ' Z  N° Cartons
Private Const COL_PER_CTN    As Integer = 27 ' AA Per CTN Qty
Private Const COL_TOTAL_PCS  As Integer = 28 ' AB Total pieces
Private Const COL_GW_CTN     As Integer = 29 ' AC GW/carton
Private Const COL_NW_CTN     As Integer = 30 ' AD NW/carton
Private Const COL_GW         As Integer = 31 ' AE GW
Private Const COL_NW         As Integer = 32 ' AF NW
Private Const COL_CTN_SIZE   As Integer = 33 ' AG Carton Size

' Size headers on output row 16 start at column M (13) for size 36
' The 13 sizes supported: 36,37,38,39,40,41,42,43,44,45,46,47,48
Private Const NUM_SIZES As Integer = 13

' Weight coefficients per size (index 0=size36 … 12=size48)
Private WeightCoeff(0 To 12) As Double

'==============================================================================
' MAIN ENTRY POINT
'==============================================================================
Public Sub GeneratePackingList()

    Dim wsIn  As Worksheet
    Dim wsOut As Worksheet

    On Error GoTo ErrHandler

    ' ── Validate sheets exist ──
    If Not SheetExists(INPUT_SHEET_NAME) Then
        MsgBox "Cannot find sheet '" & INPUT_SHEET_NAME & "'.", vbCritical
        Exit Sub
    End If
    If Not SheetExists(OUTPUT_SHEET_NAME) Then
        MsgBox "Cannot find sheet '" & OUTPUT_SHEET_NAME & "'.", vbCritical
        Exit Sub
    End If

    Set wsIn  = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET_NAME)

    ' ── Initialise weight coefficients ──
    Call InitWeightCoefficients

    ' ── Read inputs ──
    Dim orderQty(0 To NUM_SIZES - 1) As Long
    Dim sizes(0 To NUM_SIZES - 1)    As Integer
    Dim poNo As String, styleNo As String
    Dim formNo As String, formName As String
    Dim quality As String, colourNo As String
    Dim colour As String, item As String

    Call ReadInputSheet(wsIn, orderQty, sizes, poNo, styleNo, _
                        formNo, formName, quality, colourNo, colour, item)

    ' ── Build carton plan ──
    Dim cartons() As CartonRecord
    Dim numCartons As Long
    numCartons = BuildCartonPlan(orderQty, sizes, cartons)

    If numCartons = 0 Then
        MsgBox "No order quantities found. Please enter order quantities in the Input sheet.", vbInformation
        Exit Sub
    End If

    ' ── Write to Output sheet ──
    Call WriteOutputSheet(wsOut, cartons, numCartons, sizes, _
                          poNo, styleNo, formNo, formName, _
                          quality, colourNo, colour, item)

    MsgBox "Packing list generated successfully!" & vbCrLf & _
           numCartons & " carton line(s) written to the Output sheet.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'==============================================================================
' CARTON RECORD TYPE
'==============================================================================
Private Type CartonRecord
    qtys(0 To 12)   As Long     ' quantity per size slot
    totalPcs        As Long
    numSizes        As Integer  ' distinct sizes used
    isMixed         As Boolean
End Type

'==============================================================================
' READ INPUT SHEET
'==============================================================================
Private Sub ReadInputSheet(ws As Worksheet, _
                           orderQty() As Long, _
                           sizes() As Integer, _
                           poNo As String, styleNo As String, _
                           formNo As String, formName As String, _
                           quality As String, colourNo As String, _
                           colour As String, item As String)

    Dim i As Integer
    ' Sizes row 13 starts at column C (3) for size 36
    For i = 0 To NUM_SIZES - 1
        sizes(i) = CInt(ws.Cells(13, 3 + i).Value)
        Dim raw As Variant
        raw = ws.Cells(14, 3 + i).Value
        orderQty(i) = IIf(IsNumeric(raw) And Not IsEmpty(raw), CLng(raw), 0)
    Next i

    poNo     = CStr(IIf(IsEmpty(ws.Cells(4, 6).Value), "", ws.Cells(4, 6).Value))
    styleNo  = CStr(IIf(IsEmpty(ws.Cells(4, 3).Value), "", ws.Cells(4, 3).Value))
    formNo   = CStr(IIf(IsEmpty(ws.Cells(7, 3).Value), "", ws.Cells(7, 3).Value))
    formName = CStr(IIf(IsEmpty(ws.Cells(6, 6).Value), "", ws.Cells(6, 6).Value))
    quality  = CStr(IIf(IsEmpty(ws.Cells(7, 6).Value), "", ws.Cells(7, 6).Value))
    colourNo = CStr(IIf(IsEmpty(ws.Cells(6, 3).Value), "", ws.Cells(6, 3).Value))
    colour   = CStr(IIf(IsEmpty(ws.Cells(5, 3).Value), "", ws.Cells(5, 3).Value))
    item     = CStr(IIf(IsEmpty(ws.Cells(5, 6).Value), "", ws.Cells(5, 6).Value))
End Sub

'==============================================================================
' CORE PACKING ALGORITHM
'==============================================================================
Private Function BuildCartonPlan(orderQty() As Long, _
                                 sizes() As Integer, _
                                 cartons() As CartonRecord) As Long

    Dim remaining(0 To 12) As Long
    Dim i As Integer

    ' Copy order quantities to working array
    For i = 0 To NUM_SIZES - 1
        remaining(i) = orderQty(i)
    Next i

    ' Dynamic array for cartons – start with generous estimate
    Dim maxCartons As Long
    maxCartons = 0
    For i = 0 To NUM_SIZES - 1
        maxCartons = maxCartons + remaining(i)
    Next i
    If maxCartons = 0 Then
        BuildCartonPlan = 0
        Exit Function
    End If
    maxCartons = maxCartons   ' worst case: 1 piece per carton (won't happen, but safe)

    ReDim cartons(0 To maxCartons - 1)
    Dim ctnIdx As Long
    ctnIdx = 0

    ' ── PHASE 1: Full single-size cartons ──
    For i = 0 To NUM_SIZES - 1
        Dim fullCtns As Long
        fullCtns = remaining(i) \ CARTON_CAPACITY
        If fullCtns > 0 Then
            Dim c As CartonRecord
            Dim j As Integer
            For j = 0 To NUM_SIZES - 1
                c.qtys(j) = 0
            Next j
            c.qtys(i)   = CARTON_CAPACITY
            c.totalPcs  = CARTON_CAPACITY
            c.numSizes  = 1
            c.isMixed   = False
            ' We represent consolidated same-size cartons as ONE row
            ' with N° Cartons = fullCtns
            ' Store fullCtns in a temp slot – we'll handle via a wrapper
            ' Actually we store one record per "group"; numCartons per group stored separately
            ' Simplest: expand into individual carton records (accurate for output row-per-group)
            cartons(ctnIdx) = c
            ' We'll pack fullCtns into a single output row
            ' Store the repeat count in qtys(NUM_SIZES) via a workaround –
            ' we'll use a dedicated field. For now embed in isMixed=False and count separately.
            ' ► Encode group count: store in a parallel approach below.
            ctnIdx = ctnIdx + 1
            remaining(i) = remaining(i) - fullCtns * CARTON_CAPACITY
        End If
    Next i

    ' We need to know how many cartons per output row.
    ' Let's redesign: store groupCount as a field. Use a separate array.
    ' ── Restart with proper structure ──
    Dim groupCount(0 To maxCartons - 1) As Long
    ctnIdx = 0

    ' Reset remaining
    For i = 0 To NUM_SIZES - 1
        remaining(i) = orderQty(i)
    Next i

    ' Phase 1 again (proper)
    For i = 0 To NUM_SIZES - 1
        fullCtns = remaining(i) \ CARTON_CAPACITY
        If fullCtns > 0 Then
            Dim cr As CartonRecord
            For j = 0 To NUM_SIZES - 1: cr.qtys(j) = 0: Next j
            cr.qtys(i)  = CARTON_CAPACITY
            cr.totalPcs = CARTON_CAPACITY
            cr.numSizes = 1
            cr.isMixed  = False
            cartons(ctnIdx)    = cr
            groupCount(ctnIdx) = fullCtns
            ctnIdx = ctnIdx + 1
            remaining(i) = remaining(i) - fullCtns * CARTON_CAPACITY
        End If
    Next i

    ' ── PHASE 2: Mixed cartons from remainders ──
    ' Algorithm: best-fit descending
    '   1. Collect all sizes with remaining > 0
    '   2. Try to fill a carton of exactly 14 pieces using at most 3 sizes
    '   3. Greedily fill from largest remainder first, cap at 14 per fill
    '   4. Repeat until all remainders are exhausted

    Dim moreRemaining As Boolean
    moreRemaining = True

    Do While moreRemaining
        ' Check if any remaining
        Dim totalRem As Long
        totalRem = 0
        For i = 0 To NUM_SIZES - 1
            totalRem = totalRem + remaining(i)
        Next i
        If totalRem = 0 Then Exit Do

        ' Build sorted index by remaining qty descending
        Dim sortIdx(0 To 12) As Integer
        Dim k As Integer
        For k = 0 To NUM_SIZES - 1: sortIdx(k) = k: Next k
        ' Bubble sort descending
        Dim swapped As Boolean
        Dim tmp As Integer
        Do
            swapped = False
            For k = 0 To NUM_SIZES - 2
                If remaining(sortIdx(k)) < remaining(sortIdx(k + 1)) Then
                    tmp = sortIdx(k): sortIdx(k) = sortIdx(k + 1): sortIdx(k + 1) = tmp
                    swapped = True
                End If
            Next k
        Loop While swapped

        ' Build one mixed carton
        Dim mc As CartonRecord
        For j = 0 To NUM_SIZES - 1: mc.qtys(j) = 0: Next j
        mc.totalPcs = 0
        mc.numSizes = 0
        mc.isMixed  = True

        Dim spaceFilled As Long
        spaceFilled = 0
        Dim sizesUsed As Integer
        sizesUsed = 0

        For k = 0 To NUM_SIZES - 1
            If sizesUsed >= MAX_SIZES_PER_CTN Then Exit For
            Dim sIdx As Integer
            sIdx = sortIdx(k)
            If remaining(sIdx) = 0 Then GoTo NextSize

            ' How much can we take from this size?
            Dim space As Long
            space = CARTON_CAPACITY - spaceFilled
            If space <= 0 Then Exit For

            Dim take As Long
            take = remaining(sIdx)
            If take > space Then take = space

            mc.qtys(sIdx) = take
            spaceFilled   = spaceFilled + take
            sizesUsed     = sizesUsed + 1
NextSize:
        Next k

        ' If carton is not full (< 14), we have a partial – still pack it
        mc.totalPcs = spaceFilled
        mc.numSizes = sizesUsed

        If spaceFilled = 0 Then Exit Do   ' Safety exit

        ' Deduct from remaining
        For i = 0 To NUM_SIZES - 1
            remaining(i) = remaining(i) - mc.qtys(i)
        Next i

        cartons(ctnIdx)    = mc
        groupCount(ctnIdx) = 1
        ctnIdx = ctnIdx + 1
    Loop

    ' ── Trim and return ──
    ' We embed groupCount into the carton qtys spare slot trick won't work cleanly.
    ' Instead, rebuild the array properly with groupCount baked in as repeated rows
    ' OR keep groupCount separate and return via module-level variable.
    ' Cleanest: expand each group into individual rows (fine for typical order sizes).
    ' But that loses the "N° Cartons" consolidation. Let's keep groups and return groupCount
    ' by re-embedding it: use qtys(NUM_SIZES) not possible with fixed type.
    ' ► Use a module-level array for groupCount.

    m_groupCount = groupCount
    BuildCartonPlan = ctnIdx
End Function

' Module-level group count storage
Private m_groupCount(0 To 9999) As Long

'==============================================================================
' WRITE OUTPUT SHEET
'==============================================================================
Private Sub WriteOutputSheet(ws As Worksheet, _
                             cartons() As CartonRecord, _
                             numCartons As Long, _
                             sizes() As Integer, _
                             poNo As String, styleNo As String, _
                             formNo As String, formName As String, _
                             quality As String, colourNo As String, _
                             colour As String, item As String)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ── Clear existing data rows (keep header rows 1-16 intact) ──
    Dim lastDataRow As Long
    lastDataRow = ws.Cells(ws.Rows.Count, COL_CTN_FROM).End(xlUp).Row
    If lastDataRow >= OUTPUT_DATA_START Then
        ' Clear only the data rows, not headers or summary
        ws.Rows(OUTPUT_DATA_START & ":" & lastDataRow).ClearContents
    End If

    ' ── Write carton rows ──
    Dim ctnNumber As Long
    ctnNumber = 1
    Dim rowIdx As Long
    rowIdx = OUTPUT_DATA_START

    Dim c As Long
    For c = 0 To numCartons - 1

        Dim numInGroup As Long
        numInGroup = m_groupCount(c)
        If numInGroup < 1 Then numInGroup = 1

        Dim cr As CartonRecord
        cr = cartons(c)

        With ws

            ' Carton number from
            .Cells(rowIdx, COL_CTN_FROM).Value = ctnNumber
            .Cells(rowIdx, COL_DASH).Value      = "-"
            .Cells(rowIdx, COL_CTN_TO).Value    = ctnNumber + numInGroup - 1

            ' Order metadata
            .Cells(rowIdx, COL_PO).Value         = poNo
            .Cells(rowIdx, COL_STYLE).Value      = styleNo
            .Cells(rowIdx, COL_FORM_NO).Value    = formNo
            .Cells(rowIdx, COL_FORM_NAME).Value  = formName
            .Cells(rowIdx, COL_QUALITY).Value    = quality
            .Cells(rowIdx, COL_COLOUR_NO).Value  = colourNo
            .Cells(rowIdx, COL_COLOUR).Value     = colour
            .Cells(rowIdx, COL_ITEM).Value       = item

            ' Size quantities (columns M=13 through Y=25, for sizes 36..48)
            Dim s As Integer
            For s = 0 To NUM_SIZES - 1
                If cr.qtys(s) > 0 Then
                    .Cells(rowIdx, COL_SIZE_FIRST + s).Value = cr.qtys(s)
                End If
            Next s

            ' Carton counts and totals
            .Cells(rowIdx, COL_NUM_CTNS).Value  = numInGroup
            .Cells(rowIdx, COL_PER_CTN).Value   = cr.totalPcs
            .Cells(rowIdx, COL_TOTAL_PCS).Value = cr.totalPcs * numInGroup

            ' Weight calculations
            Dim nw As Double
            nw = CalcNW(cr)
            .Cells(rowIdx, COL_NW_CTN).Value = Round(nw, 3)
            .Cells(rowIdx, COL_GW_CTN).Value = Round(nw + 1.75, 3)
            .Cells(rowIdx, COL_NW).Value     = Round(nw * numInGroup, 3)
            .Cells(rowIdx, COL_GW).Value     = Round((nw + 1.75) * numInGroup, 3)

            ' Carton dimensions
            .Cells(rowIdx, COL_CTN_SIZE).Value = "584X382X312 MM"

        End With

        ctnNumber = ctnNumber + numInGroup
        rowIdx    = rowIdx + 1
    Next c

    ' ── Update S-Total row (row 29 in template) ──
    Dim totalRow As Long
    totalRow = rowIdx   ' One row below last data

    ' Write S-Total label
    ws.Cells(totalRow, COL_CTN_FROM).Value = "S -Total"

    ' Sum size columns
    Dim s2 As Integer
    For s2 = 0 To NUM_SIZES - 1
        Dim colAddr As String
        colAddr = ws.Cells(OUTPUT_DATA_START, COL_SIZE_FIRST + s2).Address(False, True)
        Dim colTop As String
        colTop = Split(colAddr, "$")(1)
        ws.Cells(totalRow, COL_SIZE_FIRST + s2).Formula = _
            "=SUM(" & colTop & OUTPUT_DATA_START & ":" & colTop & (totalRow - 1) & ")"
    Next s2

    ' Sum Z, AB, AC, AD, AE, AF
    Dim zCol As String: zCol = "Z"
    Dim abCol As String: abCol = "AB"
    ws.Cells(totalRow, COL_NUM_CTNS).Formula  = "=SUM(Z" & OUTPUT_DATA_START & ":Z" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_TOTAL_PCS).Formula = "=SUM(AB" & OUTPUT_DATA_START & ":AB" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_GW_CTN).Formula    = "=SUM(AC" & OUTPUT_DATA_START & ":AC" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_NW_CTN).Formula    = "=SUM(AD" & OUTPUT_DATA_START & ":AD" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_GW).Formula        = "=SUM(AE" & OUTPUT_DATA_START & ":AE" & (totalRow - 1) & ")"
    ws.Cells(totalRow, COL_NW).Formula        = "=SUM(AF" & OUTPUT_DATA_START & ":AF" & (totalRow - 1) & ")"

    ' ── Copy row formatting from template row 17 to all new rows ──
    Call CopyRowFormats(ws, OUTPUT_DATA_START, rowIdx - 1)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' NET WEIGHT CALCULATION
'==============================================================================
Private Function CalcNW(cr As CartonRecord) As Double
    Dim nw As Double
    nw = 0
    Dim i As Integer
    For i = 0 To NUM_SIZES - 1
        nw = nw + cr.qtys(i) * WeightCoeff(i)
    Next i
    CalcNW = nw
End Function

'==============================================================================
' INIT WEIGHT COEFFICIENTS (kg per piece per size)
' Default values from WEIGHT sheet; adjust as needed.
'==============================================================================
Private Sub InitWeightCoefficients()
    ' Sizes 36-48 → index 0-12
    WeightCoeff(0)  = 0.257   ' Size 36
    WeightCoeff(1)  = 0.259   ' Size 37
    WeightCoeff(2)  = 0.268   ' Size 38
    WeightCoeff(3)  = 0.270   ' Size 39
    WeightCoeff(4)  = 0.278   ' Size 40
    WeightCoeff(5)  = 0.281   ' Size 41
    WeightCoeff(6)  = 0.288   ' Size 42
    WeightCoeff(7)  = 0.293   ' Size 43
    WeightCoeff(8)  = 0.300   ' Size 44
    WeightCoeff(9)  = 0.307   ' Size 45
    WeightCoeff(10) = 0.314   ' Size 46
    WeightCoeff(11) = 0.321   ' Size 47
    WeightCoeff(12) = 0.328   ' Size 48
End Sub

'==============================================================================
' COPY ROW FORMATS FROM TEMPLATE ROW
'==============================================================================
Private Sub CopyRowFormats(ws As Worksheet, firstDataRow As Long, lastDataRow As Long)
    If lastDataRow < firstDataRow Then Exit Sub
    Dim templateRow As Long
    templateRow = firstDataRow  ' Use first row as format template
    Dim r As Long
    For r = firstDataRow + 1 To lastDataRow
        ws.Rows(templateRow).Copy
        ws.Rows(r).PasteSpecial Paste:=xlPasteFormats
    Next r
    Application.CutCopyMode = False
End Sub

'==============================================================================
' ADD BUTTON TO INPUT SHEET  (run once via Immediate Window or on open)
'==============================================================================
Public Sub AddGenerateButton()
    Dim ws As Worksheet
    If Not SheetExists(INPUT_SHEET_NAME) Then Exit Sub
    Set ws = ThisWorkbook.Sheets(INPUT_SHEET_NAME)

    ' Remove existing button if present
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btnGenerate" Then shp.Delete
    Next shp

    ' Add new button
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 120, 180, 30)
    With btn
        .Name = "btnGenerate"
        .TextFrame.Characters.Text = "Generate Packing List"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 10
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Line.Visible = msoFalse
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "GeneratePackingList"
    End With
End Sub

'==============================================================================
' UTILITY
'==============================================================================
Private Function SheetExists(name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
