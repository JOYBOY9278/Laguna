Option Explicit

'==============================================================================
' ALDENO UNIVERSAL CARTON STICKER GENERATOR  v2.0
' Works with ANY Aldeno packing list regardless of:
'   - Sheet names (any color/style name as sheet tab)
'   - OC Number location (auto-detects K12, M12, or scans row 12)
'   - Size columns (S/M/L/XL/XXL  OR  2XL/3XL/4XL/5XL/6XL/7XL/8XL)
'   - Total Cartons in F14 or blank (auto-counts from data rows)
'   - Carton# in column A with #REF! errors (uses row sequence instead)
'   - Multiple packing sheets per workbook (generates stickers for ALL)
'==============================================================================

' ---- Sticker output sheet name ----
Private Const STICKER_SHEET As String = "Sticker"

' ---- Sticker layout constants ----
Private Const STICKER_BLOCK_ROWS As Long = 7
Private Const STICKER_GAP_ROWS   As Long = 1
Private Const STICKER_FIRST_ROW  As Long = 2
Private Const ROW_OC      As Long = 0
Private Const ROW_STYLE   As Long = 1
Private Const ROW_PRODUCT As Long = 2
Private Const ROW_COLOR   As Long = 3
Private Const ROW_SIZE    As Long = 4
Private Const ROW_QTY     As Long = 5
Private Const ROW_CTN     As Long = 6
Private Const STICKER_VAL_COL As Long = 3  ' Column C

' ---- Sheets to SKIP (not packing data) ----
Private Const SKIP_SHEETS As String = "|Sticker|CUT QTY|⚡ SETUP INSTRUCTIONS|► VBA CODE|ShipmentData|Sheet1|Sheet2|Sheet3|Sheet4|Sheet5|Sheet6|Sheet7|Sheet8|ALDENO Status|stickers|"

'==============================================================================
' DATA TYPE
'==============================================================================
Private Type CartonInfo
    CartonNum As Long
    TotalCtns As Long
    Color     As String
    SizeText  As String   ' pre-built size string e.g. "S=10 M=20"
    TotalQty  As Long
    OcNumber  As String
    StyleName As String
    ProductID As String
End Type

'==============================================================================
' MAIN ENTRY POINT
'==============================================================================
Public Sub GenerateStickers()

    Dim wb          As Workbook
    Dim wsSticker   As Worksheet
    Dim calcMode    As XlCalculation
    Dim i           As Long

    On Error GoTo ErrHandler

    calcMode = Application.Calculation
    Application.Calculation    = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    Set wb = ThisWorkbook

    '--- Validate sticker output sheet ----------------------------------------
    If Not SheetExists(wb, STICKER_SHEET) Then
        MsgBox "ERROR: Sheet '" & STICKER_SHEET & "' not found." & vbNewLine & _
               "Please make sure a sheet named 'Sticker' exists in this workbook.", _
               vbCritical, "Sticker Sheet Missing"
        GoTo Cleanup
    End If
    Set wsSticker = wb.Worksheets(STICKER_SHEET)

    '--- Collect all packing sheets -------------------------------------------
    Dim packingSheets() As String
    Dim packCount       As Long
    packCount = 0
    ReDim packingSheets(1 To wb.Sheets.Count)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If IsPackingSheet(ws) Then
            packCount = packCount + 1
            packingSheets(packCount) = ws.Name
        End If
    Next ws

    If packCount = 0 Then
        MsgBox "No packing list sheets found in this workbook." & vbNewLine & _
               "Sheets named 'Sticker', 'CUT QTY', 'Sheet1-8' etc. are skipped automatically.", _
               vbExclamation, "No Packing Sheets"
        GoTo Cleanup
    End If

    '--- Read ALL cartons from ALL packing sheets ------------------------------
    Dim allCartons()  As CartonInfo
    Dim totalCartons  As Long
    totalCartons = 0
    ReDim allCartons(1 To 2000)

    Dim sheetList As String
    sheetList = ""
    Dim s As Long

    For s = 1 To packCount
        Dim wsPL As Worksheet
        Set wsPL = wb.Worksheets(packingSheets(s))

        Dim sheetCartons() As CartonInfo
        Dim sheetCount     As Long
        sheetCount = ReadCartonSheet(wsPL, sheetCartons)

        If sheetCount > 0 Then
            sheetList = sheetList & packingSheets(s) & " (" & sheetCount & " cartons)" & vbNewLine

            Dim c As Long
            For c = 1 To sheetCount
                totalCartons = totalCartons + 1
                If totalCartons > 2000 Then
                    MsgBox "More than 2000 carton rows detected. Please check the packing list.", _
                           vbCritical, "Too Many Cartons"
                    GoTo Cleanup
                End If
                allCartons(totalCartons) = sheetCartons(c)
            Next c
        End If
    Next s

    If totalCartons = 0 Then
        MsgBox "No valid carton data found in any packing sheet.", vbExclamation, "No Data"
        GoTo Cleanup
    End If

    '--- Confirm --------------------------------------------------------------
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Ready to generate " & totalCartons & " stickers from " & packCount & " sheet(s):" & _
                 vbNewLine & vbNewLine & sheetList & vbNewLine & _
                 "This will CLEAR the '" & STICKER_SHEET & "' sheet. Continue?", _
                 vbQuestion + vbYesNo, "Generate Stickers")
    If ans <> vbYes Then GoTo Cleanup

    '--- Write stickers -------------------------------------------------------
    wsSticker.Cells.ClearContents
    
    For i = 1 To totalCartons
        Dim baseRow As Long
        baseRow = STICKER_FIRST_ROW + (i - 1) * (STICKER_BLOCK_ROWS + STICKER_GAP_ROWS)

        With allCartons(i)
            ' Labels
            wsSticker.Cells(baseRow + ROW_OC,      2).Value = "OC NUMBER"
            wsSticker.Cells(baseRow + ROW_STYLE,   2).Value = "STYLE NAME"
            wsSticker.Cells(baseRow + ROW_PRODUCT, 2).Value = "PRODUCT ID"
            wsSticker.Cells(baseRow + ROW_COLOR,   2).Value = "COLOR"
            wsSticker.Cells(baseRow + ROW_SIZE,    2).Value = "SIZE"
            wsSticker.Cells(baseRow + ROW_QTY,     2).Value = "QUANTITY"
            wsSticker.Cells(baseRow + ROW_CTN,     2).Value = "CARTON NUMBER"
            ' Values
            wsSticker.Cells(baseRow + ROW_OC,      STICKER_VAL_COL).Value = .OcNumber
            wsSticker.Cells(baseRow + ROW_STYLE,   STICKER_VAL_COL).Value = .StyleName
            wsSticker.Cells(baseRow + ROW_PRODUCT, STICKER_VAL_COL).Value = .ProductID
            wsSticker.Cells(baseRow + ROW_COLOR,   STICKER_VAL_COL).Value = .Color
            wsSticker.Cells(baseRow + ROW_SIZE,    STICKER_VAL_COL).Value = .SizeText
            wsSticker.Cells(baseRow + ROW_QTY,     STICKER_VAL_COL).Value = .TotalQty
            wsSticker.Cells(baseRow + ROW_CTN,     STICKER_VAL_COL).Value = .CartonNum & " OF " & .TotalCtns
        End With
    Next i

    '--- Formatting & page setup ----------------------------------------------
    Call ApplyStickerFormatting(wsSticker, totalCartons)

    '--- Print area -----------------------------------------------------------
    Dim lastPR As Long
    lastPR = STICKER_FIRST_ROW + totalCartons * (STICKER_BLOCK_ROWS + STICKER_GAP_ROWS) - STICKER_GAP_ROWS - 1
    wsSticker.PageSetup.PrintArea = "$A$1:$C$" & lastPR

    Application.ScreenUpdating = True
    wsSticker.Activate
    wsSticker.Range("B2").Select

    MsgBox "SUCCESS! " & totalCartons & " stickers generated across " & packCount & " colour/style sheet(s)." & _
           vbNewLine & "The '" & STICKER_SHEET & "' sheet is ready to print.", _
           vbInformation, "Stickers Generated"

Cleanup:
    Application.Calculation    = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Exit Sub

ErrHandler:
    Application.Calculation    = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    MsgBox "ERROR " & Err.Number & ": " & Err.Description & vbNewLine & _
           "In: " & Err.Source, vbCritical, "Runtime Error"
End Sub

'==============================================================================
' READ ALL CARTON ROWS FROM ONE PACKING SHEET
' Returns count; fills cartonData()
'==============================================================================
Private Function ReadCartonSheet(wsPL As Worksheet, ByRef cartonData() As CartonInfo) As Long

    On Error GoTo ErrOut

    Dim count As Long
    count = 0
    ReDim cartonData(1 To 500)

    '--- Detect header row (row with "CTN No" or "CTN NO" in col A) -----------
    Dim headerRow As Long
    headerRow = FindHeaderRow(wsPL)
    If headerRow = 0 Then
        ReadCartonSheet = 0
        Exit Function
    End If

    Dim dataStartRow As Long
    dataStartRow = headerRow + 1

    '--- Detect size columns from header row ----------------------------------
    Dim sizeMap(1 To 12) As String   ' size label per column index (relative to E=5)
    Dim qtyCol   As Long             ' column of "NO OF PCS /CTN"
    Dim totalCtnCol As Long          ' column of "Total CTN"
    Dim colorCol As Long             ' column of COLOR (usually C=3)
    colorCol = 3

    Call DetectColumns(wsPL, headerRow, sizeMap, qtyCol, totalCtnCol)

    If qtyCol = 0 Then
        ReadCartonSheet = 0
        Exit Function
    End If

    '--- Read header fields (OC, Style, Product) ------------------------------
    Dim ocNumber  As String
    Dim styleName As String
    Dim productID As String

    ocNumber  = FindOcNumber(wsPL)
    styleName = SafeStr(wsPL.Range("F10").Value)
    productID = SafeStr(wsPL.Range("F11").Value)

    If ocNumber = "" Then ocNumber = "[OC NOT FOUND]"

    '--- Read total cartons from F14 ------------------------------------------
    Dim totalCtnsHeader As Long
    totalCtnsHeader = SafeLong(wsPL.Range("F14").Value)

    '--- Scan data rows -------------------------------------------------------
    Dim lastRow As Long
    lastRow = wsPL.Cells(wsPL.Rows.Count, colorCol).End(xlUp).Row

    Dim r         As Long
    Dim seqNum    As Long
    seqNum = 0

    For r = dataStartRow To lastRow

        Dim colorVal As Variant
        colorVal = wsPL.Cells(r, colorCol).Value

        ' Skip rows with no colour or error in colour
        If IsError(colorVal) Or IsEmpty(colorVal) Then GoTo NextRow
        Dim colorStr As String
        colorStr = Trim(CStr(colorVal))
        If colorStr = "" Or colorStr = "TOTAL" Or colorStr = "Total" Then GoTo NextRow

        ' Skip summary/total rows by checking if qty col has a non-numeric
        Dim qtyVal As Variant
        qtyVal = wsPL.Cells(r, qtyCol).Value
        If IsError(qtyVal) Or IsEmpty(qtyVal) Then GoTo NextRow
        If Not IsNumeric(qtyVal) Then GoTo NextRow
        If CLng(qtyVal) <= 0 Then GoTo NextRow

        ' Determine carton number
        Dim ctnNum As Long
        Dim ctnCell As Variant
        ctnCell = wsPL.Cells(r, 1).Value   ' Column A
        If IsError(ctnCell) Or IsEmpty(ctnCell) Or Not IsNumeric(ctnCell) Then
            seqNum = seqNum + 1
            ctnNum = seqNum
        Else
            ctnNum = CLng(ctnCell)
            If ctnNum <= 0 Then
                seqNum = seqNum + 1
                ctnNum = seqNum
            Else
                seqNum = ctnNum
            End If
        End If

        ' Build size text from detected size columns
        Dim sizeText As String
        sizeText = BuildSizeTextDynamic(wsPL, r, sizeMap, qtyCol)

        count = count + 1
        With cartonData(count)
            .CartonNum = ctnNum
            .Color     = colorStr
            .SizeText  = sizeText
            .TotalQty  = CLng(qtyVal)
            .OcNumber  = ocNumber
            .StyleName = styleName
            .ProductID = productID
        End With

NextRow:
    Next r

    ' Fill TotalCtns: use header value if valid, else use max carton number found
    Dim maxCtn As Long
    maxCtn = 0
    Dim j As Long
    For j = 1 To count
        If cartonData(j).CartonNum > maxCtn Then maxCtn = cartonData(j).CartonNum
    Next j

    Dim finalTotal As Long
    If totalCtnsHeader > 0 Then
        finalTotal = totalCtnsHeader
    Else
        finalTotal = count  ' fallback: count of rows = total cartons
    End If

    For j = 1 To count
        cartonData(j).TotalCtns = finalTotal
    Next j

    ReadCartonSheet = count
    Exit Function

ErrOut:
    ReadCartonSheet = 0
End Function

'==============================================================================
' FIND THE HEADER ROW (row containing "CTN No" in col A or nearby)
'==============================================================================
Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long
    For r = 15 To 25
        Dim v As Variant
        v = ws.Cells(r, 1).Value
        If Not IsError(v) Then
            If InStr(1, CStr(v), "CTN", vbTextCompare) > 0 Then
                FindHeaderRow = r
                Exit Function
            End If
        End If
    Next r
    FindHeaderRow = 0
End Function

'==============================================================================
' DETECT SIZE COLUMNS AND QTY COLUMN FROM HEADER ROW
'==============================================================================
Private Sub DetectColumns(ws As Worksheet, headerRow As Long, _
                           ByRef sizeMap() As String, _
                           ByRef qtyCol As Long, ByRef totalCtnCol As Long)
    Dim c As Long
    qtyCol = 0
    totalCtnCol = 0

    For c = 1 To 20
        Dim v As Variant
        v = ws.Cells(headerRow, c).Value
        If Not IsError(v) And Not IsEmpty(v) Then
            Dim hdr As String
            hdr = Trim(UCase(CStr(v)))
            ' Size columns: S M L XL XXL 2XL 3XL 4XL 5XL 6XL 7XL 8XL
            If hdr = "S" Or hdr = "M" Or hdr = "L" Or hdr = "XL" Or hdr = "XXL" Or _
               hdr = "2XL" Or hdr = "3XL" Or hdr = "4XL" Or hdr = "5XL" Or _
               hdr = "6XL" Or hdr = "7XL" Or hdr = "8XL" Then
                If c <= 14 Then sizeMap(c) = Trim(CStr(ws.Cells(headerRow, c).Value))
            End If
            If InStr(hdr, "NO OF PCS") > 0 Or InStr(hdr, "PCS /CTN") > 0 Or hdr = "PCS" Then
                qtyCol = c
            End If
            If InStr(hdr, "TOTAL CTN") > 0 Then
                totalCtnCol = c
            End If
        End If
    Next c
End Sub

'==============================================================================
' BUILD SIZE TEXT DYNAMICALLY from detected columns
'==============================================================================
Private Function BuildSizeTextDynamic(ws As Worksheet, dataRow As Long, _
                                       sizeMap() As String, qtyCol As Long) As String
    Dim result As String
    result = ""
    Dim c As Long

    For c = 1 To 14
        If sizeMap(c) <> "" Then
            Dim v As Variant
            v = ws.Cells(dataRow, c).Value
            If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
                If CLng(v) > 0 Then
                    If result <> "" Then result = result & " "
                    result = result & sizeMap(c) & "=" & CLng(v)
                End If
            End If
        End If
    Next c

    BuildSizeTextDynamic = result
End Function

'==============================================================================
' FIND OC NUMBER — searches row 12 across multiple possible columns
'==============================================================================
Private Function FindOcNumber(ws As Worksheet) As String
    Dim c As Long
    For c = 8 To 16   ' scan columns H to P in row 12
        Dim v As Variant
        v = ws.Cells(12, c).Value
        If Not IsError(v) And Not IsEmpty(v) Then
            Dim s As String
            s = Trim(CStr(v))
            ' OC numbers match pattern like LC/ALD/25/12987 or LC/ALD/25/13929.1
            If Len(s) > 5 And (InStr(s, "/") > 0 Or InStr(s, "\") > 0) Then
                FindOcNumber = s
                Exit Function
            End If
        End If
    Next c
    ' Fallback: also check F12
    v = ws.Cells(12, 6).Value
    If Not IsError(v) And Not IsEmpty(v) Then
        s = Trim(CStr(v))
        If Len(s) > 5 And InStr(s, "/") > 0 Then
            FindOcNumber = s
            Exit Function
        End If
    End If
    FindOcNumber = ""
End Function

'==============================================================================
' DETERMINE IF A SHEET IS A PACKING DATA SHEET
'==============================================================================
Private Function IsPackingSheet(ws As Worksheet) As Boolean
    ' Skip sheets in the exclusion list
    If InStr(1, SKIP_SHEETS, "|" & ws.Name & "|", vbTextCompare) > 0 Then
        IsPackingSheet = False
        Exit Function
    End If
    ' Must have "ALDENO PACKING LIST" in A1
    Dim v As Variant
    On Error Resume Next
    v = ws.Range("A1").Value
    On Error GoTo 0
    If IsError(v) Or IsEmpty(v) Then
        IsPackingSheet = False
        Exit Function
    End If
    IsPackingSheet = (InStr(1, CStr(v), "ALDENO PACKING LIST", vbTextCompare) > 0)
End Function

'==============================================================================
' APPLY FORMATTING TO STICKER SHEET
'==============================================================================
Private Sub ApplyStickerFormatting(wsSticker As Worksheet, cartonCount As Long)
    Dim i As Long, baseRow As Long, r As Long

    For i = 1 To cartonCount
        baseRow = STICKER_FIRST_ROW + (i - 1) * (STICKER_BLOCK_ROWS + STICKER_GAP_ROWS)

        For r = baseRow To baseRow + STICKER_BLOCK_ROWS - 1
            If r = baseRow + ROW_CTN Then
                wsSticker.Rows(r).RowHeight = 47.25
            Else
                wsSticker.Rows(r).RowHeight = 46.5
            End If
            With wsSticker.Cells(r, 2)
                .Font.Name   = "Calibri"
                .Font.Size   = 36
                .Font.Bold   = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment   = xlCenter
            End With
            With wsSticker.Cells(r, 3)
                .Font.Name   = "Calibri"
                .Font.Size   = 36
                .Font.Bold   = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment   = xlCenter
            End With
        Next r

        wsSticker.Rows(baseRow + STICKER_BLOCK_ROWS).RowHeight = 15
    Next i

    wsSticker.Columns("A").ColumnWidth = 6.43
    wsSticker.Columns("B").ColumnWidth = 60
    wsSticker.Columns("C").ColumnWidth = 136.14

    With wsSticker.PageSetup
        .Orientation        = xlLandscape
        .PaperSize          = xlPaperA4
        .Zoom               = 69
        .LeftMargin         = Application.InchesToPoints(0.47)
        .RightMargin        = Application.InchesToPoints(0)
        .TopMargin          = Application.InchesToPoints(0.75)
        .BottomMargin       = Application.InchesToPoints(0)
        .CenterHorizontally = False
        .CenterVertically   = False
    End With
End Sub

'==============================================================================
' HELPERS
'==============================================================================
Private Function SafeStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
End Function

Private Function SafeLong(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        SafeLong = 0
    ElseIf IsNumeric(v) Then
        SafeLong = CLng(v)
    Else
        SafeLong = 0
    End If
End Function

Private Function SheetExists(wb As Workbook, shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function
