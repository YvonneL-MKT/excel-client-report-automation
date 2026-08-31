Attribute VB_Name = "ExcelClientReportAutomation"
Option Explicit

' ============================================================
' Excel Client Report Automation
' Portfolio Version
'
' This source code is an anonymized portfolio version of a
' production workflow used to prepare monthly client reports.
'
' Client names, client codes, mapping-sheet names, and local
' file paths have been replaced with representative values.
'
' Intended for portfolio review and implementation reference.
' It is not packaged as a standalone application.
' ============================================================


' ============================================================
' STEP 1: CLEAN AND STANDARDIZE SOURCE ORDER DATA
' ============================================================

Public Sub PrepareOrderData()

    Dim newSheet As Worksheet

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Remove the previous working sheet if it already exists.
    On Error Resume Next
    Worksheets("Sheet2").Delete
    On Error GoTo ErrHandler

    ' Create a working copy of the monthly source data.
    Worksheets("Sheet1").Copy After:=Worksheets(Worksheets.Count)
    Set newSheet = ActiveSheet
    newSheet.Name = "Sheet2"

    ' Remove source formatting that is not required in the client report.
    newSheet.Cells.Interior.Pattern = xlNone

    ' Remove fields that are not required in client-facing reports.
    DeleteColumnByHeader newSheet, "Subtotal"
    DeleteColumnByHeader newSheet, "Item Code"
    DeleteColumnByHeader newSheet, "Note"

    ' Standardize order status values.
    ReplaceValueByHeader newSheet, "Status", "Returned", "Canceled"

    ' Standardize date formatting.
    FormatDateColumnByHeader newSheet, "Completed Date"
    FormatDateColumnByHeader newSheet, "Cancelled Date"

    ' Remove non-order charge rows from the client report dataset.
    DeleteRowsByItemValues newSheet, "Item", Array("Rush Fee", "Extra Charge")

    ' Remove records without a valid status.
    DeleteBlankRowsByHeader newSheet, "Status"

    ' Standardize currency formatting.
    FormatCurrencyColumnByHeader newSheet, "U.Price"
    FormatCurrencyColumnByHeader newSheet, "Rush fee"
    FormatCurrencyColumnByHeader newSheet, "Extra Charge"

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Order data preparation completed.", vbInformation
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "An error occurred while preparing order data: " & Err.Description, vbExclamation

End Sub


' ============================================================
' STEP 2: GENERATE CLIENT-SPECIFIC WORKSHEETS
' ============================================================

Public Sub GenerateClientWorksheets()

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long
    Dim clientCode As String
    Dim cell As Range
    Dim targetRow As Long
    Dim columnsToWrap As Variant

    columnsToWrap = Array(2, 3, 4, 8, 15)

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set wsSource = wb.Sheets("Sheet2")
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "The working sheet 'Sheet2' does not exist.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    For Each cell In wsSource.Range("A2:A" & lastRow)

        clientCode = Trim(CStr(cell.Value))

        If clientCode <> "" Then

            ' Some production client codes are billed under the same
            ' parent company. Portfolio identifiers are used here.
            Select Case LCase(clientCode)

                Case "client_a", "client_a_alt"
                    Set wsTarget = GetOrCreateSheet(wb, "Client A", wsSource)

                Case "client_b", "client_b_alt"
                    Set wsTarget = GetOrCreateSheet(wb, "Client B", wsSource)

                Case "client_c", "client_c_alt1", "client_c_alt2", "client_c_alt3"
                    Set wsTarget = GetOrCreateSheet(wb, "Client C", wsSource)

                Case Else
                    Set wsTarget = GetOrCreateSheet(wb, clientCode, wsSource)

            End Select

            targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
            wsSource.Rows(cell.Row).Copy wsTarget.Rows(targetRow)

        End If

    Next cell

    FormatAllClientSheets wb, wsSource.Name, columnsToWrap

    Application.ScreenUpdating = True

    MsgBox "Client worksheets have been generated.", vbInformation

End Sub


' ============================================================
' STEP 3: MAP INTERNAL CLIENT CODES TO CLIENT-FACING NAMES
' ============================================================

Public Sub MapClientNames()

    Const MAPPING_SHEET_NAME As String = "Client Mapping"
    Const CLIENT_CODE_COLUMN As String = "B"
    Const COMPANY_NAME_COLUMN As String = "A"
    Const DATA_START_ROW As Long = 2

    Dim ws As Worksheet
    Dim wsMapping As Worksheet
    Dim clientDictionary As Object
    Dim clientCode As String
    Dim companyName As String
    Dim lastRow As Long
    Dim rowIndex As Long

    Set clientDictionary = CreateObject("Scripting.Dictionary")

    On Error GoTo MissingMappingSheet
    Set wsMapping = ActiveWorkbook.Sheets(MAPPING_SHEET_NAME)
    On Error GoTo 0

    lastRow = wsMapping.Cells(wsMapping.Rows.Count, CLIENT_CODE_COLUMN).End(xlUp).Row

    For rowIndex = DATA_START_ROW To lastRow

        clientCode = Trim(CStr(wsMapping.Cells(rowIndex, CLIENT_CODE_COLUMN).Value))
        companyName = Trim(CStr(wsMapping.Cells(rowIndex, COMPANY_NAME_COLUMN).Value))

        If clientCode <> "" And companyName <> "" Then
            If Not clientDictionary.Exists(clientCode) Then
                clientDictionary.Add clientCode, companyName
            End If
        End If

    Next rowIndex

    If clientDictionary.Count = 0 Then
        MsgBox "No client mapping records were found.", vbExclamation
        Exit Sub
    End If

    For Each ws In ActiveWorkbook.Worksheets

        If ws.Name <> wsMapping.Name Then

            clientCode = ws.Name

            If clientDictionary.Exists(clientCode) Then

                companyName = clientDictionary.Item(clientCode)

                If ws.Name <> companyName Then
                    On Error GoTo RenameError
                    ws.Name = Left(companyName, 31)
                    On Error GoTo 0
                End If

            End If

        End If

    Next ws

    Set clientDictionary = Nothing

    MsgBox "Client-facing worksheet names have been updated.", vbInformation
    Exit Sub

MissingMappingSheet:
    MsgBox "The mapping sheet '" & MAPPING_SHEET_NAME & "' could not be found.", vbCritical
    Set clientDictionary = Nothing
    Exit Sub

RenameError:
    MsgBox "Unable to rename worksheet '" & clientCode & _
           "' to '" & companyName & "': " & Err.Description, vbExclamation
    Resume Next

End Sub


' ============================================================
' STEP 4: APPLY PRINT-READY PAGE FORMATTING
' ============================================================

Public Sub FormatClientReports()

    Dim ws As Worksheet
    Dim col As Range
    Dim colWidth As Double

    For Each ws In ActiveWorkbook.Worksheets

        If ws.Name <> "Sheet1" _
           And ws.Name <> "Sheet2" _
           And ws.Name <> "Client Mapping" Then

            With ws.PageSetup
                .PaperSize = xlPaperA4
                .Orientation = xlLandscape
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False

                .TopMargin = Application.InchesToPoints(0.75)
                .BottomMargin = Application.InchesToPoints(0.75)
                .LeftMargin = Application.InchesToPoints(0.7)
                .RightMargin = Application.InchesToPoints(0.7)

                .CenterHorizontally = True
                .CenterVertically = False
            End With

            For Each col In ws.UsedRange.Columns
                colWidth = col.ColumnWidth

                If colWidth > 20 Then
                    col.WrapText = True
                Else
                    col.WrapText = False
                End If
            Next col

            ws.Cells.EntireRow.AutoFit

        End If

    Next ws

    MsgBox "Client report formatting completed.", vbInformation

End Sub


' ============================================================
' STEP 5: EXPORT CLIENT REPORTS TO PDF
' ============================================================

Public Sub ExportClientReportsToPDF()

    Dim ws As Worksheet
    Dim folderPath As String
    Dim pdfFileName As String
    Dim reportMonth As String
    Dim filePath As String

    reportMonth = Format(DateAdd("m", -1, Date), "yyyy-mm")

    ' Portfolio-safe output location.
    ' Production local paths are intentionally not included.
    folderPath = ActiveWorkbook.Path & Application.PathSeparator & _
                 reportMonth & "-Client-Reports" & Application.PathSeparator

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    For Each ws In ActiveWorkbook.Worksheets

        If ws.Name <> "Sheet1" _
           And ws.Name <> "Sheet2" _
           And ws.Name <> "Client Mapping" Then

            pdfFileName = ws.Name & "_" & reportMonth & "_Breakdown.pdf"
            filePath = folderPath & pdfFileName

            On Error Resume Next
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath

            If Err.Number <> 0 Then
                MsgBox "Unable to export worksheet '" & ws.Name & _
                       "' as PDF: " & Err.Description, vbExclamation
                Err.Clear
            End If

            On Error GoTo 0

        End If

    Next ws

    MsgBox "PDF export completed.", vbInformation

End Sub


' ============================================================
' SUPPORTING FUNCTIONS
' ============================================================

Private Function GetColumnByHeader(ws As Worksheet, headerName As String) As Long

    Dim headerCell As Range

    Set headerCell = ws.Rows(1).Find( _
        What:=headerName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If Not headerCell Is Nothing Then
        GetColumnByHeader = headerCell.Column
    Else
        GetColumnByHeader = 0
    End If

End Function


Private Sub DeleteColumnByHeader(ws As Worksheet, headerName As String)

    Dim columnNumber As Long

    columnNumber = GetColumnByHeader(ws, headerName)

    If columnNumber > 0 Then
        ws.Columns(columnNumber).Delete
    End If

End Sub


Private Sub ReplaceValueByHeader( _
    ws As Worksheet, _
    headerName As String, _
    oldValue As String, _
    newValue As String)

    Dim columnNumber As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    columnNumber = GetColumnByHeader(ws, headerName)

    If columnNumber = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    For rowIndex = 2 To lastRow

        If Trim(CStr(ws.Cells(rowIndex, columnNumber).Value)) = oldValue Then
            ws.Cells(rowIndex, columnNumber).Value = newValue
        End If

    Next rowIndex

End Sub


Private Sub FormatDateColumnByHeader(ws As Worksheet, headerName As String)

    Dim columnNumber As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    columnNumber = GetColumnByHeader(ws, headerName)

    If columnNumber = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    For rowIndex = 2 To lastRow

        If IsDate(ws.Cells(rowIndex, columnNumber).Value) Then
            ws.Cells(rowIndex, columnNumber).Value = _
                CDate(ws.Cells(rowIndex, columnNumber).Value)
        End If

    Next rowIndex

    ws.Columns(columnNumber).NumberFormat = "mm/dd/yyyy"

End Sub


Private Sub FormatCurrencyColumnByHeader(ws As Worksheet, headerName As String)

    Dim columnNumber As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    columnNumber = GetColumnByHeader(ws, headerName)

    If columnNumber = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    For rowIndex = 2 To lastRow

        If Trim(CStr(ws.Cells(rowIndex, columnNumber).Value)) <> "" Then

            If IsNumeric(ws.Cells(rowIndex, columnNumber).Value) Then
                ws.Cells(rowIndex, columnNumber).Value = _
                    CDbl(ws.Cells(rowIndex, columnNumber).Value)
            End If

        End If

    Next rowIndex

    ws.Columns(columnNumber).NumberFormat = "$#,##0.00"

End Sub


Private Sub DeleteRowsByItemValues( _
    ws As Worksheet, _
    headerName As String, _
    valuesToDelete As Variant)

    Dim columnNumber As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim valueToDelete As Variant

    columnNumber = GetColumnByHeader(ws, headerName)

    If columnNumber = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    For rowIndex = lastRow To 2 Step -1

        For Each valueToDelete In valuesToDelete

            If Trim(CStr(ws.Cells(rowIndex, columnNumber).Value)) = _
               CStr(valueToDelete) Then

                ws.Rows(rowIndex).Delete
                Exit For

            End If

        Next valueToDelete

    Next rowIndex

End Sub


Private Sub DeleteBlankRowsByHeader(ws As Worksheet, headerName As String)

    Dim columnNumber As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    columnNumber = GetColumnByHeader(ws, headerName)

    If columnNumber = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

    For rowIndex = lastRow To 2 Step -1

        If Trim(CStr(ws.Cells(rowIndex, columnNumber).Value)) = "" Then
            ws.Rows(rowIndex).Delete
        End If

    Next rowIndex

End Sub


Private Function GetOrCreateSheet( _
    wb As Workbook, _
    sheetName As String, _
    wsSource As Worksheet) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = Left(sheetName, 31)

        wsSource.Rows(1).Copy ws.Rows(1)

    End If

    Set GetOrCreateSheet = ws

End Function


Private Sub FormatAllClientSheets( _
    wb As Workbook, _
    sourceSheetName As String, _
    columnsToWrap As Variant)

    Dim ws As Worksheet
    Dim index As Long

    For Each ws In wb.Worksheets

        If ws.Name <> "Sheet1" _
           And ws.Name <> sourceSheetName _
           And ws.Name <> "Client Mapping" Then

            With ws.Range("A1").CurrentRegion.Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With

            ws.Columns.AutoFit

            For index = LBound(columnsToWrap) To UBound(columnsToWrap)

                If ws.Columns(columnsToWrap(index)).ColumnWidth > 20 Then
                    ws.Columns(columnsToWrap(index)).WrapText = True
                    ws.Columns(columnsToWrap(index)).ColumnWidth = 20
                End If

            Next index

        End If

    Next ws

End Sub
