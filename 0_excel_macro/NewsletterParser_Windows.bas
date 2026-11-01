Attribute VB_Name = "NewsletterParser"
Option Explicit

Public Const FINAL_SHEET_NAME As String = "최종"
Public Const BUSINESS_DB_SHEET As String = "DB-사업안내형"
Public Const ACADEMY_DB_SHEET As String = "DB-아카데미"
Public Const BANNER_DB_SHEET As String = "DB-배너형"
Public Const ALIMTALK_DB_SHEET As String = "DB-알림톡"
Public Const CALENDAR_SHEET As String = "뉴스레터 · 알림톡 게재 일정"

Private Const DB_HEADER_ROW As Long = 3
Private Const DB_FIRST_ROW As Long = 4
Private Const BUSINESS_CAPACITY As Long = 8
Private Const ACADEMY_CAPACITY As Long = 4
Private Const BANNER_CAPACITY As Long = 2
Private Const ALIMTALK_CAPACITY As Long = 2
Private Const FINAL_BANNER_PLACEHOLDER As String = "해외정보팀 담당 영역"
Private Const FORM_WARNING_SUFFIX As String = "은 보안상 사용이 금지된 설문조사 플랫폼입니다. 지원요청자에게 KOTRA 내부 설문조사 플랫폼 등으로 URL 수정을 요청하세요."

Private Const BUSINESS_VISIBLE_COLS As Long = 9
Private Const BUSINESS_SEQ_COL As Long = 10
Private Const BUSINESS_START_COL As Long = 11
Private Const BUSINESS_END_COL As Long = 12
Private Const BUSINESS_DEADLINE_COL As Long = 13
Private Const BUSINESS_PUB_START_COL As Long = 14
Private Const BUSINESS_PUB_END_COL As Long = 15

Private Const ACADEMY_VISIBLE_COLS As Long = 5
Private Const ACADEMY_SEQ_COL As Long = 6
Private Const ACADEMY_START_COL As Long = 7
Private Const ACADEMY_END_COL As Long = 8
Private Const ACADEMY_DEADLINE_COL As Long = 9

Private Const BANNER_VISIBLE_COLS As Long = 8
Private Const BANNER_SEQ_COL As Long = 9
Private Const BANNER_PUB_START_COL As Long = 10
Private Const BANNER_PUB_END_COL As Long = 11

Private Const ALIMTALK_VISIBLE_COLS As Long = 9
Private Const ALIMTALK_SEQ_COL As Long = 10
Private Const ALIMTALK_START_COL As Long = 11
Private Const ALIMTALK_END_COL As Long = 12
Private Const ALIMTALK_DEADLINE_COL As Long = 13
Private Const ALIMTALK_PUB_START_COL As Long = 14
Private Const ALIMTALK_PUB_END_COL As Long = 15

Private Const CALENDAR_FIRST_BLOCK_ROW As Long = 5
Private Const CALENDAR_ALIMTALK_ROWS As Long = 10
Private Const CALENDAR_BUSINESS_ROWS As Long = 20
Private Const CALENDAR_BANNER_ROWS As Long = 10
Private Const CALENDAR_BLOCK_ROWS As Long = 43
Private Const CALENDAR_MARKER As String = "NEWSLETTER_CALENDAR_V3"


Public Sub Auto_Open()
    On Error GoTo OpenFailed
    ConfigureWorkbook
    RefreshAllCore
    If TypeName(ActiveSheet) = "Worksheet" And TypeName(Selection) = "Range" Then
        UpdatePasteShortcut ActiveSheet, Selection
    End If
    Exit Sub

OpenFailed:
    MsgBox "뉴스레터 DB를 새로고침하지 못했습니다." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub


Public Sub Auto_Close()
    On Error Resume Next
    RemovePasteShortcut
End Sub


Private Sub ConfigureWorkbook()
    With ThisWorkbook.Worksheets(BUSINESS_DB_SHEET)
        .Columns("J:O").Hidden = True
        .Rows(1).RowHeight = 28
        .Rows(2).RowHeight = 42
        .Rows(3).RowHeight = 32
    End With
    With ThisWorkbook.Worksheets(ACADEMY_DB_SHEET)
        .Columns("F:I").Hidden = True
        .Rows(1).RowHeight = 28
        .Rows(2).RowHeight = 42
        .Rows(3).RowHeight = 32
    End With
    With ThisWorkbook.Worksheets(BANNER_DB_SHEET)
        .Columns("I:K").Hidden = True
        .Rows(1).RowHeight = 28
        .Rows(2).RowHeight = 42
        .Rows(3).RowHeight = 32
    End With
    With ThisWorkbook.Worksheets(ALIMTALK_DB_SHEET)
        .Columns("J:O").Hidden = True
        .Rows(1).RowHeight = 28
        .Rows(2).RowHeight = 42
        .Rows(3).RowHeight = 32
        .Tab.Color = RGB(255, 192, 0)
    End With
End Sub


Public Sub UpdatePasteShortcut(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo RemoveShortcut
    If Not (ws.Parent Is ThisWorkbook) Then GoTo RemoveShortcut
    If target Is Nothing Then GoTo RemoveShortcut
    If target.Cells.CountLarge <> 1 Then GoTo RemoveShortcut
    If target.Row < DB_FIRST_ROW Then GoTo RemoveShortcut

    If ws.Name = BANNER_DB_SHEET And target.Column = 4 Then
        Application.OnKey "^v", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PasteBannerImageAtSelection"
        Application.OnKey "+{INSERT}", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PasteBannerImageAtSelection"
        Exit Sub
    End If

    If IsDbSheet(ws.Name) And target.Column = 2 Then
        Application.OnKey "^v", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PasteRawTextAtSelection"
        Application.OnKey "+{INSERT}", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PasteRawTextAtSelection"
        Exit Sub
    End If

RemoveShortcut:
    RemovePasteShortcut
End Sub


Public Sub RemovePasteShortcut()
    Application.OnKey "^v"
    Application.OnKey "+{INSERT}"
End Sub


Private Function IsDbSheet(ByVal sheetName As String) As Boolean
    IsDbSheet = _
        sheetName = BUSINESS_DB_SHEET Or _
        sheetName = ACADEMY_DB_SHEET Or _
        sheetName = BANNER_DB_SHEET Or _
        sheetName = ALIMTALK_DB_SHEET
End Function


Public Sub PasteRawTextAtSelection()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim rawText As String
    Dim oldEvents As Boolean

    On Error GoTo Failed
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveSheet
    Set targetCell = Selection.Cells(1, 1)
    If Not IsDbSheet(ws.Name) Then Exit Sub
    If targetCell.Column <> 2 Or targetCell.Row < DB_FIRST_ROW Then Exit Sub

    rawText = ReadClipboardText()
    If Len(Trim$(rawText)) = 0 Then Exit Sub

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    RegisterRawTextAtRow ws, rawText, targetCell.Row
    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    Application.EnableEvents = oldEvents
    MsgBox "등록하지 못했습니다." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub


Private Function ReadClipboardText() As String
    Dim clipboard As Object
    On Error GoTo ClipboardFailed
    Set clipboard = CreateObject("Forms.DataObject.1")
    clipboard.GetFromClipboard
    ReadClipboardText = clipboard.GetText(1)
    Exit Function

ClipboardFailed:
    Err.Raise vbObjectError + 1090, , "클립보드의 텍스트를 읽지 못했습니다. DB 시트의 B열 셀을 더블클릭한 뒤 붙여넣어 주세요."
End Function


Public Sub RegisterRawTextAtRow(ByVal ws As Worksheet, ByVal rawText As String, ByVal destinationRow As Long)
    If destinationRow < DB_FIRST_ROW Then
        Err.Raise vbObjectError + 1000, , "표 머리글 아래의 빈 행에 붙여넣으세요."
    End If

    Select Case ws.Name
        Case BUSINESS_DB_SHEET
            AppendBusinessRaw rawText
        Case ACADEMY_DB_SHEET
            AppendAcademyRaw rawText
        Case BANNER_DB_SHEET
            AppendBannerRaw rawText
        Case ALIMTALK_DB_SHEET
            AppendAlimtalkRaw rawText
        Case Else
            Err.Raise vbObjectError + 1001, , "DB-사업안내형, DB-아카데미, DB-배너형 또는 DB-알림톡 시트에 붙여넣으세요."
    End Select

    RefreshAllCore
End Sub


Public Function LooksLikeRawText(ByVal value As String) As Boolean
    LooksLikeRawText = _
        (InStr(1, value, "사업명", vbTextCompare) > 0 And InStr(1, value, "http", vbTextCompare) > 0) Or _
        IsAcademyRawText(value)
End Function


Private Function IsAcademyRawText(ByVal value As String) As Boolean
    IsAcademyRawText = _
        InStr(1, value, "교육기간", vbTextCompare) > 0 And _
        InStr(1, value, "신청마감", vbTextCompare) > 0 And _
        InStr(1, value, "http", vbTextCompare) > 0
End Function


Public Sub 선택행_원문등록()
    Dim ws As Worksheet
    Dim sourceCell As Range
    Dim rawText As String

    If TypeName(ActiveSheet) <> "Worksheet" Or TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    Set sourceCell = Selection.Cells(1, 1)
    If Not IsDbSheet(ws.Name) Or sourceCell.Column <> 2 Then
        MsgBox "DB 시트에서 등록할 행의 B열 셀을 선택하세요.", vbExclamation
        Exit Sub
    End If
    rawText = CStr(sourceCell.Value2)
    If Len(Trim$(rawText)) = 0 Then
        MsgBox "선택한 셀에 원문을 먼저 붙여넣으세요.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Failed
    Application.EnableEvents = False
    sourceCell.ClearContents
    RegisterRawTextAtRow ws, rawText, sourceCell.Row
    Application.EnableEvents = True
    Exit Sub

Failed:
    Application.EnableEvents = True
    MsgBox "등록하지 못했습니다." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub


Public Sub 전체_새로고침()
    On Error GoTo Failed
    RefreshAllCore
    Exit Sub

Failed:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "새로고침하지 못했습니다." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub


Private Sub RefreshAllCore()
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo RefreshFailed

    ConfigureWorkbook
    EnsureCalendarStructure
    RebuildBusinessDb
    RebuildAcademyDb
    RebuildBannerDb
    RebuildAlimtalkDb
    RefreshFinalSheet
    RefreshCalendar

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

RefreshFailed:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function NextNewsletterDate(Optional ByVal baseDate As Date = 0) As Date
    Dim candidate As Date
    If baseDate = 0 Then baseDate = Date
    candidate = baseDate
    Do
        If (Weekday(candidate, vbMonday) = 2 Or Weekday(candidate, vbMonday) = 5) And HolidayName(candidate) = "" Then
            NextNewsletterDate = candidate
            Exit Function
        End If
        candidate = candidate + 1
    Loop
End Function


Public Function NextAlimtalkDate(Optional ByVal baseDate As Date = 0) As Date
    Dim candidate As Date
    If baseDate = 0 Then baseDate = Date
    candidate = baseDate
    Do
        If (Weekday(candidate, vbMonday) = 1 Or Weekday(candidate, vbMonday) = 3) And HolidayName(candidate) = "" Then
            NextAlimtalkDate = candidate
            Exit Function
        End If
        candidate = candidate + 1
    Loop
End Function


Private Function NextSequence(ByVal ws As Worksheet, ByVal sequenceColumn As Long) As Long
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim value As Long
    lastRow = LastDataRow(ws, sequenceColumn)
    For rowNumber = DB_FIRST_ROW To lastRow
        If IsNumeric(ws.Cells(rowNumber, sequenceColumn).Value2) Then
            value = CLng(ws.Cells(rowNumber, sequenceColumn).Value2)
            If value >= NextSequence Then NextSequence = value + 1
        End If
    Next rowNumber
    If NextSequence = 0 Then NextSequence = 1
End Function


Private Function LastDataRow(ByVal ws As Worksheet, ByVal keyColumn As Long) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).Row
    If LastDataRow < DB_FIRST_ROW Then LastDataRow = DB_FIRST_ROW - 1
End Function


Private Function NextAppendRow(ByVal ws As Worksheet, ByVal sequenceColumn As Long, ByVal minimumRow As Long) As Long
    Dim lastRow As Long
    lastRow = LastDataRow(ws, sequenceColumn)
    If lastRow < minimumRow Then
        NextAppendRow = minimumRow
    Else
        NextAppendRow = lastRow + 1
    End If
End Function


Private Sub AppendBusinessRaw(ByVal rawText As String)
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim title As String
    Dim businessText As String
    Dim applicationText As String
    Dim publicationText As String
    Dim targetURL As String
    Dim businessStart As Date
    Dim businessEnd As Date
    Dim applicationStart As Date
    Dim applicationEnd As Date
    Dim publicationStart As Date
    Dim publicationEnd As Date
    Dim baseYear As Long

    title = GetFieldValue(rawText, "사업명")
    businessText = GetFieldValue(rawText, "사업기간")
    applicationText = GetFieldValue(rawText, "사업 신청기간")
    publicationText = GetFieldValue(rawText, "희망 게재기간")
    targetURL = ExtractURL(GetFieldValue(rawText, "연결 링크 (url)"))

    If title = "" Then Err.Raise vbObjectError + 1101, , "'사업명:' 항목을 찾지 못했습니다."
    If businessText = "" Then Err.Raise vbObjectError + 1102, , "'사업기간:' 항목을 찾지 못했습니다."
    If applicationText = "" Then Err.Raise vbObjectError + 1103, , "'사업 신청기간:' 항목을 찾지 못했습니다."
    If publicationText = "" Then Err.Raise vbObjectError + 1104, , "'희망 게재기간:' 항목을 찾지 못했습니다."
    If targetURL = "" Then Err.Raise vbObjectError + 1105, , "'연결 링크 (url):' 항목에서 URL을 찾지 못했습니다."

    ParseDateRangeWithDefaultYear publicationText, Year(Date), publicationStart, publicationEnd
    baseYear = Year(publicationStart)
    ParseDateRangeWithDefaultYear businessText, baseYear, businessStart, businessEnd
    ParseDateRangeWithDefaultYear applicationText, baseYear, applicationStart, applicationEnd

    Set ws = ThisWorkbook.Worksheets(BUSINESS_DB_SHEET)
    rowNumber = NextAppendRow(ws, BUSINESS_SEQ_COL, DB_FIRST_ROW + BUSINESS_CAPACITY)
    CopyRowFormat ws, DB_FIRST_ROW, rowNumber, BUSINESS_VISIBLE_COLS
    ws.Cells(rowNumber, 1).Value = CleanDisplayText(title)
    ws.Cells(rowNumber, 2).Value = FormatDbPeriod(businessStart, businessEnd)
    ws.Cells(rowNumber, 3).Value = FormatDbDate(applicationEnd)
    ws.Cells(rowNumber, 4).Value = FormatDbPeriod(publicationStart, publicationEnd)
    ws.Cells(rowNumber, 5).Value = targetURL
    ws.Cells(rowNumber, BUSINESS_SEQ_COL).Value = NextSequence(ws, BUSINESS_SEQ_COL)
    ws.Cells(rowNumber, BUSINESS_START_COL).Value = businessStart
    ws.Cells(rowNumber, BUSINESS_END_COL).Value = businessEnd
    ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value = applicationEnd
    ws.Cells(rowNumber, BUSINESS_PUB_START_COL).Value = publicationStart
    ws.Cells(rowNumber, BUSINESS_PUB_END_COL).Value = publicationEnd
    ApplyFormWarning ws, rowNumber, 5, targetURL, BUSINESS_VISIBLE_COLS
End Sub


Private Sub AppendBannerRaw(ByVal rawText As String)
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim title As String
    Dim publicationText As String
    Dim targetURL As String
    Dim publicationStart As Date
    Dim publicationEnd As Date

    title = GetFieldValue(rawText, "사업명")
    publicationText = GetFieldValue(rawText, "희망 게재기간")
    targetURL = ExtractURL(GetFieldValue(rawText, "연결 링크 (url)"))

    If title = "" Then Err.Raise vbObjectError + 1111, , "'사업명:' 항목을 찾지 못했습니다."
    If publicationText = "" Then Err.Raise vbObjectError + 1112, , "'희망 게재기간:' 항목을 찾지 못했습니다."
    If targetURL = "" Then Err.Raise vbObjectError + 1113, , "'연결 링크 (url):' 항목에서 URL을 찾지 못했습니다."

    ParseDateRangeWithDefaultYear publicationText, Year(Date), publicationStart, publicationEnd
    Set ws = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    rowNumber = NextAppendRow(ws, BANNER_SEQ_COL, DB_FIRST_ROW + BANNER_CAPACITY)
    CopyRowFormat ws, DB_FIRST_ROW, rowNumber, BANNER_VISIBLE_COLS
    ws.Cells(rowNumber, 1).Value = CleanDisplayText(title)
    ws.Cells(rowNumber, 2).Value = FormatDbPeriod(publicationStart, publicationEnd)
    ws.Cells(rowNumber, 3).Value = targetURL
    ws.Cells(rowNumber, BANNER_SEQ_COL).Value = NextSequence(ws, BANNER_SEQ_COL)
    ws.Cells(rowNumber, BANNER_PUB_START_COL).Value = publicationStart
    ws.Cells(rowNumber, BANNER_PUB_END_COL).Value = publicationEnd
    ApplyFormWarning ws, rowNumber, 3, targetURL, BANNER_VISIBLE_COLS
End Sub


Private Sub AppendAlimtalkRaw(ByVal rawText As String)
    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim title As String
    Dim businessText As String
    Dim applicationText As String
    Dim publicationText As String
    Dim targetURL As String
    Dim businessStart As Date
    Dim businessEnd As Date
    Dim applicationStart As Date
    Dim applicationEnd As Date
    Dim publicationStart As Date
    Dim publicationEnd As Date
    Dim baseYear As Long

    title = GetFieldValue(rawText, "사업명")
    businessText = GetFieldValue(rawText, "사업기간")
    applicationText = GetFieldValue(rawText, "사업 신청기간")
    If applicationText = "" Then applicationText = GetFieldValue(rawText, "신청기간")
    publicationText = GetFieldValue(rawText, "희망 발송일")
    If publicationText = "" Then publicationText = GetFieldValue(rawText, "희망 발송기간")
    If publicationText = "" Then publicationText = GetFieldValue(rawText, "발송 희망일")
    If publicationText = "" Then publicationText = GetFieldValue(rawText, "알림톡 발송일")
    If publicationText = "" Then publicationText = GetFieldValue(rawText, "알림톡 발송 희망일")
    If publicationText = "" Then publicationText = GetFieldValue(rawText, "희망 게재기간")
    targetURL = ExtractURL(GetFieldValue(rawText, "연결 링크 (url)"))

    If title = "" Then Err.Raise vbObjectError + 1131, , "'사업명:' 항목을 찾지 못했습니다."
    If businessText = "" Then Err.Raise vbObjectError + 1132, , "'사업기간:' 항목을 찾지 못했습니다."
    If applicationText = "" Then Err.Raise vbObjectError + 1133, , "'사업 신청기간:' 항목을 찾지 못했습니다."
    If publicationText = "" Then Err.Raise vbObjectError + 1134, , "'희망 발송일:' 또는 '희망 게재기간:' 항목을 찾지 못했습니다."
    If targetURL = "" Then Err.Raise vbObjectError + 1135, , "'연결 링크 (url):' 항목에서 URL을 찾지 못했습니다."

    ParseDateRangeWithDefaultYear publicationText, Year(Date), publicationStart, publicationEnd
    baseYear = Year(publicationStart)
    ParseDateRangeWithDefaultYear businessText, baseYear, businessStart, businessEnd
    ParseDateRangeWithDefaultYear applicationText, baseYear, applicationStart, applicationEnd

    Set ws = ThisWorkbook.Worksheets(ALIMTALK_DB_SHEET)
    rowNumber = NextAppendRow(ws, ALIMTALK_SEQ_COL, DB_FIRST_ROW + ALIMTALK_CAPACITY)
    CopyRowFormat ws, DB_FIRST_ROW, rowNumber, ALIMTALK_VISIBLE_COLS
    ws.Cells(rowNumber, 1).Value = CleanDisplayText(title)
    ws.Cells(rowNumber, 2).Value = FormatDbPeriod(businessStart, businessEnd)
    ws.Cells(rowNumber, 3).Value = FormatDbDate(applicationEnd)
    ws.Cells(rowNumber, 4).Value = FormatDbPeriod(publicationStart, publicationEnd)
    ws.Cells(rowNumber, 5).Value = targetURL
    ws.Cells(rowNumber, ALIMTALK_SEQ_COL).Value = NextSequence(ws, ALIMTALK_SEQ_COL)
    ws.Cells(rowNumber, ALIMTALK_START_COL).Value = businessStart
    ws.Cells(rowNumber, ALIMTALK_END_COL).Value = businessEnd
    ws.Cells(rowNumber, ALIMTALK_DEADLINE_COL).Value = applicationEnd
    ws.Cells(rowNumber, ALIMTALK_PUB_START_COL).Value = publicationStart
    ws.Cells(rowNumber, ALIMTALK_PUB_END_COL).Value = publicationEnd
    ApplyFormWarning ws, rowNumber, 5, targetURL, ALIMTALK_VISIBLE_COLS
End Sub


Private Sub AppendAcademyRaw(ByVal rawText As String)
    Dim ws As Worksheet
    Dim records As Collection
    Dim record As Variant
    Dim issueDates As Object
    Dim issueKey As String
    Dim rowNumber As Long
    Dim sequenceValue As Long
    Dim educationStart As Date
    Dim educationEnd As Date
    Dim deadlineStart As Date
    Dim deadlineEnd As Date
    Dim periodDisplay As String
    Dim deadlineDisplay As String
    Dim targetURL As String
    Dim item As Variant
    Dim rowIndex As Long
    Dim lastRow As Long

    If Not IsAcademyRawText(rawText) Then
        Err.Raise vbObjectError + 1120, , "[월.일], + 과정명, 교육기간, 신청마감, URL이 포함된 아카데미 원문을 붙여넣으세요."
    End If

    Set records = ParseAcademyGroupedRecords(rawText)
    Set issueDates = CreateObject("Scripting.Dictionary")
    For Each item In records
        issueKey = Format$(CDate(item(0)), "yyyy-mm-dd")
        If Not issueDates.Exists(issueKey) Then issueDates.Add issueKey, 0
        issueDates(issueKey) = CLng(issueDates(issueKey)) + 1
        If CLng(issueDates(issueKey)) > ACADEMY_CAPACITY Then
            Err.Raise vbObjectError + 1121, , issueKey & " 아카데미 항목이 4건을 초과합니다."
        End If
    Next item

    Set ws = ThisWorkbook.Worksheets(ACADEMY_DB_SHEET)
    lastRow = LastDataRow(ws, ACADEMY_SEQ_COL)
    For rowIndex = lastRow To DB_FIRST_ROW Step -1
        If IsDate(ws.Cells(rowIndex, 1).Value) Then
            issueKey = Format$(CDate(ws.Cells(rowIndex, 1).Value), "yyyy-mm-dd")
            If issueDates.Exists(issueKey) Then ws.Rows(rowIndex).Delete
        End If
    Next rowIndex

    sequenceValue = NextSequence(ws, ACADEMY_SEQ_COL)
    For Each record In records
        periodDisplay = FormatAcademyPeriod(CStr(record(2)), Year(CDate(record(0))), educationStart, educationEnd)
        ParseAcademyDeadline CStr(record(3)), educationStart, deadlineStart, deadlineEnd
        deadlineDisplay = FormatNewsletterDate(deadlineEnd, Year(CDate(record(0))))
        targetURL = ExtractURL(CStr(record(4)))
        If targetURL = "" Then Err.Raise vbObjectError + 1122, , "아카데미 URL을 찾지 못했습니다: " & CStr(record(1))

        rowNumber = NextAppendRow(ws, ACADEMY_SEQ_COL, DB_FIRST_ROW)
        CopyRowFormat ws, DB_FIRST_ROW, rowNumber, ACADEMY_VISIBLE_COLS
        ws.Cells(rowNumber, 1).Value = CDate(record(0))
        ws.Cells(rowNumber, 1).NumberFormat = "yyyy.mm.dd.(aaa)"
        ws.Cells(rowNumber, 2).Value = CleanDisplayText(CStr(record(1)))
        ws.Cells(rowNumber, 3).Value = periodDisplay
        ws.Cells(rowNumber, 4).Value = deadlineDisplay
        ws.Cells(rowNumber, 5).Value = targetURL
        ws.Cells(rowNumber, ACADEMY_SEQ_COL).Value = sequenceValue
        ws.Cells(rowNumber, ACADEMY_START_COL).Value = educationStart
        ws.Cells(rowNumber, ACADEMY_END_COL).Value = educationEnd
        ws.Cells(rowNumber, ACADEMY_DEADLINE_COL).Value = deadlineEnd
        ApplyFormWarning ws, rowNumber, 5, targetURL, ACADEMY_VISIBLE_COLS
        sequenceValue = sequenceValue + 1
    Next record
End Sub


Private Function ParseAcademyGroupedRecords(ByVal rawText As String) As Collection
    Dim records As New Collection
    Dim normalized As String
    Dim lines() As String
    Dim line As Variant
    Dim currentLine As String
    Dim issueDate As Date
    Dim hasIssueDate As Boolean
    Dim title As String
    Dim educationPeriod As String
    Dim deadlineText As String
    Dim targetURL As String

    normalized = Replace(rawText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, "：", ":")
    lines = Split(normalized, vbLf)

    For Each line In lines
        currentLine = Trim$(Replace(CStr(line), ChrW(160), " "))
        If IsAcademyIssueHeader(currentLine) Then
            AddAcademyGroupedRecord records, hasIssueDate, issueDate, title, educationPeriod, deadlineText, targetURL
            title = "": educationPeriod = "": deadlineText = "": targetURL = ""
            issueDate = ParseAcademyIssueDate(currentLine)
            hasIssueDate = True
        ElseIf Left$(currentLine, 1) = "+" Then
            AddAcademyGroupedRecord records, hasIssueDate, issueDate, title, educationPeriod, deadlineText, targetURL
            title = Trim$(Mid$(currentLine, 2))
            educationPeriod = "": deadlineText = "": targetURL = ""
        ElseIf title <> "" Then
            If AcademyFieldValue(currentLine, "교육기간") <> "" Then
                educationPeriod = AcademyFieldValue(currentLine, "교육기간")
            ElseIf AcademyFieldValue(currentLine, "신청마감") <> "" Then
                deadlineText = AcademyFieldValue(currentLine, "신청마감")
            ElseIf targetURL = "" And InStr(1, currentLine, "http", vbTextCompare) > 0 Then
                targetURL = ExtractURL(currentLine)
            End If
        End If
    Next line

    AddAcademyGroupedRecord records, hasIssueDate, issueDate, title, educationPeriod, deadlineText, targetURL
    If records.Count = 0 Then
        Err.Raise vbObjectError + 1123, , "[월.일] 아래에서 '+ 과정명'으로 시작하는 아카데미 항목을 찾지 못했습니다."
    End If
    Set ParseAcademyGroupedRecords = records
End Function


Private Sub AddAcademyGroupedRecord( _
    ByVal records As Collection, _
    ByVal hasIssueDate As Boolean, _
    ByVal issueDate As Date, _
    ByVal title As String, _
    ByVal educationPeriod As String, _
    ByVal deadlineText As String, _
    ByVal targetURL As String)

    If Trim$(title) = "" Then Exit Sub
    If Not hasIssueDate Then Err.Raise vbObjectError + 1124, , "아카데미 과정 앞의 [월.일] 발행일을 찾지 못했습니다."
    If Trim$(educationPeriod) = "" Then Err.Raise vbObjectError + 1125, , title & " 과정에 '교육기간:'이 없습니다."
    If Trim$(deadlineText) = "" Then Err.Raise vbObjectError + 1126, , title & " 과정에 '신청마감:'이 없습니다."
    If ExtractURL(targetURL) = "" Then Err.Raise vbObjectError + 1127, , title & " 과정의 URL을 찾지 못했습니다."
    records.Add Array(issueDate, title, educationPeriod, deadlineText, targetURL)
End Sub


Private Function IsAcademyIssueHeader(ByVal value As String) As Boolean
    Dim text As String
    text = Trim$(value)
    IsAcademyIssueHeader = Left$(text, 1) = "[" And Right$(text, 1) = "]" And ExtractNumbers(text).Count >= 2
End Function


Private Function ParseAcademyIssueDate(ByVal value As String) As Date
    Dim numbers As Collection
    Dim yearValue As Long
    Dim monthValue As Long
    Dim dayValue As Long
    Dim candidate As Date

    Set numbers = ExtractNumbers(value)
    If numbers.Count >= 3 And IsYearToken(CLng(numbers(1))) Then
        yearValue = ExpandYear(CLng(numbers(1)))
        monthValue = CLng(numbers(2))
        dayValue = CLng(numbers(3))
    Else
        yearValue = Year(Date)
        monthValue = CLng(numbers(1))
        dayValue = CLng(numbers(2))
    End If
    candidate = StrictDate(yearValue, monthValue, dayValue, value)
    If numbers.Count < 3 Or Not IsYearToken(CLng(numbers(1))) Then
        If candidate < Date - 183 Then candidate = DateSerial(yearValue + 1, monthValue, dayValue)
        If candidate > Date + 183 Then candidate = DateSerial(yearValue - 1, monthValue, dayValue)
    End If
    ParseAcademyIssueDate = candidate
End Function


Private Sub RebuildBusinessDb()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearLastRow As Long
    Dim sourceRow As Long
    Dim col As Long
    Dim count As Long
    Dim data() As Variant
    Dim nextSeq As Long
    Dim issueDate As Date
    Dim eligible() As Long
    Dim eligibleCount As Long
    Dim selected() As Boolean
    Dim remaining() As Long
    Dim remainingCount As Long
    Dim index As Long
    Dim outputRow As Long
    Dim selectedCount As Long

    Set ws = ThisWorkbook.Worksheets(BUSINESS_DB_SHEET)
    lastRow = MaxLastRow(ws, 1, BUSINESS_SEQ_COL)
    If lastRow < DB_FIRST_ROW Then lastRow = DB_FIRST_ROW - 1
    ReDim data(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1), 1 To BUSINESS_PUB_END_COL)
    nextSeq = NextSequence(ws, BUSINESS_SEQ_COL)

    For sourceRow = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(sourceRow, 1).Value2))) > 0 Then
            If Not IsNumeric(ws.Cells(sourceRow, BUSINESS_SEQ_COL).Value2) Then
                ws.Cells(sourceRow, BUSINESS_SEQ_COL).Value = nextSeq
                nextSeq = nextSeq + 1
            End If
            BackfillBusinessDates ws, sourceRow
            If Not BusinessRowExpired(ws, sourceRow, Date) Then
                count = count + 1
                For col = 1 To BUSINESS_PUB_END_COL
                    data(count, col) = ws.Cells(sourceRow, col).Value
                Next col
            End If
        End If
    Next sourceRow

    issueDate = NextNewsletterDate(Date)
    ReDim eligible(1 To Application.Max(1, count))
    ReDim selected(1 To Application.Max(1, count))
    For index = 1 To count
        If BusinessRecordEligible(data, index, issueDate) Then
            eligibleCount = eligibleCount + 1
            eligible(eligibleCount) = index
        End If
    Next index
    If eligibleCount > 1 Then SortIndexArray eligible, eligibleCount, data, BUSINESS_DEADLINE_COL, BUSINESS_SEQ_COL
    selectedCount = Application.Min(BUSINESS_CAPACITY, eligibleCount)
    For index = 1 To selectedCount
        selected(eligible(index)) = True
    Next index

    ReDim remaining(1 To Application.Max(1, count))
    For index = 1 To count
        If Not selected(index) Then
            remainingCount = remainingCount + 1
            remaining(remainingCount) = index
        End If
    Next index
    If remainingCount > 1 Then SortIndexArray remaining, remainingCount, data, BUSINESS_SEQ_COL, BUSINESS_SEQ_COL

    clearLastRow = Application.Max(lastRow, DB_FIRST_ROW + BUSINESS_CAPACITY + count + 5)
    ws.Range(ws.Cells(DB_FIRST_ROW, 1), ws.Cells(clearLastRow, BUSINESS_PUB_END_COL)).ClearContents
    outputRow = DB_FIRST_ROW
    For index = 1 To selectedCount
        WriteDataRow ws, outputRow, data, eligible(index), BUSINESS_PUB_END_COL, BUSINESS_VISIBLE_COLS
        ApplyFormWarning ws, outputRow, 5, CStr(data(eligible(index), 5)), BUSINESS_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    Do While outputRow < DB_FIRST_ROW + BUSINESS_CAPACITY
        CopyRowFormat ws, DB_FIRST_ROW, outputRow, BUSINESS_VISIBLE_COLS
        outputRow = outputRow + 1
    Loop
    ApplyCapacityBorder ws, DB_FIRST_ROW + BUSINESS_CAPACITY - 1, BUSINESS_VISIBLE_COLS
    For index = 1 To remainingCount
        WriteDataRow ws, outputRow, data, remaining(index), BUSINESS_PUB_END_COL, BUSINESS_VISIBLE_COLS
        ApplyFormWarning ws, outputRow, 5, CStr(data(remaining(index), 5)), BUSINESS_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    CopyRowFormat ws, DB_FIRST_ROW, outputRow, BUSINESS_VISIBLE_COLS
    ws.Range("A2:I2").MergeArea.Cells(1, 1).Value = _
        "다음 발행일: " & Format$(issueDate, "yyyy.mm.dd.(aaa)") & _
        " | 구분선 위 최대 8건이 최종 시트에 반영됩니다. 마감 임박 순, 동률은 등록 순입니다."
End Sub


Private Sub RebuildAcademyDb()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearLastRow As Long
    Dim sourceRow As Long
    Dim col As Long
    Dim count As Long
    Dim data() As Variant
    Dim order() As Long
    Dim index As Long
    Dim outputRow As Long
    Dim nextSeq As Long
    Dim selectedIssue As Date

    Set ws = ThisWorkbook.Worksheets(ACADEMY_DB_SHEET)
    lastRow = MaxLastRow(ws, 2, ACADEMY_SEQ_COL)
    If lastRow < DB_FIRST_ROW Then lastRow = DB_FIRST_ROW - 1
    ReDim data(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1), 1 To ACADEMY_DEADLINE_COL)
    nextSeq = NextSequence(ws, ACADEMY_SEQ_COL)

    For sourceRow = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(sourceRow, 2).Value2))) > 0 Then
            If Not IsNumeric(ws.Cells(sourceRow, ACADEMY_SEQ_COL).Value2) Then
                ws.Cells(sourceRow, ACADEMY_SEQ_COL).Value = nextSeq
                nextSeq = nextSeq + 1
            End If
            If IsDate(ws.Cells(sourceRow, 1).Value) And CDate(ws.Cells(sourceRow, 1).Value) >= Date Then
                count = count + 1
                For col = 1 To ACADEMY_DEADLINE_COL
                    data(count, col) = ws.Cells(sourceRow, col).Value
                Next col
            End If
        End If
    Next sourceRow

    ReDim order(1 To Application.Max(1, count))
    For index = 1 To count
        order(index) = index
    Next index
    If count > 1 Then SortIndexArray order, count, data, 1, ACADEMY_SEQ_COL

    clearLastRow = Application.Max(lastRow, DB_FIRST_ROW + count + ACADEMY_CAPACITY + 5)
    ws.Range(ws.Cells(DB_FIRST_ROW, 1), ws.Cells(clearLastRow, ACADEMY_DEADLINE_COL)).ClearContents
    outputRow = DB_FIRST_ROW
    For index = 1 To count
        WriteDataRow ws, outputRow, data, order(index), ACADEMY_DEADLINE_COL, ACADEMY_VISIBLE_COLS
        ws.Cells(outputRow, 1).NumberFormat = "yyyy.mm.dd.(aaa)"
        ApplyFormWarning ws, outputRow, 5, CStr(data(order(index), 5)), ACADEMY_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    CopyRowFormat ws, DB_FIRST_ROW, outputRow, ACADEMY_VISIBLE_COLS

    selectedIssue = SelectedAcademyIssueDate(ws, Date)
    If selectedIssue > 0 Then
        ws.Range("A2:E2").MergeArea.Cells(1, 1).Value = _
            "최종 반영 발행일: " & Format$(selectedIssue, "yyyy.mm.dd.(aaa)") & _
            " | 해당 발행일의 최대 4건이 최종 시트에 자동 반영됩니다."
    Else
        ws.Range("A2:E2").MergeArea.Cells(1, 1).Value = _
            "아카데미 메일의 두 발행일 묶음을 한 번에 붙여넣으세요. 지난 발행일 자료는 파일을 열 때 자동 정리됩니다."
    End If
End Sub


Private Sub RebuildBannerDb()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearLastRow As Long
    Dim sourceRow As Long
    Dim col As Long
    Dim count As Long
    Dim data() As Variant
    Dim issueDate As Date
    Dim eligible() As Long
    Dim eligibleCount As Long
    Dim selected() As Boolean
    Dim remaining() As Long
    Dim remainingCount As Long
    Dim index As Long
    Dim outputRow As Long
    Dim selectedCount As Long
    Dim seqValue As Long
    Dim nextSeq As Long

    Set ws = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    lastRow = MaxLastRow(ws, 1, BANNER_SEQ_COL)
    If lastRow < DB_FIRST_ROW Then lastRow = DB_FIRST_ROW - 1
    ReDim data(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1), 1 To BANNER_PUB_END_COL)
    nextSeq = NextSequence(ws, BANNER_SEQ_COL)

    For sourceRow = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(sourceRow, 1).Value2))) > 0 Then
            If Not IsNumeric(ws.Cells(sourceRow, BANNER_SEQ_COL).Value2) Then
                ws.Cells(sourceRow, BANNER_SEQ_COL).Value = nextSeq
                nextSeq = nextSeq + 1
            End If
            BackfillBannerDates ws, sourceRow
            seqValue = CLng(ws.Cells(sourceRow, BANNER_SEQ_COL).Value2)
            If BannerRowExpired(ws, sourceRow, Date) Then
                DeleteShapeIfExists ws, BannerSourceShapeName(seqValue)
            Else
                count = count + 1
                For col = 1 To BANNER_PUB_END_COL
                    data(count, col) = ws.Cells(sourceRow, col).Value
                Next col
            End If
        End If
    Next sourceRow

    issueDate = NextNewsletterDate(Date)
    ReDim eligible(1 To Application.Max(1, count))
    ReDim selected(1 To Application.Max(1, count))
    For index = 1 To count
        If BannerRecordEligible(data, index, issueDate) Then
            eligibleCount = eligibleCount + 1
            eligible(eligibleCount) = index
        End If
    Next index
    If eligibleCount > 1 Then SortIndexArray eligible, eligibleCount, data, BANNER_SEQ_COL, BANNER_SEQ_COL
    selectedCount = Application.Min(BANNER_CAPACITY, eligibleCount)
    For index = 1 To selectedCount
        selected(eligible(index)) = True
    Next index

    ReDim remaining(1 To Application.Max(1, count))
    For index = 1 To count
        If Not selected(index) Then
            remainingCount = remainingCount + 1
            remaining(remainingCount) = index
        End If
    Next index
    If remainingCount > 1 Then SortIndexArray remaining, remainingCount, data, BANNER_SEQ_COL, BANNER_SEQ_COL

    clearLastRow = Application.Max(lastRow, DB_FIRST_ROW + BANNER_CAPACITY + count + 5)
    ws.Range(ws.Cells(DB_FIRST_ROW, 1), ws.Cells(clearLastRow, BANNER_PUB_END_COL)).ClearContents
    outputRow = DB_FIRST_ROW
    For index = 1 To selectedCount
        WriteDataRow ws, outputRow, data, eligible(index), BANNER_PUB_END_COL, BANNER_VISIBLE_COLS
        ApplyFormWarning ws, outputRow, 3, CStr(data(eligible(index), 3)), BANNER_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    Do While outputRow < DB_FIRST_ROW + BANNER_CAPACITY
        CopyRowFormat ws, DB_FIRST_ROW, outputRow, BANNER_VISIBLE_COLS
        outputRow = outputRow + 1
    Loop
    ApplyCapacityBorder ws, DB_FIRST_ROW + BANNER_CAPACITY - 1, BANNER_VISIBLE_COLS
    For index = 1 To remainingCount
        WriteDataRow ws, outputRow, data, remaining(index), BANNER_PUB_END_COL, BANNER_VISIBLE_COLS
        ApplyFormWarning ws, outputRow, 3, CStr(data(remaining(index), 3)), BANNER_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    CopyRowFormat ws, DB_FIRST_ROW, outputRow, BANNER_VISIBLE_COLS
    RepositionBannerSourceImages
    DeleteOrphanBannerShapes
    ws.Range("A2:H2").MergeArea.Cells(1, 1).Value = _
        "다음 발행일: " & Format$(issueDate, "yyyy.mm.dd.(aaa)") & _
        " | 구분선 위 2건이 최종 시트의 배너 1과 3에 반영됩니다. 배너 2는 해외정보팀 담당 영역입니다."
End Sub


Private Sub RebuildAlimtalkDb()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearLastRow As Long
    Dim sourceRow As Long
    Dim col As Long
    Dim count As Long
    Dim data() As Variant
    Dim nextSeq As Long
    Dim issueDate As Date
    Dim eligible() As Long
    Dim eligibleCount As Long
    Dim selected() As Boolean
    Dim remaining() As Long
    Dim remainingCount As Long
    Dim index As Long
    Dim outputRow As Long
    Dim selectedCount As Long

    Set ws = ThisWorkbook.Worksheets(ALIMTALK_DB_SHEET)
    lastRow = MaxLastRow(ws, 1, ALIMTALK_SEQ_COL)
    If lastRow < DB_FIRST_ROW Then lastRow = DB_FIRST_ROW - 1
    ReDim data(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1), 1 To ALIMTALK_PUB_END_COL)
    nextSeq = NextSequence(ws, ALIMTALK_SEQ_COL)

    For sourceRow = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(sourceRow, 1).Value2))) > 0 Then
            If Not IsNumeric(ws.Cells(sourceRow, ALIMTALK_SEQ_COL).Value2) Then
                ws.Cells(sourceRow, ALIMTALK_SEQ_COL).Value = nextSeq
                nextSeq = nextSeq + 1
            End If
            BackfillBusinessDates ws, sourceRow
            If Not BusinessRowExpired(ws, sourceRow, Date) Then
                count = count + 1
                For col = 1 To ALIMTALK_PUB_END_COL
                    data(count, col) = ws.Cells(sourceRow, col).Value
                Next col
            End If
        End If
    Next sourceRow

    issueDate = NextAlimtalkDate(Date)
    ReDim eligible(1 To Application.Max(1, count))
    ReDim selected(1 To Application.Max(1, count))
    For index = 1 To count
        If BusinessRecordEligible(data, index, issueDate) Then
            eligibleCount = eligibleCount + 1
            eligible(eligibleCount) = index
        End If
    Next index
    If eligibleCount > 1 Then SortIndexArray eligible, eligibleCount, data, ALIMTALK_DEADLINE_COL, ALIMTALK_SEQ_COL
    selectedCount = Application.Min(ALIMTALK_CAPACITY, eligibleCount)
    For index = 1 To selectedCount
        selected(eligible(index)) = True
    Next index

    ReDim remaining(1 To Application.Max(1, count))
    For index = 1 To count
        If Not selected(index) Then
            remainingCount = remainingCount + 1
            remaining(remainingCount) = index
        End If
    Next index
    If remainingCount > 1 Then SortIndexArray remaining, remainingCount, data, ALIMTALK_SEQ_COL, ALIMTALK_SEQ_COL

    clearLastRow = Application.Max(lastRow, DB_FIRST_ROW + ALIMTALK_CAPACITY + count + 5)
    ws.Range(ws.Cells(DB_FIRST_ROW, 1), ws.Cells(clearLastRow, ALIMTALK_PUB_END_COL)).ClearContents
    outputRow = DB_FIRST_ROW
    For index = 1 To selectedCount
        WriteDataRow ws, outputRow, data, eligible(index), ALIMTALK_PUB_END_COL, ALIMTALK_VISIBLE_COLS
        ApplyFormWarning ws, outputRow, 5, CStr(data(eligible(index), 5)), ALIMTALK_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    Do While outputRow < DB_FIRST_ROW + ALIMTALK_CAPACITY
        CopyRowFormat ws, DB_FIRST_ROW, outputRow, ALIMTALK_VISIBLE_COLS
        outputRow = outputRow + 1
    Loop
    ApplyCapacityBorder ws, DB_FIRST_ROW + ALIMTALK_CAPACITY - 1, ALIMTALK_VISIBLE_COLS
    For index = 1 To remainingCount
        WriteDataRow ws, outputRow, data, remaining(index), ALIMTALK_PUB_END_COL, ALIMTALK_VISIBLE_COLS
        ApplyFormWarning ws, outputRow, 5, CStr(data(remaining(index), 5)), ALIMTALK_VISIBLE_COLS
        outputRow = outputRow + 1
    Next index
    CopyRowFormat ws, DB_FIRST_ROW, outputRow, ALIMTALK_VISIBLE_COLS
    ws.Range("A2:I2").MergeArea.Cells(1, 1).Value = _
        "다음 발송일: " & Format$(issueDate, "yyyy.mm.dd.(aaa)") & _
        " | 구분선 위 최대 2건이 달력에 반영됩니다. 마감 임박 순, 동률은 등록 순입니다."
End Sub


Private Function MaxLastRow(ByVal ws As Worksheet, ByVal firstColumn As Long, ByVal secondColumn As Long) As Long
    MaxLastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, firstColumn).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, secondColumn).End(xlUp).Row)
End Function


Private Sub WriteDataRow( _
    ByVal ws As Worksheet, _
    ByVal outputRow As Long, _
    ByRef data As Variant, _
    ByVal dataRow As Long, _
    ByVal totalColumns As Long, _
    ByVal visibleColumns As Long)

    Dim col As Long
    CopyRowFormat ws, DB_FIRST_ROW, outputRow, visibleColumns
    For col = 1 To totalColumns
        ws.Cells(outputRow, col).Value = data(dataRow, col)
    Next col
End Sub


Private Sub SortIndexArray( _
    ByRef indices() As Long, _
    ByVal count As Long, _
    ByRef data As Variant, _
    ByVal primaryColumn As Long, _
    ByVal secondaryColumn As Long)

    Dim i As Long
    Dim j As Long
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim temp As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            leftIndex = indices(i)
            rightIndex = indices(j)
            If SortValue(data(rightIndex, primaryColumn)) < SortValue(data(leftIndex, primaryColumn)) Or _
                (SortValue(data(rightIndex, primaryColumn)) = SortValue(data(leftIndex, primaryColumn)) And _
                SortValue(data(rightIndex, secondaryColumn)) < SortValue(data(leftIndex, secondaryColumn))) Then
                temp = indices(i)
                indices(i) = indices(j)
                indices(j) = temp
            End If
        Next j
    Next i
End Sub


Private Function SortValue(ByVal value As Variant) As Double
    If IsDate(value) Then
        SortValue = CDbl(CDate(value))
    ElseIf IsNumeric(value) Then
        SortValue = CDbl(value)
    Else
        SortValue = 9.9E+307
    End If
End Function


Private Function BusinessRecordEligible(ByRef data As Variant, ByVal rowIndex As Long, ByVal issueDate As Date) As Boolean
    If Not IsDate(data(rowIndex, BUSINESS_START_COL)) Then Exit Function
    If Not IsDate(data(rowIndex, BUSINESS_END_COL)) Then Exit Function
    If Not IsDate(data(rowIndex, BUSINESS_DEADLINE_COL)) Then Exit Function
    If Not IsDate(data(rowIndex, BUSINESS_PUB_START_COL)) Then Exit Function
    If Not IsDate(data(rowIndex, BUSINESS_PUB_END_COL)) Then Exit Function
    BusinessRecordEligible = _
        CDate(data(rowIndex, BUSINESS_PUB_START_COL)) <= issueDate And _
        CDate(data(rowIndex, BUSINESS_PUB_END_COL)) >= issueDate And _
        CDate(data(rowIndex, BUSINESS_DEADLINE_COL)) >= issueDate And _
        CDate(data(rowIndex, BUSINESS_END_COL)) >= issueDate
End Function


Private Function BannerRecordEligible(ByRef data As Variant, ByVal rowIndex As Long, ByVal issueDate As Date) As Boolean
    If Not IsDate(data(rowIndex, BANNER_PUB_START_COL)) Then Exit Function
    If Not IsDate(data(rowIndex, BANNER_PUB_END_COL)) Then Exit Function
    BannerRecordEligible = _
        CDate(data(rowIndex, BANNER_PUB_START_COL)) <= issueDate And _
        CDate(data(rowIndex, BANNER_PUB_END_COL)) >= issueDate
End Function


Private Function BusinessRowExpired(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal todayValue As Date) As Boolean
    If IsDate(ws.Cells(rowNumber, BUSINESS_END_COL).Value) Then
        If CDate(ws.Cells(rowNumber, BUSINESS_END_COL).Value) < todayValue Then BusinessRowExpired = True: Exit Function
    End If
    If IsDate(ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value) Then
        If CDate(ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value) < todayValue Then BusinessRowExpired = True: Exit Function
    End If
    If IsDate(ws.Cells(rowNumber, BUSINESS_PUB_END_COL).Value) Then
        If CDate(ws.Cells(rowNumber, BUSINESS_PUB_END_COL).Value) < todayValue Then BusinessRowExpired = True
    End If
End Function


Private Function BannerRowExpired(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal todayValue As Date) As Boolean
    If IsDate(ws.Cells(rowNumber, BANNER_PUB_END_COL).Value) Then
        BannerRowExpired = CDate(ws.Cells(rowNumber, BANNER_PUB_END_COL).Value) < todayValue
    End If
End Function


Private Sub BackfillBusinessDates(ByVal ws As Worksheet, ByVal rowNumber As Long)
    Dim startDate As Date
    Dim endDate As Date
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_START_COL).Value) Then
        ParseDateRangeWithDefaultYear CStr(ws.Cells(rowNumber, 2).Value2), Year(Date), startDate, endDate
        ws.Cells(rowNumber, BUSINESS_START_COL).Value = startDate
        ws.Cells(rowNumber, BUSINESS_END_COL).Value = endDate
    End If
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value) Then
        ParseDateRangeWithDefaultYear CStr(ws.Cells(rowNumber, 3).Value2), Year(Date), startDate, endDate
        ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value = endDate
    End If
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_PUB_START_COL).Value) Then
        ParseDateRangeWithDefaultYear CStr(ws.Cells(rowNumber, 4).Value2), Year(Date), startDate, endDate
        ws.Cells(rowNumber, BUSINESS_PUB_START_COL).Value = startDate
        ws.Cells(rowNumber, BUSINESS_PUB_END_COL).Value = endDate
    End If
End Sub


Private Sub BackfillBannerDates(ByVal ws As Worksheet, ByVal rowNumber As Long)
    Dim startDate As Date
    Dim endDate As Date
    If Not IsDate(ws.Cells(rowNumber, BANNER_PUB_START_COL).Value) Then
        ParseDateRangeWithDefaultYear CStr(ws.Cells(rowNumber, 2).Value2), Year(Date), startDate, endDate
        ws.Cells(rowNumber, BANNER_PUB_START_COL).Value = startDate
        ws.Cells(rowNumber, BANNER_PUB_END_COL).Value = endDate
    End If
End Sub


Private Function SelectedAcademyIssueDate(ByVal ws As Worksheet, ByVal todayValue As Date) As Date
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim value As Date
    lastRow = MaxLastRow(ws, 2, ACADEMY_SEQ_COL)
    For rowNumber = DB_FIRST_ROW To lastRow
        If IsDate(ws.Cells(rowNumber, 1).Value) Then
            value = CDate(ws.Cells(rowNumber, 1).Value)
            If value >= todayValue Then
                If SelectedAcademyIssueDate = 0 Or value < SelectedAcademyIssueDate Then SelectedAcademyIssueDate = value
            End If
        End If
    Next rowNumber
End Function


Private Sub RefreshFinalSheet()
    Dim finalWs As Worksheet
    Dim businessWs As Worksheet
    Dim academyWs As Worksheet
    Dim bannerWs As Worksheet
    Dim issueDate As Date
    Dim academyIssue As Date
    Dim rowNumber As Long
    Dim outputRow As Long
    Dim itemCount As Long
    Dim sourceSeq As Long

    Set finalWs = ThisWorkbook.Worksheets(FINAL_SHEET_NAME)
    Set businessWs = ThisWorkbook.Worksheets(BUSINESS_DB_SHEET)
    Set academyWs = ThisWorkbook.Worksheets(ACADEMY_DB_SHEET)
    Set bannerWs = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    issueDate = NextNewsletterDate(Date)

    finalWs.Range("A6:E13").ClearContents
    For rowNumber = 6 To 13
        CopyRowFormat finalWs, 6, rowNumber, 5
        finalWs.Cells(rowNumber, 1).Value = rowNumber - 5
    Next rowNumber
    itemCount = 0
    rowNumber = DB_FIRST_ROW
    Do While itemCount < BUSINESS_CAPACITY And Len(Trim$(CStr(businessWs.Cells(rowNumber, 1).Value2))) > 0
        If BusinessWorksheetRowEligible(businessWs, rowNumber, issueDate) Then
            outputRow = 6 + itemCount
            finalWs.Cells(outputRow, 2).Value = businessWs.Cells(rowNumber, 1).Value
            finalWs.Cells(outputRow, 3).Value = FormatPeriodForNewsletter( _
                CDate(businessWs.Cells(rowNumber, BUSINESS_START_COL).Value), _
                CDate(businessWs.Cells(rowNumber, BUSINESS_END_COL).Value), _
                Year(issueDate))
            finalWs.Cells(outputRow, 4).Value = FormatNewsletterDate( _
                CDate(businessWs.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value), _
                Year(issueDate))
            finalWs.Cells(outputRow, 5).Value = FirstLineURL(CStr(businessWs.Cells(rowNumber, 5).Value2))
            ApplyFormWarning finalWs, outputRow, 5, CStr(finalWs.Cells(outputRow, 5).Value2), 5
            itemCount = itemCount + 1
        End If
        rowNumber = rowNumber + 1
    Loop

    finalWs.Range("A18:E21").ClearContents
    For rowNumber = 18 To 21
        CopyRowFormat finalWs, 18, rowNumber, 5
        finalWs.Cells(rowNumber, 1).Value = rowNumber - 17
    Next rowNumber
    academyIssue = SelectedAcademyIssueDate(academyWs, Date)
    itemCount = 0
    If academyIssue > 0 Then
        For rowNumber = DB_FIRST_ROW To MaxLastRow(academyWs, 2, ACADEMY_SEQ_COL)
            If IsDate(academyWs.Cells(rowNumber, 1).Value) Then
                If CDate(academyWs.Cells(rowNumber, 1).Value) = academyIssue Then
                    If itemCount < ACADEMY_CAPACITY Then
                        outputRow = 18 + itemCount
                        finalWs.Cells(outputRow, 2).Value = academyWs.Cells(rowNumber, 2).Value
                        finalWs.Cells(outputRow, 3).Value = academyWs.Cells(rowNumber, 3).Value
                        finalWs.Cells(outputRow, 4).Value = academyWs.Cells(rowNumber, 4).Value
                        finalWs.Cells(outputRow, 5).Value = FirstLineURL(CStr(academyWs.Cells(rowNumber, 5).Value2))
                        ApplyFormWarning finalWs, outputRow, 5, CStr(finalWs.Cells(outputRow, 5).Value2), 5
                        itemCount = itemCount + 1
                    End If
                End If
            End If
        Next rowNumber
    End If

    finalWs.Range("A26:D28").ClearContents
    For rowNumber = 26 To 28
        CopyRowFormat finalWs, 26, rowNumber, 4
    Next rowNumber
    finalWs.Cells(26, 1).Value = 1
    finalWs.Cells(27, 1).Value = 2
    finalWs.Cells(28, 1).Value = 3
    SetBannerPlaceholder finalWs

    DeleteFinalBannerShapes finalWs
    finalWs.Range("A30").Value = "배너 1 이미지"
    finalWs.Range("C30").Value = "배너 3 이미지"
    itemCount = 0
    rowNumber = DB_FIRST_ROW
    Do While itemCount < BANNER_CAPACITY And Len(Trim$(CStr(bannerWs.Cells(rowNumber, 1).Value2))) > 0
        If BannerWorksheetRowEligible(bannerWs, rowNumber, issueDate) Then
            If itemCount = 0 Then
                outputRow = 26
            Else
                outputRow = 28
            End If
            finalWs.Cells(outputRow, 2).Value = bannerWs.Cells(rowNumber, 1).Value
            finalWs.Cells(outputRow, 3).Value = Format$(CDate(bannerWs.Cells(rowNumber, BANNER_PUB_END_COL).Value), "yyyy-mm-dd") & "까지 게재"
            finalWs.Cells(outputRow, 4).Value = FirstLineURL(CStr(bannerWs.Cells(rowNumber, 3).Value2))
            ApplyFormWarning finalWs, outputRow, 4, CStr(finalWs.Cells(outputRow, 4).Value2), 4
            sourceSeq = CLng(bannerWs.Cells(rowNumber, BANNER_SEQ_COL).Value2)
            If itemCount = 0 Then
                CopyBannerImageToFinal sourceSeq, finalWs, 1
            Else
                CopyBannerImageToFinal sourceSeq, finalWs, 3
            End If
            itemCount = itemCount + 1
        End If
        rowNumber = rowNumber + 1
    Loop
End Sub


Private Function BusinessWorksheetRowEligible(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal issueDate As Date) As Boolean
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_END_COL).Value) Then Exit Function
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value) Then Exit Function
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_PUB_START_COL).Value) Then Exit Function
    If Not IsDate(ws.Cells(rowNumber, BUSINESS_PUB_END_COL).Value) Then Exit Function
    BusinessWorksheetRowEligible = _
        CDate(ws.Cells(rowNumber, BUSINESS_PUB_START_COL).Value) <= issueDate And _
        CDate(ws.Cells(rowNumber, BUSINESS_PUB_END_COL).Value) >= issueDate And _
        CDate(ws.Cells(rowNumber, BUSINESS_DEADLINE_COL).Value) >= issueDate And _
        CDate(ws.Cells(rowNumber, BUSINESS_END_COL).Value) >= issueDate
End Function


Private Function BannerWorksheetRowEligible(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal issueDate As Date) As Boolean
    If Not IsDate(ws.Cells(rowNumber, BANNER_PUB_START_COL).Value) Then Exit Function
    If Not IsDate(ws.Cells(rowNumber, BANNER_PUB_END_COL).Value) Then Exit Function
    BannerWorksheetRowEligible = _
        CDate(ws.Cells(rowNumber, BANNER_PUB_START_COL).Value) <= issueDate And _
        CDate(ws.Cells(rowNumber, BANNER_PUB_END_COL).Value) >= issueDate
End Function


Private Sub SetBannerPlaceholder(ByVal ws As Worksheet)
    ws.Cells(27, 2).Value = FINAL_BANNER_PLACEHOLDER
    ws.Range("A27:D27").Interior.Color = RGB(229, 231, 235)
    ws.Range("A27:D27").Font.Color = RGB(100, 116, 139)
End Sub


Public Sub PasteBannerImageAtSelection()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim sequenceValue As Long
    Dim shapeCount As Long
    Dim pastedShape As Shape

    On Error GoTo Failed
    If TypeName(ActiveSheet) <> "Worksheet" Or TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    Set targetCell = Selection.Cells(1, 1)
    If ws.Name <> BANNER_DB_SHEET Or targetCell.Column <> 4 Or targetCell.Row < DB_FIRST_ROW Then Exit Sub
    If Len(Trim$(CStr(ws.Cells(targetCell.Row, 1).Value2))) = 0 Then
        MsgBox "먼저 같은 행에 배너형 원문을 등록하세요.", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(ws.Cells(targetCell.Row, BANNER_SEQ_COL).Value2) Then
        ws.Cells(targetCell.Row, BANNER_SEQ_COL).Value = NextSequence(ws, BANNER_SEQ_COL)
    End If
    sequenceValue = CLng(ws.Cells(targetCell.Row, BANNER_SEQ_COL).Value2)
    DeleteShapeIfExists ws, BannerSourceShapeName(sequenceValue)

    shapeCount = ws.Shapes.Count
    ws.Paste
    If ws.Shapes.Count <= shapeCount Then Err.Raise vbObjectError + 1160, , "클립보드에서 이미지를 붙여넣지 못했습니다."
    Set pastedShape = ws.Shapes(ws.Shapes.Count)
    pastedShape.Name = BannerSourceShapeName(sequenceValue)
    FitShapeToRange pastedShape, targetCell
    pastedShape.Placement = xlMoveAndSize
    RefreshAllCore
    Exit Sub

Failed:
    MsgBox "배너 이미지를 등록하지 못했습니다." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub


Private Function BannerSourceShapeName(ByVal sequenceValue As Long) As String
    BannerSourceShapeName = "DB_BANNER_" & CStr(sequenceValue)
End Function


Private Sub RepositionBannerSourceImages()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim sequenceValue As Long
    Dim rowNumber As Long
    Set ws = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    For Each shp In ws.Shapes
        If Left$(shp.Name, 10) = "DB_BANNER_" Then
            sequenceValue = CLng(Val(Mid$(shp.Name, 11)))
            rowNumber = FindSequenceRow(ws, BANNER_SEQ_COL, sequenceValue)
            If rowNumber > 0 Then
                FitShapeToRange shp, ws.Cells(rowNumber, 4)
                shp.Placement = xlMoveAndSize
            End If
        End If
    Next shp
End Sub


Private Sub DeleteOrphanBannerShapes()
    Dim ws As Worksheet
    Dim index As Long
    Dim sequenceValue As Long
    Set ws = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    For index = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(index).Name, 10) = "DB_BANNER_" Then
            sequenceValue = CLng(Val(Mid$(ws.Shapes(index).Name, 11)))
            If FindSequenceRow(ws, BANNER_SEQ_COL, sequenceValue) = 0 Then ws.Shapes(index).Delete
        End If
    Next index
End Sub


Private Function FindSequenceRow(ByVal ws As Worksheet, ByVal sequenceColumn As Long, ByVal sequenceValue As Long) As Long
    Dim rowNumber As Long
    Dim lastRow As Long
    lastRow = LastDataRow(ws, sequenceColumn)
    For rowNumber = DB_FIRST_ROW To lastRow
        If IsNumeric(ws.Cells(rowNumber, sequenceColumn).Value2) Then
            If CLng(ws.Cells(rowNumber, sequenceColumn).Value2) = sequenceValue Then
                FindSequenceRow = rowNumber
                Exit Function
            End If
        End If
    Next rowNumber
End Function


Private Sub CopyBannerImageToFinal(ByVal sequenceValue As Long, ByVal finalWs As Worksheet, ByVal finalSlot As Long)
    Dim sourceWs As Worksheet
    Dim sourceShape As Shape
    Dim pastedShape As Shape
    Dim targetRange As Range
    Dim shapeCount As Long
    Set sourceWs = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    On Error Resume Next
    Set sourceShape = sourceWs.Shapes(BannerSourceShapeName(sequenceValue))
    On Error GoTo 0
    If sourceShape Is Nothing Then Exit Sub

    If finalSlot = 1 Then
        Set targetRange = finalWs.Range("A30:B36")
        finalWs.Range("A30").ClearContents
    Else
        Set targetRange = finalWs.Range("C30:D36")
        finalWs.Range("C30").ClearContents
    End If
    sourceShape.Copy
    shapeCount = finalWs.Shapes.Count
    finalWs.Paste
    If finalWs.Shapes.Count > shapeCount Then
        Set pastedShape = finalWs.Shapes(finalWs.Shapes.Count)
        pastedShape.Name = "AUTO_BANNER_" & CStr(finalSlot)
        FitShapeToRange pastedShape, targetRange
        pastedShape.Placement = xlMoveAndSize
    End If
End Sub


Private Sub DeleteFinalBannerShapes(ByVal ws As Worksheet)
    Dim index As Long
    For index = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(index).Name, 12) = "AUTO_BANNER_" Then ws.Shapes(index).Delete
    Next index
End Sub


Private Sub FitShapeToRange(ByVal shp As Shape, ByVal targetRange As Range)
    Dim availableWidth As Double
    Dim availableHeight As Double
    availableWidth = targetRange.Width - 8
    availableHeight = targetRange.Height - 8
    shp.LockAspectRatio = msoTrue
    If shp.Width / shp.Height > availableWidth / availableHeight Then
        shp.Width = availableWidth
    Else
        shp.Height = availableHeight
    End If
    shp.Left = targetRange.Left + (targetRange.Width - shp.Width) / 2
    shp.Top = targetRange.Top + (targetRange.Height - shp.Height) / 2
End Sub


Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub


Private Sub EnsureCalendarStructure()
    Dim ws As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim monday As Date
    Dim weekIndex As Long
    Dim blockRow As Long
    Dim rowNumber As Long
    Dim col As Long
    Dim weekdays As Variant

    Set ws = ThisWorkbook.Worksheets(CALENDAR_SHEET)
    If CStr(ws.Range("H2").Value2) = CALENDAR_MARKER Then Exit Sub

    If IsDate(ws.Range("B2").Value) Then
        startDate = CDate(ws.Range("B2").Value)
    Else
        startDate = Date - Weekday(Date, vbMonday) + 1
    End If
    If IsDate(ws.Range("E2").Value) Then
        endDate = CDate(ws.Range("E2").Value)
    ElseIf Month(Date) <= 3 Then
        endDate = DateSerial(Year(Date), 3, 31)
    Else
        endDate = DateSerial(Year(Date) + 1, 3, 31)
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    ws.Cells.UnMerge
    ws.Cells.ClearOutline
    On Error GoTo 0
    Application.DisplayAlerts = True
    ws.Cells.Clear
    ws.Cells.Font.Name = "맑은 고딕"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlCenter
    ActiveWindow.DisplayGridlines = False

    ws.Range("A1:G1").Merge
    ws.Range("A1").Value = CALENDAR_SHEET
    With ws.Range("A1:G1")
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 14
        .RowHeight = 30
    End With
    ws.Cells(2, 1).Value = "시작일 (월요일)"
    ws.Cells(2, 2).Value = startDate
    ws.Cells(2, 3).Value = ""
    ws.Cells(2, 4).Value = "종료일"
    ws.Cells(2, 5).Value = endDate
    ws.Cells(2, 6).Value = ""
    ws.Cells(2, 7).Value = "파일을 열면 지난 주차 자동 접기"
    ws.Range("B2").NumberFormat = "yyyy-mm-dd"
    ws.Range("E2").NumberFormat = "yyyy-mm-dd"
    ws.Rows(2).RowHeight = 24
    ws.Cells(3, 1).Value = "범례"
    ws.Cells(3, 2).Value = "정기 발행: 화·금"
    ws.Cells(3, 3).Value = "정원 내"
    ws.Cells(3, 4).Value = ""
    ws.Cells(3, 5).Value = "정원 초과"
    ws.Cells(3, 6).Value = ""
    ws.Cells(3, 7).Value = "공휴일"
    ws.Range("C3:D3").Interior.Color = RGB(226, 240, 217)
    ws.Range("E3:F3").Interior.Color = RGB(252, 232, 230)
    ws.Range("G3").Interior.Color = RGB(229, 231, 235)
    ws.Rows(3).RowHeight = 22
    ws.Range("A4:G4").Merge
    ws.Range("A4").Value = "공휴일 기준: 우주항공청 2026·2027년 월력요항 (한국천문연구원 월력요항)"
    ws.Range("A4:G4").Font.Italic = True
    ws.Range("A4:G4").Font.Color = RGB(100, 116, 139)
    ws.Rows(4).RowHeight = 20

    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C:G").ColumnWidth = 31
    ws.Columns("H").Hidden = True
    ws.Range("H2").Value = CALENDAR_MARKER
    weekdays = Array("월요일", "화요일", "수요일", "목요일", "금요일")

    weekIndex = 0
    monday = startDate
    Do While monday <= endDate
        blockRow = CALENDAR_FIRST_BLOCK_ROW + weekIndex * CALENDAR_BLOCK_ROWS
        ws.Cells(blockRow, 1).Value = Format$(monday, "yyyy.m.d") & " ~ " & Format$(monday + 4, "m.d")
        ws.Cells(blockRow, 2).Value = "순번"
        For col = 0 To 4
            ws.Cells(blockRow, col + 3).Value = weekdays(col)
            If monday + col <= endDate Then
                ws.Cells(blockRow + 1, col + 3).Value = Format$(monday + col, "yyyy-mm-dd")
            End If
        Next col
        StyleCalendarHeader ws.Range(ws.Cells(blockRow, 1), ws.Cells(blockRow, 7))
        StyleCalendarDateRow ws.Range(ws.Cells(blockRow + 1, 1), ws.Cells(blockRow + 1, 7))

        For rowNumber = 1 To CALENDAR_ALIMTALK_ROWS
            ws.Cells(blockRow + 1 + rowNumber, 1).Value = IIf(rowNumber = 1, "알림톡", "")
            ws.Cells(blockRow + 1 + rowNumber, 2).Value = rowNumber
            StyleCalendarAlimtalkRow ws.Range(ws.Cells(blockRow + 1 + rowNumber, 1), ws.Cells(blockRow + 1 + rowNumber, 7))
        Next rowNumber
        ApplyCalendarCapacityBorder ws, blockRow + 1 + ALIMTALK_CAPACITY, RGB(255, 192, 0)

        For rowNumber = 1 To CALENDAR_BUSINESS_ROWS
            ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + rowNumber, 1).Value = IIf(rowNumber = 1, "뉴스레터(사업안내형)", "")
            ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + rowNumber, 2).Value = rowNumber
            StyleCalendarDataRow ws.Range(ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + rowNumber, 1), ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + rowNumber, 7))
        Next rowNumber
        ApplyCalendarCapacityBorder ws, blockRow + 1 + CALENDAR_ALIMTALK_ROWS + BUSINESS_CAPACITY, RGB(31, 78, 120)

        For rowNumber = 1 To CALENDAR_BANNER_ROWS
            ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + rowNumber, 1).Value = IIf(rowNumber = 1, "뉴스레터(배너형)", "")
            ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + rowNumber, 2).Value = rowNumber
            StyleCalendarDataRow ws.Range(ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + rowNumber, 1), ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + rowNumber, 7))
        Next rowNumber
        ApplyCalendarCapacityBorder ws, blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + BANNER_CAPACITY, RGB(31, 78, 120)
        ws.Rows(blockRow + CALENDAR_BLOCK_ROWS - 1).RowHeight = 10

        monday = monday + 7
        weekIndex = weekIndex + 1
    Loop
    ws.Outline.SummaryRow = xlAbove
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub


Private Sub RefreshCalendar()
    Dim ws As Worksheet
    Dim businessWs As Worksheet
    Dim bannerWs As Worksheet
    Dim alimtalkWs As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim monday As Date
    Dim weekIndex As Long
    Dim blockRow As Long
    Dim blockEnd As Long
    Dim issueDate As Date
    Dim dayOffset As Long
    Dim targetColumn As Long
    Dim businessRows() As Long
    Dim bannerRows() As Long
    Dim alimtalkRows() As Long
    Dim businessCount As Long
    Dim bannerCount As Long
    Dim alimtalkCount As Long
    Dim maxBusiness As Long
    Dim maxBanner As Long
    Dim maxAlimtalk As Long
    Dim displayCount As Long
    Dim itemIndex As Long
    Dim rowNumber As Long
    Dim holidayText As String
    Dim currentMonday As Date

    Set ws = ThisWorkbook.Worksheets(CALENDAR_SHEET)
    Set businessWs = ThisWorkbook.Worksheets(BUSINESS_DB_SHEET)
    Set bannerWs = ThisWorkbook.Worksheets(BANNER_DB_SHEET)
    Set alimtalkWs = ThisWorkbook.Worksheets(ALIMTALK_DB_SHEET)
    startDate = CDate(ws.Range("B2").Value)
    endDate = CDate(ws.Range("E2").Value)
    currentMonday = Date - Weekday(Date, vbMonday) + 1

    On Error Resume Next
    ws.Cells.ClearOutline
    On Error GoTo 0
    ws.Outline.SummaryRow = xlAbove

    weekIndex = 0
    monday = startDate
    Do While monday <= endDate
        blockRow = CALENDAR_FIRST_BLOCK_ROW + weekIndex * CALENDAR_BLOCK_ROWS
        blockEnd = blockRow + CALENDAR_BLOCK_ROWS - 1
        ws.Rows(blockRow & ":" & blockEnd).Hidden = False

        If monday < currentMonday Then
            ws.Rows((blockRow + 1) & ":" & blockEnd).Rows.Group
            ws.Rows((blockRow + 1) & ":" & blockEnd).Hidden = True
        Else
            ws.Range(ws.Cells(blockRow + 2, 3), ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + CALENDAR_BANNER_ROWS, 7)).ClearContents
            ResetCalendarBodyColors ws, blockRow
            maxAlimtalk = 0
            maxBusiness = 0
            maxBanner = 0

            For dayOffset = 0 To 4
                issueDate = monday + dayOffset
                targetColumn = dayOffset + 3
                ws.Cells(blockRow + 1, targetColumn).Value = IIf(issueDate <= endDate, Format$(issueDate, "yyyy-mm-dd"), "")
                holidayText = HolidayName(issueDate)
                If holidayText <> "" Then
                    ws.Cells(blockRow + 1, targetColumn).Interior.Color = RGB(229, 231, 235)
                    ws.Cells(blockRow + 1, targetColumn).Font.Color = RGB(192, 0, 0)
                    ws.Cells(blockRow + 2, targetColumn).Value = "공휴일: " & holidayText
                    ws.Cells(blockRow + 2, targetColumn).Interior.Color = RGB(229, 231, 235)
                Else
                    If issueDate <= endDate And (Weekday(issueDate, vbMonday) = 1 Or Weekday(issueDate, vbMonday) = 3) Then
                        GetAlimtalkCandidateRows alimtalkWs, issueDate, alimtalkRows, alimtalkCount
                        If alimtalkCount > maxAlimtalk Then maxAlimtalk = alimtalkCount
                        displayCount = Application.Min(alimtalkCount, CALENDAR_ALIMTALK_ROWS)
                        For itemIndex = 1 To displayCount
                            ws.Cells(blockRow + 1 + itemIndex, targetColumn).Value = CalendarBusinessTitle(alimtalkWs, alimtalkRows(itemIndex))
                            If itemIndex > ALIMTALK_CAPACITY Then ws.Cells(blockRow + 1 + itemIndex, targetColumn).Interior.Color = RGB(252, 232, 230)
                        Next itemIndex
                        If alimtalkCount > CALENDAR_ALIMTALK_ROWS Then
                            ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS, targetColumn).Value = _
                                ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS, targetColumn).Value & _
                                " 외 " & CStr(alimtalkCount - CALENDAR_ALIMTALK_ROWS) & "건"
                        End If
                    End If

                    If issueDate <= endDate And (Weekday(issueDate, vbMonday) = 2 Or Weekday(issueDate, vbMonday) = 5) Then
                        GetBusinessCandidateRows businessWs, issueDate, businessRows, businessCount
                        GetBannerCandidateRows bannerWs, issueDate, bannerRows, bannerCount
                        If businessCount > maxBusiness Then maxBusiness = businessCount
                        If bannerCount > maxBanner Then maxBanner = bannerCount

                        displayCount = Application.Min(businessCount, CALENDAR_BUSINESS_ROWS)
                        For itemIndex = 1 To displayCount
                            rowNumber = blockRow + 1 + CALENDAR_ALIMTALK_ROWS + itemIndex
                            ws.Cells(rowNumber, targetColumn).Value = CalendarBusinessTitle(businessWs, businessRows(itemIndex))
                            If itemIndex > BUSINESS_CAPACITY Then ws.Cells(rowNumber, targetColumn).Interior.Color = RGB(252, 232, 230)
                        Next itemIndex
                        If businessCount > CALENDAR_BUSINESS_ROWS Then
                            rowNumber = blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS
                            ws.Cells(rowNumber, targetColumn).Value = ws.Cells(rowNumber, targetColumn).Value & _
                                " 외 " & CStr(businessCount - CALENDAR_BUSINESS_ROWS) & "건"
                        End If

                        displayCount = Application.Min(bannerCount, CALENDAR_BANNER_ROWS)
                        For itemIndex = 1 To displayCount
                            rowNumber = blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + itemIndex
                            ws.Cells(rowNumber, targetColumn).Value = CalendarBannerTitle(bannerWs, bannerRows(itemIndex))
                            If itemIndex > BANNER_CAPACITY Then ws.Cells(rowNumber, targetColumn).Interior.Color = RGB(252, 232, 230)
                        Next itemIndex
                        If bannerCount > CALENDAR_BANNER_ROWS Then
                            rowNumber = blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + CALENDAR_BANNER_ROWS
                            ws.Cells(rowNumber, targetColumn).Value = ws.Cells(rowNumber, targetColumn).Value & _
                                " 외 " & CStr(bannerCount - CALENDAR_BANNER_ROWS) & "건"
                        End If
                    End If
                End If
            Next dayOffset

            ShowCalendarRows ws, blockRow, maxAlimtalk, maxBusiness, maxBanner
        End If

        monday = monday + 7
        weekIndex = weekIndex + 1
    Loop
End Sub


Private Sub GetAlimtalkCandidateRows( _
    ByVal ws As Worksheet, _
    ByVal issueDate As Date, _
    ByRef rows() As Long, _
    ByRef count As Long)

    Dim rowNumber As Long
    Dim lastRow As Long
    count = 0
    lastRow = MaxLastRow(ws, 1, ALIMTALK_SEQ_COL)
    ReDim rows(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1))
    For rowNumber = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(rowNumber, 1).Value2))) > 0 Then
            If BusinessWorksheetRowEligible(ws, rowNumber, issueDate) Then
                count = count + 1
                rows(count) = rowNumber
            End If
        End If
    Next rowNumber
    If count > 1 Then SortWorksheetRows rows, count, ws, ALIMTALK_DEADLINE_COL, ALIMTALK_SEQ_COL
End Sub


Private Sub GetBusinessCandidateRows( _
    ByVal ws As Worksheet, _
    ByVal issueDate As Date, _
    ByRef rows() As Long, _
    ByRef count As Long)

    Dim rowNumber As Long
    Dim lastRow As Long
    count = 0
    lastRow = MaxLastRow(ws, 1, BUSINESS_SEQ_COL)
    ReDim rows(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1))
    For rowNumber = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(rowNumber, 1).Value2))) > 0 Then
            If BusinessWorksheetRowEligible(ws, rowNumber, issueDate) Then
                count = count + 1
                rows(count) = rowNumber
            End If
        End If
    Next rowNumber
    If count > 1 Then SortWorksheetRows rows, count, ws, BUSINESS_DEADLINE_COL, BUSINESS_SEQ_COL
End Sub


Private Sub GetBannerCandidateRows( _
    ByVal ws As Worksheet, _
    ByVal issueDate As Date, _
    ByRef rows() As Long, _
    ByRef count As Long)

    Dim rowNumber As Long
    Dim lastRow As Long
    count = 0
    lastRow = MaxLastRow(ws, 1, BANNER_SEQ_COL)
    ReDim rows(1 To Application.Max(1, lastRow - DB_FIRST_ROW + 1))
    For rowNumber = DB_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(rowNumber, 1).Value2))) > 0 Then
            If BannerWorksheetRowEligible(ws, rowNumber, issueDate) Then
                count = count + 1
                rows(count) = rowNumber
            End If
        End If
    Next rowNumber
    If count > 1 Then SortWorksheetRows rows, count, ws, BANNER_SEQ_COL, BANNER_SEQ_COL
End Sub


Private Sub SortWorksheetRows( _
    ByRef rows() As Long, _
    ByVal count As Long, _
    ByVal ws As Worksheet, _
    ByVal primaryColumn As Long, _
    ByVal secondaryColumn As Long)

    Dim i As Long
    Dim j As Long
    Dim temp As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            If SortValue(ws.Cells(rows(j), primaryColumn).Value) < SortValue(ws.Cells(rows(i), primaryColumn).Value) Or _
                (SortValue(ws.Cells(rows(j), primaryColumn).Value) = SortValue(ws.Cells(rows(i), primaryColumn).Value) And _
                SortValue(ws.Cells(rows(j), secondaryColumn).Value) < SortValue(ws.Cells(rows(i), secondaryColumn).Value)) Then
                temp = rows(i)
                rows(i) = rows(j)
                rows(j) = temp
            End If
        Next j
    Next i
End Sub


Private Function CalendarBusinessTitle(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    CalendarBusinessTitle = CStr(ws.Cells(rowNumber, 1).Value2)
    If Len(Trim$(CStr(ws.Cells(rowNumber, 7).Value2))) > 0 Then
        CalendarBusinessTitle = CalendarBusinessTitle & "(" & CStr(ws.Cells(rowNumber, 7).Value2) & ")"
    End If
End Function


Private Function CalendarBannerTitle(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    CalendarBannerTitle = CStr(ws.Cells(rowNumber, 1).Value2)
    If Len(Trim$(CStr(ws.Cells(rowNumber, 6).Value2))) > 0 Then
        CalendarBannerTitle = CalendarBannerTitle & "(" & CStr(ws.Cells(rowNumber, 6).Value2) & ")"
    End If
End Function


Private Sub ShowCalendarRows( _
    ByVal ws As Worksheet, _
    ByVal blockRow As Long, _
    ByVal alimtalkCount As Long, _
    ByVal businessCount As Long, _
    ByVal bannerCount As Long)

    Dim alimtalkVisible As Long
    Dim businessVisible As Long
    Dim bannerVisible As Long
    Dim rowNumber As Long
    alimtalkVisible = Application.Max(ALIMTALK_CAPACITY, Application.Min(alimtalkCount, CALENDAR_ALIMTALK_ROWS))
    businessVisible = Application.Max(BUSINESS_CAPACITY, Application.Min(businessCount, CALENDAR_BUSINESS_ROWS))
    bannerVisible = Application.Max(BANNER_CAPACITY, Application.Min(bannerCount, CALENDAR_BANNER_ROWS))

    For rowNumber = 1 To CALENDAR_ALIMTALK_ROWS
        ws.Rows(blockRow + 1 + rowNumber).Hidden = rowNumber > alimtalkVisible
    Next rowNumber
    For rowNumber = 1 To CALENDAR_BUSINESS_ROWS
        ws.Rows(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + rowNumber).Hidden = rowNumber > businessVisible
    Next rowNumber
    For rowNumber = 1 To CALENDAR_BANNER_ROWS
        ws.Rows(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + rowNumber).Hidden = rowNumber > bannerVisible
    Next rowNumber
End Sub


Private Sub ResetCalendarBodyColors(ByVal ws As Worksheet, ByVal blockRow As Long)
    ws.Range(ws.Cells(blockRow + 1, 1), ws.Cells(blockRow + 1, 7)).Interior.Color = RGB(247, 249, 252)
    ws.Range(ws.Cells(blockRow + 1, 1), ws.Cells(blockRow + 1, 7)).Font.Color = RGB(31, 41, 55)
    ws.Range(ws.Cells(blockRow + 2, 1), ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS, 7)).Interior.Color = RGB(255, 253, 242)
    ws.Range(ws.Cells(blockRow + 2 + CALENDAR_ALIMTALK_ROWS, 1), ws.Cells(blockRow + 1 + CALENDAR_ALIMTALK_ROWS + CALENDAR_BUSINESS_ROWS + CALENDAR_BANNER_ROWS, 7)).Interior.Color = vbWhite
End Sub


Private Function HolidayName(ByVal value As Date) As String
    Select Case Format$(value, "yyyy-mm-dd")
        Case "2026-12-25": HolidayName = "성탄절"
        Case "2027-01-01": HolidayName = "신정"
        Case "2027-02-08": HolidayName = "설날 연휴"
        Case "2027-02-09": HolidayName = "설날 대체공휴일"
        Case "2027-03-01": HolidayName = "삼일절"
    End Select
End Function


Private Sub StyleCalendarHeader(ByVal targetRange As Range)
    With targetRange
        .Interior.Color = RGB(217, 234, 247)
        .Font.Bold = True
        .Font.Color = RGB(23, 54, 93)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(159, 186, 208)
        .RowHeight = 22
    End With
End Sub


Private Sub StyleCalendarDateRow(ByVal targetRange As Range)
    With targetRange
        .Interior.Color = RGB(247, 249, 252)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(214, 222, 230)
        .RowHeight = 22
    End With
End Sub


Private Sub StyleCalendarDataRow(ByVal targetRange As Range)
    With targetRange
        .Interior.Color = vbWhite
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(214, 222, 230)
        .RowHeight = 30
    End With
    targetRange.Cells(1, 2).HorizontalAlignment = xlCenter
End Sub


Private Sub StyleCalendarAlimtalkRow(ByVal targetRange As Range)
    StyleCalendarDataRow targetRange
    targetRange.Interior.Color = RGB(255, 253, 242)
End Sub


Private Sub ApplyCalendarCapacityBorder(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal colorValue As Long)
    With ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 7)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = colorValue
    End With
End Sub


Private Function FormatAcademyPeriod( _
    ByVal value As String, _
    ByVal baseYear As Long, _
    ByRef firstDate As Date, _
    ByRef lastDate As Date) As String

    Dim normalized As String
    Dim parts() As String
    Dim index As Long
    Dim currentDefaultYear As Long
    Dim segmentStart As Date
    Dim segmentEnd As Date
    Dim previousDate As Date
    Dim segmentText As String
    Dim result As String
    Dim hasFirstDate As Boolean

    normalized = Replace(value, ChrW(183), "|")
    normalized = Replace(normalized, ChrW(12685), "|")
    normalized = Replace(normalized, ChrW(12539), "|")

    If InStr(1, normalized, "|", vbBinaryCompare) = 0 Then
        ParseDateRangeWithDefaultYear value, baseYear, firstDate, lastDate
        FormatAcademyPeriod = FormatPeriodForNewsletter(firstDate, lastDate, baseYear)
        Exit Function
    End If

    parts = Split(normalized, "|")
    currentDefaultYear = baseYear
    For index = LBound(parts) To UBound(parts)
        segmentText = Trim$(parts(index))
        If segmentText <> "" Then
            ParseDateRangeWithDefaultYear segmentText, currentDefaultYear, segmentStart, segmentEnd
            If index > LBound(parts) And Not HasExplicitYear(segmentText) And segmentStart < previousDate Then
                ParseDateRangeWithDefaultYear segmentText, Year(previousDate) + 1, segmentStart, segmentEnd
            End If
            If result <> "" Then result = result & " · "
            result = result & FormatPeriodForNewsletter(segmentStart, segmentEnd, baseYear)
            If Not hasFirstDate Then
                firstDate = segmentStart
                hasFirstDate = True
            End If
            lastDate = segmentEnd
            previousDate = segmentEnd
            currentDefaultYear = Year(segmentEnd)
        End If
    Next index

    If result = "" Then Err.Raise vbObjectError + 1058, , "교육기간에서 날짜를 찾지 못했습니다: " & value
    FormatAcademyPeriod = result
End Function


Private Sub ParseAcademyDeadline( _
    ByVal value As String, _
    ByVal educationStart As Date, _
    ByRef startDate As Date, _
    ByRef endDate As Date)

    ParseDateRangeWithDefaultYear value, Year(educationStart), startDate, endDate
    If Not HasExplicitYear(value) And endDate > educationStart Then
        ParseDateRangeWithDefaultYear value, Year(educationStart) - 1, startDate, endDate
    End If
End Sub


Private Function HasExplicitYear(ByVal value As String) As Boolean
    Dim numbers As Collection
    Set numbers = ExtractNumbers(RemoveParentheticalText(value))
    If numbers.Count < 3 Then Exit Function
    HasExplicitYear = IsYearToken(CLng(numbers(1)))
End Function


Private Function AcademyFieldValue(ByVal value As String, ByVal wantedLabel As String) As String
    Dim colonPosition As Long
    Dim key As String
    colonPosition = InStr(1, value, ":", vbBinaryCompare)
    If colonPosition = 0 Then Exit Function
    key = Left$(value, colonPosition - 1)
    If NormalizeKey(key) = NormalizeKey(wantedLabel) Then AcademyFieldValue = Trim$(Mid$(value, colonPosition + 1))
End Function


Private Function FormatPeriodForNewsletter(ByVal startDate As Date, ByVal endDate As Date, ByVal baseYear As Long) As String
    FormatPeriodForNewsletter = FormatNewsletterDate(startDate, baseYear)
    If endDate <> startDate Then
        FormatPeriodForNewsletter = FormatPeriodForNewsletter & "  -  " & FormatNewsletterDate(endDate, baseYear)
    End If
End Function


Private Function FormatNewsletterDate(ByVal value As Date, ByVal baseYear As Long) As String
    Dim yearPrefix As String
    Dim weekdayText As String
    If Year(value) <> baseYear Then yearPrefix = Format$(Year(value) Mod 100, "00") & "."
    weekdayText = Choose(Weekday(value, vbSunday), "일", "월", "화", "수", "목", "금", "토")
    FormatNewsletterDate = yearPrefix & Month(value) & "." & Day(value) & ".(" & weekdayText & ")"
End Function


Private Function FormatDbDate(ByVal value As Date) As String
    FormatDbDate = Format$(value, "yyyy.m.d")
End Function


Private Function FormatDbPeriod(ByVal startDate As Date, ByVal endDate As Date) As String
    FormatDbPeriod = FormatDbDate(startDate)
    If endDate <> startDate Then FormatDbPeriod = FormatDbPeriod & " ~ " & FormatDbDate(endDate)
End Function


Private Function FirstLineURL(ByVal value As String) As String
    Dim lineBreak As Long
    value = Trim$(value)
    lineBreak = InStr(1, value, vbLf, vbBinaryCompare)
    If lineBreak > 0 Then value = Left$(value, lineBreak - 1)
    FirstLineURL = ExtractURL(value)
End Function


Private Sub CopyRowFormat( _
    ByVal ws As Worksheet, _
    ByVal templateRow As Long, _
    ByVal destinationRow As Long, _
    ByVal visibleColumns As Long)

    ws.Range(ws.Cells(templateRow, 1), ws.Cells(templateRow, visibleColumns)).Copy
    ws.Range(ws.Cells(destinationRow, 1), ws.Cells(destinationRow, visibleColumns)).PasteSpecial xlPasteFormats
    ws.Rows(destinationRow).RowHeight = ws.Rows(templateRow).RowHeight
    Application.CutCopyMode = False

    If IsDbSheet(ws.Name) Then
        With ws.Range(ws.Cells(destinationRow, 1), ws.Cells(destinationRow, visibleColumns))
            .Interior.Color = vbWhite
            .Font.Name = "맑은 고딕"
            .Font.Size = 10
            .Font.Bold = False
            .Font.Color = RGB(31, 41, 55)
            .WrapText = True
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(214, 222, 230)
            .Borders.Weight = xlThin
        End With
        ws.Rows(destinationRow).RowHeight = 34
        If ws.Name = BANNER_DB_SHEET Then
            ws.Cells(destinationRow, 4).Interior.Color = RGB(247, 249, 252)
            ws.Cells(destinationRow, 4).Font.Color = RGB(100, 116, 139)
        End If
    End If
End Sub


Private Sub ApplyCapacityBorder(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal visibleColumns As Long)
    With ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, visibleColumns)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(31, 78, 120)
    End With
End Sub


Private Sub ApplyFormWarning( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal urlColumn As Long, _
    ByVal targetURL As String, _
    ByVal visibleColumns As Long)

    Dim provider As String
    Dim rowRange As Range
    provider = FormProviderName(targetURL)
    If provider = "" Then Exit Sub

    Set rowRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, visibleColumns))
    rowRange.Interior.Color = RGB(255, 199, 206)
    rowRange.Font.Color = RGB(156, 0, 6)
    ws.Cells(rowNumber, urlColumn).Value = FirstLineURL(targetURL) & vbLf & provider & FORM_WARNING_SUFFIX
    ws.Cells(rowNumber, urlColumn).WrapText = True
    ws.Cells(rowNumber, urlColumn).Font.Bold = True
    ws.Rows(rowNumber).AutoFit
    If ws.Rows(rowNumber).RowHeight < 60 Then ws.Rows(rowNumber).RowHeight = 60
End Sub


Private Function FormProviderName(ByVal value As String) As String
    Dim hostValue As String
    Dim lowerValue As String
    hostValue = URLHost(FirstLineURL(value))
    lowerValue = LCase$(value)
    If HostMatches(hostValue, "moaform.com") Or HostMatches(hostValue, "surveyl.ink") Then
        FormProviderName = "모아폼"
    ElseIf HostMatches(hostValue, "forms.gle") Or HostMatches(hostValue, "forms.google.com") Then
        FormProviderName = "구글폼"
    ElseIf HostMatches(hostValue, "docs.google.com") And InStr(1, lowerValue, "/forms", vbTextCompare) > 0 Then
        FormProviderName = "구글폼"
    ElseIf HostMatches(hostValue, "form.naver.com") Then
        FormProviderName = "네이버폼"
    End If
End Function


Private Function HostMatches(ByVal hostValue As String, ByVal domainValue As String) As Boolean
    hostValue = LCase$(hostValue)
    domainValue = LCase$(domainValue)
    If hostValue = domainValue Then
        HostMatches = True
    ElseIf Len(hostValue) > Len(domainValue) Then
        HostMatches = Right$(hostValue, Len(domainValue) + 1) = "." & domainValue
    End If
End Function


Private Function URLHost(ByVal value As String) As String
    Dim lowerValue As String
    Dim schemePosition As Long
    Dim hostStart As Long
    Dim hostEnd As Long
    Dim candidatePosition As Long
    Dim colonPosition As Long
    Dim hostValue As String
    lowerValue = LCase$(Trim$(value))
    schemePosition = InStr(1, lowerValue, "://", vbBinaryCompare)
    If schemePosition = 0 Then Exit Function
    hostStart = schemePosition + 3
    hostEnd = Len(lowerValue) + 1
    candidatePosition = InStr(hostStart, lowerValue, "/")
    If candidatePosition > 0 And candidatePosition < hostEnd Then hostEnd = candidatePosition
    candidatePosition = InStr(hostStart, lowerValue, "?")
    If candidatePosition > 0 And candidatePosition < hostEnd Then hostEnd = candidatePosition
    candidatePosition = InStr(hostStart, lowerValue, "#")
    If candidatePosition > 0 And candidatePosition < hostEnd Then hostEnd = candidatePosition
    hostValue = Mid$(lowerValue, hostStart, hostEnd - hostStart)
    colonPosition = InStr(1, hostValue, ":", vbBinaryCompare)
    If colonPosition > 0 Then hostValue = Left$(hostValue, colonPosition - 1)
    URLHost = hostValue
End Function


Private Function GetFieldValue(ByVal rawText As String, ByVal wantedLabel As String) As String
    Dim normalized As String
    Dim lines() As String
    Dim line As Variant
    Dim colonPosition As Long
    Dim key As String
    normalized = Replace(rawText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, "：", ":")
    lines = Split(normalized, vbLf)
    For Each line In lines
        colonPosition = InStr(1, CStr(line), ":", vbBinaryCompare)
        If colonPosition > 0 Then
            key = Left$(CStr(line), colonPosition - 1)
            If NormalizeKey(key) = NormalizeKey(wantedLabel) Then
                GetFieldValue = Trim$(Mid$(CStr(line), colonPosition + 1))
                Exit Function
            End If
        End If
    Next line
End Function


Private Function NormalizeKey(ByVal value As String) As String
    Dim digitValue As Long
    value = LCase$(value)
    value = Replace(value, " ", "")
    value = Replace(value, vbTab, "")
    value = Replace(value, "*", "")
    value = Replace(value, "ㅇ", "")
    value = Replace(value, "○", "")
    value = Replace(value, "●", "")
    value = Replace(value, ChrW(8226), "")
    value = Replace(value, "-", "")
    value = Replace(value, "(", "")
    value = Replace(value, ")", "")
    value = Replace(value, ".", "")
    For digitValue = 0 To 9
        value = Replace(value, CStr(digitValue), "")
    Next digitValue
    NormalizeKey = Trim$(value)
End Function


Private Function CleanDisplayText(ByVal value As String) As String
    value = Replace(value, ChrW(160), " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    CleanDisplayText = Trim$(value)
End Function


' 별도 참조 설정 없이 Windows Excel에서 동작하도록 정규식 대신 숫자 토큰을 직접 읽습니다.
Private Sub ParseDateRangeWithDefaultYear( _
    ByVal value As String, _
    ByVal defaultYear As Long, _
    ByRef startDate As Date, _
    ByRef endDate As Date)

    ParseFlexibleDateRange value, defaultYear, startDate, endDate
End Sub


Private Sub ParseFlexibleDateRange( _
    ByVal value As String, _
    ByVal defaultYear As Long, _
    ByRef startDate As Date, _
    ByRef endDate As Date)

    Dim cleanedValue As String
    Dim numbers As Collection
    Dim numberCount As Long
    Dim nextIndex As Long
    Dim remainingCount As Long
    Dim startYear As Long
    Dim startMonth As Long
    Dim startDay As Long
    Dim endYear As Long
    Dim endMonth As Long
    Dim endDay As Long

    cleanedValue = RemoveParentheticalText(value)
    Set numbers = ExtractNumbers(cleanedValue)
    numberCount = numbers.Count

    If numberCount < 2 Then
        Err.Raise vbObjectError + 1034, , "월과 일을 포함한 날짜를 찾지 못했습니다: " & value
    End If

    startYear = defaultYear
    nextIndex = 1
    If numberCount >= 3 And IsYearToken(CLng(numbers(1))) Then
        startYear = ExpandYear(CLng(numbers(1)))
        nextIndex = 2
    End If

    If nextIndex + 1 > numberCount Then
        Err.Raise vbObjectError + 1031, , "첫 번째 날짜가 완전하지 않습니다: " & value
    End If

    startMonth = CLng(numbers(nextIndex))
    startDay = CLng(numbers(nextIndex + 1))
    startDate = StrictDate(startYear, startMonth, startDay, value)
    nextIndex = nextIndex + 2

    If nextIndex > numberCount Or Not ContainsRangeMarker(cleanedValue) Then
        endDate = startDate
        Exit Sub
    End If

    remainingCount = numberCount - nextIndex + 1
    If remainingCount >= 3 And IsYearToken(CLng(numbers(nextIndex))) Then
        endYear = ExpandYear(CLng(numbers(nextIndex)))
        endMonth = CLng(numbers(nextIndex + 1))
        endDay = CLng(numbers(nextIndex + 2))
        endDate = StrictDate(endYear, endMonth, endDay, value)
        Exit Sub
    End If

    endYear = startYear
    If remainingCount >= 2 Then
        endMonth = CLng(numbers(nextIndex))
        endDay = CLng(numbers(nextIndex + 1))
        endDate = StrictDate(endYear, endMonth, endDay, value)
        If endDate < startDate Then endDate = StrictDate(endYear + 1, endMonth, endDay, value)
    Else
        endMonth = startMonth
        endDay = CLng(numbers(nextIndex))
        endDate = StrictDate(endYear, endMonth, endDay, value)
        If endDate < startDate Then
            endMonth = startMonth + 1
            If endMonth > 12 Then
                endMonth = 1
                endYear = endYear + 1
            End If
            endDate = StrictDate(endYear, endMonth, endDay, value)
        End If
    End If
End Sub


Private Function IsYearToken(ByVal value As Long) As Boolean
    IsYearToken = (value >= 1900 And value <= 2100) Or (value >= 20 And value <= 99)
End Function


Private Function ExpandYear(ByVal value As Long) As Long
    If value >= 20 And value <= 99 Then
        ExpandYear = 2000 + value
    Else
        ExpandYear = value
    End If
End Function


Private Function RemoveParentheticalText(ByVal value As String) As String
    Dim result As String
    Dim index As Long
    Dim character As String
    Dim depth As Long

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If character = "(" Or character = ChrW(65288) Then
            depth = depth + 1
        ElseIf character = ")" Or character = ChrW(65289) Then
            If depth > 0 Then depth = depth - 1
        ElseIf depth = 0 Then
            result = result & character
        End If
    Next index
    RemoveParentheticalText = result
End Function


Private Function ExtractNumbers(ByVal value As String) As Collection
    Dim result As New Collection
    Dim currentNumber As String
    Dim index As Long
    Dim character As String

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If character >= "0" And character <= "9" Then
            currentNumber = currentNumber & character
        ElseIf currentNumber <> "" Then
            result.Add CLng(currentNumber)
            currentNumber = ""
        End If
    Next index
    If currentNumber <> "" Then result.Add CLng(currentNumber)
    Set ExtractNumbers = result
End Function


Private Function ContainsRangeMarker(ByVal value As String) As Boolean
    ContainsRangeMarker = _
        InStr(value, "~") > 0 Or InStr(value, "～") > 0 Or _
        InStr(value, ChrW(8764)) > 0 Or InStr(value, ChrW(12316)) > 0 Or _
        InStr(value, "부터") > 0 Or InStr(value, "까지") > 0 Or _
        InStr(value, "에서") > 0 Or InStr(value, "-") > 0 Or _
        InStr(value, ChrW(8722)) > 0 Or InStr(value, ChrW(8211)) > 0 Or _
        InStr(value, ChrW(8212)) > 0 Or InStr(value, ChrW(8213)) > 0 Or _
        InStr(value, ChrW(65293)) > 0
End Function


Private Function StrictDate(ByVal yearValue As Long, ByVal monthValue As Long, ByVal dayValue As Long, ByVal sourceText As String) As Date
    Dim parsed As Date
    On Error GoTo InvalidDate
    parsed = DateSerial(yearValue, monthValue, dayValue)
    If Year(parsed) <> yearValue Or Month(parsed) <> monthValue Or Day(parsed) <> dayValue Then GoTo InvalidDate
    StrictDate = parsed
    Exit Function

InvalidDate:
    Err.Raise vbObjectError + 1033, , "올바르지 않은 날짜입니다: " & sourceText
End Function


Private Function ExtractURL(ByVal value As String) As String
    Dim lowerValue As String
    Dim httpPosition As Long
    Dim httpsPosition As Long
    Dim startPosition As Long
    Dim index As Long
    Dim character As String
    Dim result As String

    lowerValue = LCase$(value)
    httpPosition = InStr(1, lowerValue, "http://", vbBinaryCompare)
    httpsPosition = InStr(1, lowerValue, "https://", vbBinaryCompare)
    If httpPosition = 0 Then
        startPosition = httpsPosition
    ElseIf httpsPosition = 0 Then
        startPosition = httpPosition
    ElseIf httpPosition < httpsPosition Then
        startPosition = httpPosition
    Else
        startPosition = httpsPosition
    End If
    If startPosition = 0 Then Exit Function

    For index = startPosition To Len(value)
        character = Mid$(value, index, 1)
        If character = " " Or character = ChrW(160) Or character = vbTab Or character = vbCr Or character = vbLf _
            Or character = ")" Or character = "]" Or character = ">" Then Exit For
        result = result & character
    Next index
    Do While Len(result) > 0 And InStr(".,;", Right$(result, 1)) > 0
        result = Left$(result, Len(result) - 1)
    Loop
    ExtractURL = NormalizeKotraURL(result)
End Function


Private Function NormalizeKotraURL(ByVal value As String) As String
    Dim lowerValue As String
    Dim schemeLength As Long
    Dim hostStart As Long
    Dim hostEnd As Long
    Dim questionPosition As Long
    Dim hashPosition As Long
    Dim hostValue As String

    lowerValue = LCase$(value)
    If Left$(lowerValue, 8) = "https://" Then
        schemeLength = 8
    ElseIf Left$(lowerValue, 7) = "http://" Then
        schemeLength = 7
    Else
        NormalizeKotraURL = value
        Exit Function
    End If

    hostStart = schemeLength + 1
    hostEnd = InStr(hostStart, lowerValue, "/")
    questionPosition = InStr(hostStart, lowerValue, "?")
    hashPosition = InStr(hostStart, lowerValue, "#")
    If hostEnd = 0 Or (questionPosition > 0 And questionPosition < hostEnd) Then hostEnd = questionPosition
    If hostEnd = 0 Or (hashPosition > 0 And hashPosition < hostEnd) Then hostEnd = hashPosition
    If hostEnd = 0 Then hostEnd = Len(value) + 1

    hostValue = Mid$(lowerValue, hostStart, hostEnd - hostStart)
    If hostValue = "kotra.or.kr" Then
        NormalizeKotraURL = Left$(value, schemeLength) & "www." & Mid$(value, hostStart)
    Else
        NormalizeKotraURL = value
    End If
End Function
