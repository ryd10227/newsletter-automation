Attribute VB_Name = "NewsletterParser"
Option Explicit

Public Const OUTPUT_SHEET_NAME As String = "최종 업로드"

Private Const MAX_BUSINESS_ROWS As Long = 8
Private Const MAX_ACADEMY_ROWS As Long = 4
Private Const MAX_BANNER_ROWS As Long = 3
Private Const BANNER_PLACEHOLDER_TITLE As String = "해외정보팀 담당 영역"
Private Const BANNER_PLACEHOLDER_STATUS As String = ""
Private Const FORM_WARNING_SUFFIX As String = "은 보안상 사용이 금지된 설문조사 플랫폼입니다. 지원요청자에게 KOTRA 내부 설문조사 플랫폼 등으로 URL 수정을 요청하세요."
Private Const BUSINESS_INPUT_GUIDE As String = "지원요청 본문의 '1.'부터 마지막 URL까지 그대로 복사해 연번 1의 B열에 붙여넣으세요."
Private Const ACADEMY_INPUT_GUIDE As String = "아카데미 메일에서 해당 발행일(e.g. [10.20]) 줄부터 4번째 항목의 URL까지 그대로 복사해 연번 1의 B열에 붙여넣으세요. 메일 원문 페이지에서는 메일 내용이 복사되지 않으니, ""답장"" 창을 띄운 뒤 내용을 복사하세요."
Private Const BANNER_INPUT_GUIDE As String = "지원요청 본문의 '1.'부터 마지막 URL까지 그대로 복사해 연번 1의 B열에 붙여넣으세요."


' 지정된 B열 입력 셀을 한 번 클릭한 상태에서 Ctrl+V를 누르면
' 여러 줄 원문이 아래 셀로 분리되기 전에 클립보드에서 직접 읽어 등록합니다.
Public Sub UpdatePasteShortcut(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo RemoveShortcut

    If Not (ws.Parent Is ThisWorkbook) Then GoTo RemoveShortcut
    If ws.Name <> OUTPUT_SHEET_NAME Then GoTo RemoveShortcut
    If target Is Nothing Then GoTo RemoveShortcut
    If target.Cells.CountLarge <> 1 Then GoTo RemoveShortcut
    If target.Column <> 2 Then GoTo RemoveShortcut
    If SectionForInputRow(ws, target.Row) = "" Then GoTo RemoveShortcut

    Application.OnKey "^v", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PasteRawTextAtSelection"
    Application.OnKey "+{INSERT}", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PasteRawTextAtSelection"
    Exit Sub

RemoveShortcut:
    RemovePasteShortcut
End Sub


Public Sub RemovePasteShortcut()
    Application.OnKey "^v"
    Application.OnKey "+{INSERT}"
End Sub


Public Sub Auto_Open()
    On Error Resume Next
    If ActiveWorkbook Is ThisWorkbook Then
        If TypeName(ActiveSheet) = "Worksheet" And TypeName(Selection) = "Range" Then
            UpdatePasteShortcut ActiveSheet, Selection
        End If
    End If
End Sub


Public Sub Auto_Close()
    On Error Resume Next
    RemovePasteShortcut
End Sub


Public Sub PasteRawTextAtSelection()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim rawText As String
    Dim oldEvents As Boolean
    Dim eventsChanged As Boolean

    On Error GoTo Failed
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveSheet
    Set targetCell = Selection.Cells(1, 1)
    If ws.Name <> OUTPUT_SHEET_NAME Then Exit Sub
    If targetCell.Column <> 2 Then Exit Sub
    If SectionForInputRow(ws, targetCell.Row) = "" Then Exit Sub

    rawText = ReadClipboardText()
    If Len(rawText) = 0 Then Exit Sub

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    eventsChanged = True

    If LooksLikeRawText(rawText) Then
        RegisterRawTextAtRow ws, rawText, targetCell.Row
    Else
        targetCell.Value2 = rawText
    End If

    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    If eventsChanged Then Application.EnableEvents = oldEvents
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
    Err.Raise vbObjectError + 1090, , _
        "클립보드의 텍스트를 읽지 못했습니다. B열 셀을 더블클릭한 뒤 붙여넣어 주세요."
End Function


' 시트 이벤트를 사용하지 않을 때 선택한 B열 셀의 원문을 수동 등록합니다.
Public Sub 선택행_원문등록()
    Dim ws As Worksheet
    Dim sourceCell As Range
    Dim rawText As String
    Dim oldEvents As Boolean

    Set ws = ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME)
    If ActiveSheet.Name <> ws.Name Then
        MsgBox "'" & OUTPUT_SHEET_NAME & "' 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sourceCell = Selection.Cells(1, 1)
    If sourceCell.Column <> 2 Then
        MsgBox "등록할 행의 B열(행사명) 셀을 선택하세요.", vbExclamation
        Exit Sub
    End If
    rawText = CStr(sourceCell.Value2)

    If Len(Trim$(rawText)) = 0 Then
        MsgBox "선택한 셀에 원문을 먼저 붙여넣으세요.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    sourceCell.ClearContents
    RegisterRawTextAtRow ws, rawText, sourceCell.Row
    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    sourceCell.Value2 = rawText
    Application.EnableEvents = oldEvents
    MsgBox "등록하지 못했습니다." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub


' Worksheet_Change 이벤트가 호출하는 진입점입니다.
Public Sub RegisterRawTextAtRow(ByVal ws As Worksheet, ByVal rawText As String, ByVal destinationRow As Long)
    Dim title As String
    Dim businessPeriod As String
    Dim applicationPeriod As String
    Dim publicationPeriod As String
    Dim targetURL As String
    Dim targetSection As String

    targetSection = SectionForInputRow(ws, destinationRow)

    If targetSection = "" Then
        Err.Raise vbObjectError + 1000, , "사업안내형, 아카데미 또는 배너형 구역의 데이터 행에 붙여넣으세요."
    End If

    If IsAcademyRawText(rawText) Then
        If targetSection <> "academy" Then
            Err.Raise vbObjectError + 1009, , "아카데미 원문은 아카데미 연번 1의 B열에 붙여넣으세요."
        End If
        AddAcademyRecords ws, rawText, destinationRow
        Exit Sub
    End If

    title = GetFieldValue(rawText, "사업명")
    businessPeriod = GetFieldValue(rawText, "사업기간")
    applicationPeriod = GetFieldValue(rawText, "사업 신청기간")
    publicationPeriod = GetFieldValue(rawText, "희망 게재기간")
    targetURL = ExtractURL(GetFieldValue(rawText, "연결 링크 (url)"))

    If title = "" Then
        Err.Raise vbObjectError + 1001, , "'사업명:' 항목을 찾지 못했습니다."
    End If
    If targetURL = "" Then
        Err.Raise vbObjectError + 1002, , "'연결 링크 (url):' 항목에서 http 또는 https 주소를 찾지 못했습니다."
    End If

    If businessPeriod <> "" Or applicationPeriod <> "" Then
        If targetSection <> "business" Then
            Err.Raise vbObjectError + 1006, , "사업안내형 원문은 사업안내형 구역의 B열(행사명)에 붙여넣으세요."
        End If
        If businessPeriod = "" Then
            Err.Raise vbObjectError + 1003, , "사업안내형에 '사업기간:'이 없습니다."
        End If
        If applicationPeriod = "" Then
            Err.Raise vbObjectError + 1004, , "사업안내형에 '사업 신청기간:'이 없습니다."
        End If
        If publicationPeriod = "" Then
            Err.Raise vbObjectError + 1008, , "사업안내형에 '희망 게재기간:'이 없습니다. 기준 연도를 정하려면 이 항목이 필요합니다."
        End If
        AddBusinessRecord ws, destinationRow, title, businessPeriod, applicationPeriod, publicationPeriod, targetURL
    Else
        If targetSection <> "banner" Then
            Err.Raise vbObjectError + 1007, , "배너형 원문은 배너형 구역의 B열(사업명)에 붙여넣으세요."
        End If
        If publicationPeriod = "" Then
            Err.Raise vbObjectError + 1005, , "배너형에 '희망 게재기간:'이 없습니다."
        End If
        AddBannerRecord ws, destinationRow, title, publicationPeriod, targetURL
    End If
End Sub


Public Function LooksLikeRawText(ByVal value As String) As Boolean
    LooksLikeRawText = _
        ((InStr(1, value, "사업명", vbTextCompare) > 0) Or IsAcademyRawText(value)) And _
        (InStr(1, value, ":", vbBinaryCompare) > 0 Or InStr(1, value, "：", vbBinaryCompare) > 0)
End Function


Private Function IsAcademyRawText(ByVal value As String) As Boolean
    IsAcademyRawText = _
        InStr(1, value, "교육기간", vbTextCompare) > 0 And _
        InStr(1, value, "신청마감", vbTextCompare) > 0 And _
        InStr(1, value, "http", vbTextCompare) > 0
End Function


Private Sub AddAcademyRecords( _
    ByVal ws As Worksheet, _
    ByVal rawText As String, _
    ByVal destinationRow As Long)

    Dim records As Collection
    Dim record As Variant
    Dim titles(1 To MAX_ACADEMY_ROWS) As String
    Dim periodValues(1 To MAX_ACADEMY_ROWS) As String
    Dim deadlineValues(1 To MAX_ACADEMY_ROWS) As String
    Dim urlValues(1 To MAX_ACADEMY_ROWS) As String
    Dim firstStartDate As Date
    Dim startDate As Date
    Dim endDate As Date
    Dim deadlineStart As Date
    Dim deadlineDate As Date
    Dim baseYear As Long
    Dim headerRow As Long
    Dim rowNumber As Long
    Dim index As Long

    headerRow = FindHeaderRow(ws, "아카데미")
    If destinationRow <= headerRow Or destinationRow > headerRow + MAX_ACADEMY_ROWS Then
        Err.Raise vbObjectError + 1050, , "아카데미 연번 1의 B열에 붙여넣으세요."
    End If

    Set records = ParseAcademyRecords(rawText)
    If records.Count > MAX_ACADEMY_ROWS Then
        Err.Raise vbObjectError + 1051, , "아카데미 원문은 최대 " & MAX_ACADEMY_ROWS & "건까지 등록할 수 있습니다."
    End If

    record = records(1)
    ParseDateRangeWithDefaultYear CStr(record(1)), Year(Date), firstStartDate, endDate
    baseYear = AcademyBaseYear(rawText, firstStartDate)

    ' 기존 행을 바꾸기 전에 모든 원문을 먼저 검증하고 표시값을 만듭니다.
    For index = 1 To records.Count
        record = records(index)
        periodValues(index) = FormatAcademyPeriod(CStr(record(1)), baseYear, startDate)
        ParseAcademyDeadline CStr(record(2)), startDate, deadlineStart, deadlineDate

        titles(index) = CleanDisplayText(CStr(record(0)))
        deadlineValues(index) = FormatNewsletterDate(deadlineDate, baseYear)
        urlValues(index) = ExtractURL(CStr(record(3)))
        If urlValues(index) = "" Then
            Err.Raise vbObjectError + 1052, , "아카데미 " & index & "번에서 http 또는 https URL을 찾지 못했습니다."
        End If
    Next index

    ' 어느 아카데미 행에 붙여넣어도 1번부터 차례대로 다시 채웁니다.
    For index = 1 To MAX_ACADEMY_ROWS
        rowNumber = headerRow + index
        CopyTemplateFormat ws, headerRow + 1, rowNumber, 1, 5
        ResetDataRowAlertStyle ws, rowNumber, 5
        ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 5)).ClearContents

        If index <= records.Count Then
            ws.Cells(rowNumber, 1).Value = index
            ws.Cells(rowNumber, 2).Value = titles(index)
            ws.Cells(rowNumber, 3).Value = periodValues(index)
            ws.Cells(rowNumber, 4).Value = deadlineValues(index)
            ws.Cells(rowNumber, 5).Value = urlValues(index)
            ApplyFormWarning ws, rowNumber, 5, urlValues(index)
        End If
    Next index
End Sub


Private Function FormatAcademyPeriod( _
    ByVal value As String, _
    ByVal baseYear As Long, _
    ByRef firstDate As Date) As String

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

    ' 연속되지 않은 여러 교육일 표기(예: 11.20·11.27·12.4)도 모두 보존합니다.
    normalized = Replace(value, ChrW(183), "|")
    normalized = Replace(normalized, ChrW(12685), "|")
    normalized = Replace(normalized, ChrW(12539), "|")

    If InStr(1, normalized, "|", vbBinaryCompare) = 0 Then
        ParseDateRangeWithDefaultYear value, baseYear, firstDate, segmentEnd
        FormatAcademyPeriod = FormatPeriodForNewsletter(firstDate, segmentEnd, baseYear)
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
            previousDate = segmentEnd
            currentDefaultYear = Year(segmentEnd)
        End If
    Next index

    If result = "" Then
        Err.Raise vbObjectError + 1058, , "교육기간에서 날짜를 찾지 못했습니다: " & value
    End If
    FormatAcademyPeriod = result
End Function


Private Sub ParseAcademyDeadline( _
    ByVal value As String, _
    ByVal educationStart As Date, _
    ByRef startDate As Date, _
    ByRef endDate As Date)

    ParseDateRangeWithDefaultYear value, Year(educationStart), startDate, endDate

    ' 1월 교육에 대해 신청마감이 12월처럼 연도 없이 오면 전년 마감으로 봅니다.
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


Private Function ParseAcademyRecords(ByVal rawText As String) As Collection
    Dim records As New Collection
    Dim normalized As String
    Dim lines() As String
    Dim line As Variant
    Dim currentLine As String
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
        If Left$(currentLine, 1) = "+" Then
            If title <> "" Then
                AddAcademyRecordToCollection records, title, educationPeriod, deadlineText, targetURL
            End If
            title = Trim$(Mid$(currentLine, 2))
            educationPeriod = ""
            deadlineText = ""
            targetURL = ""
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

    If title <> "" Then
        AddAcademyRecordToCollection records, title, educationPeriod, deadlineText, targetURL
    End If
    If records.Count = 0 Then
        Err.Raise vbObjectError + 1053, , "'+ 제목'으로 시작하는 아카데미 항목을 찾지 못했습니다."
    End If
    Set ParseAcademyRecords = records
End Function


Private Sub AddAcademyRecordToCollection( _
    ByVal records As Collection, _
    ByVal title As String, _
    ByVal educationPeriod As String, _
    ByVal deadlineText As String, _
    ByVal targetURL As String)

    Dim itemNumber As Long
    itemNumber = records.Count + 1
    If Trim$(title) = "" Then
        Err.Raise vbObjectError + 1054, , "아카데미 " & itemNumber & "번 제목이 비어 있습니다."
    End If
    If Trim$(educationPeriod) = "" Then
        Err.Raise vbObjectError + 1055, , "아카데미 " & itemNumber & "번에 '교육기간:'이 없습니다."
    End If
    If Trim$(deadlineText) = "" Then
        Err.Raise vbObjectError + 1056, , "아카데미 " & itemNumber & "번에 '신청마감:'이 없습니다."
    End If
    If ExtractURL(targetURL) = "" Then
        Err.Raise vbObjectError + 1057, , "아카데미 " & itemNumber & "번에서 http 또는 https URL을 찾지 못했습니다."
    End If
    records.Add Array(title, educationPeriod, deadlineText, targetURL)
End Sub


Private Function AcademyFieldValue(ByVal value As String, ByVal wantedLabel As String) As String
    Dim colonPosition As Long
    Dim key As String

    colonPosition = InStr(1, value, ":", vbBinaryCompare)
    If colonPosition = 0 Then Exit Function
    key = Left$(value, colonPosition - 1)
    If NormalizeKey(key) = NormalizeKey(wantedLabel) Then
        AcademyFieldValue = Trim$(Mid$(value, colonPosition + 1))
    End If
End Function


Private Function AcademyBaseYear(ByVal rawText As String, ByVal firstEducationDate As Date) As Long
    Dim normalized As String
    Dim lines() As String
    Dim line As Variant
    Dim currentLine As String
    Dim numbers As Collection
    Dim currentYear As Long

    normalized = Replace(rawText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    lines = Split(normalized, vbLf)
    For Each line In lines
        currentLine = Trim$(CStr(line))
        If Left$(currentLine, 1) = "[" And Right$(currentLine, 1) = "]" Then
            Set numbers = ExtractNumbers(currentLine)
            If numbers.Count >= 3 Then
                If CLng(numbers(1)) >= 1900 And CLng(numbers(1)) <= 2100 Then
                    AcademyBaseYear = CLng(numbers(1))
                    Exit Function
                ElseIf CLng(numbers(1)) >= 20 And CLng(numbers(1)) <= 99 Then
                    AcademyBaseYear = 2000 + CLng(numbers(1))
                    Exit Function
                End If
            End If
        End If
    Next line

    currentYear = Year(Date)
    If Abs(Year(firstEducationDate) - currentYear) > 1 Then
        AcademyBaseYear = Year(firstEducationDate)
    Else
        AcademyBaseYear = currentYear
    End If
End Function


Private Sub AddBusinessRecord( _
    ByVal ws As Worksheet, _
    ByVal destinationRow As Long, _
    ByVal title As String, _
    ByVal periodText As String, _
    ByVal applicationText As String, _
    ByVal publicationText As String, _
    ByVal targetURL As String)

    Dim startDate As Date
    Dim endDate As Date
    Dim applicationStart As Date
    Dim deadline As Date
    Dim publicationStart As Date
    Dim publicationEnd As Date
    Dim baseYear As Long
    Dim headerRow As Long
    Dim slotNumber As Long

    ParseDateRange periodText, startDate, endDate
    ParseDateRange applicationText, applicationStart, deadline
    ParseDateRange publicationText, publicationStart, publicationEnd
    baseYear = Year(publicationStart)

    headerRow = FindHeaderRow(ws, "사업안내형")
    slotNumber = destinationRow - headerRow
    If slotNumber < 1 Or slotNumber > MAX_BUSINESS_ROWS Then
        Err.Raise vbObjectError + 1010, , "사업안내형은 머리글 아래 1~" & MAX_BUSINESS_ROWS & "번째 행에 붙여넣으세요."
    End If
    If Not RowIsAvailable(ws, destinationRow, 5) Then
        Err.Raise vbObjectError + 1011, , "선택한 사업안내형 행에 이미 데이터가 있습니다. 빈 행을 사용하세요."
    End If

    CopyTemplateFormat ws, headerRow + 1, destinationRow, 1, 5
    ResetDataRowAlertStyle ws, destinationRow, 5

    ws.Cells(destinationRow, 1).Value = slotNumber
    ws.Cells(destinationRow, 2).Value = CleanDisplayText(title)
    ws.Cells(destinationRow, 3).Value = FormatPeriodForNewsletter(startDate, endDate, baseYear)
    ws.Cells(destinationRow, 4).Value = FormatNewsletterDate(deadline, baseYear)
    ws.Cells(destinationRow, 5).Value = targetURL
    ApplyFormWarning ws, destinationRow, 5, targetURL
End Sub


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


Private Sub AddBannerRecord( _
    ByVal ws As Worksheet, _
    ByVal destinationRow As Long, _
    ByVal title As String, _
    ByVal publicationText As String, _
    ByVal targetURL As String)

    Dim publicationStart As Date
    Dim publicationEnd As Date
    Dim headerRow As Long
    Dim slotNumber As Long

    ParseDateRange publicationText, publicationStart, publicationEnd

    headerRow = FindHeaderRow(ws, "배너형")
    slotNumber = destinationRow - headerRow
    If slotNumber < 1 Or slotNumber > MAX_BANNER_ROWS Then
        Err.Raise vbObjectError + 1020, , "배너형은 머리글 아래 1~" & MAX_BANNER_ROWS & "번째 행에 붙여넣으세요."
    End If
    If Not RowIsAvailable(ws, destinationRow, 4) And Not IsBannerPlaceholderRow(ws, destinationRow) Then
        Err.Raise vbObjectError + 1021, , "선택한 배너형 행에 이미 데이터가 있습니다. 빈 행을 사용하세요."
    End If

    CopyTemplateFormat ws, headerRow + 1, destinationRow, 1, 4
    ResetDataRowAlertStyle ws, destinationRow, 4

    ws.Cells(destinationRow, 1).Value = slotNumber
    ws.Cells(destinationRow, 2).Value = CleanDisplayText(title)
    ws.Cells(destinationRow, 3).Value = Format$(publicationEnd, "yyyy-mm-dd") & "까지 게재"
    ws.Cells(destinationRow, 4).Value = targetURL
    ApplyFormWarning ws, destinationRow, 4, targetURL
End Sub


' 새 회차를 시작할 때만 사용합니다. 형식은 유지하고 사업안내/아카데미/배너 데이터를 지웁니다.
Public Sub 새회차_사업배너_초기화()
    Dim ws As Worksheet
    Dim answer As VbMsgBoxResult
    Dim businessHeader As Long
    Dim academyHeader As Long
    Dim bannerHeader As Long
    Dim businessRange As Range
    Dim academyRange As Range
    Dim bannerRange As Range

    answer = MsgBox("사업안내형, 아카데미, 배너형의 기존 데이터가 지워집니다." & vbCrLf & "계속할까요?", vbYesNo + vbExclamation)
    If answer <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME)

    businessHeader = FindHeaderRow(ws, "사업안내형")
    Set businessRange = ws.Range(ws.Cells(businessHeader + 1, 1), ws.Cells(businessHeader + MAX_BUSINESS_ROWS, 5))
    businessRange.ClearContents
    ResetDataRangeAlertStyle businessRange
    ws.Rows((businessHeader + 1) & ":" & (businessHeader + MAX_BUSINESS_ROWS)).RowHeight = 34
    SetSectionNumbers ws, businessHeader, MAX_BUSINESS_ROWS
    SetInputGuideRow ws, businessHeader - 1, BUSINESS_INPUT_GUIDE

    academyHeader = FindHeaderRow(ws, "아카데미")
    Set academyRange = ws.Range(ws.Cells(academyHeader + 1, 1), ws.Cells(academyHeader + MAX_ACADEMY_ROWS, 5))
    academyRange.ClearContents
    ResetDataRangeAlertStyle academyRange
    ws.Rows((academyHeader + 1) & ":" & (academyHeader + MAX_ACADEMY_ROWS)).RowHeight = 34
    SetSectionNumbers ws, academyHeader, MAX_ACADEMY_ROWS
    SetInputGuideRow ws, academyHeader - 1, ACADEMY_INPUT_GUIDE

    bannerHeader = FindHeaderRow(ws, "배너형")
    Set bannerRange = ws.Range(ws.Cells(bannerHeader + 1, 1), ws.Cells(bannerHeader + MAX_BANNER_ROWS, 4))
    bannerRange.ClearContents
    ResetDataRangeAlertStyle bannerRange
    ws.Range(ws.Cells(bannerHeader, 5), ws.Cells(bannerHeader + MAX_BANNER_ROWS, 5)).Clear
    ws.Rows((bannerHeader + 1) & ":" & (bannerHeader + MAX_BANNER_ROWS)).RowHeight = 34
    SetSectionNumbers ws, bannerHeader, MAX_BANNER_ROWS
    SetInputGuideRow ws, bannerHeader - 1, BANNER_INPUT_GUIDE
    SetBannerPlaceholder ws, bannerHeader + 2

    MsgBox "사업안내형, 아카데미, 배너형 데이터를 비웠습니다.", vbInformation
End Sub


Private Sub SetSectionNumbers(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal itemCount As Long)
    Dim index As Long
    For index = 1 To itemCount
        ws.Cells(headerRow + index, 1).Value = index
    Next index
End Sub


Private Sub SetInputGuideRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal guideText As String)
    Dim guideRange As Range

    Set guideRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 5))
    guideRange.UnMerge
    guideRange.ClearContents
    guideRange.Merge
    guideRange.Interior.Color = RGB(255, 242, 204)
    guideRange.Font.Color = RGB(127, 96, 0)
    guideRange.Font.Italic = True
    ws.Cells(rowNumber, 1).Value = guideText
    ws.Cells(rowNumber, 1).WrapText = True
    ws.Cells(rowNumber, 1).HorizontalAlignment = xlLeft
    ws.Cells(rowNumber, 1).VerticalAlignment = xlCenter
    If guideText = ACADEMY_INPUT_GUIDE Then
        ws.Rows(rowNumber).RowHeight = 52
    Else
        ws.Rows(rowNumber).RowHeight = 38
    End If
End Sub


Private Sub ApplyFormWarning(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal urlColumn As Long, ByVal targetURL As String)
    Dim provider As String
    Dim rowRange As Range

    provider = FormProviderName(targetURL)
    If provider = "" Then Exit Sub

    Set rowRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, urlColumn))
    rowRange.Interior.Color = RGB(255, 199, 206)
    rowRange.Font.Color = RGB(156, 0, 6)
    ws.Cells(rowNumber, urlColumn).Value = targetURL & vbLf & provider & FORM_WARNING_SUFFIX
    ws.Cells(rowNumber, urlColumn).WrapText = True
    ws.Cells(rowNumber, urlColumn).Font.Bold = True
    ws.Rows(rowNumber).AutoFit
    If ws.Rows(rowNumber).RowHeight < 60 Then ws.Rows(rowNumber).RowHeight = 60
End Sub


Private Sub ResetDataRowAlertStyle(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal lastColumn As Long)
    ResetDataRangeAlertStyle ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, lastColumn))
    ws.Rows(rowNumber).RowHeight = 34
End Sub


Private Sub ResetDataRangeAlertStyle(ByVal targetRange As Range)
    targetRange.Interior.Pattern = xlNone
    targetRange.Font.Color = RGB(31, 41, 55)
    targetRange.Font.Bold = False
    targetRange.Font.Italic = False
End Sub


Private Function FormProviderName(ByVal value As String) As String
    Dim hostValue As String
    Dim lowerValue As String

    hostValue = URLHost(value)
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
        Exit Function
    End If
    If Len(hostValue) > Len(domainValue) Then
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


Private Function IsBannerPlaceholderRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    IsBannerPlaceholderRow = _
        CLng(Val(CStr(ws.Cells(rowNumber, 1).Value2))) = 2 And _
        CStr(ws.Cells(rowNumber, 3).Value2) = BANNER_PLACEHOLDER_STATUS And _
        Len(Trim$(CStr(ws.Cells(rowNumber, 4).Value2))) = 0
End Function


Private Function IsInputGuideRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim guideText As String

    guideText = CStr(ws.Cells(rowNumber, 2).Value2)
    If guideText = "" Then guideText = CStr(ws.Cells(rowNumber, 1).Value2)
    IsInputGuideRow = _
        (guideText = BUSINESS_INPUT_GUIDE Or guideText = ACADEMY_INPUT_GUIDE Or guideText = BANNER_INPUT_GUIDE) And _
        Len(Trim$(CStr(ws.Cells(rowNumber, 3).Value2))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNumber, 4).Value2))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNumber, 5).Value2))) = 0
End Function


Private Sub SetBannerPlaceholder(ByVal ws As Worksheet, ByVal rowNumber As Long)
    Dim placeholderRange As Range

    CopyTemplateFormat ws, rowNumber - 1, rowNumber, 1, 4
    ws.Cells(rowNumber, 1).Value = 2
    ws.Cells(rowNumber, 2).Value = BANNER_PLACEHOLDER_TITLE
    ws.Cells(rowNumber, 3).ClearContents
    ws.Cells(rowNumber, 4).ClearContents

    Set placeholderRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 4))
    placeholderRange.Interior.Color = RGB(229, 231, 235)
    placeholderRange.Font.Color = RGB(107, 114, 128)
    placeholderRange.Font.Italic = True
End Sub


Private Function GetFieldValue(ByVal rawText As String, ByVal wantedLabel As String) As String
    Dim normalized As String
    Dim lines() As String
    Dim line As Variant
    Dim colonPosition As Long
    Dim key As String
    Dim value As String

    normalized = Replace(rawText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, "：", ":")
    lines = Split(normalized, vbLf)

    For Each line In lines
        colonPosition = InStr(1, CStr(line), ":", vbBinaryCompare)
        If colonPosition > 0 Then
            key = Left$(CStr(line), colonPosition - 1)
            If NormalizeKey(key) = NormalizeKey(wantedLabel) Then
                value = Mid$(CStr(line), colonPosition + 1)
                GetFieldValue = Trim$(value)
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


' 별도 참조 설정 없이 Windows Excel에서 동작하도록 VBScript.RegExp를 사용하지 않습니다.
Private Sub ParseDateRange(ByVal value As String, ByRef startDate As Date, ByRef endDate As Date)
    ParseFlexibleDateRange value, Year(Date), startDate, endDate
End Sub


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
        ' 12.31.~1.1.처럼 종료 연도가 생략된 연말/연초 범위입니다.
        If endDate < startDate Then
            endDate = StrictDate(endYear + 1, endMonth, endDay, value)
        End If
    Else
        ' 26.10.11~13처럼 종료 월까지 생략하면 시작 월을 이어받습니다.
        endMonth = startMonth
        endDay = CLng(numbers(nextIndex))
        endDate = StrictDate(endYear, endMonth, endDay, value)
        ' 10.31~2처럼 종료 일이 더 작으면 다음 달로 해석합니다.
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
    IsYearToken = _
        (value >= 1900 And value <= 2100) Or _
        (value >= 20 And value <= 99)
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
        InStr(value, "~") > 0 Or _
        InStr(value, "～") > 0 Or _
        InStr(value, ChrW(8764)) > 0 Or _
        InStr(value, ChrW(12316)) > 0 Or _
        InStr(value, "부터") > 0 Or _
        InStr(value, "까지") > 0 Or _
        InStr(value, "에서") > 0 Or _
        InStr(value, "-") > 0 Or _
        InStr(value, ChrW(8722)) > 0 Or _
        InStr(value, ChrW(8211)) > 0 Or _
        InStr(value, ChrW(8212)) > 0 Or _
        InStr(value, ChrW(8213)) > 0 Or _
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
            Or character = ")" Or character = "]" Or character = ">" Then
            Exit For
        End If
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


Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal sectionName As String) As Long
    Dim sectionRow As Long
    Dim rowNumber As Long
    Dim lastRow As Long

    sectionRow = FindSectionRow(ws, sectionName)
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 10
    For rowNumber = sectionRow + 1 To lastRow
        If NormalizeKey(CStr(ws.Cells(rowNumber, 1).Value2)) = NormalizeKey("연번") Then
            FindHeaderRow = rowNumber
            Exit Function
        End If
        If IsSectionLabel(CStr(ws.Cells(rowNumber, 1).Value2)) Then Exit For
    Next rowNumber
    Err.Raise vbObjectError + 1040, , sectionName & " 구역의 '연번' 머리글을 찾지 못했습니다."
End Function


Private Function FindSectionRow(ByVal ws As Worksheet, ByVal sectionName As String) As Long
    Dim rowNumber As Long
    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 10
    For rowNumber = 1 To lastRow
        If NormalizeKey(CStr(ws.Cells(rowNumber, 1).Value2)) = NormalizeKey(sectionName) Then
            FindSectionRow = rowNumber
            Exit Function
        End If
    Next rowNumber
    Err.Raise vbObjectError + 1041, , "첫 시트에서 '" & sectionName & "' 구역을 찾지 못했습니다."
End Function


Private Function FindNextSectionRow(ByVal ws As Worksheet, ByVal afterRow As Long) As Long
    Dim rowNumber As Long
    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 20
    For rowNumber = afterRow + 1 To lastRow
        If IsSectionLabel(CStr(ws.Cells(rowNumber, 1).Value2)) Then
            FindNextSectionRow = rowNumber
            Exit Function
        End If
    Next rowNumber
    FindNextSectionRow = lastRow
End Function


Private Function IsSectionLabel(ByVal value As String) As Boolean
    Dim normalized As String
    normalized = NormalizeKey(value)
    IsSectionLabel = _
        normalized = NormalizeKey("사업안내형") Or _
        normalized = NormalizeKey("배너형") Or _
        normalized = NormalizeKey("아카데미")
End Function


Public Function SectionForInputRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As String
    Dim headerRow As Long
    Dim nextSectionRow As Long

    headerRow = FindHeaderRow(ws, "사업안내형")
    nextSectionRow = FindNextSectionRow(ws, headerRow)
    If rowNumber > headerRow And rowNumber <= headerRow + MAX_BUSINESS_ROWS And rowNumber < nextSectionRow Then
        SectionForInputRow = "business"
        Exit Function
    End If

    headerRow = FindHeaderRow(ws, "아카데미")
    nextSectionRow = FindNextSectionRow(ws, headerRow)
    If rowNumber > headerRow And rowNumber <= headerRow + MAX_ACADEMY_ROWS And rowNumber < nextSectionRow Then
        SectionForInputRow = "academy"
        Exit Function
    End If

    headerRow = FindHeaderRow(ws, "배너형")
    nextSectionRow = FindNextSectionRow(ws, headerRow)
    If rowNumber > headerRow And rowNumber <= headerRow + MAX_BANNER_ROWS And rowNumber < nextSectionRow Then
        SectionForInputRow = "banner"
    End If
End Function


Private Function RowIsAvailable(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal lastColumn As Long) As Boolean
    Dim columnNumber As Long
    If IsInputGuideRow(ws, rowNumber) Then
        RowIsAvailable = True
        Exit Function
    End If
    ' A열의 기존 연번은 허용하고, 실제 데이터가 들어가는 B열 이후만 확인합니다.
    For columnNumber = 2 To lastColumn
        If Len(Trim$(CStr(ws.Cells(rowNumber, columnNumber).Value2))) > 0 Then Exit Function
    Next columnNumber
    RowIsAvailable = True
End Function


Private Sub CopyTemplateFormat( _
    ByVal ws As Worksheet, _
    ByVal templateRow As Long, _
    ByVal destinationRow As Long, _
    ByVal firstColumn As Long, _
    ByVal lastColumn As Long)

    ws.Range(ws.Cells(templateRow, firstColumn), ws.Cells(templateRow, lastColumn)).Copy
    ws.Range(ws.Cells(destinationRow, firstColumn), ws.Cells(destinationRow, lastColumn)).PasteSpecial xlPasteFormats
    ws.Rows(destinationRow).RowHeight = ws.Rows(templateRow).RowHeight
    Application.CutCopyMode = False
End Sub
