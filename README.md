Sub 데이터무작위섞기()
    '선택셀 무작위 섞기
    Dim varData As Variant, varTemp As Variant
    Dim lngI As Long, lngJ As Long, lngK As Long, lngL As Long
    Dim lngRow As Long, lngCol As Long
    
    On Error GoTo ErrorHandler
    
    varData = Selection.Value
    lngRow = UBound(varData, 1)
    lngCol = UBound(varData, 2)
    
    For lngJ = 1 To lngCol
      For lngI = 1 To lngRow
        lngK = lngK + 1
        lngL = Int(Rnd * (lngRow * lngCol - lngK + 1)) + lngK
        varTemp = varData(lngI, lngJ)
        varData(lngI, lngJ) = varData((lngL - 1) Mod _
          lngRow + 1, Int((lngL - 1) / lngRow) + 1)
        varData((lngL - 1) Mod lngRow + 1, _
          Int((lngL - 1) / lngRow) + 1) = varTemp
      Next lngI
    Next lngJ
    
    Selection.Value = varData
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 9  ' 첨자 범위 오류
            MsgBox "데이터 범위를 다시 확인해 주세요.", vbCritical
        Case 13 ' 형식 불일치
            MsgBox "데이터 범위를 다시 확인해 주세요.", vbCritical
        Case Else
            MsgBox "오류: " & Err.Description, vbCritical
    End Select
End Sub


Sub 나열하여_붙여넣기()

'좌석배치도 시트의 다중선택셀 한열로 나열하여 붙여넣기

    Dim selectedRange As Range
    Set selectedRange = Selection
    
    ' "참석자좌석명단" 시트의 G열 4번째 줄부터 나열하여 붙여넣기
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("참석자좌석명단")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row + 1
    
    Dim cell As Range
    
    For Each cell In selectedRange
        ' 빈 셀은 건너뛰기
        If Not IsEmpty(cell) Then
            ws.Cells(lastRow, "G").Value = cell.Value
            lastRow = lastRow + 1
        End If
    Next cell
    Sheets("참석자좌석명단").Activate


End Sub


Sub 삭제()
'
' 입력 데이터 초기화

    Sheets("raw_data").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Sheets("접수확인명단").Select
    Range("B4:D203").Select
    Selection.ClearContents
    Range("G4:N203").Select
    Selection.ClearContents
    
    Sheets("출석명단").Select
    Range("B4:D203").Select
    Selection.ClearContents
    
    Sheets("참석자좌석명단").Select
    Range("B4:C203").Select
    Selection.ClearContents
    Range("G4:G203").Select
    Selection.ClearContents
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "초성"
    Range("B4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[4]"
    Range("B4").Select
    Selection.AutoFill Destination:=Range("B4:B203"), Type:=xlFillDefault
    Range("B4:B203").Select
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
End Sub



Sub SetPrintAreaAndPreview()

'참석자좌석명단 시트 인쇄설정 및 미리보기

    Dim lastRow As Long
    Dim ws As Worksheet
    
    ' 원하는 시트를 지정
    Set ws = ThisWorkbook.Sheets("참석자좌석명단")
    
    ' "B"열의 마지막 데이터가 있는 행을 찾기
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' 인쇄 영역을 설정 (A, B, C, G 열을 포함).
    ws.PageSetup.PrintArea = ws.Range("A3:G" & lastRow).Address
    
    ' 미리보기 열기
    ws.PrintPreview
End Sub



Sub 접수확인명단인쇄미리보기()

'접수확인명단 시트 인쇄설정 및 미리보기

    Dim lastRow As Long
    Dim ws As Worksheet
    
    ' 원하는 시트를 지정하기
    Set ws = ThisWorkbook.Sheets("접수확인명단")
    
    ' "G"열의 마지막 데이터가 있는 행을 찾기
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' 인쇄 영역을 설정 (A부터 N열을 포함).
    ws.PageSetup.PrintArea = ws.Range("A3:M" & lastRow).Address
    
    ' 미리보기 열기
    ws.PrintPreview
End Sub



Sub 좌석배치도인쇄미리보기()

  ActiveSheet.PrintPreview
  
End Sub


Sub 출석명단미리보기()

'출석명단 시트 인쇄설정 및 미리보기

    Dim lastRow As Long
    Dim ws As Worksheet
    
    ' 원하는 시트를 지정하기
    Set ws = ThisWorkbook.Sheets("출석명단")
    
    ' "B"열의 마지막 데이터가 있는 행을 찾기
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' 인쇄 영역을 설정 (A부터 D열을 포함).
    ws.PageSetup.PrintArea = ws.Range("A3:D" & lastRow).Address
    
    ' 미리보기 열기
    ws.PrintPreview
End Sub






Sub CreateSeatLayout_Final()

'현재 시트에서 좌석배치 설정 및 표 그리기

    Dim ws As Worksheet
    Dim colCount As Long, rowCount As Long
    Dim i As Long, j As Long
    Dim startCell As Range
    Dim response As Variant
    Dim rowLabel As String
    Dim koreanLabels As Variant
    Dim useKorean As VbMsgBoxResult
    Dim doSectioning As VbMsgBoxResult
    Dim doZigZag As VbMsgBoxResult
    Dim sectionSize As Long: sectionSize = 0
    Dim offsetX As Long
    Dim colOffset As Long

    ' ▼ 현재 활성 시트 사용
    Set ws = ActiveSheet
    ws.Range("A3:AZ500").ClearContents

    ' ▼ 행 라벨 방식 선택
    useKorean = MsgBox("좌석 행 이름 방식을 선택하세요." & vbCrLf & _
                       "[예] 한글 (가, 나, 다...)" & vbCrLf & _
                       "[아니오] 알파벳 (A, B, C...)" & vbCrLf & _
                       "[취소] 작업 종료", _
                       vbYesNoCancel + vbQuestion, "행 라벨 선택")
    If useKorean = vbCancel Then Exit Sub

    ' ▼ 분단 나누기 여부
    doSectioning = MsgBox("좌석을 분단 단위로 나누시겠습니까?" & vbCrLf & _
                          "(예: 4열마다 1열 공백 삽입)", _
                          vbYesNoCancel + vbQuestion, "분단 설정")
    If doSectioning = vbCancel Then Exit Sub
    If doSectioning = vbYes Then
        response = InputBox("몇 열마다 분단을 나눌까요? (예: 4)", "분단 칸 수", 4)
        If Not IsNumeric(response) Or Trim(response) = "" Then Exit Sub
        sectionSize = CLng(response)
        If sectionSize <= 0 Then
            MsgBox "0보다 큰 숫자를 입력해야 합니다.", vbExclamation
            Exit Sub
        End If
    End If

    ' ▼ 엇갈림 배치 여부
    doZigZag = MsgBox("짝수 번째 행을 한 칸 오른쪽으로 밀어 엇갈리게 배치할까요?" & vbCrLf & _
                      "(예: 계단형 좌석)", _
                      vbYesNo + vbQuestion, "엇갈림 설정")

    ' ▼ 가로 좌석 수 입력
    response = InputBox("가로 좌석 수를 입력하세요 (최대 52)", "열 개수", 10)
    If Not IsNumeric(response) Or Trim(response) = "" Then Exit Sub
    colCount = CLng(response)
    If colCount > 52 Then colCount = 52

    ' ▼ 세로 좌석 수 입력
    response = InputBox("세로 좌석 수를 입력하세요 (최대 50)", "행 개수", 5)
    If Not IsNumeric(response) Or Trim(response) = "" Then Exit Sub
    rowCount = CLng(response)
    If rowCount > 50 Then rowCount = 50

    ' ▼ 시작 셀
    Set startCell = ws.Range("B7")

    ' ▼ 한글 라벨 배열
    koreanLabels = Array("가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하", _
                         "거", "너", "더", "러", "머", "버", "서", "어", "저", "처", "커", "터", "퍼", "허")

    ' ▼ 좌석 배치
    For i = 1 To rowCount
        offsetX = 0

        For j = 1 To colCount
            ' 분단 공백 계산
            If sectionSize > 0 Then
                If j > 1 And ((j - 1) Mod sectionSize = 0) Then
                    offsetX = offsetX + 1
                End If
            End If

            ' 엇갈림 적용
            colOffset = j - 1 + offsetX
            If doZigZag = vbYes And i Mod 2 = 0 Then
                colOffset = colOffset + 1
            End If

            ' 행 라벨 결정
            If useKorean = vbYes Then
                If i <= UBound(koreanLabels) + 1 Then
                    rowLabel = koreanLabels(i - 1)
                Else
                    rowLabel = "?"
                End If
            Else
                rowLabel = Chr(64 + i)
            End If

            ' 좌석명 입력 (한 칸 아래는 비움)
            ws.Cells(startCell.Row + (i - 1) * 2, startCell.Column + colOffset).Value = rowLabel & " " & j
        Next j
    Next i

    ' ▼ 완료 메시지
    MsgBox "좌석 배치 완료!" & vbCrLf & _
           "행: " & rowCount & ", 열: " & colCount & vbCrLf & _
           "행 라벨: " & IIf(useKorean = vbYes, "한글", "알파벳") & vbCrLf & _
           "분단 구분: " & IIf(sectionSize > 0, "O", "X") & ", 엇갈림: " & IIf(doZigZag = vbYes, "O", "X"), vbInformation
End Sub








Sub ApplyAllFormatting()

    Dim ws As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim targetRange As Range
    Dim formulaText As String
    Dim formulaCondition As String
    Dim refSheet As Worksheet

    ' ▼ 참가자좌석명단 시트 G4:G205 초기화
    On Error Resume Next
    Set refSheet = ThisWorkbook.Sheets("참석자좌석명단")
    If Not refSheet Is Nothing Then
        refSheet.Range("G4:G205").ClearContents
    Else
        MsgBox "'참석자좌석명단' 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' ▼ 현재 시트 포맷 적용
    Set ws = ActiveSheet
    Set rng = ws.Range("A3:AZ100")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 기존 조건부 서식 제거
    rng.FormatConditions.Delete

    For Each cell In rng
        If cell.Value <> "" Then
            ' 조건부 서식: 아래 셀에 값이 있는 경우 현재 셀 노란색
            If cell.Row < ws.Rows.Count Then
                formulaCondition = "=AND(" & cell.Address(False, False) & "<>""""," & _
                                            cell.Offset(1, 0).Address(False, False) & "<>"""")"
                cell.FormatConditions.Add Type:=xlExpression, Formula1:=formulaCondition
                cell.FormatConditions(cell.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)

                ' 아래 셀에 INDEX+MATCH 함수 삽입
                formulaText = "=IFERROR(INDEX(참석자좌석명단!$A:$G," & _
                               "MATCH(" & cell.Address(False, False) & ",참석자좌석명단!$G:$G,0),2),"""")"
                cell.Offset(1, 0).formula = formulaText
            End If

            ' 테두리 적용: 현재 셀과 아래 셀
            Set targetRange = Union(cell, cell.Offset(1, 0))
            With targetRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End If
    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "작업 완료되었습니다!", vbInformation
End Sub








Sub PrintPreview_OnePage_SafeBlock()

'기본배치도시트 인쇄영역 자동설정 및 미리보기
'기본배치도 시트 복사 후 사용할 수 있게 활성시트로 변경


    Dim ws As Worksheet
    Dim dataRange As Range
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim cell As Range
    Dim found As Boolean

    Set ws = ActiveSheet
    Set dataRange = ws.Range("A3:AZ500")

    firstRow = 0
    lastRow = 0
    firstCol = 0
    lastCol = 0
    found = False

    ' 모든 셀을 검사해서 값이 있는 경우 좌표 기록 (공백 무시)
    For Each cell In dataRange
        If Trim(cell.Value) <> "" Then
            r = cell.Row
            c = cell.Column

            If Not found Then
                firstRow = r
                lastRow = r
                firstCol = c
                lastCol = c
                found = True
            Else
                If r < firstRow Then firstRow = r
                If r > lastRow Then lastRow = r
                If c < firstCol Then firstCol = c
                If c > lastCol Then lastCol = c
            End If
        End If
    Next cell

    If Not found Then
        MsgBox "A3:AZ500 범위에 인쇄할 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 인쇄 설정: 직사각형 범위 지정
    With ws
        .PageSetup.PrintArea = .Range(.Cells(firstRow, firstCol), .Cells(lastRow, lastCol)).Address
        With .PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .CenterHorizontally = True
            .CenterVertically = True
            .PrintGridlines = True
        End With
        .PrintPreview
    End With
End Sub






Sub AssignSeats_WithBiasProperly()
    Dim wsLayout As Worksheet
    Dim wsList As Worksheet
    Dim seatList As Collection
    Dim seatRange As Range
    Dim validSeats() As String
    Dim selectedSeats() As String
    Dim attendeeCount As Long
    Dim totalSeats As Long
    Dim i As Long
    Dim biasFactor As Double
    Dim response As Variant

    Dim allowedKor() As String
    allowedKor = Split("가,나,다,라,마,바,사,아,자,차,카,타,파,하,거,너,더,러,머,버,서,어,저,처,커,터,퍼,허", ",")

    Set wsLayout = ActiveSheet  ' 현재 시트를 사용
    Set wsList = ThisWorkbook.Sheets("참석자좌석명단")
    Set seatList = New Collection

    ' ▼ 유효한 좌석 수집
    For Each seatRange In wsLayout.Range("A3:AZ500")
        Dim txt As String
        txt = Trim(seatRange.Value)
        If txt <> "" Then
            Dim parts() As String
            parts = Split(txt, " ")
            If UBound(parts) = 1 Then
                If IsNumeric(parts(1)) Then
                    If IsInArray(parts(0), allowedKor) Or _
                       (Len(parts(0)) = 1 And Asc(UCase(parts(0))) >= 65 And Asc(UCase(parts(0))) <= 90) Then
                        seatList.Add txt
                    End If
                End If
            End If
        End If
    Next seatRange

    totalSeats = seatList.Count
    If totalSeats = 0 Then
        MsgBox "유효한 좌석명이 없습니다.", vbExclamation
        Exit Sub
    End If

    ' ▼ 참석자 수 계산
    attendeeCount = WorksheetFunction.CountA(wsList.Range("B4:B203"))
    If attendeeCount = 0 Then
        MsgBox "참석자 명단이 없습니다.", vbExclamation
        Exit Sub
    End If

    If totalSeats < attendeeCount Then
        MsgBox "좌석 수가 참석자 수보다 부족합니다.", vbCritical
        Exit Sub
    End If

    ' ▼ 정렬된 좌석 배열
    ReDim validSeats(1 To totalSeats)
    For i = 1 To totalSeats
        validSeats(i) = seatList(i)
    Next i
    SortSeatArray validSeats

    ' ▼ 사용자 입력으로 편향도 설정
    response = InputBox("좌석 편향도 설정 (0 = 앞쪽 집중, 1 = 균등, 2 = 뒤쪽 집중)", "좌석 편향도", 1)
    If Not IsNumeric(response) Then Exit Sub
    biasFactor = CDbl(response)
    If biasFactor < 0 Then biasFactor = 0
    If biasFactor > 3 Then biasFactor = 3

    ' ▼ 참석자 수 만큼 우선순위 좌석 추출
    ReDim selectedSeats(1 To attendeeCount)
    Dim pos As Double
    Dim usedIndex() As Boolean
    ReDim usedIndex(1 To totalSeats)

    For i = 1 To attendeeCount
        If biasFactor = 0 Then
            pos = i / attendeeCount
        Else
            pos = (i / attendeeCount) ^ (1 / biasFactor)
        End If
        Dim idx As Long
        idx = WorksheetFunction.RoundUp(pos * totalSeats, 0)
        If idx > totalSeats Then idx = totalSeats

        ' 중복 방지
        Do While usedIndex(idx) And idx < totalSeats
            idx = idx + 1
        Loop
        If usedIndex(idx) Then
            idx = 1
            Do While usedIndex(idx) And idx < totalSeats
                idx = idx + 1
            Loop
        End If
        If Not usedIndex(idx) Then
            selectedSeats(i) = validSeats(idx)
            usedIndex(idx) = True
        End If
    Next i

    ' ▼ 결과 반영
    wsList.Range("G4:G1000").ClearContents
    For i = 1 To attendeeCount
        wsList.Cells(3 + i, "G").Value = selectedSeats(i)
    Next i

    ' ▼ 완료 메시지
    MsgBox "좌석 배정 완료!" & vbCrLf & _
           "전체 좌석 수: " & totalSeats & vbCrLf & _
           "배정된 참석자 수: " & attendeeCount & vbCrLf & _
           "좌석 편향도: " & biasFactor, vbInformation
End Sub

' ▼ 좌석 정렬 함수 (한글 순서 반영)
Sub SortSeatArray(arr() As String)
    Dim i As Long, j As Long, temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CompareSeatName(arr(i), arr(j)) > 0 Then
                temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        Next j
    Next i
End Sub

Function CompareSeatName(a As String, b As String) As Long
    Dim aPart() As String, bPart() As String
    Dim labelOrder As Variant
    Dim iA As Long, iB As Long
    Dim foundA As Boolean, foundB As Boolean

    aPart = Split(a, " ")
    bPart = Split(b, " ")

    labelOrder = Split("가,나,다,라,마,바,사,아,자,차,카,타,파,하,거,너,더,러,머,버,서,어,저,처,커,터,퍼,허", ",")

    foundA = False: foundB = False
    For iA = 0 To UBound(labelOrder)
        If labelOrder(iA) = aPart(0) Then foundA = True: Exit For
    Next iA
    For iB = 0 To UBound(labelOrder)
        If labelOrder(iB) = bPart(0) Then foundB = True: Exit For
    Next iB

    If foundA And foundB Then
        If iA = iB Then
            CompareSeatName = CLng(aPart(1)) - CLng(bPart(1))
        Else
            CompareSeatName = iA - iB
        End If
    Else
        ' 한글 라벨에 없는 경우 일반 텍스트 비교
        If aPart(0) = bPart(0) Then
            CompareSeatName = CLng(aPart(1)) - CLng(bPart(1))
        Else
            CompareSeatName = StrComp(aPart(0), bPart(0), vbTextCompare)
        End If
    End If
End Function

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then IsInArray = True: Exit Function
    Next i
    IsInArray = False
End Function


