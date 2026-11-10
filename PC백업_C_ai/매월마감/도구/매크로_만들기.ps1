# 매월마감_실행.xlsm (버튼 파일) 다시 만들기 — 버튼·매크로를 고칠 때만 실행
$ErrorActionPreference = 'Stop'
$out = 'C:\ai\매월마감\매월마감_실행.xlsm'
$vba = @'
Option Explicit
Private Const ROOT_DIR As String = "C:\ai\매월마감"
Private Const PS_FILE As String = "C:\ai\매월마감\도구\수출마감_자동.ps1"
Private Const PDF_PS As String = "C:\ai\매월마감\도구\수출서류_PDF합치기.ps1"
Private Const PDF_ROOT As String = "C:\ai\매월마감\해외pdf"
Private Const LOG_ROW As Long = 30

Private Function MonthText() As String
    MonthText = Trim(CStr(ThisWorkbook.Worksheets(1).Range("C4").Value))
End Function

Private Function MonthFolder(ByVal ym As String) As String
    MonthFolder = ROOT_DIR & "\" & ym
End Function

Private Function CheckMonth(ByVal ym As String) As Boolean
    If Len(ym) = 6 And IsNumeric(ym) Then
        If Val(Right(ym, 2)) >= 1 And Val(Right(ym, 2)) <= 12 Then CheckMonth = True: Exit Function
    End If
    MsgBox "마감월을 6자리 숫자로 넣어 주세요. 예: 202611", vbExclamation, "수출 마감"
End Function

Public Sub Open_Folder()
    Dim ym As String: ym = MonthText()
    If Not CheckMonth(ym) Then Exit Sub
    Dim f As String: f = MonthFolder(ym)
    If Dir(f, vbDirectory) = "" Then MkDir f
    Shell "explorer.exe """ & f & """", vbNormalFocus
End Sub

Public Sub Run_Check()
    RunScript True
End Sub

Public Sub Run_Build()
    RunScript False
End Sub

Public Sub Open_PdfFolder()
    Dim ym As String: ym = MonthText()
    If Not CheckMonth(ym) Then Exit Sub
    Dim f As String: f = PDF_ROOT & "\" & ym & "해외pdf"
    If Dir(PDF_ROOT, vbDirectory) = "" Then MkDir PDF_ROOT
    If Dir(f, vbDirectory) = "" Then MkDir f
    Shell "explorer.exe """ & f & """", vbNormalFocus
End Sub

Public Sub Run_Pdf()
    Dim ym As String: ym = MonthText()
    If Not CheckMonth(ym) Then Exit Sub
    Dim f As String: f = MonthFolder(ym)
    If Dir(f & "\*월반영.xlsx") = "" Then MsgBox "이 달 결과 파일이 없습니다. '③ 결과 만들기'를 먼저 눌러 주세요.", vbExclamation, "수출 마감": Exit Sub
    If Dir(PDF_PS) = "" Then MsgBox "스크립트가 없습니다: " & PDF_PS, vbCritical, "수출 마감": Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A" & LOG_ROW & ":A2000").EntireRow.Delete
    ws.Range("C6").Value = "PDF 합치는 중... (검은 창이 닫힐 때까지 기다려 주세요, 1~3분)"
    DoEvents
    Dim rc As Long: rc = CreateObject("WScript.Shell").Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & PDF_PS & """ -YM " & ym, 1, True)
    ShowLog f, rc, 2
End Sub

Public Sub Open_Result()
    Dim ym As String: ym = MonthText()
    If Not CheckMonth(ym) Then Exit Sub
    Dim f As String: f = Dir(MonthFolder(ym) & "\*월반영.xlsx")
    If f = "" Then MsgBox "아직 결과 파일이 없습니다. '③ 결과 만들기'를 먼저 눌러 주세요.", vbInformation, "수출 마감": Exit Sub
    Workbooks.Open MonthFolder(ym) & "\" & f
End Sub

Private Function CloseOpenFiles(ByVal folder As String) As Boolean
    Dim wb As Workbook, names As String, i As Long
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name And LCase(wb.Path) = LCase(folder) Then names = names & vbCrLf & "  - " & wb.Name
    Next
    If names = "" Then CloseOpenFiles = True: Exit Function
    If MsgBox("이 달 폴더의 파일이 엑셀에서 열려 있습니다." & names & vbCrLf & vbCrLf & _
              "저장하고 닫은 뒤 계속할까요? (입력한 내용은 저장된 다음 새 결과 파일에 되살립니다)", vbOKCancel + vbQuestion, "수출 마감") <> vbOK Then Exit Function
    For i = Application.Workbooks.Count To 1 Step -1
        Set wb = Application.Workbooks(i)
        If wb.Name <> ThisWorkbook.Name And LCase(wb.Path) = LCase(folder) Then
            If wb.ReadOnly Then wb.Close SaveChanges:=False Else wb.Close SaveChanges:=True
        End If
    Next
    CloseOpenFiles = True
End Function

Private Function ReadUtf8(ByVal p As String) As String
    If Dir(p) = "" Then Exit Function
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.LoadFromFile p
    ReadUtf8 = st.ReadText: st.Close
End Function

Private Sub RunScript(ByVal checkOnly As Boolean)
    Dim ym As String: ym = MonthText()
    If Not CheckMonth(ym) Then Exit Sub
    Dim f As String: f = MonthFolder(ym)
    If Dir(f, vbDirectory) = "" Then
        MsgBox "폴더가 없습니다: " & f & vbCrLf & "'① 자료 폴더 열기'를 눌러 폴더를 만들고 기초자료를 넣어 주세요.", vbExclamation, "수출 마감": Exit Sub
    End If
    If Dir(PS_FILE) = "" Then MsgBox "스크립트가 없습니다: " & PS_FILE, vbCritical, "수출 마감": Exit Sub
    If Not CloseOpenFiles(f) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A" & LOG_ROW & ":A2000").EntireRow.Delete
    ws.Range("C6").Value = IIf(checkOnly, "파일 확인 중...", "만드는 중... (검은 창이 닫힐 때까지 기다려 주세요, 1~3분)")
    DoEvents
    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & PS_FILE & """ -YM " & ym
    If checkOnly Then cmd = cmd & " -CheckOnly"
    Dim rc As Long: rc = CreateObject("WScript.Shell").Run(cmd, 1, True)
    ShowLog f, rc, IIf(checkOnly, 0, 1)
End Sub

Private Sub ShowLog(ByVal folder As String, ByVal rc As Long, ByVal mode As Long)
    ' mode 0 = 파일 확인, 1 = 결과 만들기, 2 = PDF 합치기
    Dim checkOnly As Boolean: checkOnly = (mode = 0)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(1)
    Dim pre As String: pre = IIf(mode = 2, "PDF_", "")
    Dim logTxt As String: logTxt = ReadUtf8(folder & "\작업\" & pre & "실행기록.txt")
    Dim res As String: res = ReadUtf8(folder & "\작업\" & pre & "결과.txt")
    Dim lines() As String, i As Long, r As Long, s As String
    lines = Split(Replace(logTxt, vbCr, ""), vbLf)
    r = LOG_ROW
    ws.Cells(r, 1).Value = "실행 기록 (" & Format(Now, "yyyy-mm-dd hh:nn") & ")": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    For i = 0 To UBound(lines)
        s = lines(i)
        If Len(s) > 0 Then
            ws.Cells(r, 1).Value = "'" & s
            If InStr(s, "[에러]") > 0 Or InStr(s, "[문제]") > 0 Or InStr(s, "멈췄") > 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(255, 206, 199)
            ElseIf InStr(s, "[알림]") > 0 Or InStr(s, "[확인]") > 0 Or InStr(s, "[빠짐]") > 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(255, 255, 153)
            ElseIf Left(s, 1) = "①" Or Left(s, 1) = "②" Or Left(s, 1) = "④" Or Left(s, 1) = "⑤" Or Left(s, 1) = "⑥" Or Left(s, 2) = "══" Then
                ws.Cells(r, 1).Font.Bold = True
            End If
            r = r + 1
        End If
    Next
    Dim ok As Boolean: ok = (rc = 0 And Left(res, 2) = "OK")
    If ok Then
        ws.Range("C6").Value = IIf(checkOnly, "파일 확인 끝 - 문제 없음", "완료 - " & Format(Now, "yyyy-mm-dd hh:nn"))
        ws.Range("C6").Interior.Color = RGB(198, 239, 206)
    Else
        ws.Range("C6").Value = "문제가 있습니다 - 아래 빨간 줄을 확인해 주세요"
        ws.Range("C6").Interior.Color = RGB(255, 206, 199)
    End If
    Dim msg As String: msg = Mid(res, InStr(res & vbLf, vbLf) + 1)
    If Len(msg) > 900 Then msg = Left(msg, 900) & vbCrLf & "... (아래 실행 기록에 전부 있음)"
    If ok And mode = 2 Then
        Dim pdf As String: pdf = PDF_ROOT & "\인쇄용_" & MonthText() & "\" & MonthText() & "_해외_수출서류_인쇄용.pdf"
        If MsgBox("PDF를 합쳤습니다." & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & "빠진 서류와 확인할 것은 같은 폴더의 '인쇄순서' 엑셀에 있습니다." & vbCrLf & "PDF를 열까요?", vbYesNo + vbInformation, "수출 마감") = vbYes Then Shell "explorer.exe """ & pdf & """", vbNormalFocus
    ElseIf ok And Not checkOnly Then
        If MsgBox("다 만들었습니다." & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & "결과 파일을 열까요?", vbYesNo + vbInformation, "수출 마감") = vbYes Then Open_Result
    ElseIf ok Then
        MsgBox "파일 확인 끝 - 문제 없습니다. '③ 결과 만들기'를 누르면 됩니다.", vbInformation, "수출 마감"
    Else
        If msg = "" Then msg = "결과를 읽지 못했습니다. 검은 창에 나온 내용을 알려 주세요."
        MsgBox msg, vbExclamation, "수출 마감"
    End If
End Sub
'@

$key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$old = (Get-ItemProperty $key -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty $key -Name AccessVBOM -Value 1 -Type DWord
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add(); while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $ws = $wb.Worksheets.Item(1); $ws.Name = '수출마감'
  $ws.Cells.Font.Name = '맑은 고딕'; $ws.Cells.Font.Size = 10
  $ws.Columns('A').ColumnWidth = 3; $ws.Columns('B').ColumnWidth = 16; $ws.Columns('C').ColumnWidth = 70; $ws.Columns('D:H').ColumnWidth = 12
  $ws.Range('B2').Value2 = '수출 마감 자동 만들기 (해외 시트 + 기초자료 검증)'; $ws.Range('B2').Font.Size = 16; $ws.Range('B2').Font.Bold = $true
  $ws.Range('B4').Value2 = '마감월'; $ws.Range('C4').Formula = '=TEXT(EDATE(TODAY(),-1),"yyyymm")'
  $ws.Range('C4').Interior.Color = 0x99FFFF; $ws.Range('C4').Font.Size = 14; $ws.Range('C4').Font.Bold = $true; $ws.Range('C4').HorizontalAlignment = -4131
  $ws.Range('D4').Value2 = '← 자동으로 지난달. 다른 달이면 6자리로 직접 입력 (예: 202611)'; $ws.Range('D4').Font.Color = 0x808080
  $ws.Range('B5').Value2 = '자료 폴더'; $ws.Range('C5').Formula = '="C:\ai\매월마감\"&C4'
  $ws.Range('B6').Value2 = '상태'; $ws.Range('C6').Value2 = '대기'
  $ws.Range('B4:B6').Font.Bold = $true
  # 버튼
  $top = $ws.Range('B8').Top; $h = 34; $w = 150; $x0 = $ws.Range('B8').Left
  $btns = @(@('① 자료 폴더 열기', 'Open_Folder'), @('② 파일 확인만', 'Run_Check'), @('③ 결과 만들기', 'Run_Build'), @('④ 결과 파일 열기', 'Open_Result'), @('⑤ 수출서류 PDF 합치기', 'Run_Pdf'), @('⑥ PDF 폴더 열기', 'Open_PdfFolder'))
  for ($i = 0; $i -lt $btns.Count; $i++) { $b = $ws.Buttons().Add($x0 + ($i % 4) * ($w + 10), $top + [Math]::Floor($i / 4) * ($h + 8), $w, $h); $b.Characters().Text = $btns[$i][0]; $b.OnAction = $btns[$i][1]; $b.Font.Size = 12; $b.Font.Bold = $true; $b.Name = $btns[$i][1] }
  $help = @(
    '쓰는 방법',
    '1. 마감월 확인 → ① 자료 폴더 열기 → 아래 파일을 그 폴더에 넣기',
    '     · ■YYYY년 M월 매출마감_라인별_인보이스별 상세■.xlsx  (해외영업팀)',
    '     · ■YY년 M월 수출마감_해외영업팀(라인별 매출)■.xlsx  (해외영업팀)',
    '     · 신코_지출품의서 마감.xlsx / 코마츠_지출품의서 마감.xlsx  (그 달 시트 신코YYYYMM·코마츠YYYYMM만 읽음)',
    '     · YYYY년 수출 (YYYY.MM).xlsx  (회계 수출 파일 — 없으면 지난달 결과 파일을 바탕으로 씀)',
    '2. ② 파일 확인만 → 빠진 파일·시트, 다른 달 파일, 열려 있는 파일을 알려 줌',
    '3. ③ 결과 만들기 → "…_M월반영.xlsx" 새 파일 (원본은 그대로, 백업 폴더에 복사본)',
    '     · 다시 만들어도 유니패스·메모·면장번호 입력은 인보이스별로 되살림 (실패하면 기존 결과 파일은 손대지 않음)',
    '     · 환율은 서울외국환중개 매매기준율로 비교 → 틀린 칸은 값은 그대로 두고 빨간색만 칠함 (기초자료·결과 둘 다)',
    '     · 확인할 것은 결과 파일의 "M월 기초자료 검증" 시트와 아래 실행 기록(분홍 = 에러, 노랑 = 알림)',
    '4. ⑥ PDF 폴더 열기 → 그 달 해외 PDF 폴더(예: 202611해외pdf)를 열어 면장·인보이스·패킹·BL 폴더를 넣기 (지난달 폴더도 지우지 말 것: 이월분을 거기서 찾음)',
    '5. ⑤ 수출서류 PDF 합치기 → 해외pdf\인쇄용_YYYYMM 에 인쇄용 PDF 하나 + 인쇄순서 엑셀(빠진 서류·확인할 것)',
    '     · 대상: 결과 파일의 전월이월(녹색)·당월·툴링 (다음 달 이월·관세 줄·신코·코마츠·ATLAS 제외)',
    '     · 인보이스마다 면장 → 인보이스 → 패킹리스트 → BL, 여러 버전이면 금액이 맞는 최신본'
  )
  for ($i = 0; $i -lt $help.Count; $i++) { $ws.Cells(13 + $i, 2).Value2 = $help[$i] }
  $ws.Cells(13, 2).Font.Bold = $true
  $m = $wb.VBProject.VBComponents.Add(1); $m.Name = 'ExportClose'; $m.CodeModule.AddFromString($vba)
  $xl.ActiveWindow.DisplayGridlines = $false; $ws.Range('C4').Select() | Out-Null
  if (Test-Path $out) { New-Item -ItemType Directory -Force 'C:\ai\매월마감\도구\백업' | Out-Null; Copy-Item $out ("C:\ai\매월마감\도구\백업\매월마감_실행_이전_$(Get-Date -f yyyyMMdd_HHmmss).xlsm") }
  $wb.SaveAs($out, 52); $wb.Close($false)
  "저장: $out"
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null; [GC]::Collect(); Start-Sleep -Seconds 4
  if ($null -eq $old) { Remove-ItemProperty $key -Name AccessVBOM -ErrorAction SilentlyContinue } else { Set-ItemProperty $key -Name AccessVBOM -Value $old -Type DWord }
  "VBA 접근 설정 되돌림: $(if ($null -eq $old) { '원래 없음 → 지움' } else { "원래 값 $old" })"
}
