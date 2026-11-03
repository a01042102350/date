# 검사자 제출용 정리본 만들기 (원본은 그대로, 복사본만 정리)
#  - mes 시트: 이월 행 빨간 칸 → 회색 / 달러 칸 빨간 → 연한 노랑(출하기준 금액변동, 참고) / 단수 차이·국해외(아틀라스) 빨간 → 해제
#    → 빨간 칸 = 진짜 에러만 (MES 금액이 비어 있는 행)
#  - 숨은 작업용 시트 '오류원본' 삭제, 오류목록의 다음 달 세금계산서 참고 줄 삭제
#  - HTML: GPT 문구·복사 버튼 삭제, 이름 변경
param([string]$Dir = "C:\ai\기초자료\202609", [string]$Mes = "mes9월.xls", [string]$MonthLabel = "2026년9월")
$GRAY = 15921906; $YEL = 10284031; $RED = 255
$out = Join-Path $Dir "제출본"; New-Item -ItemType Directory -Force $out | Out-Null
$xlsOut = Join-Path $out ($Mes -replace '\.xls$', '_제출본.xls'); Copy-Item (Join-Path $Dir $Mes) $xlsOut -Force
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($xlsOut); if ($wb.ReadOnly) { throw "제출본이 읽기 전용으로 열림" }
  $xl.ScreenUpdating = $false; $xl.Calculation = -4135
  $ws = $wb.Worksheets.Item(1); $nc = $ws.UsedRange.Columns.Count; $n = $ws.UsedRange.Rows.Count
  $C = @{}; for ($j = 1; $j -le $nc; $j++) { $h = "$($ws.Cells(1,$j).Value2)".Trim(); if ($h -and -not $C.ContainsKey($h)) { $C[$h] = $j } }
  $six = @('사형', '자동', '금형 주조', '가공1(2공장)', '가공2(5공장)', '가공3(3공장)') | % { $C[$_] } | ? { $_ }
  $cnt = @{ Gray = 0; Yel = 0; Clear = 0; Keep = 0 }
  for ($i = 2; $i -lt $n; $i++) {
    $carry = "$($ws.Cells($i,$C['BL날짜']).Text)" -eq '이월'
    $amtEmpty = -not "$($ws.Cells($i,$C['금액']).Value2)".Trim()
    $bg = $ws.Cells($i, $C['출고일']).Interior; $bgNone = $bg.ColorIndex -eq -4142; $bgC = [int]$bg.Color
    for ($j = 1; $j -le $nc; $j++) {
      $cell = $ws.Cells($i, $j); if ([int]$cell.Interior.Color -ne $RED) { continue }
      if ($carry) { $cell.Interior.Color = $GRAY; $cnt.Gray++ }
      elseif ($j -eq $C['달러']) { $cell.Interior.Color = $YEL; $cnt.Yel++ }
      elseif ($six -contains $j -and $amtEmpty) { $cnt.Keep++ }                       # 진짜 에러: MES 금액 없음
      elseif ($six -contains $j -or $j -eq $C['수출면장 번호']) { if ($bgNone) { $cell.Interior.ColorIndex = -4142 } else { $cell.Interior.Color = $bgC }; $cnt.Clear++ }   # 단수 차이·국해외(면장 없음 정상)
      else { $cnt.Keep++ }
    }
  }
  "mes 시트: 회색(이월) $($cnt.Gray) / 연한 노랑(달러) $($cnt.Yel) / 해제 $($cnt.Clear) / 빨간색 유지(에러) $($cnt.Keep)"
  foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq '오류원본') { $s.Visible = -1; $s.Delete(); "숨은 시트 '오류원본' 삭제" } }
  $es = $wb.Worksheets.Item("오류목록"); $en = $es.UsedRange.Rows.Count; $del = 0; $delRaw = 0
  for ($i = $en; $i -ge 2; $i--) { if ("$($es.Cells($i,10).Text)" -match '이번 달 비교에서 제외') { $delRaw += [double]("0" + $es.Cells($i,6).Value2); $es.Range("A${i}:J${i}").Delete(-4162) | Out-Null; $del++ } }   # 오른쪽 요약표는 그대로 두고 A:J만 위로
  for ($i = 1; $i -le 12; $i++) { if ("$($es.Cells($i,12).Text)" -eq '참고') { $es.Cells($i,13).Value2 = [double]($es.Cells($i,13).Value2 - $del); $es.Cells($i,14).Value2 = [double]($es.Cells($i,14).Value2 - $delRaw) } }
  "오류목록: 다음 달 세금계산서 참고 $($del)줄 삭제"
  $wb.Worksheets.Item(1).Activate() | Out-Null
  $xl.Calculation = -4105; $xl.ScreenUpdating = $true; $wb.Save(); $wb.Close($false)
} finally { $xl.Quit() }
# HTML
$htmlOut = Join-Path $out "$($MonthLabel)_매출마감_결과.html"
& "C:\ai\도구\월마감_HTML.ps1" -XlsPath $xlsOut -Out $htmlOut | Out-Null
$h = [IO.File]::ReadAllText($htmlOut)
$h = [regex]::Replace($h, '<div class="prompt">.*?</div>', '', 'Singleline')
$h = [regex]::Replace($h, '<div class="upload">.*?</div>', '', 'Singleline')   # 검사자용에는 자료 올리기 버튼 빼기
$h = [regex]::Replace($h, "document\.getElementById\('copy'\)[^\n]*\n?", '')
$h = [regex]::Replace($h, '\(2026-10-28 작업 [^)]*\)', '')
$h = $h.Replace('빨간 칸 = 회계 기준과 다르거나 자료 없음(사유는 오류목록).', '빨간 칸 = 고쳐야 할 에러(사유는 오류목록). 엑셀에서 회색 = 다음 달 이월, 연한 노랑 = 참고(출하기준 금액변동).')
[IO.File]::WriteAllText($htmlOut, $h, (New-Object Text.UTF8Encoding $false))
"HTML: $htmlOut (GPT 문구 남은 곳: $(([regex]::Matches($h, 'GPT')).Count))"
