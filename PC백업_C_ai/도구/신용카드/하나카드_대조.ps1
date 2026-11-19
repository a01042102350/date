param(
  [string]$Book = "$env:USERPROFILE\Desktop\하나카드 통합_카드별.xlsx",                  # 카드별로 나눈 하나카드 내역
  [string]$Ledger = "$env:USERPROFILE\Desktop\gridDetailTab2_excel_2026-11-06_13-40.xlsx"  # 장부 상세(날짜|적요|대변|카드번호)
)
# 하나카드 카드별 시트의 각 줄을 장부와 1건씩 맞춰 보고, 옆 칸(O~Q)에 장부 날짜·적요를 적음. 장부에 없는 줄은 빨간색.
$ErrorActionPreference = 'Stop'
function N($x) { if ($null -eq $x -or "$x".Trim() -eq '') { 0 } else { [double]("$x" -replace ',', '') } }
function Put($s, $r, $vals, $c0 = 1) { $a = New-Object 'object[,]' 1, $vals.Count
  for ($i = 0; $i -lt $vals.Count; $i++) { $x = $vals[$i]; if ($null -ne $x) { $x = $x.psobject.BaseObject }; $a[0, $i] = $x }
  $s.Range($s.Cells($r, $c0), $s.Cells($r, $c0 + $vals.Count - 1)).Value2 = $a }
$RED = 0xC7CEFF; $YEL = 0x99FFFF; $GRY = 0xF2F2F2; $HD = 0xF2E6DD
$bkDir = Join-Path (Split-Path $PSCommandPath) '백업'; New-Item -ItemType Directory -Force $bkDir | Out-Null; Copy-Item $Book (Join-Path $bkDir ([IO.Path]::GetFileNameWithoutExtension($Book) + "_대조전_$(Get-Date -f yyyyMMdd_HHmmss).xlsx"))   # 반영 전 백업
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  # ── 장부 ──
  $wl = $xl.Workbooks.Open($Ledger, 0, $true); $v = $wl.Worksheets(1).UsedRange.Value2; $wl.Close($false)
  # 머리글에서 칸 찾기: 날짜 | 적요 | 대변(또는 금액) | 카드번호(없으면 적요의 #번호)
  $hc = @{}; for ($c = 1; $c -le $v.GetLength(1); $c++) { $hc["$($v[1,$c])".Trim()] = $c }
  $cD = $hc['날짜']; $cT = $hc['적요']; $cA = $(if ($hc['대변']) { $hc['대변'] } else { 3 }); $cK = $hc['카드번호']
  $led = for ($r = 2; $r -le $v.GetLength(0); $r++) { $d = "$($v[$r,$cD])"; if ($d -notmatch '^\d{4}-\d\d-\d\d' -or $null -eq $v[$r,$cA] -or "$($v[$r,$cA])" -eq '') { continue }
    $t = "$($v[$r,$cT])"; $k = $(if ($cK) { "$($v[$r,$cK])".Trim() } else { '' }); if (-not $k) { $m = [regex]::Match($t, '#_?\s*(\d{4})'); if ($m.Success) { $k = $m.Groups[1].Value } elseif ($t -match '_(\d{4})$') { $k = $Matches[1] } }
    [pscustomobject]@{ 날짜 = [datetime]$d.Substring(0, 10); 적요 = $t; 금액 = (N $v[$r,$cA]); 카드 = $k; 짝 = $null } }
  $ledSum = ($led | Measure-Object 금액 -Sum).Sum

  # ── 카드 시트 ──
  $wb = $xl.Workbooks.Open($Book)
  $rows = New-Object Collections.Generic.List[object]
  foreach ($s in $wb.Worksheets) { if ($s.Name -eq '카드별 요약') { continue }
    $last = $s.Cells($s.Rows.Count, 4).End(-4162).Row - 1   # 날짜 칸 맨 아래 = 합계 줄, 그 바로 위까지
    $s.Range("O1:R$($last + 1)").Clear() | Out-Null; $s.Range("A2:N$last").Interior.ColorIndex = -4142; $s.Range("A2:N$last").Font.ColorIndex = -4105
    $cv = $s.Range("A1:N$last").Value2
    for ($r = 2; $r -le $last; $r++) {
      $rows.Add([pscustomobject]@{ 시트 = $s; 행 = $r; 카드 = $s.Name; 날짜 = [datetime]::ParseExact("$($cv[$r,4])", 'yyyy.MM.dd', $null)
        거래처 = "$($cv[$r,5])"; 금액 = (N $cv[$r,11]); 잔액 = (N $cv[$r,14]); 수수료 = (N $cv[$r,12]); 취소 = "$($cv[$r,8])"; 결과 = ''; 장부날짜 = ''; 장부금액 = $null; 장부적요 = ''; 비고 = '' }) } }

  # 0원(같은 달 승인 후 취소)은 비교 대상 아님
  foreach ($x in $rows | ? { $_.금액 -eq 0 }) { $x.결과 = '취소'; $x.장부적요 = '0원(승인 후 바로 취소)'; $x.비고 = '비교 안 함' }
  # 1) 같은 카드·같은 금액 ±7일  2) 카드번호가 다르게 적힌 것: 같은 금액 ±3일  3) 할부: 장부에 전체 금액으로 적힌 것
  # 0) 장부 안에서 같은 카드·같은 금액 +/− 로 서로 지워지는 것(전표 취소)은 짝지어 뺌
  foreach ($x in $led | ? { $_.금액 -lt 0 -and -not $_.짝 }) {
    $y = $led | ? { -not $_.짝 -and $_.카드 -eq $x.카드 -and $_.금액 -eq - $x.금액 } | Sort-Object { [Math]::Abs(($_.날짜 - $x.날짜).Days) } | Select-Object -First 1
    if ($y) { $x.짝 = 'X'; $y.짝 = 'X' } }
  # 1) 같은 카드·같은 금액 ±7일  2) 카드번호가 다르게 적힌 것: 같은 금액 ±3일  3) 할부: 장부에 전체 금액  4) 해외 결제: 같은 카드, 금액 3% 안, ±3일
  foreach ($pass in 1, 2, 3, 4) {
    foreach ($l in $led | ? { -not $_.짝 }) {
      $cand = foreach ($c in $rows) { if ($c.결과) { continue }
        $dd = [Math]::Abs(($c.날짜 - $l.날짜).Days)
        $ok = if ($pass -eq 1) { ($dd -le 7 -or $c.금액 -lt 0) -and $c.금액 -eq $l.금액 -and $c.카드 -eq $l.카드 }   # 지난달 사용분 환불(마이너스)은 날짜 안 봄
          elseif ($pass -eq 2) { $dd -le 3 -and $c.금액 -eq $l.금액 }
          elseif ($pass -eq 3) { $dd -le 7 -and $c.잔액 -gt 0 -and $c.카드 -eq $l.카드 -and [Math]::Abs($l.금액 - ($c.금액 - $c.수수료 + $c.잔액)) -le 2 }
          else { $dd -le 3 -and $c.카드 -eq $l.카드 -and $l.금액 -gt 0 -and [Math]::Abs($c.금액 - $l.금액) -le $l.금액 * 0.03 -and $c.거래처 -match '[A-Z]{3}' }
        if ($ok) { $c } }
      $b = $cand | Sort-Object { [Math]::Abs(($_.날짜 - $l.날짜).Days) } | Select-Object -First 1
      if ($b) { $l.짝 = $b; $b.결과 = '있음'; $b.장부금액 = $l.금액; $b.장부날짜 = $l.날짜.ToString('yyyy-MM-dd'); $b.장부적요 = $l.적요
        if ($pass -eq 2) { $b.비고 = "장부에는 카드번호가 " + $(if ($l.카드) { "#$($l.카드)" } else { '없이' }) + " 적혀 있음 → #$($b.카드)로 고쳐야 함" }
        if ($pass -eq 3) { $b.비고 = "할부 선결제: 장부는 전체 $('{0:N0}' -f $l.금액)원, 이번 달 청구는 1회차 $('{0:N0}' -f $b.금액)원(수수료 포함) — 남은 금액은 선결제로 처리" }
        if ($pass -eq 4) { $b.비고 = "해외 결제 환율 차이: 장부 $('{0:N0}' -f $l.금액)원, 청구 $('{0:N0}' -f $b.금액)원 ($('{0:N0}' -f ($b.금액 - $l.금액))원 차이)" } } } }
  # 5) 마이너스 전표 하나 = 같은 카드의 취소 2줄 합계
  foreach ($l in $led | ? { -not $_.짝 -and $_.금액 -lt 0 }) {
    $neg = @($rows | ? { -not $_.결과 -and $_.금액 -lt 0 -and $_.카드 -eq $l.카드 })
    :f for ($i = 0; $i -lt $neg.Count; $i++) { for ($j = $i + 1; $j -lt $neg.Count; $j++) { if ($neg[$i].금액 + $neg[$j].금액 -eq $l.금액) {
      $neg[$i].장부금액 = $l.금액; foreach ($b in $neg[$i], $neg[$j]) { $b.결과 = '있음'; $b.장부날짜 = $l.날짜.ToString('yyyy-MM-dd'); $b.장부적요 = $l.적요; $b.비고 = "취소 2줄이 장부에는 한 전표 $('{0:N0}' -f $l.금액)원으로 합쳐 있음" }
      $l.짝 = $neg[$i]; break f } } } }
  foreach ($b in $rows | ? { -not $_.결과 }) { $b.결과 = '없음'; $b.장부적요 = '없음'
    $b.비고 = if ($b.금액 -lt 0) { "지난달 사용분 취소(환불) $('{0:N0}' -f $b.금액)원이 장부에 없음" } else { '장부에 없음' } }

  # ── 시트에 쓰기 ──
  foreach ($g in $rows | Group-Object 카드) { $s = $g.Group[0].시트
    Put $s 1 @('장부 날짜', '장부 금액', '장부 적요', '비고') 15; $s.Range('O1:R1').Interior.Color = $HD; $s.Range('O1:R1').Font.Bold = $true
    foreach ($x in $g.Group) { $r = $x.행; Put $s $r @($x.장부날짜, $x.장부금액, $x.장부적요, $x.비고) 15
      if ($x.결과 -eq '없음') { $s.Range("A${r}:R$r").Interior.Color = $RED; $s.Range("A${r}:R$r").Font.Color = 0x0000C0 }
      elseif ($x.비고 -and $x.결과 -eq '있음') { $s.Range("O${r}:R$r").Interior.Color = $YEL }
      elseif ($x.결과 -eq '취소') { $s.Range("A${r}:R$r").Font.Color = 0x808080 } }
    $last = $g.Group.Count + 1; $s.Range("O2:O$last").NumberFormat = 'yyyy-mm-dd'
    $s.Range("O1:R$last").Borders.LineStyle = 1; $s.Range("O1:R$last").Borders.Color = 0xBFBFBF; $s.Range("P2:P$($last + 1)").NumberFormat = '#,##0;-#,##0;0'
    $s.Cells($last + 1, 15).Value2 = '장부 합계'; $s.Cells($last + 1, 16).Formula = "=SUM(P2:P$last)"; $s.Range("O$($last + 1):P$($last + 1)").Font.Bold = $true; $s.Range("O$($last + 1):P$($last + 1)").Interior.Color = $GRY
    $s.Columns('O').ColumnWidth = 12; $s.Columns('P').ColumnWidth = 13; $s.Columns('Q').ColumnWidth = 48; $s.Columns('R').ColumnWidth = 40 }

  # ── 요약 시트: 통째로 새로 그림 ──
  $sum = $wb.Worksheets('카드별 요약'); $sum.Cells.Clear() | Out-Null; $sum.Cells.Font.Name = '맑은 고딕'; $sum.Cells.Font.Size = 10
  $sum.Hyperlinks.Delete() | Out-Null; $sum.Columns('A').NumberFormat = '@'
  Put $sum 1 @('하나카드 10월 청구 ↔ 장부 대조'); $sum.Cells(1,1).Font.Bold = $true; $sum.Cells(1,1).Font.Size = 13
  Put $sum 2 @("장부: $(Split-Path $Ledger -Leaf)  /  카드 시트: 결제금액 오른쪽에 장부 날짜·적요·비고, 빨간 줄 = 장부에 없는 것, 노란 칸 = 맞지만 확인할 것")
  Put $sum 4 @('카드', '카드번호', '건수', '카드 청구 합계', '장부 금액', '차이(카드 - 장부)', '장부에 없는 건수', '설명')
  $sum.Range('A4:H4').Interior.Color = $HD; $sum.Range('A4:H4').Font.Bold = $true
  $r = 5
  foreach ($g in $rows | Group-Object 카드 | Sort-Object Name) { $k = $g.Name
    $ca = ($g.Group | Measure-Object 금액 -Sum).Sum; $la = ($led | ? { $_.카드 -eq $k } | Measure-Object 금액 -Sum).Sum; if ($null -eq $la) { $la = 0 }
    $miss = @($g.Group | ? 결과 -eq '없음').Count; $note = @($g.Group | ? { $_.결과 -eq '있음' -and $_.비고 } | % 비고 | Sort-Object -Unique) -join ' / '
    if (-not $note -and $ca -eq $la) { $note = '일치' }
    Put $sum $r @($k, $(if ($g.Group[0].시트.Cells(2, 3).Text -match '-') { $g.Group[0].시트.Cells(2, 3).Text } else { "4201-****-****-$k" }), $g.Count, $ca, $la, ($ca - $la), $miss, $note)
    $sum.Hyperlinks.Add($sum.Cells($r, 1), '', "'$k'!A1") | Out-Null
    $sum.Range("A${r}:H$r").Interior.Color = $(if ($miss) { $RED } elseif ($ca -ne $la) { $YEL } else { 0xCEEFC6 }); $r++ }
  Put $sum $r @('합계', '', "=SUM(C5:C$($r-1))", "=SUM(D5:D$($r-1))", "=SUM(E5:E$($r-1))", "=SUM(F5:F$($r-1))", "=SUM(G5:G$($r-1))")
  foreach ($c in 3..7) { $sum.Cells($r, $c).Formula = $sum.Cells($r, $c).Value2 }
  $sum.Range("A${r}:H$r").Interior.Color = $GRY; $sum.Range("A${r}:H$r").Font.Bold = $true
  $sum.Range("C5:G$r").NumberFormat = '#,##0;-#,##0;0'; $sum.Range("A4:H$r").Borders.LineStyle = 1; $sum.Range("A4:H$r").Borders.Color = 0xBFBFBF
  $only = @($led | ? { -not $_.짝 -and $_.카드 })
  $r += 2; Put $sum $r @("장부에만 있고 하나카드 내역에 없는 것 ($($only.Count)건)"); $sum.Cells($r, 1).Font.Bold = $true; $r++
  if ($only.Count) { Put $sum $r @('카드', '장부 날짜', '금액', '적요'); $sum.Range("A${r}:D$r").Interior.Color = $HD; $r++
    foreach ($l in $only | Sort-Object 날짜) { Put $sum $r @($l.카드, $l.날짜.ToString('yyyy-MM-dd'), $l.금액, $l.적요); $sum.Range("A${r}:D$r").Interior.Color = $YEL; $sum.Cells($r,3).NumberFormat = '#,##0'; $r++ } }
  else { Put $sum $r @('없음 — 장부의 카드 전표는 전부 카드 내역과 맞음'); $r++ }
  $sum.Columns('A').ColumnWidth = 8; $sum.Columns('B').ColumnWidth = 22; $sum.Columns('C').ColumnWidth = 7; $sum.Columns('D:F').ColumnWidth = 15; $sum.Columns('G').ColumnWidth = 9; $sum.Columns('H').ColumnWidth = 90
  $sum.Range("H5:H20").WrapText = $true
  $sum.Activate(); $wb.Save(); $wb.Close()

  "장부 $(@($led).Count)건 {0:N0} / 카드 $($rows.Count)줄" -f $ledSum
  $rows | Group-Object 결과 | % { "  $($_.Name) $($_.Count)건" }
  "=== 장부에 없는 카드 줄"; $rows | ? 결과 -eq '없음' | % { "  #{0} {1:MM-dd} {2,10:N0} {3}" -f $_.카드, $_.날짜, $_.금액, $_.거래처 }
  "=== 카드번호·할부 비고"; $rows | ? { $_.결과 -eq '있음' -and $_.비고 } | % { "  #{0} {1:N0} {2} / {3}" -f $_.카드, $_.금액, $_.거래처, $_.비고 }
  "=== 장부에만 있는 것"; $only | % { "  {0:MM-dd} #{1} {2,10:N0} {3}" -f $_.날짜, $_.카드, $_.금액, $_.적요 }
} finally { $xl.Quit() }
