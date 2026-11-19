param(
  [string]$Bill = "$env:USERPROFILE\Desktop\10월 기업은행전체 데이터.xlsx",   # 청구서(청구세부내역 시트)
  [string]$Ledger = "$env:USERPROFILE\Downloads\gridDetailTab2_excel_2026-11-05_10-49.xlsx",  # 기업비씨카드 계정 상세(적요에 _#카드4자리)
  [string]$Total = "$env:USERPROFILE\Desktop\gridDetailTab1_excel_2026-11-05_13-31.xlsx"     # 거래처별 합계(참고)
)
# 기업은행 결제 카드 청구 세부내역을 장부(적요의 #카드번호)와 1건씩 맞춰 보고, 카드 끝 4자리별 시트에 '없는 것'을 적음
$ErrorActionPreference = 'Stop'
# 원본은 그대로 두고 "원본이름_대조결과.xlsx"에 결과 시트를 더함
$bk = Join-Path (Split-Path $Bill) ([IO.Path]::GetFileNameWithoutExtension($Bill) + "_대조결과.xlsx")
Copy-Item $Bill $bk -Force

$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false; $xl.Visible = $false
try {
  # ── 장부 ──
  $wl = $xl.Workbooks.Open($Ledger, 0, $true); $v = $wl.Worksheets(1).UsedRange.Value2; $wl.Close($false)
  $led = for ($r = 2; $r -le $v.GetLength(0); $r++) {
    $d = "$($v[$r,1])"; $amt = $v[$r, 4]; if (-not $d -or -not $amt) { continue }
    $t = "$($v[$r,2])"; $m = [regex]::Match($t, '#_?\s*(\d{4})')
    [pscustomobject]@{ 날짜 = [datetime]$d; 적요 = $t; 금액 = [double]$amt; 카드 = $(if ($m.Success) { $m.Groups[1].Value } else { '' }); 짝 = $null }
  }
  $ledSum = ($led | Measure-Object 금액 -Sum).Sum
  # ── 거래처별 합계(참고) ──
  $tot = @{}
  if (Test-Path $Total) { $wt = $xl.Workbooks.Open($Total, 0, $true); $tv = $wt.Worksheets(1).UsedRange.Value2; $wt.Close($false)
    for ($r = 2; $r -le $tv.GetLength(0); $r++) { $n = "$($tv[$r,2])"; if ($n -match '기업' -and $n -match '카드|비씨') { $tot[$n.Trim()] = [double]("0" + $tv[$r, 5]) } } }

  # ── 청구서 ──
  $wb = $xl.Workbooks.Open($bk); $v = $wb.Worksheets('청구세부내역').UsedRange.Value2
  $rows = for ($r = 2; $r -le $v.GetLength(0); $r++) {
    if ("$($v[$r,3])" -ne '기업은행') { continue }
    $c = "$($v[$r,2])"; $m = [regex]::Match($c, '^\d{4}-\d{4}-\d{4}-(\d{4})(.*)$')
    [pscustomobject]@{ 행 = $r; 카드전체 = $c; 끝4 = $m.Groups[1].Value; 종류 = $m.Groups[2].Value.Trim(); 날짜 = [datetime]::ParseExact("$($v[$r,7])", 'yyyy.MM.dd', $null)
      금액 = [double]$v[$r, 8]; 거래처 = "$($v[$r,11])"; 구분 = "$($v[$r,5])"; 결과 = ''; 내용 = ''; 비고 = '' }
  }
  # 1) 같은 카드·같은 거래처의 +/− 취소 짝
  foreach ($x in $rows | ? { $_.금액 -lt 0 }) {
    $y = $rows | ? { -not $_.결과 -and $_.끝4 -eq $x.끝4 -and $_.거래처 -eq $x.거래처 -and $_.금액 -eq - $x.금액 } | Sort-Object { [Math]::Abs(($_.날짜 - $x.날짜).Days) } | Select-Object -First 1
    if ($y) { foreach ($z in $x, $y) { $z.결과 = '취소'; $z.내용 = '승인 후 취소되어 서로 없어짐(정상)' } }
  }
  # 2) 같은 카드 + 같은 금액 (날짜 가까운 순, 최대 ±7일)  3) 카드번호가 다르게 적힌 것: 같은 금액·같은 날
  foreach ($pass in 1, 2) {
    foreach ($l in $led | ? { -not $_.짝 }) {
      $cand = $rows | ? { -not $_.결과 -and $_.금액 -eq $l.금액 -and [Math]::Abs(($_.날짜 - $l.날짜).Days) -le $(if ($pass -eq 1) { 7 } else { 3 }) -and ($pass -eq 2 -or $_.끝4 -eq $l.카드) }
      $b = $cand | Sort-Object { [Math]::Abs(($_.날짜 - $l.날짜).Days) } | Select-Object -First 1
      if ($b) { $l.짝 = $b; $b.결과 = '있음'; $b.내용 = $l.적요
        if ($pass -eq 2) { $b.비고 = "장부 적요에는 카드번호가 " + $(if ($l.카드) { "#$($l.카드)" } else { '없게' }) + " 적혀 있음 → #$($b.끝4)로 고쳐야 함" } }
    }
  }
  foreach ($b in $rows | ? { -not $_.결과 }) { $b.결과 = '없음'; $b.내용 = '없음' }

  # ── 기존 결과 시트 지우기 ──
  foreach ($ws in @($wb.Worksheets)) { if ($ws.Name -notin 'Sheet2', '청구세부내역') { $ws.Delete() } }
  $last = $wb.Worksheets($wb.Worksheets.Count)
  function NewSheet($name) { $s = $wb.Worksheets.Add([Type]::Missing, $script:last); $s.Name = $name; $script:last = $s; $s.Cells.Font.Name = '맑은 고딕'; $s.Cells.Font.Size = 10; $s }
  function Put($s, $r, $vals) {   # 한 줄을 배열로 한 번에 씀 (칸마다 쓰면 PowerShell이 글자/숫자 형식을 섞어 오류)
    $a = New-Object 'object[,]' 1, $vals.Count
    for ($i = 0; $i -lt $vals.Count; $i++) { $x = $vals[$i]; if ($null -ne $x) { $x = $x.psobject.BaseObject }; $a[0, $i] = $x }
    $s.Range($s.Cells($r, 1), $s.Cells($r, $vals.Count)).Value2 = $a }
  $red = 0xC7CEFF; $yel = 0x99FFFF; $gry = 0xEEEEEE; $head = 0xF2E6DD

  # ── 요약 시트 ──
  $sum = NewSheet '카드별 요약'
  Put $sum 1 @('기업은행 카드 10월 청구 ↔ 장부(기업비씨카드 계정) 대조')
  $sum.Cells(1, 1).Font.Size = 13; $sum.Cells(1, 1).Font.Bold = $true
  Put $sum 2 @("청구서: $(Split-Path $Bill -Leaf) 청구세부내역(결제은행 기업은행)  /  장부: $(Split-Path $Ledger -Leaf)  /  빨간 줄 = 장부에 없는 것")
  Put $sum 4 @('카드 끝 4자리', '카드 종류', '청구 합계(노란색 금액)', '장부에 있는 금액', '장부에 없는 금액', '없는 건수', '확인할 내용')
  $sum.Range('A4:G4').Interior.Color = $head; $sum.Range('A4:G4').Font.Bold = $true
  $r = 5
  $groups = $rows | Group-Object 카드전체 | Sort-Object { $g = $_.Group[0]; @{ '' = 0; '두산' = 1; 'LNG' = 2; '후불하이패스' = 3 }[$g.종류] }, Name
  foreach ($g in $groups) {
    $f = $g.Group[0]; $all = ($g.Group | Measure-Object 금액 -Sum).Sum
    $no = @($g.Group | ? 결과 -eq '없음'); $noSum = ($no | Measure-Object 금액 -Sum).Sum; if (-not $noSum) { $noSum = 0 }
    $okSum = (@($g.Group | ? 결과 -eq '있음') | Measure-Object 금액 -Sum).Sum; if (-not $okSum) { $okSum = 0 }
    $kind = if ($f.종류) { $f.종류 } else { '법인카드' }
    $note = if (-not $no.Count) { '모두 맞음' } elseif ($no.Count -eq $g.Count) { '장부에 이 카드 사용 내역이 하나도 없음' } else { "장부에 없는 것 $($no.Count)건 → '$($f.끝4)' 시트 빨간 줄" }
    Put $sum $r @("'$($f.끝4)", $kind, $all, $okSum, $noSum, $no.Count, $note)
    $sum.Cells($r, 3).Interior.Color = $yel
    if ($no.Count) { $sum.Range("A${r}:G$r").Interior.Color = $red; $sum.Cells($r, 3).Interior.Color = $yel }
    $r++
    # ── 카드별 시트 ──
    $sname = $f.끝4 + $(if ($f.종류 -eq '후불하이패스') { '하이패스' } elseif ($f.종류) { $f.종류 } else { '' })
    $s = NewSheet $sname
    Put $s 1 @("카드 $($f.카드전체)  ($kind)"); $s.Cells(1, 1).Font.Bold = $true; $s.Cells(1, 1).Font.Size = 12
    Put $s 2 @('청구 합계', $all, '장부에 없는 금액', $noSum, '없는 건수', $no.Count); $s.Cells(2, 2).Interior.Color = $yel
    foreach ($c in 2, 4) { $s.Cells(2, $c).NumberFormat = '#,##0;-#,##0;0' }; $s.Range('A2,C2,E2').Font.Bold = $true
    Put $s 4 @('승인일', '금액', '거래처', '내용(장부 적요)', '결과', '비고')
    $s.Range('A4:F4').Interior.Color = $head; $s.Range('A4:F4').Font.Bold = $true
    $rr = 5
    foreach ($b in $g.Group | Sort-Object @{e = { @{ '없음' = 0; '있음' = 1; '취소' = 2 }[$_.결과] } }, 날짜, 행) {
      Put $s $rr @($b.날짜.ToString('yyyy-MM-dd'), $b.금액, $b.거래처, $b.내용, $b.결과, $b.비고)
      if ($b.결과 -eq '없음') { $s.Range("A${rr}:F$rr").Interior.Color = $red; $s.Range("A${rr}:F$rr").Font.Color = 0x00009C }
      elseif ($b.결과 -eq '취소') { $s.Range("A${rr}:F$rr").Interior.Color = $gry }
      if ($b.비고) { $s.Cells($rr, 6).Interior.Color = $yel }
      $rr++
    }
    $s.Range("B5:B$rr").NumberFormat = '#,##0;-#,##0;0'; $s.Range("A5:A$rr").NumberFormat = 'yyyy-mm-dd'; $s.Range("A5:A$rr").HorizontalAlignment = -4108
    $s.Columns('A').ColumnWidth = 11; $s.Columns('B').ColumnWidth = 12; $s.Columns('C').ColumnWidth = 34; $s.Columns('D').ColumnWidth = 60; $s.Columns('E').ColumnWidth = 7; $s.Columns('F').ColumnWidth = 45
    $s.Range("A4:F$($rr-1)").Borders.LineStyle = 1; $s.Range("A4:F$($rr-1)").Borders.Color = 0xBFBFBF
    $s.Activate(); $xl.ActiveWindow.SplitRow = 4; $xl.ActiveWindow.FreezePanes = $true
  }
  $billSum = ($rows | Measure-Object 금액 -Sum).Sum
  Put $sum $r @('합계', '', "=SUM(C5:C$($r-1))", "=SUM(D5:D$($r-1))", "=SUM(E5:E$($r-1))", "=SUM(F5:F$($r-1))")
  $sum.Range("A${r}:G$r").Font.Bold = $true; $sum.Range("A${r}:G$r").Borders.Item(8).Weight = -4138
  $sum.Range("C5:E$r").NumberFormat = '#,##0;-#,##0;0'
  $sum.Range("A4:G$r").Borders.LineStyle = 1; $sum.Range("A4:G$r").Borders.Color = 0xBFBFBF
  $r += 2

  # ── 장부에만 있고 기업은행 청구서에 없는 것 ──
  $only = @($led | ? { -not $_.짝 })
  Put $sum $r @('장부에는 있는데 기업은행 청구서에 없는 것'); $sum.Cells($r, 1).Font.Bold = $true; $r++
  Put $sum $r @('날짜', '카드', '금액', '장부 적요', '', '', '확인할 내용'); $sum.Range("A${r}:G$r").Interior.Color = $head; $r++
  if (-not $only.Count) { Put $sum $r @('없음'); $r++ }
  foreach ($l in $only) {
    $why = if ($l.카드 -eq '7006') { '#7006은 경남은행 카드 → 경남카드 청구서 대상, 기업비씨카드 계정에 잘못 들어갔는지 확인' } else { '기업은행 청구서에 같은 금액이 없음' }
    Put $sum $r @($l.날짜.ToString('yyyy-MM-dd'), "#$($l.카드)", $l.금액, $l.적요, '', '', $why); $sum.Range("A${r}:G$r").Interior.Color = $red; $sum.Cells($r, 3).NumberFormat = '#,##0'; $sum.Cells($r, 1).NumberFormat = 'yyyy-mm-dd'; $r++
  }
  $r++
  Put $sum $r @('금액 맞추기'); $sum.Cells($r, 1).Font.Bold = $true; $r++
  $lines = @(
    @('기업은행 청구 합계(노란색 총합계)', $billSum),
    @('장부 상세(기업비씨카드 10월 대변 합계)', $ledSum)
  )
  foreach ($k in $tot.Keys | Sort-Object) { $lines += , @("거래처별 합계 파일 '$k' 10월 증가", $tot[$k]) }
  foreach ($ln in $lines) { Put $sum $r @($ln[0], '', $ln[1]); $sum.Cells($r, 3).NumberFormat = '#,##0'; $r++ }
  $k2 = $tot.Keys | ? { $_ -eq '기업비씨카드' }
  if ($k2 -and $tot[$k2] -ne $ledSum) {
    Put $sum $r @("※ 거래처별 합계 파일(기업비씨카드 $('{0:N0}' -f $tot[$k2]))과 장부 상세($('{0:N0}' -f $ledSum))가 $('{0:N0}' -f ($tot[$k2] - $ledSum))원 다릅니다. 상세 파일을 받은 뒤에 전표가 더 입력된 것 같습니다. 최신 상세를 다시 받아 주시면 다시 맞춰 보겠습니다.")
    $sum.Range("A${r}:G$r").Interior.Color = $yel; $r++
  }
  $sum.Columns('A').ColumnWidth = 14; $sum.Columns('B').ColumnWidth = 12; $sum.Columns('C').ColumnWidth = 20; $sum.Columns('D').ColumnWidth = 18; $sum.Columns('E').ColumnWidth = 16; $sum.Columns('F').ColumnWidth = 10; $sum.Columns('G').ColumnWidth = 55
  $sum.Activate()
  $wb.Save(); $wb.Close($true)
  "결과: $bk"
  "청구 합계 {0:N0} / 장부 상세 {1:N0}" -f $billSum, $ledSum
  $rows | Group-Object 끝4 | % { $n = @($_.Group | ? 결과 -eq '없음'); "{0}: {1}건 중 없음 {2}건 {3:N0}원" -f $_.Name, $_.Count, $n.Count, (($n | Measure-Object 금액 -Sum).Sum) }
  '--- 카드번호 다르게 적힌 것'; $rows | ? 비고 | % { "$($_.끝4) $($_.금액) $($_.비고)" }
  '--- 장부에만 있음'; $only | % { "$($_.날짜.ToString('MM-dd')) #$($_.카드) $($_.금액) $($_.적요)" }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
