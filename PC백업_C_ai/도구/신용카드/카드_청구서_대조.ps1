param(
  [Parameter(Mandatory)][string]$Bill,   # 카드사 청구서 엑셀 (원본내역 시트)
  [Parameter(Mandatory)][string]$Cmp,    # 비교 데이터 (날짜·적요·대변, 카드번호 칸이 없으면 적요의 #끝4자리)
  [Parameter(Mandatory)][string]$Out,
  [string]$Title = '카드 청구서 대조'
)
# 청구서 거래 ↔ 비교 데이터를 카드 끝 4자리로 맞춰 봄. 금액이 틀리거나 한쪽에 없는 줄 = 빨간 줄.
$ErrorActionPreference = 'Stop'
function Num($x) { $s = "$x" -replace '[,\s원]', ''; if ($s -match '^-?\d+(\.\d+)?$') { [double]$s } else { $null } }
function Dt($x) { $s = "$x".Trim() -replace '[./]', '-'; if ($s -match '^(\d{4})-(\d{1,2})-(\d{1,2})') { '{0}-{1:00}-{2:00}' -f $matches[1], [int]$matches[2], [int]$matches[3] } else { '' } }
function DayGap($a, $b) { [Math]::Abs(([datetime]$a - [datetime]$b).TotalDays) }

$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false; $xl.Visible = $false
try {
  # ── 청구서 ──
  $wb = $xl.Workbooks.Open($Bill, 0, $true); $ws = $wb.Worksheets.Item('원본내역'); $u = $ws.UsedRange
  $nR = $u.Rows.Count; $nC = $u.Columns.Count; $H = @{}
  for ($c = 1; $c -le $nC; $c++) { $H["$($u.Cells(1,$c).Text)".Trim()] = $c }
  $bills = New-Object Collections.Generic.List[object]
  for ($r = 2; $r -le $nR; $r++) {
    $card = "$($u.Cells($r,$H['카드번호']).Text)".Trim(); if (-not $card) { continue }
    $bills.Add([pscustomobject]@{ Card = $card.Substring($card.Length - 4); Date = Dt $u.Cells($r, $H['승인일자']).Text
      Amt = Num $u.Cells($r, $H['이용금액']).Text; Shop = "$($u.Cells($r,$H['가맹점명/국가명']).Text)".Trim()
      Kind = "$($u.Cells($r,$H['구분']).Text)".Trim(); Row = $r; M = $null })
  }
  $wb.Close($false)
  # ── 비교 데이터 ──
  $wb = $xl.Workbooks.Open($Cmp, 0, $true)
  $ws = @($wb.Worksheets) | ? { $hr1 = $_.UsedRange.Rows.Item(1).Value2; "$hr1" -match '적요' -or (@($hr1) -contains '적요') } | Select-Object -First 1
  if (-not $ws) { throw "비교 데이터에서 '적요' 머리글이 있는 시트를 못 찾음" }
  $u = $ws.UsedRange
  $nR = $u.Rows.Count; $nC = $u.Columns.Count; $H = @{}
  for ($c = 1; $c -le $nC; $c++) { $H["$($u.Cells(1,$c).Text)".Trim()] = $c }
  $cmps = New-Object Collections.Generic.List[object]
  for ($r = 2; $r -le $nR; $r++) {
    $d = Dt $u.Cells($r, $H['날짜']).Text; $a = Num $u.Cells($r, $H['대변']).Text
    if (-not $d -or $null -eq $a) { continue }   # 이월·월계·카드대금 결제(차변) 줄은 뺌
    $memo = "$($u.Cells($r,$H['적요']).Text)".Trim(); $card = ''
    if ($H['카드번호']) { $card = "$($u.Cells($r,$H['카드번호']).Text)".Trim() }
    if (-not $card -and $memo -match '#\s*(\d{4})') { $card = $matches[1] }
    if ($card.Length -gt 4) { $card = $card.Substring($card.Length - 4) }
    $cmps.Add([pscustomobject]@{ Card = $card; Date = $d; Amt = $a; Memo = $memo; Row = $r; M = $null })
  }
  $wb.Close($false)

  # ── 맞추기 ──  결과 줄: 청구서 쪽(B) + 비교 쪽(C), 상태, 설명
  $res = New-Object Collections.Generic.List[object]
  function AddRes($b, $c, $st, $why, $cAmtShow = $true) {
    $res.Add([pscustomobject]@{ B = $b; C = $c; St = $st; Why = $why; ShowC = $cAmtShow })
  }
  # 0) 취소: 같은 카드·날짜·적요의 +/− 금액
  foreach ($c in $cmps) { if ($c.M -or $c.Amt -ge 0) { continue }
    $p = $cmps | ? { -not $_.M -and $_ -ne $c -and $_.Card -eq $c.Card -and $_.Date -eq $c.Date -and $_.Memo -eq $c.Memo -and $_.Amt -eq -$c.Amt } | Select-Object -First 1
    if ($p) { $c.M = 'x'; $p.M = 'x'; AddRes $null $p '취소(상계)' '비교 데이터에서 같은 날 +/− 금액으로 취소됨 → 청구서에 없는 것이 정상'; AddRes $null $c '취소(상계)' '위 줄과 상계' } }
  # 1) 같은 카드·같은 금액·같은 날짜  2) 같은 카드·같은 금액·5일 안 (가까운 날 먼저)
  foreach ($gap in 0, 5) {
    foreach ($c in $cmps) { if ($c.M) { continue }
      $b = $bills | ? { -not $_.M -and $_.Card -eq $c.Card -and $_.Amt -eq $c.Amt -and (DayGap $_.Date $c.Date) -le $gap } | Sort-Object { DayGap $_.Date $c.Date }, Row | Select-Object -First 1
      if ($b) { $b.M = 'x'; $c.M = 'x'; AddRes $b $c '일치' $(if ($gap) { "금액 같음, 날짜만 다름 (청구서 $($b.Date) / 비교 $($c.Date))" } else { '' }) } } }
  # 3) 여러 건을 묶은 비교 줄: '외화' = 그 카드 해외 거래 전부, 'OPENAI·CHATGPT' = 그 카드 OPENAI 거래 전부
  foreach ($c in $cmps) { if ($c.M) { continue }
    $grp = $null; $what = ''
    if ($c.Memo -match '외화') { $grp = @($bills | ? { -not $_.M -and $_.Card -eq $c.Card -and $_.Kind -like '해외*' }); $what = '해외(외화) 거래' }
    elseif ($c.Memo -match 'OPENA|CHATGPT') { $grp = @($bills | ? { -not $_.M -and $_.Card -eq $c.Card -and $_.Shop -match 'OPENAI' }); $what = 'OPENAI 거래' }
    if (-not $grp -or $grp.Count -eq 0) { continue }
    $sum = ($grp | Measure-Object Amt -Sum).Sum; $c.M = 'x'; $first = $true
    $st = if ($sum -eq $c.Amt) { '합계로 일치' } else { '합계 다름' }
    $why = "비교 데이터 한 줄($('{0:N0}' -f $c.Amt))을 청구서 $what $($grp.Count)건 합계($('{0:N0}' -f $sum))와 비교" + $(if ($sum -ne $c.Amt) { " → $('{0:N0}' -f ($sum - $c.Amt))원 차이. 비교 데이터에 빠진 거래가 있는지 확인 필요" } else { '' })
    foreach ($b in ($grp | Sort-Object Date, Row)) { $b.M = 'x'; AddRes $b $c $st $why $first; $first = $false } }
  # 4) 같은 카드·같은 날 청구서 2~3건 합 = 비교 1줄
  foreach ($c in $cmps) { if ($c.M) { continue }
    $cand = @($bills | ? { -not $_.M -and $_.Card -eq $c.Card -and $_.Date -eq $c.Date }); $hit = $null
    for ($i = 0; $i -lt $cand.Count -and -not $hit; $i++) { for ($j = $i + 1; $j -lt $cand.Count -and -not $hit; $j++) {
      if ($cand[$i].Amt + $cand[$j].Amt -eq $c.Amt) { $hit = @($cand[$i], $cand[$j]) }
      for ($k = $j + 1; $k -lt $cand.Count -and -not $hit; $k++) { if ($cand[$i].Amt + $cand[$j].Amt + $cand[$k].Amt -eq $c.Amt) { $hit = @($cand[$i], $cand[$j], $cand[$k]) } } } }
    if ($hit) { $c.M = 'x'; $first = $true; foreach ($b in $hit) { $b.M = 'x'; AddRes $b $c '여러 건 합쳐서 일치' "같은 날 청구서 $($hit.Count)건 합계 = 비교 데이터 한 줄" $first; $first = $false } } }
  # 5) 같은 카드·같은 날 1건씩 남음 → 금액 다름
  foreach ($c in $cmps) { if ($c.M) { continue }
    $cb = @($bills | ? { -not $_.M -and $_.Card -eq $c.Card -and $_.Date -eq $c.Date })
    $cc = @($cmps | ? { -not $_.M -and $_.Card -eq $c.Card -and $_.Date -eq $c.Date })
    if ($cb.Count -eq 1 -and $cc.Count -eq 1) { $cb[0].M = 'x'; $c.M = 'x'; AddRes $cb[0] $c '금액 다름' "같은 카드·같은 날인데 금액이 $('{0:N0}' -f ($cb[0].Amt - $c.Amt))원 다름" } }
  # 6) 남은 것
  foreach ($b in $bills) { if ($b.M) { continue }
    if ($b.Amt -eq 0) { AddRes $b $null '0원(참고)' '청구 금액 0원 → 대조 필요 없음' } else { AddRes $b $null '비교 데이터에 없음' '청구서에는 있는데 비교 데이터에 없음 → 비용 처리 빠졌는지 확인' } }
  foreach ($c in $cmps) { if ($c.M) { continue }
    $dup = $cmps | ? { $_ -ne $c -and $_.Card -eq $c.Card -and $_.Amt -eq $c.Amt -and (DayGap $_.Date $c.Date) -le 5 } | Select-Object -First 1
    AddRes $null $c '청구서에 없음' $(if ($dup) { '같은 카드·같은 금액 거래가 비교 데이터에 두 번 있음 (중복 입력 또는 카드번호 잘못 적음) → 확인' } else { '비교 데이터에는 있는데 이 카드 청구서에 없음 → 카드번호·금액·청구 월 확인' }) }

  $red = '금액 다름', '비교 데이터에 없음', '청구서에 없음', '합계 다름'
  $heads = '카드 끝자리', '승인일자(청구서)', '가맹점(청구서)', '청구서 금액', '날짜(비교)', '적요(비교)', '비교 금액', '차이(청구서−비교)', '결과', '설명'
  foreach ($x in $res) {
    $x | Add-Member Card $(if ($x.B) { $x.B.Card } else { $x.C.Card })
    $x | Add-Member Key $(if ($x.B) { $x.B.Date } else { $x.C.Date })
    $x | Add-Member Ord $(if ($x.B) { $x.B.Row } else { 1000 + $x.C.Row })
  }
  function ToArr($list) {
    $a = New-Object 'object[,]' ($list.Count + 1), $heads.Count
    for ($j = 0; $j -lt $heads.Count; $j++) { $a[0, $j] = $heads[$j] }
    $i = 1
    foreach ($x in $list) {
      $a[$i, 0] = "'" + $x.Card
      if ($x.B) { $a[$i, 1] = $x.B.Date; $a[$i, 2] = $x.B.Shop; $a[$i, 3] = $x.B.Amt }
      if ($x.C) { $a[$i, 4] = $x.C.Date; $a[$i, 5] = $x.C.Memo; if ($x.ShowC) { $a[$i, 6] = $x.C.Amt } }
      if ($x.B -and $x.C -and $x.ShowC -and $x.St -ne '합계로 일치' -and $x.St -ne '합계 다름' -and $x.St -ne '여러 건 합쳐서 일치') { $a[$i, 7] = $x.B.Amt - $x.C.Amt }
      elseif ($x.B -and -not $x.C) { $a[$i, 7] = $x.B.Amt } elseif ($x.C -and -not $x.B) { $a[$i, 7] = - $x.C.Amt }
      $a[$i, 8] = $x.St; $a[$i, 9] = $x.Why; $i++
    }
    , $a
  }
  $byCard = @($res | Sort-Object Card, Key, Ord)
  $byDate = @($res | Sort-Object Key, Card, Ord)

  $ob = $xl.Workbooks.Add(); while ($ob.Worksheets.Count -lt 3) { $ob.Worksheets.Add([Type]::Missing, $ob.Worksheets.Item($ob.Worksheets.Count)) | Out-Null }
  $names = '카드번호순', '날짜순', '카드별 합계'
  for ($k = 0; $k -lt 2; $k++) {
    $s = $ob.Worksheets.Item($k + 1); $s.Name = $names[$k]; $list = @($byCard, $byDate)[$k]
    $arr = ToArr $list; $n = $list.Count + 1
    $s.Range($s.Cells(1, 1), $s.Cells($n, $heads.Count)).Value2 = $arr
    $s.Cells.Font.Name = '맑은 고딕'; $s.Cells.Font.Size = 10
    $hd = $s.Range($s.Cells(1, 1), $s.Cells(1, $heads.Count)); $hd.Font.Bold = $true; $hd.Interior.Color = 0xE6E6E6; $hd.WrapText = $true
    foreach ($col in 4, 7, 8) { $s.Columns($col).NumberFormat = '#,##0;-#,##0;0' }
    $s.Columns(1).HorizontalAlignment = -4108
    for ($i = 2; $i -le $n; $i++) {
      $x = $list[$i - 2]
      if ($x.St -in $red) { $rg = $s.Range($s.Cells($i, 1), $s.Cells($i, $heads.Count)); $rg.Interior.Color = 0xCEC7FF; $rg.Font.Color = 0x06009C }   # 연한 빨강 바탕 + 진한 빨강 글씨
      elseif ($x.St -like '취소*' -or $x.St -like '0원*') { $s.Range($s.Cells($i, 1), $s.Cells($i, $heads.Count)).Font.Color = 0x808080 }
      if ($k -eq 0 -and $i -gt 2 -and $list[$i - 3].Card -ne $x.Card) { $s.Range($s.Cells($i, 1), $s.Cells($i, $heads.Count)).Borders.Item(8).Weight = -4138 }   # 카드 바뀌는 곳 굵은 선
    }
    $s.Columns(1).ColumnWidth = 8; $s.Columns(2).ColumnWidth = 11; $s.Columns(3).ColumnWidth = 28; $s.Columns(4).ColumnWidth = 12
    $s.Columns(5).ColumnWidth = 11; $s.Columns(6).ColumnWidth = 44; $s.Columns(7).ColumnWidth = 12; $s.Columns(8).ColumnWidth = 12
    $s.Columns(9).ColumnWidth = 16; $s.Columns(10).ColumnWidth = 60
    $s.Range($s.Cells(1, 1), $s.Cells($n, $heads.Count)).AutoFilter() | Out-Null
    $s.Activate(); $xl.ActiveWindow.FreezePanes = $false; $s.Range('A2').Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true
  }
  # 카드별 합계
  $s = $ob.Worksheets.Item(3); $s.Name = $names[2]; $s.Cells.Font.Name = '맑은 고딕'; $s.Cells.Font.Size = 10
  $s.Cells(1, 1).Value2 = $Title; $s.Cells(1, 1).Font.Size = 13; $s.Cells(1, 1).Font.Bold = $true
  $s.Cells(2, 1).Value2 = '빨간 줄 = 금액이 다르거나 한쪽에 없는 거래 (카드번호순·날짜순 시트). 회색 글씨 = 취소·0원(정상). 청구서 금액은 이용금액(해외 수수료 포함).'
  $hh = '카드 끝자리', '청구서 건수', '청구서 합계', '비교 건수', '비교 합계', '차이(청구서−비교)', '빨간 줄 수'
  for ($j = 0; $j -lt $hh.Count; $j++) { $s.Cells(4, $j + 1).Value2 = $hh[$j] }
  $hr = $s.Range('A4:G4'); $hr.Font.Bold = $true; $hr.Interior.Color = 0xE6E6E6
  $cards = @(($bills.Card + $cmps.Card) | Sort-Object -Unique); $r = 5
  foreach ($cd in $cards) {
    $bb = @($bills | ? Card -eq $cd); $cc = @($cmps | ? Card -eq $cd)
    $bs = [double]($bb | Measure-Object Amt -Sum).Sum; $cs = [double]($cc | Measure-Object Amt -Sum).Sum
    $s.Cells($r, 1).Value2 = "'$cd"; $s.Cells($r, 2).Value2 = $bb.Count; $s.Cells($r, 3).Value2 = $bs; $s.Cells($r, 4).Value2 = $cc.Count; $s.Cells($r, 5).Value2 = $cs
    $s.Cells($r, 6).Formula = "=C$r-E$r"
    $s.Cells($r, 7).Formula = "=COUNTIFS('카드번호순'!A:A,A$r,'카드번호순'!I:I,""금액 다름"")+COUNTIFS('카드번호순'!A:A,A$r,'카드번호순'!I:I,""비교 데이터에 없음"")+COUNTIFS('카드번호순'!A:A,A$r,'카드번호순'!I:I,""청구서에 없음"")+COUNTIFS('카드번호순'!A:A,A$r,'카드번호순'!I:I,""합계 다름"")"
    $r++ }
  $s.Cells($r, 1).Value2 = '합계'; foreach ($col in 'B', 'C', 'D', 'E', 'F', 'G') { $s.Range("$col$r").Formula = "=SUM(${col}5:${col}$($r-1))" }
  $s.Range("A${r}:G$r").Font.Bold = $true; $s.Range("A${r}:G$r").Borders.Item(8).Weight = -4138
  $s.Range("B5:G$r").NumberFormat = '#,##0;-#,##0;0'; $s.Range("A5:A$r").HorizontalAlignment = -4108
  $s.Range("A4:G$r").Borders.Weight = 2; $s.Columns('A:G').ColumnWidth = 15
  $ob.Worksheets.Item(1).Activate()
  if (Test-Path $Out) { Remove-Item $Out }
  $ob.SaveAs($Out, 51); $ob.Close($false)
  # 화면 요약
  "[$Title] 청구서 $($bills.Count)건 $('{0:N0}' -f ($bills | Measure-Object Amt -Sum).Sum) / 비교 $($cmps.Count)건 $('{0:N0}' -f ($cmps | Measure-Object Amt -Sum).Sum)"
  $res | Group-Object St | % { "  $($_.Name): $($_.Count)줄" }
  foreach ($x in $byCard) { if ($x.St -in $red -or $x.St -like '합계*' -or $x.St -like '여러*' -or $x.Why -like '*날짜만*') {
    "  #$($x.Card) [$($x.St)] 청구 $($x.B.Date) $($x.B.Amt) $($x.B.Shop) | 비교 $($x.C.Date) $(if($x.ShowC){$x.C.Amt}) $($x.C.Memo)" } }
  "저장: $Out"
} catch { "오류 줄 $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; throw } finally { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
