#  담당자별 매출 체크 엑셀 만들기
#  MES 출고(mesN월.xls) → 담당자 1명 = 파일 1개. 메모 옆 체크 칸 4개(매출/불량대체/이월/관세 툴링)에 미리 V를 넣어 보냄.
#  미리 넣은 값은 숨긴 '처음 체크' 칸에 보관 → 돌려받으면 담당자가 바꾼 줄을 찾을 수 있음.
#  미리 체크 기준: MES 금액 빈 국내 줄·메모 Replacement = 불량대체 / BL날짜 '이월'·일별마감 이월 국내 줄 = 이월 / 나머지 = 매출 / 관세 툴링 = 담당자가 직접
param(
  [string]$MesPath = "C:\ai\기초자료\202609\mes9월.xls",
  [string]$MesSheet = "mes9월",
  [string]$DailyPath = "C:\ai\기초자료\202609\2026년 일별마감list.xlsx",
  [int]$CloseYM = 202609,
  [string]$OutDir = "C:\ai\개발\2_부서배포\견본_202609"
)
$ErrorActionPreference = 'Stop'
function Num($x) { $d = 0.0; if ([double]::TryParse("$x", [ref]$d)) { $d } else { $null } }
$kNorm = { param($x) (("$x" -replace '\(주\)|㈜|주식회사|\(유\)|\s', '' -replace '에스앤티', 'SNT' -replace '아트라스콥코코리아|아트라스콥코', 'ATLASKOREA').ToUpper()) }
$pnK = { param($x) ("$x" -replace '[^0-9A-Za-z]', '').ToUpper() }
$sameCust = { param($a, $b) $a -eq $b -or ($a.Length -ge 3 -and $b.Length -ge 3 -and ($a.StartsWith($b.Substring(0, 3)) -or $b.StartsWith($a.Substring(0, 3)))) }
# 품번: 일별마감은 앞뒤에 번호가 붙음(예: MES KWA81284 = 일별 70KWA81284, 8810110401R = 3270-R8810110401) → 포함되면 같은 것
$pnSame = { param($m, $d) if (-not $m -or -not $d) { return $false }; if ($d.Contains($m)) { return $true }; $md = $m -replace '\D', ''; $md.Length -ge 6 -and ($d -replace '\D', '').Contains($md) }
$mNo = $CloseYM % 100; $nextNo = $mNo % 12 + 1; $yy = [math]::Floor($CloseYM / 100)
$checks = @('매출', '불량대체', '이월', '관세 툴링')
# 담당자 엑셀에 보여 줄 MES 칸 (순서대로, 없으면 건너뜀)
$showCols = @('NO', '출고일', '구분', '고객사', '품명', '품번', '출고수량', '단가', '금액', '출고Lot', 'BL날짜', '인보이스 번호', '메모')

$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
  # ── 일별마감 이번 달 시트: 이월 줄 (비고 'N월이월' 또는 고객사 하늘색 소계 범위 밖)
  $carryD = New-Object System.Collections.ArrayList
  if (Test-Path $DailyPath) {
    $wbD = $xl.Workbooks.Open($DailyPath, 0, $true); $sD = $wbD.Worksheets.Item("$($CloseYM - 200000)")
    $dn = $sD.UsedRange.Rows.Count; $dv = $sD.Range("A1:I$dn").Value2; $dfm = $sD.Range("H1:H$dn").Formula; $wbD.Close($false)
    $covD = @{}; $subK = @{}
    for ($i = 2; $i -le $dn; $i++) { if ("$($dfm[$i,1])" -match '^=SUM\(H(\d+):H(\d+)\)$') { $a1 = [int]$matches[1]; $b1 = [int]$matches[2]; for ($k1 = $a1; $k1 -le $b1; $k1++) { $covD[$k1] = $true }; $subK[(& $kNorm $dv[$a1,2])] = $true } }
    for ($i = 2; $i -le $dn; $i++) {
      $c0 = "$($dv[$i,2])".Trim(); $p0 = & $pnK $dv[$i,4]; $d0 = $dv[$i,1]; if (-not $c0 -or -not $p0 -or -not ($d0 -is [double])) { continue }
      $k0 = & $kNorm $c0; $b0 = "$($dv[$i,9])".Trim()
      if ($b0 -match '불량반출') { continue }
      if ($b0 -match "^$($nextNo)월\s*이월" -or ($subK[$k0] -and -not $covD[$i])) {
        $q0 = $dv[$i,6]; [void]$carryD.Add([pscustomobject]@{ R = $i; K = $k0; P = $p0; Q = $(if ($q0 -is [double]) { $q0 } else { 0.0 }); D = [math]::Floor($d0); B = $b0; Used = $false })
      }
    }
  }

  # ── MES
  $wb = $xl.Workbooks.Open($MesPath, 0, $true); $ws = $wb.Worksheets.Item($MesSheet)
  $nc = $ws.UsedRange.Columns.Count; $last = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  $v = $ws.Range($ws.Cells(1, 1), $ws.Cells($last, $nc)).Value2; $wb.Close($false)
  $C = @{}; for ($j = 1; $j -le $nc; $j++) { $h = "$($v[1,$j])".Trim(); if ($h -and -not $C.ContainsKey($h)) { $C[$h] = $j } }
  foreach ($h in '구분', '고객사', '금액', '출고수량', '출고일', '품번', '담당자', '메모', 'BL날짜') { if (-not $C[$h]) { throw "$MesSheet 에 '$h' 열이 없음" } }
  $cols = @($showCols | ? { $C[$_] })
  $rows = @()
  for ($r = 2; $r -le $last; $r++) {
    $who = "$($v[$r,$C['담당자']])".Trim(); if (-not $who) { continue }   # TOTAL 줄 등
    $kind = "$($v[$r,$C['구분']])".Trim(); $memo = "$($v[$r,$C['메모']])"; $amt = Num $v[$r,$C['금액']]
    $pick = '매출'; $why = ''
    if ($memo -match 'replacement') { $pick = '불량대체'; $why = '메모에 Replacement(교체품)' }
    elseif ($kind -eq '국내' -and -not $amt) { $pick = '불량대체'; $why = 'MES 금액이 비어 있음(불량대체)' }
    elseif ("$($v[$r,$C['BL날짜']])".Trim() -eq '이월') { $pick = '이월'; $why = "마감 결과에서 BL날짜가 '이월'($($nextNo)월 매출)" }
    elseif ($kind -eq '국내' -and $carryD.Count) {
      $k = & $kNorm $v[$r,$C['고객사']]; $p = & $pnK $v[$r,$C['품번']]; $q = [double](Num $v[$r,$C['출고수량']]); $d = [math]::Floor([double](Num $v[$r,$C['출고일']]))
      $hit = @($carryD | ? { -not $_.Used -and $_.Q -eq $q -and (& $sameCust $_.K $k) -and (& $pnSame $p $_.P) } | sort { [math]::Abs($_.D - $d) }) | select -First 1
      if ($hit) { $hit.Used = $true; $pick = '이월'; $why = "일별마감 $($mNo)월 시트 $($hit.R)행이 이월$(if ($hit.B) { " (비고: $($hit.B))" })" }
    }
    $rows += [pscustomobject]@{ R = $r; Who = $who; Kind = $kind; Pick = $pick; Why = $why }
  }

  # ── 담당자별 파일
  New-Item -ItemType Directory -Force $OutDir | Out-Null
  $GREEN = 13561798; $REDF = 13551615; $HEAD = 14277081; $LOCKF = 15921906
  $hdr = @($cols) + $checks + @('체크 확인', '미리 체크한 이유', '담당자 의견', '처음 체크', 'MES 행')
  $nH = $hdr.Count; $c1 = $cols.Count + 1; $c4 = $cols.Count + 4; $cOk = $c4 + 1; $cWhy = $c4 + 2; $cNote = $c4 + 3; $cOrig = $c4 + 4; $cRow = $c4 + 5
  function ColL($n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
  $L1 = ColL $c1; $L4 = ColL $c4
  $made = @()
  foreach ($g in @($rows | group Who)) {
    $who = $g.Name; $list = @($g.Group); $n = $list.Count
    $arr = New-Object 'object[,]' ($n + 1), $nH
    for ($j = 0; $j -lt $nH; $j++) { $arr[0, $j] = [string]$hdr[$j] }
    for ($i = 0; $i -lt $n; $i++) {
      $x = $list[$i]
      for ($j = 0; $j -lt $cols.Count; $j++) { $arr[($i + 1), $j] = $v[$x.R, $C[$cols[$j]]] }
      $arr[($i + 1), ($cols.Count + [array]::IndexOf($checks, $x.Pick))] = 'V'
      $arr[($i + 1), ($cWhy - 1)] = [string]$x.Why; $arr[($i + 1), ($cOrig - 1)] = [string]$x.Pick; $arr[($i + 1), ($cRow - 1)] = [double]$x.R
    }
    $out = $xl.Workbooks.Add(); while ($out.Worksheets.Count -gt 1) { $out.Worksheets.Item($out.Worksheets.Count).Delete() }
    $sh = $out.Worksheets.Item(1); $sh.Name = '매출 체크'
    $sh.Cells.Font.Name = '맑은 고딕'; $sh.Cells.Font.Size = 10
    $sh.Range($sh.Cells(1, 1), $sh.Cells($n + 1, $nH)).Value2 = $arr
    $last2 = $n + 1
    # 서식
    $h1 = $sh.Range($sh.Cells(1, 1), $sh.Cells(1, $nH)); $h1.Font.Bold = $true; $h1.Interior.Color = $HEAD; $h1.WrapText = $true; $h1.HorizontalAlignment = -4108; $h1.VerticalAlignment = -4108; $sh.Rows(1).RowHeight = 32
    $sh.Range($sh.Cells(1, $c1), $sh.Cells(1, $c4)).Interior.Color = 10284031   # 체크 칸 머리글 = 연한 노랑
    foreach ($h in '출고일', 'BL날짜') { $j = [array]::IndexOf($cols, $h) + 1; if ($j) { $sh.Range($sh.Cells(2, $j), $sh.Cells($last2, $j)).NumberFormat = 'yyyy-mm-dd' } }
    foreach ($h in '출고수량', '단가', '금액') { $j = [array]::IndexOf($cols, $h) + 1; if ($j) { $sh.Range($sh.Cells(2, $j), $sh.Cells($last2, $j)).NumberFormat = '#,##0' } }
    $ck = $sh.Range("$($L1)2:$($L4)$last2"); $ck.HorizontalAlignment = -4108; $ck.Font.Bold = $true
    [void]$ck.Validation.Delete(); [void]$ck.Validation.Add(3, 1, 1, 'V'); $ck.Validation.IgnoreBlank = $true; $ck.Validation.InCellDropdown = $true
    $ck.Validation.ErrorTitle = '체크'; $ck.Validation.ErrorMessage = "V만 넣을 수 있습니다. 지우려면 Delete 키를 누르세요."
    $fc = $ck.FormatConditions.Add(1, 3, '="V"'); $fc.Interior.Color = $GREEN; $fc.Font.Color = 24832   # V = 초록
    # 체크 확인: 한 줄에 정확히 1개
    $ok = $sh.Range($sh.Cells(2, $cOk), $sh.Cells($last2, $cOk))
    $ok.Formula = "=IF(COUNTIF($($L1)2:$($L4)2,""V"")=1,""OK"",IF(COUNTIF($($L1)2:$($L4)2,""V"")=0,""체크 필요"",""하나만 체크""))"
    $ok.HorizontalAlignment = -4108
    $fo = $ok.FormatConditions.Add(1, 4, '="OK"'); $fo.Interior.Color = $REDF; $fo.Font.Color = 393372; $fo.Font.Bold = $true
    # 잠금: 체크 칸·의견 칸만 수정 가능, 나머지는 회색 바탕 없이 잠금
    $sh.Cells.Locked = $true; $ck.Locked = $false; $sh.Range($sh.Cells(2, $cNote), $sh.Cells($last2, $cNote)).Locked = $false
    $sh.Range($sh.Cells(2, $cNote), $sh.Cells($last2, $cNote)).Interior.Color = 15921906
    $grid = $sh.Range($sh.Cells(1, 1), $sh.Cells($last2, $cNote))
    foreach ($b in 7, 8, 9, 10, 11, 12) { $grid.Borders.Item($b).LineStyle = 1; $grid.Borders.Item($b).Color = 12566463 }
    $sh.Columns.AutoFit() | Out-Null
    for ($j = $c1; $j -le $c4; $j++) { $sh.Columns($j).ColumnWidth = 8 }
    $sh.Columns($cOk).ColumnWidth = 11; $sh.Columns($cWhy).ColumnWidth = 40; $sh.Columns($cNote).ColumnWidth = 30
    $jm = [array]::IndexOf($cols, '메모') + 1; if ($jm) { $sh.Columns($jm).ColumnWidth = 30 }
    $jn = [array]::IndexOf($cols, '품명') + 1; if ($jn) { $sh.Columns($jn).ColumnWidth = 24 }
    $sh.Range($sh.Cells(1, $cOrig), $sh.Cells(1, $cRow)).EntireColumn.Hidden = $true   # 처음 체크·MES 행 = 집계용
    $sh.Activate(); $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true
    [void]$sh.Range($sh.Cells(1, 1), $sh.Cells($last2, $nH)).AutoFilter()
    $sh.Protect('', $false, $true, $false, $false, $true, $true, $true, $false, $false, $false, $false, $false, $false, $true, $false)   # 서식·필터 허용, 비밀번호 없음

    # 요약 시트: 체크 항목별 줄 수·금액 (수식)
    $sm = $out.Worksheets.Add($sh); $sm.Name = '요약'; $sm.Cells.Font.Name = '맑은 고딕'
    $ja = ColL ([array]::IndexOf($cols, '금액') + 1)
    $sm.Cells(1, 1).Value2 = "$yy년 $($mNo)월 매출 체크 — 담당자: $who"; $sm.Cells(1, 1).Font.Size = 14; $sm.Cells(1, 1).Font.Bold = $true
    $sm.Cells(2, 1).Value2 = "'매출 체크' 시트에서 줄마다 매출·불량대체·이월·관세 툴링 중 하나에 V가 있는지 확인해 주세요. 미리 V를 넣어 두었습니다. 틀린 것만 고치고, 고친 이유는 '담당자 의견' 칸에 적어 주세요."
    $sm.Range('A2:D2').Merge(); $sm.Range('A2').WrapText = $true; $sm.Rows(2).RowHeight = 45; $sm.Range('A2').VerticalAlignment = -4160
    $hs = @('구분', '줄 수', '금액(원)'); for ($j = 0; $j -lt 3; $j++) { $sm.Cells(4, $j + 1).Value2 = $hs[$j] }
    $sm.Range('A4:C4').Font.Bold = $true; $sm.Range('A4:C4').Interior.Color = $HEAD
    for ($i = 0; $i -lt 4; $i++) {
      $cl = ColL ($c1 + $i); $rr = 5 + $i
      $sm.Cells($rr, 1).Value2 = $checks[$i]
      $sm.Cells($rr, 2).Formula = "=COUNTIF('매출 체크'!$($cl)2:$($cl)$last2,""V"")"
      $sm.Cells($rr, 3).Formula = "=SUMIF('매출 체크'!$($cl)2:$($cl)$last2,""V"",'매출 체크'!$($ja)2:$($ja)$last2)"
    }
    $sm.Cells(9, 1).Value2 = '합계'; $sm.Cells(9, 2).Formula = '=SUM(B5:B8)'; $sm.Cells(9, 3).Formula = '=SUM(C5:C8)'; $sm.Range('A9:C9').Font.Bold = $true
    $sm.Cells(10, 1).Value2 = '확인 필요 줄'; $sm.Cells(10, 2).Formula = "=COUNTA('매출 체크'!$(ColL $cOk)2:$(ColL $cOk)$last2)-COUNTIF('매출 체크'!$(ColL $cOk)2:$(ColL $cOk)$last2,""OK"")"
    $f10 = $sm.Range('B10').FormatConditions.Add(1, 5, '=0'); $f10.Interior.Color = $REDF
    $sm.Cells(11, 1).Value2 = '전체 줄'; $sm.Cells(11, 2).Value2 = [double]$n
    $sm.Range('B5:C11').NumberFormat = '#,##0'; $sm.Columns('A').ColumnWidth = 16; $sm.Columns('B').ColumnWidth = 10; $sm.Columns('C').ColumnWidth = 18; $sm.Columns('D').ColumnWidth = 40
    $g4 = $sm.Range('A4:C11'); foreach ($b in 7, 8, 9, 10, 11, 12) { $g4.Borders.Item($b).LineStyle = 1; $g4.Borders.Item($b).Color = 12566463 }
    $sm.Activate()

    $fn = Join-Path $OutDir ("{0}년{1}월_매출체크_{2}.xlsx" -f $yy, $mNo, $who)
    if (Test-Path $fn) { Remove-Item $fn -Force }
    $out.SaveAs($fn, 51); $out.Close($false)
    $cnt = $list | group Pick | % { "$($_.Name) $($_.Count)" }
    $made += "{0} ({1}) {2}줄: {3}" -f $who, (($list | % Kind | select -Unique) -join '·'), $n, ($cnt -join ' / ')
  }
  "만든 파일: $OutDir"; $made
} finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
