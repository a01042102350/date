$ErrorActionPreference = 'Stop'
$sp = Split-Path $MyInvocation.MyCommand.Path
$newF = 'C:\ai\매월마감\202610\2026년 수출 (2026.10)_10월반영.xlsx'
$W = Import-Csv "$sp\b_rows.csv" -Encoding UTF8
function SetV($cell, $v) { $a = New-Object 'object[,]' 1, 1; if ($null -ne $v) { $v = $v.psobject.BaseObject }; $a[0, 0] = $v; $cell.Value2 = $a }
function Blocks($rows, $col) { $rs = @($rows | % { [int]$_.행 } | Sort-Object); $out = @(); $st = $rs[0]; $pv = $rs[0]
  foreach ($x in $rs[1..($rs.Count - 1)]) { if ($x -ne $pv + 1) { $out += "$col$($st):$col$pv"; $st = $x }; $pv = $x }; $out += "$col$($st):$col$pv"; $out -join ',' }
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($newF); $ws = $wb.Worksheets.Item('26.10(NEW)'); $s9 = $wb.Worksheets.Item('26.09(NEW)')
  # 표 위치 찾기 (N칸 '부  서')
  $h = 0; for ($r = 15; $r -le 30; $r++) { if ($ws.Cells($r,14).Text -match '부\s*서') { $h = $r } }
  if (-not $h) { throw '부서 머리글 못 찾음' }
  # 머리글 복구 (9월 같은 칸 값)
  $h9 = $h - 1; foreach ($c in 14..20) { SetV $ws.Cells($h, $c) ($s9.Cells($h9, $c).Value2) }
  $rDom = $h + 1; $rOvT = $h + 2; $rOv = $h + 3; $rOvTool = $h + 4; $rAt = $h + 5; $rTot = $h + 7
  $yen = $W | ? { $_.거래처 -in 'SHINKO', 'JS TECH(코마츠)' }; $ys = @($yen | % { [int]$_.행 } | Sort-Object)
  $tool = $W | ? { $_.종류 -in '툴링', '관세' }
  $atl = $W | ? 거래처 -eq 'ATLAS KOREA'
  $prod = $W | ? { $_.종류 -in '전월', '당월' -and $_.거래처 -notin 'SHINKO', 'JS TECH(코마츠)', 'ATLAS KOREA' }
  $ws.Cells($rDom, 16).Formula = "=SUM(H$($ys[0]):H$($ys[-1]),)"; $ws.Cells($rDom, 17).Formula = "=SUM(J$($ys[0]):J$($ys[-1]))"
  $ws.Cells($rOv, 16).Formula = "=SUM($(Blocks $prod 'H'))"; $ws.Cells($rOv, 17).Formula = "=SUM($(Blocks $prod 'J'))"
  $ws.Cells($rOvTool, 16).Formula = "=SUM($(Blocks $tool 'H'))"; $ws.Cells($rOvTool, 17).Formula = "=SUM($(Blocks $tool 'J'))"
  $ws.Cells($rAt, 16).Formula = "=SUM($(Blocks $atl 'H'))"; $ws.Cells($rAt, 17).Formula = "=SUM($(Blocks $atl 'J'))"
  # 위에서 잘못 쓴 칸(머리글 줄 다음 해외영업 TOOLING 줄은 위에서 다시 씀) — 합계 줄 이름
  SetV $ws.Cells($rTot, 14) ('10월 수출매출(국내영업+해외영업) 합계')
  # 합계 줄 이월 칸 숫자 형식을 옆 칸과 같게
  $t = 0; for ($r = 100; $r -le 200; $r++) { if ($ws.Cells($r,1).Text -like '총*계*') { $t = $r; break } }
  $ws.Cells($t, 11).NumberFormat = $ws.Cells($t, 10).NumberFormat
  $xl.Calculate()
  "표:"; for ($r = $h; $r -le $rTot; $r++) { $l = @(); foreach ($c in 14..20) { $x = $ws.Cells($r,$c); $tx = $x.Text.Trim(); if ($tx) { $l += "$($x.Address(0,0))=$tx" } }; $l -join ' | ' }
  "합계줄 K: $($ws.Cells($t,11).Text)"
  $bad = @(); foreach ($sh in $ws, $wb.Worksheets.Item('10월 기초자료 검증')) { $u = $sh.UsedRange; foreach ($cell in $u.Cells) { if ($cell.Text -match '^#(N/A|REF|VALUE|DIV|NAME)') { $bad += "$($sh.Name)!$($cell.Address(0,0))" } } }
  "남은 오류 칸: $($bad.Count) $($bad -join ', ')"
  $ws.Activate(); $ws.Range('A1').Select() | Out-Null
  $wb.Save(); $wb.Close($false); '저장'
} finally { $xl.Quit() }
