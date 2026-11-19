param(
  [string]$Folder = "$env:USERPROFILE\Desktop\지출결의서",
  [string]$Stamp = (Get-Date -Format 'yyyyMMdd'),
  [switch]$Print
)
# 카드 자동전표(업로드 엑셀)를 사람이 보기 쉬운 분개 목록으로 정리 → A4 가로 인쇄
$ErrorActionPreference = 'Stop'
$upload = Join-Path $Folder "지출결의_카드자동전표_$Stamp.xlsx"
$review = Join-Path $Folder "지출결의_카드자동전표_$($Stamp)_검토.xlsx"
$outF = Join-Path $Folder "지출결의_카드자동전표_$($Stamp)_분개확인.xlsx"
$ACCNM = @{ '8110000' = '복리후생비(판)'; '5110000' = '복리후생비(제)'; '8120000' = '여비교통비(판)'; '5120000' = '여비교통비(제)'; '8130000' = '접대비(판)'; '5130000' = '접대비(제)'; '8250000' = '교육훈련비(판)'; '5250000' = '교육훈련비(제)'; '8300000' = '소모품비(판)'; '5300000' = '소모품비(제)'; '8310000' = '지급수수료(판)'; '5310000' = '지급수수료(제)'; '8340000' = '해외출장비'; '1350000' = '부가세대급금'; '2530000' = '미지급금' }
$PJNM = @{ '303' = '대신금속'; '103' = '창원공장'; '401' = '가공1(마산2)'; '402' = '가공2'; '403' = '가공3(마산3)'; '201' = '금형1'; '202' = '금형2'; '203' = '금형공장'; '101' = '사형라인'; '102' = '자동라인' }
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
try {
  # 업로드 줄 읽기 (2행 칸 코드로 위치 찾기)
  $wb = $xl.Workbooks.Open($upload, 0, $true); $sh = $wb.Worksheets.Item(1)
  $last = $sh.Cells($sh.Rows.Count, 1).End(-4162).Row; $maxC = $sh.UsedRange.Columns.Count
  $hdr = $sh.Range($sh.Cells(2, 1), $sh.Cells(2, $maxC)).Value2; $colOf = @{}; for ($c = 1; $c -le $maxC; $c++) { $cd = "$($hdr[1, $c])".Trim(); if ($cd -and -not $colOf.ContainsKey($cd)) { $colOf[$cd] = $c } }
  $vals = $sh.Range($sh.Cells(4, 1), $sh.Cells($last, $maxC)).Value2; $nr = $last - 3
  $wb.Close($false)
  function G($r, $cd) { "$($vals[$r, $colOf[$cd]])".Trim() }
  # 검토 엑셀: (날짜, 작성번호) → 문서번호
  $docOf = @{}
  $wr = $xl.Workbooks.Open($review, 0, $true); $s1 = $wr.Worksheets.Item(1); $rl = $s1.Cells($s1.Rows.Count, 1).End(-4162).Row
  $rv = $s1.Range("A2:C$rl").Value2; $wr.Close($false)
  for ($i = 1; $i -le $rl - 1; $i++) { $dv = $rv[$i, 1]; $dk = $(if ($dv -is [double]) { [DateTime]::FromOADate($dv).ToString('yyyyMMdd') } else { ("$dv" -replace '\D', '') }); $docOf["$dk|$($rv[$i, 2])"] = ("$($rv[$i, 3])" -replace '^.*-', '') }

  # 정리
  $head = @('날짜', '번호', '문서', '차대', '계정', '거래처', '사업자번호', '차변', '대변', '프로젝트', '부가세(세무·공급가액·카드)', '적요')
  $rowsOut = New-Object Collections.ArrayList; $blockEnd = New-Object Collections.ArrayList; $prevKey = ''; $sumD = 0; $sumC = 0; $nv = 0
  for ($r = 1; $r -le $nr; $r++) {
    $dt = G $r 'MENU_DT'; $sq = G $r 'MENU_SQ'; $key = "$dt|$sq"; $first = ($key -ne $prevKey)
    if ($first -and $prevKey) { [void]$blockEnd.Add($rowsOut.Count) }
    if ($first) { $nv++ }
    $dr = G $r 'DRCR_FG'; $ac = G $r 'ACCT_CD'; $am = [double](G $r 'ACCT_AM')
    $trc = G $r 'TR_CD'; $trn = G $r 'TR_NM'; $tr = $(if ($trc) { "[$trc] $trn" } else { "(새 거래처) $trn" })
    $pj = G $r 'PJT_CD'; $vat = ''
    if ((G $r 'TAX_FG')) { $cn = G $r 'CT_NB'; $vat = "세무 $(G $r 'TAX_FG') / 공급가액 {0:N0} / 결제카드 {1}" -f [double](G $r 'SUP_AM'), $cn }
    if ($dr -eq '3') { $sumD += $am } else { $sumC += $am }
    [void]$rowsOut.Add(@(
      $(if ($first) { "$($dt.Substring(4,2))-$($dt.Substring(6,2))" } else { '' }),
      $(if ($first) { $sq } else { '' }),
      $(if ($first) { $docOf[$key] } else { '' }),
      $(if ($dr -eq '3') { '차변' } else { '대변' }),
      "[$ac] $($ACCNM[$ac])",
      $tr, (G $r 'REG_NB'),
      $(if ($dr -eq '3') { $am } else { $null }), $(if ($dr -eq '4') { $am } else { $null }),
      $(if ($pj) { "$pj $($PJNM[$pj])" } else { '' }), $vat,
      (G $r 'RMK_DC')))
    $prevKey = $key
  }
  [void]$blockEnd.Add($rowsOut.Count)

  # 새 엑셀에 쓰기
  $wo = $xl.Workbooks.Add(); $so = $wo.Worksheets.Item(1); $so.Name = '분개 확인'
  $so.Range('A1').Value2 = "지출결의서(법인카드) 자동전표 분개 확인 — 전표 $nv 장 · 줄 $nr 개 · 작성 $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
  $so.Range('A1').Font.Bold = $true; $so.Range('A1').Font.Size = 12
  $n = $rowsOut.Count; $arr = New-Object 'object[,]' ($n + 1), $head.Count
  for ($c = 0; $c -lt $head.Count; $c++) { $arr[0, $c] = $head[$c] }
  for ($i = 0; $i -lt $n; $i++) { for ($c = 0; $c -lt $head.Count; $c++) { $arr[($i + 1), $c] = $rowsOut[$i][$c] } }
  foreach ($cl in 'A', 'B', 'C', 'G') { $so.Range("$($cl)3:$($cl)$($n + 3)").NumberFormat = '@' }
  $so.Range($so.Cells(3, 1), $so.Cells($n + 3, $head.Count)).Value2 = $arr
  $tot = $n + 4
  $so.Cells($tot, 7).Value2 = '합계'; $so.Cells($tot, 8).Value2 = $sumD; $so.Cells($tot, 9).Value2 = $sumC
  $so.Cells($tot, 10).Value2 = $(if ($sumD -eq $sumC) { '차변 = 대변' } else { "차이 $($sumD - $sumC)" })
  $so.Range("G$($tot):J$($tot)").Font.Bold = $true
  $so.Range('H:I').NumberFormat = '#,##0'
  $hh = $so.Range($so.Cells(3, 1), $so.Cells(3, $head.Count)); $hh.Font.Bold = $true; $hh.Interior.Color = 0xE6E6E6
  $all = $so.Range($so.Cells(3, 1), $so.Cells($n + 3, $head.Count)); $all.Font.Size = 9; $all.VerticalAlignment = -4108
  $all.Borders.Item(12).LineStyle = 1; $all.Borders.Item(12).Color = 0xDDDDDD          # 가로 안쪽 연한 선
  $all.Borders.Item(11).LineStyle = 1; $all.Borders.Item(11).Color = 0xDDDDDD          # 세로 안쪽
  foreach ($b in 7, 8, 9, 10) { $all.Borders.Item($b).LineStyle = 1; $all.Borders.Item($b).Weight = 2 }
  foreach ($e in $blockEnd) { $rr = $e + 3; $bd = $so.Range($so.Cells($rr, 1), $so.Cells($rr, $head.Count)).Borders.Item(9); $bd.LineStyle = 1; $bd.Weight = -4138; $bd.Color = 0x555555 }   # 전표마다 굵은 선
  $widths = @(6.5, 5.5, 6.5, 4.5, 21, 30, 11, 10, 10, 13, 30, 44)
  for ($c = 0; $c -lt $widths.Count; $c++) { $so.Columns.Item($c + 1).ColumnWidth = [double]$widths[$c] }
  $so.Range('F:F').WrapText = $true; $so.Range('L:L').WrapText = $true; $so.Range('K:K').WrapText = $true
  $so.Range("A4:D$($n + 3)").HorizontalAlignment = -4108
  # 인쇄 설정: A4 가로, 가로 1쪽에 맞춤, 머리글 줄 반복, 쪽 번호
  $ps = $so.PageSetup
  $ps.PrintTitleRows = '$3:$3'; $ps.Orientation = 2; $ps.PaperSize = 9; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
  $ps.LeftMargin = $xl.CentimetersToPoints(0.8); $ps.RightMargin = $xl.CentimetersToPoints(0.8); $ps.TopMargin = $xl.CentimetersToPoints(1.0); $ps.BottomMargin = $xl.CentimetersToPoints(1.0)
  $ps.CenterFooter = '&P / &N 쪽'; $ps.CenterHorizontally = $true
  $xl.ActiveWindow.DisplayGridlines = $false
  if (Test-Path -LiteralPath $outF) { Remove-Item -LiteralPath $outF -Force }
  $wo.SaveAs($outF, 51)
  $pages = $so.PageSetup.Pages.Count
  "분개 확인 엑셀: $outF (전표 $nv 장, 줄 $nr 개, 차변 {0:N0} / 대변 {1:N0}, $pages 쪽)" -f $sumD, $sumC
  if ($Print) { $so.PrintOut(); "인쇄 보냄: $pages 쪽 → 기본 프린터" }
  $wo.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
