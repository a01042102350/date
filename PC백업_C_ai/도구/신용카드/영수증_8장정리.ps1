param([string]$Workbook = '')   # 비우면 지금 열려 있는 엑셀의 현재 시트
# 엑셀 시트에 붙여넣은 영수증 그림들을 A4 한 장에 2칸 × 4줄(8장)로 정렬하고 인쇄 설정까지 맞춤
$ErrorActionPreference = 'Stop'
try {
  if ($Workbook) { $xl = New-Object -ComObject Excel.Application; $wb = $xl.Workbooks.Open($Workbook); $ws = $wb.ActiveSheet }
  else {
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { throw '열려 있는 엑셀이 없습니다. 영수증을 붙여넣은 엑셀을 열어 둔 채로 실행해 주세요.' }
    $ws = $xl.ActiveSheet; if (-not $ws) { throw '엑셀에 열린 시트가 없습니다.' }
  }
  $pics = @($ws.Shapes | ? { $_.Type -in 13, 11, 12 -or $_.Type -eq 1 -and $_.Fill.Type -eq 6 })   # 그림·붙여넣은 이미지
  if (-not $pics.Count) { $pics = @($ws.Shapes) }
  if (-not $pics.Count) { throw "'$($ws.Name)' 시트에 영수증 그림이 없습니다." }
  $pics = @($pics | Sort-Object { [Math]::Round($_.Top / 40) }, Left)   # 붙여넣은 순서(위→아래, 왼쪽→오른쪽)

  $xl.ScreenUpdating = $false
  $ws.Cells.ClearFormats() | Out-Null
  # A4(595×842pt)에서 여백 빼고: 칸 2개 × 줄 4개 = 칸 1개 가로 268pt, 세로 192pt (프린터 여백 감안)
  $cw = 268; $rh = 192; $pad = 5
  foreach ($c in 1, 2) { $col = $ws.Columns($c); $col.ColumnWidth = 40; for ($k = 0; $k -lt 6; $k++) { $col.ColumnWidth = $col.ColumnWidth * $cw / $col.Width } }
  $nRows = [Math]::Ceiling($pics.Count / 2)
  $ws.Range("A1:A$nRows").RowHeight = $rh
  for ($i = 0; $i -lt $pics.Count; $i++) {
    $p = $pics[$i]; $cell = $ws.Cells([Math]::Floor($i / 2) + 1, ($i % 2) + 1)
    $p.LockAspectRatio = -1; $p.Placement = 2
    $s = [Math]::Min(($cw - 2 * $pad) / $p.Width, ($rh - 2 * $pad) / $p.Height); $p.Width = $p.Width * $s
    $p.Left = $cell.Left + ($cw - $p.Width) / 2; $p.Top = $cell.Top + ($rh - $p.Height) / 2
  }
  # 자르기 선(연한 점선) + 인쇄 설정
  $rg = $ws.Range("A1:B$nRows"); foreach ($b in 7, 8, 9, 10, 11, 12) { $bd = $rg.Borders.Item($b); $bd.LineStyle = -4118; $bd.Color = 0xBBBBBB }
  $ps = $ws.PageSetup
  $ws.ResetAllPageBreaks()
  $ps.PrintArea = "`$A`$1:`$B`$$nRows"; $ps.PaperSize = 9; $ps.Orientation = 1; $ps.Zoom = 100
  $ps.LeftMargin = $xl.CentimetersToPoints(0.5); $ps.RightMargin = $xl.CentimetersToPoints(0.5)
  $ps.TopMargin = $xl.CentimetersToPoints(0.8); $ps.BottomMargin = $xl.CentimetersToPoints(0.5)
  $ps.HeaderMargin = 0; $ps.FooterMargin = 0; $ps.CenterHorizontally = $true; $ps.CenterHeader = ''; $ps.CenterFooter = ''
  for ($r = 5; $r -le $nRows; $r += 4) { $ws.HPageBreaks.Add($ws.Cells($r, 1)) | Out-Null }   # 4줄(8장)마다 다음 장
  $xl.ActiveWindow.DisplayGridlines = $false
  $xl.ScreenUpdating = $true
  $msg = "영수증 $($pics.Count)장을 A4 $([Math]::Ceiling($pics.Count / 8))장(한 장에 8장)으로 정리했습니다.`n이제 Ctrl + P 로 인쇄하시면 됩니다."
  if ($Workbook) { $wb.Save(); $wb.Close(); $xl.Quit(); $msg }
  else { Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show($msg, '영수증 8장 정리') | Out-Null }
} catch {
  if ($xl) { $xl.ScreenUpdating = $true }
  if ($Workbook) { throw }
  Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show("정리하지 못했습니다.`n$($_.Exception.Message)", '영수증 8장 정리') | Out-Null
}
