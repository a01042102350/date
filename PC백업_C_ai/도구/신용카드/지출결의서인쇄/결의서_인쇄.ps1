param(
  [string]$In = '',              # 결의서_모으기.ps1 결과 폴더 (manifest.json)
  [string]$Printer = '',         # 비우면 기본 프린터
  [switch]$Preview,              # 인쇄하지 않고 쪽 미리보기 그림만 만듦
  [string]$PreviewOut = '',
  [switch]$WithReceipts,         # 영수증 8장 모음을 맨 끝에 따로(2026-11-13 iR-ADV C3826에서 안 나옴 → 쓰지 말 것)
  [switch]$PerDoc                # 문서마다 결재문서 → 영수증 → 첨부 순서
)
# 문서마다: 결재문서 → 첨부(미리보기 쪽 그림) / 맨 끝: 신용카드매출전표 A4 한 장에 8장(2칸×4줄, 자르기 선, 아래 작은 글씨로 문서 표시)
$ErrorActionPreference = 'Stop'
$S = $PSScriptRoot
if (-not $In) { $In = "$env:TEMP\지출결의서인쇄\모음" }
if (-not $PreviewOut) { $PreviewOut = "$env:TEMP\지출결의서인쇄\미리보기" }
Add-Type -AssemblyName System.Drawing, System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asTaskA = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(); $task.Result }
function AwaitA($op) { $asTaskA.Invoke($null, @($op)).Wait() }
function PdfPages($path, $width) {
  $f = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
  $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($f)) ([Windows.Data.Pdf.PdfDocument])
  for ($i = 0; $i -lt $doc.PageCount; $i++) {
    $pg = $doc.GetPage($i); $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = $width
    $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream; AwaitA ($pg.RenderToStreamAsync($ms, $opt))
    $b = New-Object System.Drawing.Bitmap ([System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0)))
    , (New-Object System.Drawing.Bitmap $b); $b.Dispose(); $pg.Dispose()
  }
}
function LoadImg($p) { $b = New-Object System.Drawing.Bitmap $p; $c = New-Object System.Drawing.Bitmap $b; $b.Dispose(); return $c }

$man = @((Get-Content "$In\manifest.json" -Raw -Encoding UTF8) | ConvertFrom-Json | ForEach-Object { $_ })   # PS 5.1은 JSON 배열을 한 덩어리로 줌
$PW = 827; $PH = 1169            # A4, 1/100인치
$capFont = New-Object System.Drawing.Font('맑은 고딕', 7)
$gray = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(190, 190, 190)), 1
$gray.DashStyle = 'Dash'

# 한 쪽 그리기: 좌표는 실제 종이 기준(1/100인치)
function DrawPage($g, $pg) {
  $g.InterpolationMode = 'HighQualityBicubic'; $g.SmoothingMode = 'HighQuality'
  switch ($pg.kind) {
    'pdf' { $g.DrawImage($pg.bmp, 0, 0, $PW, $PH) }                                   # 결재문서: 종이 그대로(이미 여백 있음)
    'img' {                                                                              # 첨부: 여백 0.3인치 안에 맞춤
      $m = 30; $aw = $PW - 2 * $m; $ah = $PH - 2 * $m; $b = $pg.bmp
      $sc = [Math]::Min($aw / $b.Width, $ah / $b.Height); $w = $b.Width * $sc; $h = $b.Height * $sc
      $g.DrawImage($b, [float]($m + ($aw - $w) / 2), [float]$m, [float]$w, [float]$h)
    }
    'rcpt' {                                                                             # 영수증 8장
      $m = 30; $cw = ($PW - 2 * $m) / 2; $ch = ($PH - 2 * $m) / 4; $pad = 6; $capH = 16
      for ($k = 0; $k -lt $pg.items.Count; $k++) {
        $it = $pg.items[$k]; $cx = $m + ($k % 2) * $cw; $cy = $m + [Math]::Floor($k / 2) * $ch
        $b = $it.bmp; $aw = $cw - 2 * $pad; $ah = $ch - 2 * $pad - $capH
        $sc = [Math]::Min($aw / $b.Width, $ah / $b.Height); $w = $b.Width * $sc; $h = $b.Height * $sc
        $g.DrawImage($b, [float]($cx + ($cw - $w) / 2), [float]($cy + $pad + ($ah - $h) / 2), [float]$w, [float]$h)
        $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment = 'Center'; $fmt.Trimming = 'EllipsisCharacter'; $fmt.FormatFlags = 'NoWrap'
        $g.DrawString($it.cap, $capFont, [System.Drawing.Brushes]::DimGray, (New-Object System.Drawing.RectangleF ([float]($cx + $pad), [float]($cy + $ch - $pad - $capH + 2), [float]($cw - 2 * $pad), [float]$capH)), $fmt)
      }
      for ($c = 0; $c -le 2; $c++) { $x = $m + $c * $cw; $g.DrawLine($gray, [float]$x, [float]$m, [float]$x, [float]($PH - $m)) }
      for ($r = 0; $r -le 4; $r++) { $y = $m + $r * $ch; $g.DrawLine($gray, [float]$m, [float]$y, [float]($PW - $m), [float]$y) }
    }
  }
}

# 영수증 8장 한 쪽을 한 장짜리 그림(200dpi)으로 미리 그림 — 결재문서 쪽과 같은 방식이라 프린터에서 확실히 나옴
function RenderRcptPage($items) {
  $bmp = New-Object System.Drawing.Bitmap -ArgumentList 1654, 2339
  $bmp.SetResolution(200, 200)
  $g = [System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White)
  $g.PageUnit = [System.Drawing.GraphicsUnit]::Inch; $g.PageScale = 0.01      # 좌표 = 1/100인치
  DrawPage $g @{ kind = 'rcpt'; items = $items }
  $g.Dispose(); return $bmp
}

# 인쇄 묶음(작업) 만들기
$jobs = New-Object Collections.ArrayList
foreach ($d in $man) {
  $no = ($d.docNo -split '-')[-1]
  $pages = New-Object Collections.ArrayList
  if ($d.pdf -and (Test-Path -LiteralPath $d.pdf)) { foreach ($b in @(PdfPages $d.pdf 1654)) { [void]$pages.Add(@{ kind = 'pdf'; bmp = $b }) } }
  if ($PerDoc) {   # 문서마다 결재문서 → 그 문서 영수증(8장씩) → 첨부 (2026-11-13 사용자 지정 순서)
    $mine = @(@($d.receipts) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | ForEach-Object { @{ bmp = (LoadImg $_); cap = ('{0:00}. {1} {2}' -f $d.n, $no, $d.title) } })
    for ($i = 0; $i -lt $mine.Count; $i += 8) { [void]$pages.Add(@{ kind = 'pdf'; bmp = (RenderRcptPage @($mine[$i..([Math]::Min($i + 7, $mine.Count - 1))])) }) }
    foreach ($m in $mine) { $m.bmp.Dispose() }
  }
  foreach ($a in @($d.attach)) { foreach ($p in @($a.pages)) { if ($p -and (Test-Path -LiteralPath $p) -and (Get-Item -LiteralPath $p).Length -gt 0) { [void]$pages.Add(@{ kind = 'img'; bmp = (LoadImg $p) }) } } }
  [void]$jobs.Add(@{ name = ('지출결의서 {0:00} ({1}) {2}' -f $d.n, $no, $d.title); pages = $pages })
}
$rc = New-Object Collections.ArrayList
foreach ($d in $man) { $no = ($d.docNo -split '-')[-1]; foreach ($r in @($d.receipts)) { if ($r -and (Test-Path -LiteralPath $r)) { [void]$rc.Add(@{ bmp = (LoadImg $r); cap = ('{0:00}. {1} {2}' -f $d.n, $no, $d.title) }) } } }
if ($rc.Count -and $WithReceipts -and -not $PerDoc) {
  $pages = New-Object Collections.ArrayList
  for ($i = 0; $i -lt $rc.Count; $i += 8) { [void]$pages.Add(@{ kind = 'rcpt'; items = @($rc[$i..([Math]::Min($i + 7, $rc.Count - 1))]) }) }
  [void]$jobs.Add(@{ name = "신용카드매출전표 $($rc.Count)장 (8장씩)"; pages = $pages })
}
$totalPages = ($jobs | ForEach-Object { $_.pages.Count } | Measure-Object -Sum).Sum
"인쇄 묶음 $($jobs.Count)개, 종이 $totalPages 쪽 (영수증 $($rc.Count)장)"

if ($Preview) {
  New-Item -ItemType Directory -Force $PreviewOut | Out-Null
  $sc = 0.3; $tw = [int]($PW * $sc); $th = [int]($PH * $sc); $cols = 6; $all = @()
  foreach ($j in $jobs) { foreach ($p in $j.pages) { $all += , @($j.name, $p) } }
  $per = 18
  for ($s0 = 0; $s0 -lt $all.Count; $s0 += $per) {
    $chunk = $all[$s0..([Math]::Min($s0 + $per - 1, $all.Count - 1))]
    $rows = [Math]::Ceiling($chunk.Count / $cols)
    $sheet = New-Object System.Drawing.Bitmap ($cols * ($tw + 10) + 10), ($rows * ($th + 34) + 10)
    $g = [System.Drawing.Graphics]::FromImage($sheet); $g.Clear([System.Drawing.Color]::FromArgb(120, 120, 120))
    $lab = New-Object System.Drawing.Font('맑은 고딕', 8)
    for ($i = 0; $i -lt $chunk.Count; $i++) {
      $x = 10 + ($i % $cols) * ($tw + 10); $y = 10 + [Math]::Floor($i / $cols) * ($th + 34)
      $g.FillRectangle([System.Drawing.Brushes]::White, $x, $y, $tw, $th)
      $st = $g.Save(); $g.TranslateTransform($x, $y); $g.ScaleTransform($sc, $sc); DrawPage $g $chunk[$i][1]; $g.Restore($st)
      $g.DrawString(("{0}. {1}" -f ($s0 + $i + 1), $chunk[$i][0]), $lab, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF ([float]$x, [float]($y + $th + 2), [float]$tw, [float]30)))
    }
    $g.Dispose(); $f = "$PreviewOut\미리보기_{0:00}.png" -f ([int]($s0 / $per) + 1); $sheet.Save($f, [System.Drawing.Imaging.ImageFormat]::Png); $sheet.Dispose(); "미리보기: $f"
  }
} else {
  $ps = New-Object System.Drawing.Printing.PrinterSettings
  if ($Printer) { $ps.PrinterName = $Printer }
  if (-not $ps.IsValid) { throw "프린터를 찾을 수 없습니다: $($ps.PrinterName)" }
  "프린터: $($ps.PrinterName)"
  foreach ($j in $jobs) {
    if (-not $j.pages.Count) { "건너뜀(쪽 없음): $($j.name)"; continue }
    $pd = New-Object System.Drawing.Printing.PrintDocument
    $pd.PrinterSettings = $ps; $pd.DocumentName = $j.name
    $a4 = $ps.PaperSizes | Where-Object { $_.Kind -eq 'A4' } | Select-Object -First 1
    if ($a4) { $pd.DefaultPageSettings.PaperSize = $a4 }
    $pd.DefaultPageSettings.Landscape = $false
    $pd.PrintController = New-Object System.Drawing.Printing.StandardPrintController
    $script:jp = $j.pages; $script:ji = 0
    $pd.add_PrintPage({ param($s, $e)
      $g = $e.Graphics; $g.TranslateTransform(-$e.PageSettings.HardMarginX, -$e.PageSettings.HardMarginY)
      DrawPage $g $script:jp[$script:ji]; $script:ji++
      $e.HasMorePages = $script:ji -lt $script:jp.Count
    })
    $pd.Print(); $pd.Dispose()
    "보냄: $($j.name) ($($j.pages.Count)쪽)"
  }
}
foreach ($j in $jobs) { foreach ($p in $j.pages) { if ($p.bmp) { $p.bmp.Dispose() }; if ($p.items) { foreach ($it in $p.items) { } } } }
foreach ($r in $rc) { $r.bmp.Dispose() }
