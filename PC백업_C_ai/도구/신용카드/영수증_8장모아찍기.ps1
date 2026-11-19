param(
  [Parameter(Mandatory)][string]$In,     # 영수증 PDF·이미지 파일 또는 폴더
  [string]$Out = '',                     # 결과 PDF (비우면 입력 폴더에 '영수증_8장모아찍기.pdf')
  [int]$Cols = 2, [int]$Rows = 4,
  [switch]$Split                         # 한 페이지에 영수증이 여러 장 세로로 이어져 있으면 빈 줄 기준으로 나눔
)
# PDF 페이지 → 그림 → 흰 여백 잘라 냄 → A4 한 장에 Cols×Rows 칸 → 'Microsoft Print to PDF'로 저장
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing, System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asTaskA = [System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(); $task.Result }
function AwaitA($op) { $asTaskA.Invoke($null, @($op)).Wait() }

function PdfPages($path) {   # PDF 각 페이지를 Bitmap으로 (가로 1600px)
  $f = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
  $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($f)) ([Windows.Data.Pdf.PdfDocument])
  for ($i = 0; $i -lt $doc.PageCount; $i++) {
    $pg = $doc.GetPage($i); $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = 1600
    $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    AwaitA ($pg.RenderToStreamAsync($ms, $opt))
    $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0))
    $bmp = New-Object System.Drawing.Bitmap $net
    , (New-Object System.Drawing.Bitmap $bmp)   # 스트림과 분리
    $bmp.Dispose(); $pg.Dispose()
  }
}
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System; using System.Drawing; using System.Drawing.Imaging; using System.Collections.Generic;
public static class Ink {
  static byte[] Px(Bitmap b, out int stride) {
    var d = b.LockBits(new Rectangle(0,0,b.Width,b.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
    stride = d.Stride; var a = new byte[stride*b.Height]; System.Runtime.InteropServices.Marshal.Copy(d.Scan0, a, 0, a.Length); b.UnlockBits(d); return a; }
  static bool Dark(byte[] a, int i) { return a[i] < 235 || a[i+1] < 235 || a[i+2] < 235; }
  // 영수증 영역들: 세로로 minGap 이상 빈 줄이면 나눔, 각 영역의 좌우 여백도 잘라 냄
  public static List<Rectangle> Boxes(Bitmap b, int minGap, int pad) {
    int st; var a = Px(b, out st); int w = b.Width, h = b.Height; var rowInk = new bool[h];
    for (int y=0;y<h;y++) for (int x=0;x<w;x+=2) if (Dark(a, y*st+x*4)) { rowInk[y]=true; break; }
    var res = new List<Rectangle>(); int s=-1, e=-1, gap=0;
    for (int y=0;y<=h;y++) {
      bool ink = y<h && rowInk[y];
      if (ink) { if (s<0) s=y; e=y; gap=0; }
      else if (s>=0) { gap++; if (gap>=minGap || y==h) {
        if (e-s>=40) { int x0=w, x1=0;
          for (int yy=s;yy<=e;yy++) for (int x=0;x<w;x++) if (Dark(a, yy*st+x*4)) { if (x<x0) x0=x; if (x>x1) x1=x; }
          int X=Math.Max(0,x0-pad), Y=Math.Max(0,s-pad); res.Add(new Rectangle(X, Y, Math.Min(w,x1+pad)-X, Math.Min(h,e+pad)-Y)); }
        s=-1; } }
    }
    return res; }
}
"@
function Crops($bmp) {   # 흰 여백 잘라 낸 영역(들)
  $minGap = if ($Split) { [int]($bmp.Height * 0.025) } else { $bmp.Height + 1 }
  foreach ($r in [Ink]::Boxes($bmp, $minGap, 8)) { , $bmp.Clone($r, $bmp.PixelFormat) }
}
# ── 입력 모으기 ──
$files = if (Test-Path $In -PathType Container) { Get-ChildItem $In -File | ? { $_.Extension -match '^\.(pdf|png|jpe?g|bmp|gif|tif)$' -and $_.Name -notlike '영수증_8장모아찍기*' } | Sort-Object Name } else { @(Get-Item $In) }
if (-not $files) { throw "영수증 파일이 없습니다: $In" }
if (-not $Out) { $Out = Join-Path $(if (Test-Path $In -PathType Container) { $In } else { Split-Path $In }) '영수증_8장모아찍기.pdf' }
$pieces = New-Object Collections.Generic.List[object]
foreach ($f in $files) {
  $pages = if ($f.Extension -eq '.pdf') { @(PdfPages $f.FullName) } else { @(New-Object System.Drawing.Bitmap $f.FullName) }
  foreach ($p in $pages) { foreach ($c in @(Crops $p)) { $pieces.Add($c) } }
  "$($f.Name): 페이지 $($pages.Count)장 → 영수증 누계 $($pieces.Count)장"
}
if ($pieces.Count -eq 0) { throw '잘라 낸 영수증이 없습니다.' }

# ── A4에 배치해서 PDF로 인쇄 ──
$pd = New-Object System.Drawing.Printing.PrintDocument
$pd.PrinterSettings.PrinterName = 'Microsoft Print to PDF'; $pd.PrinterSettings.PrintToFile = $true; $pd.PrinterSettings.PrintFileName = $Out
$a4 = $pd.PrinterSettings.PaperSizes | ? { $_.Kind -eq 'A4' } | Select-Object -First 1; $pd.DefaultPageSettings.PaperSize = $a4
$pd.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins 30, 30, 30, 30   # 1/100인치 (약 7.6mm)
$pd.PrintController = New-Object System.Drawing.Printing.StandardPrintController
$script:idx = 0; $per = $Cols * $Rows
$pd.add_PrintPage({ param($s, $e)
  $g = $e.Graphics; $g.InterpolationMode = 'HighQualityBicubic'; $m = $e.MarginBounds
  $cw = $m.Width / $Cols; $ch = $m.Height / $Rows; $gutter = 6
  for ($k = 0; $k -lt $per -and $script:idx -lt $pieces.Count; $k++) {
    $img = $pieces[$script:idx]; $script:idx++
    $cx = $m.X + ($k % $Cols) * $cw; $cy = $m.Y + [Math]::Floor($k / $Cols) * $ch
    $sc = [Math]::Min(($cw - 2 * $gutter) / $img.Width, ($ch - 2 * $gutter) / $img.Height)
    $dw = $img.Width * $sc; $dh = $img.Height * $sc
    $g.DrawImage($img, [float]($cx + ($cw - $dw) / 2), [float]($cy + ($ch - $dh) / 2), [float]$dw, [float]$dh)
    $g.DrawRectangle([System.Drawing.Pens]::LightGray, [float]$cx, [float]$cy, [float]$cw, [float]$ch)   # 자르기 선
  }
  $e.HasMorePages = $script:idx -lt $pieces.Count
})
if (Test-Path $Out) { Remove-Item $Out }
$pd.Print()
"완료: 영수증 $($pieces.Count)장 → A4 $([Math]::Ceiling($pieces.Count / $per))장 ($($Cols)칸×$($Rows)줄)`n저장: $Out"
