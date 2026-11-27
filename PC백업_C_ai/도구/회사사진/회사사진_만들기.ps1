# 회사사진_공급망사업자.pptx: 1공장~5공장·표준공장 페이지(1공장 페이지 모양) + 페이지마다 사진 칸 4개(2x2)
#  사진 틀 = 1장 '회사 위치도'의 사진 틀과 같은 모양: 양쪽 모서리가 둥근 사각형(152) + 주황 테두리(#EF4423, 1.75pt) + 그림 채우기
#  사진 = 바탕화면\이화성 폴더. 'N공장입구.jpg'가 있으면 그 공장 첫 칸에 넣음(가운데를 틀 비율로 잘라 넣음). 나머지 칸은 빈 자리
param([hashtable]$Photos = @{})   # 예: @{ '1공장' = @('1공장입구.jpg', 'a.jpg', $null, 'b.jpg') } — 칸 순서대로, $null = 빈 칸
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$W = 'C:\Users\user\AppData\Local\Temp\claude\d3df15a30ffc05d5\ac480345-177e-4ecb-8db7-e819891326b4\scratchpad\photo_deck'
$SRC = Join-Path $W '원본.pptx'; $WORK = Join-Path $W 'work.pptx'; $IMGDIR = Join-Path $W 'img'; $PNG = Join-Path $W 'png'
$PHOTO = 'C:\Users\user\Desktop\이화성'
New-Item -ItemType Directory -Force $IMGDIR, $PNG | Out-Null; Get-ChildItem $PNG -Filter *.png | Remove-Item
$FACT = @('1공장', '2공장', '3공장', '4공장', '5공장', '표준공장')
# 사진 칸 자리 (슬라이드 960x540, 제목 띠 아래)
$X0 = 105; $Y0 = 88; $GAP = 18; $FW = 366; $FH = 209
$ORANGE = 0x2344EF   # PowerPoint RGB(BGR) = #EF4423

function Prep-Photo([string]$file, [string]$key) {
  # 사진 방향 맞추고, 틀 비율(FW:FH)로 가운데 자르고, 1400px로 줄여 JPEG로 저장
  $im0 = [Drawing.Image]::FromFile($file)
  try {
    $ori = 1; try { $ori = [BitConverter]::ToUInt16($im0.GetPropertyItem(0x0112).Value, 0) } catch {}
    switch ($ori) { 3 { $im0.RotateFlip('Rotate180FlipNone') } 6 { $im0.RotateFlip('Rotate90FlipNone') } 8 { $im0.RotateFlip('Rotate270FlipNone') } }
    $ar = $FW / $FH; $sw = $im0.Width; $shh = $im0.Height
    if ($sw / $shh -gt $ar) { $cw = [int]($shh * $ar); $ch = $shh; $cx = [int](($sw - $cw) / 2); $cy = 0 } else { $cw = $sw; $ch = [int]($sw / $ar); $cx = 0; $cy = [int](($shh - $ch) / 2) }
    $ow = 1400; $oh = [int](1400 / $ar)
    $bmp = New-Object Drawing.Bitmap $ow, $oh; $g = [Drawing.Graphics]::FromImage($bmp); $g.InterpolationMode = 'HighQualityBicubic'; $g.SmoothingMode = 'HighQuality'
    $g.DrawImage($im0, (New-Object Drawing.Rectangle 0, 0, $ow, $oh), (New-Object Drawing.Rectangle $cx, $cy, $cw, $ch), [Drawing.GraphicsUnit]::Pixel); $g.Dispose()
    $enc = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }; $ep = New-Object Drawing.Imaging.EncoderParameters 1; $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([Drawing.Imaging.Encoder]::Quality), 88L
    $out = Join-Path $IMGDIR "$key.jpg"; $bmp.Save($out, $enc, $ep); $bmp.Dispose(); return $out
  } finally { $im0.Dispose() }
}

# 칸마다 넣을 사진 정하기: 직접 준 것($Photos) 우선, 없으면 'N공장입구.jpg'를 첫 칸에
$PLAN = [ordered]@{}
foreach ($f in $FACT) {
  $slots = @($null, $null, $null, $null)
  if ($Photos.ContainsKey($f)) { $given = @($Photos[$f]); for ($i = 0; $i -lt 4 -and $i -lt $given.Count; $i++) { $slots[$i] = $given[$i] } }
  else { $ent = Join-Path $PHOTO "$($f)입구.jpg"; if (Test-Path -LiteralPath $ent) { $slots[0] = "$($f)입구.jpg" } }
  $PLAN[$f] = $slots
}
$prepared = @{}
foreach ($f in $FACT) { for ($i = 0; $i -lt 4; $i++) { $nm = $PLAN[$f][$i]; if ($nm) { $full = Join-Path $PHOTO $nm; if (-not (Test-Path -LiteralPath $full)) { throw "사진이 없음: $full" }; $prepared["$f|$i"] = Prep-Photo $full ("{0}_{1}" -f $f, ($i + 1)) } } }

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw '파워포인트가 열려 있음 → 닫아 주세요' }
Copy-Item -LiteralPath $SRC $WORK -Force
$pp = New-Object -ComObject PowerPoint.Application; $mine = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
try {
  $pr = $pp.Presentations.Open($WORK, 0, 0, 0)
  if ($pr.Slides.Count -ne 2) { throw "슬라이드가 2장이 아님($($pr.Slides.Count)) — 원본이 바뀐 듯" }
  $base = $pr.Slides.Item(2)
  if ($base.Shapes.Item('TextBox 2').TextFrame.TextRange.Text.Trim() -ne '1공장') { throw '2장 제목이 1공장이 아님' }
  # 1공장 페이지를 복사해 2공장~표준공장 (1공장 바로 뒤로 차례대로)
  $prev = $base
  for ($k = 1; $k -lt $FACT.Count; $k++) { $dup = $prev.Duplicate().Item(1); $dup.Shapes.Item('TextBox 2').TextFrame.TextRange.Text = $FACT[$k]; $prev = $dup }   # 바로 앞 장을 복사하면 그 뒤에 붙음
  $base.Shapes.Item('TextBox 2').TextFrame.TextRange.Text = '1공장'
  for ($k = 0; $k -lt $FACT.Count; $k++) {
    $sl = $pr.Slides.Item(2 + $k)
    for ($i = 0; $i -lt 4; $i++) {
      $l = $X0 + ($i % 2) * ($FW + $GAP); $t = $Y0 + [math]::Floor($i / 2) * ($FH + $GAP)
      $sh = $sl.Shapes.AddShape(152, $l, $t, $FW, $FH)
      $sh.Name = "사진{0}" -f ($i + 1)
      $sh.Adjustments.Item(1) = 0.09704
      $sh.Line.Visible = -1; $sh.Line.ForeColor.RGB = $ORANGE; $sh.Line.Weight = 1.75
      $sh.Shadow.Visible = 0
      $key = "$($FACT[$k])|$i"
      if ($prepared.ContainsKey($key)) { $sh.Fill.UserPicture($prepared[$key]); $sh.TextFrame.TextRange.Text = '' }
      else {
        $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = 0xF2F2F2
        $tr = $sh.TextFrame.TextRange; $tr.Text = "사진 {0}" -f ($i + 1); $tr.Font.Name = 'KoPub돋움체 Bold'; $tr.Font.Size = 16; $tr.Font.Color.RGB = 0xA6A6A6; $tr.ParagraphFormat.Alignment = 2
        $sh.TextFrame.VerticalAnchor = 3
      }
    }
  }
  $pr.Save()
  $h = [int](1600 * $pr.PageSetup.SlideHeight / $pr.PageSetup.SlideWidth)
  foreach ($sl in $pr.Slides) { $sl.Export((Join-Path $PNG ("s{0:D2}.png" -f $sl.SlideIndex)), 'PNG', 1600, $h) }
  "슬라이드: " + ((@($pr.Slides) | ForEach-Object { "$($_.SlideIndex) $($_.Shapes.Item('TextBox 2').TextFrame.TextRange.Text.Trim())" }) -join ' / ')
  $pr.Close()
} finally {
  foreach ($q in @($pp.Presentations)) { if ($q) { try { $q.Saved = -1; $q.Close() } catch {} } }
  if ($pp.Presentations.Count -eq 0) { $pp.Quit() }
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pp); $pp = $null; [GC]::Collect()
}
Start-Sleep -Seconds 3
foreach ($p in @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $mine -contains $_.Id })) { if ($p.MainWindowHandle -eq 0) { Stop-Process -Id $p.Id -Force } }
"사진 넣은 칸: " + (($PLAN.Keys | ForEach-Object { $f = $_; "$f " + ((0..3 | ForEach-Object { if ($PLAN[$f][$_]) { $PLAN[$f][$_] } else { '-' } }) -join ',') }) -join ' | ')
"파일 크기: {0:N0} KB" -f ((Get-Item $WORK).Length / 1KB)
