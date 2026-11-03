# Excel 없이 xlsx 읽기 (첫 시트 → 행 배열). 사용: . read_xlsx.ps1; $rows = Read-Xlsx $path
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Read-Xlsx($path, $sheetIndex = 1) {
  $tmp = Join-Path $env:TEMP ("x_" + [guid]::NewGuid().ToString('N') + ".xlsx"); Copy-Item $path $tmp
  $zip = [IO.Compression.ZipFile]::OpenRead($tmp)
  try {
    function Entry($name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); try { [xml]$sr.ReadToEnd() } finally { $sr.Close() } }
    $ss = @(); $sx = Entry 'xl/sharedStrings.xml'
    if ($sx) { foreach ($si in $sx.sst.si) { if ($si.t -ne $null) { $ss += [string]$(if ($si.t.'#text') { $si.t.'#text' } else { $si.t }) } else { $ss += (($si.r | % { if ($_.t.'#text') { $_.t.'#text' } else { $_.t } }) -join '') } } }
    $wbx = Entry 'xl/workbook.xml'; $rel = Entry 'xl/_rels/workbook.xml.rels'
    $sh = @($wbx.workbook.sheets.sheet)[$sheetIndex - 1]; $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $target = ($rel.Relationships.Relationship | ? { $_.Id -eq $rid }).Target; if ($target -notmatch '^xl/') { $target = 'xl/' + ($target -replace '^/', '' -replace '^xl/', '') }
    $sxml = Entry $target
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $sxml.worksheet.sheetData.row) {
      $cells = @{}
      foreach ($c in @($row.c)) {
        $ref = $c.r; $colL = ($ref -replace '\d', ''); $ci = 0; foreach ($ch in $colL.ToCharArray()) { $ci = $ci * 26 + ([int]$ch - 64) }
        $val = if ($c.t -eq 's') { $ss[[int]$c.v] } elseif ($c.t -eq 'inlineStr') { $c.is.t } else { $c.v }
        $cells[$ci] = $val
      }
      [void]$rows.Add([pscustomobject]@{ R = [int]$row.r; Cells = $cells })
    }
    [pscustomobject]@{ Sheet = $sh.name; Rows = $rows }
  } finally { $zip.Dispose(); Remove-Item $tmp -Force }
}
