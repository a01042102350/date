$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = 'Tls12'
$sp = Split-Path $MyInvocation.MyCommand.Path
$out = @()
$d = [DateTime]'2026-09-28'
while ($d -le [DateTime]'2026-11-08') { $k = $d.ToString('yyyy-MM-dd')
  try { $t = (Invoke-WebRequest "http://www.smbs.biz/Flash/TodayExRate_flash.jsp?tr_date=$k" -UseBasicParsing -TimeoutSec 20).Content
    $usd = if ($t -match 'USD=([\d,\.]+)') { [double]($matches[1] -replace ',', '') } else { $null }
    $jpy = if ($t -match 'JPY=([\d,\.]+)') { [double]($matches[1] -replace ',', '') / 100 } else { $null }
    $out += [pscustomobject]@{ 날짜 = $k; 요일 = $d.ToString('ddd'); USD = $usd; JPY = $jpy } } catch { $out += [pscustomobject]@{ 날짜 = $k; 요일 = $d.ToString('ddd'); USD = 'ERR'; JPY = $_.Exception.Message } }
  $d = $d.AddDays(1) }
$out | Export-Csv "$sp\smbs.csv" -NoTypeInformation -Encoding UTF8
$out | % { "$($_.날짜) $($_.요일) USD $($_.USD) JPY $($_.JPY)" }
