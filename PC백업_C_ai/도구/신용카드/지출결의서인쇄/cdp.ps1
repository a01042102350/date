# CDP helper for the work Edge (port 9333). ASCII only (PowerShell 5.1 without BOM).
Add-Type -AssemblyName System.Web.Extensions
$script:Port = 9333
$script:Ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$script:Ser.MaxJsonLength = [int]::MaxValue
$script:Ser.RecursionLimit = 1000
$script:Events = New-Object System.Collections.ArrayList

$ErrorActionPreference = 'Stop'
function Get-Tabs { (Invoke-RestMethod "http://127.0.0.1:$($script:Port)/json/list") | ForEach-Object { $_ } | Where-Object { $_.type -eq 'page' } }
function New-Tab($url = 'about:blank') { Invoke-RestMethod -Method Put "http://127.0.0.1:$($script:Port)/json/new?$url" }
function Close-Tab($id) { try { $null = Invoke-RestMethod "http://127.0.0.1:$($script:Port)/json/close/$id" } catch {} }

function Connect-Tab($tab) {
  $ws = New-Object System.Net.WebSockets.ClientWebSocket
  $ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(20)
  $ws.ConnectAsync([Uri]$tab.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()
  $script:ws = $ws; $script:cid = 0
}
function Disconnect-Tab { try { $script:ws.CloseAsync('NormalClosure', 'bye', [Threading.CancellationToken]::None).Wait(3000) | Out-Null } catch {}; $script:ws = $null }

function Receive-One($timeoutSec) {
  if (-not $script:ws -or $script:ws.State -ne 'Open') { throw "websocket not open" }
  $buf = New-Object byte[] 1048576
  $ms = New-Object IO.MemoryStream
  do {
    $r = $script:ws.ReceiveAsync([ArraySegment[byte]]$buf, [Threading.CancellationToken]::None)
    if (-not $r.Wait($timeoutSec * 1000)) { return $null }
    $ms.Write($buf, 0, $r.Result.Count)
  } while (-not $r.Result.EndOfMessage)
  return $script:Ser.DeserializeObject([Text.Encoding]::UTF8.GetString($ms.ToArray()))
}

function Cdp($method, $params = @{}, $timeoutSec = 60) {
  $script:cid++; $myId = $script:cid
  $msg = $script:Ser.Serialize(@{ id = $myId; method = $method; params = $params })
  $b = [Text.Encoding]::UTF8.GetBytes($msg)
  $script:ws.SendAsync([ArraySegment[byte]]$b, 'Text', $true, [Threading.CancellationToken]::None).Wait()
  $end = (Get-Date).AddSeconds($timeoutSec)
  while ($true) {
    $left = [Math]::Max(1, [int]($end - (Get-Date)).TotalSeconds)
    $o = Receive-One $left
    if ($null -eq $o) { throw "timeout: $method" }
    if ($o.ContainsKey('id') -and $o['id'] -eq $myId) {
      if ($o.ContainsKey('error')) { throw "$method failed: $($o['error']['message'])" }
      return $o['result']
    }
    elseif ($o.ContainsKey('method')) { [void]$script:Events.Add($o) }
    if ((Get-Date) -gt $end) { throw "timeout: $method" }
  }
}

function Js($expr, $timeoutSec = 60) {
  $r = Cdp 'Runtime.evaluate' @{ expression = $expr; returnByValue = $true; awaitPromise = $true } $timeoutSec
  if ($r.ContainsKey('exceptionDetails')) { $d = $r['exceptionDetails']; $m = $d['text']; if ($d.ContainsKey('exception')) { $m += ' ' + $d['exception']['description'] }; throw "JS error: $m" }
  $res = $r['result']; if ($res.ContainsKey('value')) { return $res['value'] } else { return $null }
}

function WaitJs($cond, $sec = 20) { for ($i = 0; $i -lt $sec * 4; $i++) { try { if (Js $cond 10) { return $true } } catch {}; Start-Sleep -Milliseconds 250 }; return $false }
