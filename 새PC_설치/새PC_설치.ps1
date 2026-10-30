# 새 PC 설치 (처음 한 번만) — 이 폴더(새PC_설치)를 새 PC 아무 곳에 두고 실행
#  1) 엑셀 확인  2) Git 설치(없으면)  3) C:\ai\도구·CLAUDE.md·클로드 메모리 복사
#  4) GitHub 저장소 받기(C:\ai_github\date) → C:\ai\기초자료 로 복사  5) 10분 자동 동기화 등록
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoUrl = "https://github.com/a01042102350/date"; $repo = "C:\ai_github\date"

"1) 엑셀 확인"
if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application')) { throw "이 PC에 엑셀이 없습니다. 엑셀을 먼저 설치해 주세요." }
"   엑셀 있음"

"2) Git 확인"
$git = "C:\Program Files\Git\cmd\git.exe"
if (-not (Test-Path $git)) { "   Git 설치 중..."; winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements --silent; if (-not (Test-Path $git)) { throw "Git 설치 실패 — https://git-scm.com 에서 직접 설치해 주세요." } }
"   Git 있음"

"3) 도구·규칙·메모리 복사"
New-Item -ItemType Directory -Force "C:\ai\도구", "C:\ai\기초자료" | Out-Null
Copy-Item "$here\도구\*" "C:\ai\도구" -Force
Copy-Item "$here\CLAUDE.md" "C:\ai\CLAUDE.md" -Force
$mem = Join-Path $env:USERPROFILE ".claude\projects\c--ai\memory"; New-Item -ItemType Directory -Force $mem | Out-Null
Copy-Item "$here\클로드메모리\*" $mem -Force
"   C:\ai\도구, C:\ai\CLAUDE.md, 클로드 메모리($mem) 복사 완료"

"4) GitHub 저장소 받기"
if (-not (Test-Path "$repo\.git")) { New-Item -ItemType Directory -Force (Split-Path $repo) | Out-Null; & $git clone $repoUrl $repo }
else { & $git -C $repo pull --ff-only origin main }
& $git -C $repo config user.name "a01042102350"; & $git -C $repo config core.quotepath false
if (Test-Path "$repo\기초자료") { robocopy "$repo\기초자료" "C:\ai\기초자료" /E /XO /XF "~$*" /NFL /NDL /NJH /NJS /NP | Out-Null }
if (Test-Path "$repo\matching.md") { Copy-Item "$repo\matching.md" "C:\ai\matching.md" -Force }
"   C:\ai\기초자료 에 자료 복사 완료"

"5) 10분 자동 동기화 등록"
$p = "C:\ai\도구\깃허브_자동동기화.ps1"
$act = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$p`""
$trg = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10)
$set = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
Register-ScheduledTask -TaskName "월마감_깃허브_동기화" -Action $act -Trigger $trg -Settings $set -Description "C:\ai\기초자료 ↔ GitHub 10분마다 동기화" -Force | Out-Null
"   등록 완료"
""
"설치 끝. 처음 GitHub에 올릴 때 로그인 창이 뜨면 GitHub 계정으로 로그인해 주세요."
"VS Code에서 C:\ai 폴더를 열고 클로드에게 말하면 됩니다. 예) 202610 반영해줘"
