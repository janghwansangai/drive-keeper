# 드라이브 지킴이 설치 (관리자 권한 필요 없음)
# 프로그램을 %LOCALAPPDATA%\Programs\DriveKeeper 에 복사하고 바탕화면·시작 메뉴 바로가기를 만든다.
# -Quiet: 창을 띄우지 않고 진행 (자동 점검용)
param([switch]$Quiet)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
$Title = "드라이브 지킴이 설치"

function Ask($msg, $buttons = "OK", $icon = "Information", $quietAnswer = "OK") {
    if ($Quiet) { return $quietAnswer }
    return [string][System.Windows.Forms.MessageBox]::Show($msg, $Title, $buttons, $icon)
}

try {
    $src = $PSScriptRoot                                   # ...\드라이브 지킴이\app
    $exe = Join-Path $src "DriveKeeper.exe"
    if (-not (Test-Path $exe)) { Ask "DriveKeeper.exe를 찾을 수 없어요.`n압축을 모두 푼 뒤 '설치하기'를 다시 실행해 주세요." "OK" "Error" | Out-Null; exit 1 }
    $dest = Join-Path $env:LOCALAPPDATA "Programs\DriveKeeper"

    while (Get-Process -Name DriveKeeper -ErrorAction SilentlyContinue) {
        if ($Quiet) { throw "드라이브 지킴이가 켜져 있음" }
        if ((Ask "드라이브 지킴이가 켜져 있어요.`n프로그램을 닫은 뒤 [확인]을 눌러 주세요." "OKCancel" "Warning") -ne "OK") { exit 1 }
    }

    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    # 새 버전으로 바꿀 때 옛 파일이 남지 않게 통째로 맞춘다 (사용자 데이터는 다른 폴더라 안전)
    & robocopy $src $dest /MIR /NFL /NDL /NJH /NJS /NP /R:2 /W:1 | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "파일 복사 실패 (robocopy $LASTEXITCODE)" }
    Get-ChildItem $dest -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

    $target = Join-Path $dest "DriveKeeper.exe"
    $shell = New-Object -ComObject WScript.Shell
    foreach ($dir in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs"))) {
        $lnk = $shell.CreateShortcut((Join-Path $dir "드라이브 지킴이.lnk"))
        $lnk.TargetPath = $target
        $lnk.WorkingDirectory = $dest
        $lnk.IconLocation = "$target,0"
        $lnk.Description = "구글 드라이브를 비밀번호로 지키는 프로그램"
        $lnk.Save()
    }

    # Windows [설정] → [앱] 목록에 보이게 (이 사용자에게만)
    $ver = (Get-Item $target).VersionInfo.ProductVersion
    $key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DriveKeeper"
    New-Item -Path $key -Force | Out-Null
    $un = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$dest\uninstall.ps1`""
    Set-ItemProperty $key -Name DisplayName -Value "드라이브 지킴이"
    Set-ItemProperty $key -Name DisplayIcon -Value $target
    Set-ItemProperty $key -Name Publisher -Value "다있쌤 로디"
    Set-ItemProperty $key -Name InstallLocation -Value $dest
    Set-ItemProperty $key -Name UninstallString -Value $un
    Set-ItemProperty $key -Name QuietUninstallString -Value "$un -Quiet"
    Set-ItemProperty $key -Name NoModify -Value 1 -Type DWord
    Set-ItemProperty $key -Name NoRepair -Value 1 -Type DWord
    if ($ver) { Set-ItemProperty $key -Name DisplayVersion -Value $ver }

    if ((Ask "설치가 끝났어요! 🎉`n`n바탕화면의 '드라이브 지킴이'를 두 번 누르면 시작해요.`n지금 바로 켤까요?" "YesNo" "Information" "No") -eq "Yes") {
        Start-Process $target -WorkingDirectory $dest
    }
} catch {
    Ask "설치하지 못했어요.`n`n$($_.Exception.Message)" "OK" "Error" | Out-Null
    if ($Quiet) { Write-Output "ERROR: $($_.Exception.Message)" }
    exit 1
}
