# 드라이브 지킴이 지우기. 구글 드라이브의 파일은 건드리지 않는다.
# -Quiet: 창을 띄우지 않고 프로그램만 지운다 (로그인 정보·작업 폴더는 남김)
param([switch]$Quiet)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
$Title = "드라이브 지킴이 지우기"

function Ask($msg, $buttons = "OK", $icon = "Information", $quietAnswer = "OK", $default = "Button1") {
    if ($Quiet) { return $quietAnswer }
    return [string][System.Windows.Forms.MessageBox]::Show($msg, $Title, $buttons, $icon, $default)
}

try {
    if ((Ask "드라이브 지킴이를 이 컴퓨터에서 지울까요?`n(구글 드라이브의 파일은 그대로 있어요)" "YesNo" "Question" "Yes") -ne "Yes") { exit 0 }
    while (Get-Process -Name DriveKeeper -ErrorAction SilentlyContinue) {
        if ($Quiet) { throw "드라이브 지킴이가 켜져 있음" }
        if ((Ask "드라이브 지킴이가 켜져 있어요.`n프로그램을 닫은 뒤 [확인]을 눌러 주세요." "OKCancel" "Warning") -ne "OK") { exit 1 }
    }
    $dest = Join-Path $env:LOCALAPPDATA "Programs\DriveKeeper"
    foreach ($dir in @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs"))) {
        Remove-Item (Join-Path $dir "드라이브 지킴이.lnk") -ErrorAction SilentlyContinue
    }
    Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DriveKeeper" -Recurse -ErrorAction SilentlyContinue

    $data = Join-Path $env:USERPROFILE "DriveKeeper"
    if (Test-Path $data) {
        $msg = "로그인 정보와 작업 폴더(받아 둔 파일 사본)도 지울까요?`n`n$data`n`n아직 드라이브에 올리지 못한 변경이 있다면 [아니요]를 누르고, 프로그램으로 먼저 올려 주세요."
        if ((Ask $msg "YesNo" "Warning" "No" "Button2") -eq "Yes") { Remove-Item $data -Recurse -Force -ErrorAction SilentlyContinue }
    }
    # 실행 중인 이 스크립트가 들어 있는 폴더는 잠시 뒤에 지운다
    if (Test-Path $dest) {
        Start-Process cmd.exe -ArgumentList "/c timeout /t 2 /nobreak >nul & rmdir /s /q `"$dest`"" -WindowStyle Hidden
    }
    Ask "다 지웠어요. 그동안 써 주셔서 고마워요!" | Out-Null
} catch {
    Ask "지우지 못했어요.`n`n$($_.Exception.Message)" "OK" "Error" | Out-Null
    if ($Quiet) { Write-Output "ERROR: $($_.Exception.Message)" }
    exit 1
}
