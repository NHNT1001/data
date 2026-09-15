param (
    [string]$TargetFolder = ""
)

# Tu dong xac dinh thu muc: Uu tien D:\Antigravity Tai\Github\data git neu ton tai, neu khong lay thu muc chua file script
if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
    if (Test-Path "D:\Antigravity Tai\Github\data") {
        $TargetFolder = "D:\Antigravity Tai\Github\data"
    } else {
        $TargetFolder = $PSScriptRoot
    }
}

if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
    $TargetFolder = (Get-Location).Path
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$logFile = Join-Path $TargetFolder "auto_push.log"

function Write-Log([string]$msg, [string]$color = "White") {
    $timestamp = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
    $logMsg = "[$timestamp] $msg"
    try {
        Write-Host $logMsg -ForegroundColor $color
    } catch {}
    try {
        Add-Content -Path $logFile -Value $logMsg -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

Write-Log "========================================================" "Cyan"
Write-Log ">>> BAT DAU THEO DOI THU MUC: $TargetFolder" "Yellow"

Set-Location -Path $TargetFolder
$gitCheck = git status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Log "[X] Loi Git: $gitCheck" "Red"
    exit 1
}

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $TargetFolder
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite

$filterExtensions = @('.xlsx', '.xls', '.csv')

$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    $fileName = [System.IO.Path]::GetFileName($path)
    $extension = [System.IO.Path]::GetExtension($path).ToLower()

    if ($path.Contains("\.git\") -or $fileName.StartsWith("~$") -or $fileName -eq "auto_push.log") {
        return
    }

    if ($filterExtensions -contains $extension) {
        $timestamp = (Get-Date).ToString("HH:mm:ss")
        $global:hasChanges = $true
        $global:lastChangedFile = $fileName
        $global:lastChangeTime = [DateTime]::Now
    }
}

Register-ObjectEvent $watcher 'Changed' -Action $action | Out-Null
Register-ObjectEvent $watcher 'Created' -Action $action | Out-Null
Register-ObjectEvent $watcher 'Deleted' -Action $action | Out-Null
Register-ObjectEvent $watcher 'Renamed' -Action $action | Out-Null

$global:hasChanges = $false
$global:lastChangedFile = ""
$global:lastChangeTime = [DateTime]::MinValue
$debounceSeconds = 3

Write-Log ">>> He thong san sang! Dang cho thay doi file Excel..." "Green"

try {
    while ($true) {
        Start-Sleep -Seconds 1

        if ($global:hasChanges) {
            $elapsed = ([DateTime]::Now - $global:lastChangeTime).TotalSeconds
            if ($elapsed -ge $debounceSeconds) {
                $global:hasChanges = $false
                $changedFile = $global:lastChangedFile
                Write-Log "Phat hien file thay doi: $changedFile" "Magenta"
                Write-Log "Dang kiem tra git status..." "Yellow"
                
                $status = git status --porcelain
                if ([string]::IsNullOrWhiteSpace($status)) {
                    Write-Log "Khong co thay doi thuc su nao can commit." "Gray"
                    continue
                }

                $timeStr = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
                Write-Log "Dang tu dong Push len GitHub..." "Cyan"
                git add .
                git commit -m "Auto sync data: $timeStr"
                
                $pushOutput = git push origin main 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "[OK] DA DONG BO THANH CONG LEN GITHUB (main)!" "Green"
                } else {
                    $pushMaster = git push origin master 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "[OK] DA DONG BO THANH CONG LEN GITHUB (master)!" "Green"
                    } else {
                        Write-Log "[X] Loi khi Push: $pushOutput" "Red"
                    }
                }
                Write-Log "--------------------------------------------------------" "DarkGray"
            }
        }
    }
}
finally {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Get-EventSubscriber | Unregister-Event
    Write-Log "Da dung theo doi." "Yellow"
}
