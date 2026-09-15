param (
    [string]$TargetFolder = $PSScriptRoot
)

if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
    $TargetFolder = (Get-Location).Path
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ">>> DANG THEO DOI THU MUC DU LIEU:" -ForegroundColor Yellow
Write-Host "    $TargetFolder" -ForegroundColor Green
Write-Host ">>> Khi co file Excel thay doi, script se tu dong push." -ForegroundColor Gray
Write-Host "    (Nhan Ctrl + C de dung)" -ForegroundColor DarkGray
Write-Host "========================================================" -ForegroundColor Cyan

Set-Location -Path $TargetFolder
$gitCheck = git status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "[X] Loi: Thu muc nay chua duoc ket noi Git!" -ForegroundColor Red
    Write-Host $gitCheck -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    exit
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

    if ($path.Contains("\.git\") -or $fileName.StartsWith("~$")) {
        return
    }

    if ($filterExtensions -contains $extension) {
        $timestamp = (Get-Date).ToString("HH:mm:ss")
        Write-Host "[$timestamp] Phat hien thay doi: $fileName ($changeType)" -ForegroundColor Magenta
        $global:hasChanges = $true
        $global:lastChangeTime = [DateTime]::Now
    }
}

Register-ObjectEvent $watcher 'Changed' -Action $action | Out-Null
Register-ObjectEvent $watcher 'Created' -Action $action | Out-Null
Register-ObjectEvent $watcher 'Deleted' -Action $action | Out-Null
Register-ObjectEvent $watcher 'Renamed' -Action $action | Out-Null

$global:hasChanges = $false
$global:lastChangeTime = [DateTime]::MinValue
$debounceSeconds = 3

try {
    while ($true) {
        Start-Sleep -Seconds 1

        if ($global:hasChanges) {
            $elapsed = ([DateTime]::Now - $global:lastChangeTime).TotalSeconds
            if ($elapsed -ge $debounceSeconds) {
                $global:hasChanges = $false
                $timeStr = (Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
                $curTime = (Get-Date).ToString("HH:mm:ss")
                Write-Host "[$curTime] Dang kiem tra thay doi..." -ForegroundColor Yellow
                
                $status = git status --porcelain
                if ([string]::IsNullOrWhiteSpace($status)) {
                    Write-Host "[$curTime] Khong co thay doi nao can commit." -ForegroundColor Gray
                    continue
                }

                Write-Host "[$curTime] Dang tu dong day (Push) du lieu len GitHub..." -ForegroundColor Cyan
                git add .
                git commit -m "Auto sync data: $timeStr"
                
                $pushOutput = git push origin main 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "[$curTime] [OK] DA DONG BO THANH CONG LEN GITHUB!" -ForegroundColor Green
                } else {
                    $pushMaster = git push origin master 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "[$curTime] [OK] DA DONG BO THANH CONG (nhanh master)!" -ForegroundColor Green
                    } else {
                        Write-Host "[$curTime] [X] Loi khi Push:" -ForegroundColor Red
                        Write-Host $pushOutput -ForegroundColor Red
                    }
                }
                Write-Host "--------------------------------------------------------" -ForegroundColor DarkGray
            }
        }
    }
}
finally {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Get-EventSubscriber | Unregister-Event
    Write-Host ""
    Write-Host "Da dung theo doi." -ForegroundColor Yellow
}
