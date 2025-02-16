param (
    [string]$sourceDirectory,
    [string]$backupDirectory,
    [int]$threads = 8,
    [bool]$updateOneDrive = $true,
    [string]$password = ""
)


Import-Module Microsoft.PowerShell.Utility

write-host "***********************************"
write-host "*** Starting file backup script ***"
write-host "***********************************"


# check dirs
if (!(Test-Path $sourceDirectory)){
    Write-Host "Source directory does not exist"
    exit
}

if (!(Test-Path $backupDirectory)){
    Write-Host "Bakup directory does not exist"
    exit
}


# Define the 7-Zip executable path
$sevenZipPathDir = "C:\'Program Files'\7-zip\"

$sevenZipPath = $sevenZipPathDir + "7z.exe"
$currentDir = $PSScriptRoot


$newBackups = 0
if (-not [string]::IsNullOrEmpty($password)) {
    $pSwitch = "-p$password"
} else {
    $filePath = Join-Path -Path $currentDir -ChildPath "pwd.txt"
    $passString = Get-Content -Path $filePath -ErrorAction SilentlyContinue

    # Only include the -p parameter if $passString is not empty
    if (-not [string]::IsNullOrEmpty($passString)) {
        $pSwitch = "-p$passString"
    } else {
        $pSwitch = ""
    }
}

$newBackups = 0
Get-ChildItem -Path $sourceDirectory -Directory | Sort-Object LastWriteTime -Descending | ForEach-Object {
    $zipFile = Join-Path -Path $backupDirectory -ChildPath ($_.Name + ".zip")
    $SubdirectoryPath = $_.FullName
    Write-Output "Processing directory: $SubdirectoryPath"
    
    $process = $false
    if (-Not [System.IO.File]::Exists($zipFile)) {
        $process = $true
    } else {
        $latest = Get-ChildItem $SubdirectoryPath | Sort-Object -Descending -Property LastWriteTime -Top 1 
        if ($latest.LastWriteTime -gt (Get-Item $zipFile).LastWriteTime) { 
            $process = $true
            if (Test-Path $zipFile) { Remove-Item $zipFile }
        } 
    }

    if ($process -eq $true) {
        $ignoreList = Get-ChildItem -Path $SubdirectoryPath -Recurse -File -Filter "BACKUP.ignore" -Depth 5 
        $7ZipIgnore = $ignoreList | Join-String -Property { $_.DirectoryName.Substring($SubdirectoryPath.Length).TrimStart('\') } -Separator ' -xr!'
        if ($7ZipIgnore.length -gt 0) { $7ZipIgnore = "-xr!$7ZipIgnore" }


        $cmd = "$sevenZipPath a `"$zipFile`" `"$SubdirectoryPath\*`" -tzip -mx=$threads $pSwitch -mem=AES256 -bsp1 -uq0 $7ZipIgnore"
        
        write-host $cmd
        Invoke-Expression $cmd   
        $newBackups += 1     
    }
}

Write-Output "Total new backups created: $newBackups"


# Define the onedrive path
if($updateOneDrive) {
    Write-Host "stopping ms services for sync"
    $processesToStop = @(
        "csisyncclinet",
        "groove",
        "msosync",
        "msouc",
        "sysdrive"
    )

    foreach ($process in $processesToStop) {
        $runningProcesses = Get-Process | Where-Object { $_.Name -eq $process }
        if($runningProcesses) {
            Stop-Process -Name $process -Force
        }
    }
    
    Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache\*" -Force -Recurse

    Start-Process -FilePath (Get-Process OneDrive -ErrorAction SilentlyContinue | Select-Object -First 1).Path -ArgumentList "/sync"

}