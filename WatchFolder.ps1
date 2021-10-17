

$FolderPath = "C:\STATUtilities\CSVOrderImport"

$watcher = New-Object System.IO.FileSystemWatcher

$watcher.IncludeSubdirectories = $false
$watcher.Path = "$FolderPath\watch\"
$watcher.EnableRaisingEvents = $true

$action = {
    $path = $event.SourceEventArgs.FullPath
    $changetype = $event.SourceEventArgs.ChangeType
    $script = "$FolderPath\ImportOrdersFromCsv.ps1"

    if((Get-Item $path).Extension -eq ".csv") {
        try {

            $params = @('-Path', """$path""", '-WorkingDir', """$FolderPath""")
        
            Invoke-Expression "& `"$script`" $params"
        
        }
        catch {
            Write-Host "*ERROR*  *ERROR*  *ERROR*"
            Write-Host "`r`n`r`n"
            Write-Host $_.Exception
        
            Write-Host "`r`n`r`n"
        }
    }

    Write-Host "$path was $changetype at $(get-date)"
}

Register-ObjectEvent $watcher 'Created' -Action $action

while(1) { sleep 1 }