$folder = '.\INPUT' # Enter the root path you want to monitor. 
$filter = '*.txt'  # You can enter a wildcard filter here. 

$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property  @{IncludeSubdirectories = $true;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'}

Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
    $testing = 'C:\Users\aizat\Desktop\TOT 2\maybank.exe'
    $Running = Get-Process decrypt -ErrorAction SilentlyContinue
    $Start = {([wmiclass]"win32_process").Create($decrypt)} 
    $name = $Event.SourceEventArgs.Name 
    $changeType = $Event.SourceEventArgs.ChangeType 
    $timeStamp = $Event.TimeGenerated 
    Write-Host "The file '$name' was $changeType at $timeStamp" -fore green
    #Start-Process -FilePath "decrypt.exe" -Wait -WindowStyle Hidden
    if($Running -eq $null){ # evaluating if the program is running
        & $Start
    }
}