# Path to the shared file
$FilePath = "C:\Users\61840\Desktop\teste.xlsx"
$IdleThresholdMinutes = 1

# Normalize the path
$NormalizedFilePath = [System.IO.Path]::GetFullPath($FilePath)

Write-Host "versao 4 - Monitoring file: $NormalizedFilePath"

While ($true) {
    if (Test-Path $NormalizedFilePath) {
        $FileItem = Get-Item $NormalizedFilePath

        $FileLastModified = $FileItem.LastWriteTime
        $FileLastAccessed = $FileItem.LastAccessTime

        $TimeSinceLastModified = (Get-Date) - $FileLastModified
        $TimeSinceLastAccessed = (Get-Date) - $FileLastAccessed

        Write-Host "Time since last modified: $([math]::Round($TimeSinceLastModified.TotalMinutes,2)) minutes"
        Write-Host "Time since last accessed: $([math]::Round($TimeSinceLastAccessed.TotalMinutes,2)) minutes"

        if ($TimeSinceLastModified.TotalMinutes -ge $IdleThresholdMinutes -and $TimeSinceLastAccessed.TotalMinutes -ge $IdleThresholdMinutes) {
            Write-Host "Threshold exceeded for both modified and accessed. Searching for Excel processes..."

            # Get all excel.exe processes
            $ExcelProcesses = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "EXCEL.EXE" }

            $TargetProcessFound = $false

            foreach ($Process in $ExcelProcesses) {
                if ($Process.CommandLine) {
                    Write-Host "Process command line: $($Process.CommandLine)"

                    # Check if the process is using the file
                    if ($Process.CommandLine -like "*$($NormalizedFilePath)*") {
                        Write-Host "Target Excel process found. PID: $($Process.ProcessId). Closing it..."

                        # Close only that process
                        Stop-Process -Id $Process.ProcessId -Force -ErrorAction SilentlyContinue

                        $TargetProcessFound = $true
                    }
                }
            }

            if (-not $TargetProcessFound) {
                Write-Host "No Excel process is using the target file."
            }

            # After acting (closing or not), proceed without waiting too long
            Write-Host "Waiting 10 seconds before next check after action..."
            Start-Sleep -Seconds 10
        }
        else {
            Write-Host "File modified or accessed recently. No action needed."
            Start-Sleep -Seconds 10
        }
    }
    else {
        Write-Host "File path does not exist: $NormalizedFilePath"
        Start-Sleep -Seconds 10
    }
}
