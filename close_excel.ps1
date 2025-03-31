# Path to the shared file
$FilePath = "\\nas10\s_estomatologia_Escalas_urgencias$\0. BLOCOS 2025\AL.xlxs"

# Monitor Excel processes
While ($true) {
    # Check if the shared file exists
    if (Test-Path $FilePath) {
        # Get the last modified time of the file
        $FileLastModified = (Get-Item $FilePath).LastWriteTime
        $TimeSinceLastModified = (Get-Date) - $FileLastModified

        # Get running Excel processes
        $ExcelProcesses = Get-Process -Name excel -ErrorAction SilentlyContinue

        if ($ExcelProcesses) {
            # If the file hasn't been modified within the threshold, save changes and close Excel
            if ($TimeSinceLastModified.TotalMinutes -ge 1) { # Replace 1 with your desired idle threshold
                Write-Host "File has not been modified for $([math]::Round($TimeSinceLastModified.TotalMinutes, 2)) minutes. Saving changes and closing Excel."

                # Interact with Excel COM object
                $ExcelApp = New-Object -ComObject Excel.Application
                $ExcelApp.Visible = $false # Run Excel in the background

                # Loop through open workbooks
                foreach ($Workbook in $ExcelApp.Workbooks) {
                    try {
                        $Workbook.Save() # Save the workbook
                        Write-Host "Saved workbook: $($Workbook.FullName)"
                    } catch {
                        Write-Host "Failed to save workbook: $($Workbook.FullName). Error: $_"
                    }
                }

                # Quit Excel and clean up
                $ExcelApp.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null
                Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
            } else {
                Write-Host "File was modified recently ($([math]::Round($TimeSinceLastModified.TotalMinutes, 2)) minutes ago). Excel remains open."
            }
        } else {
            Write-Host "No Excel processes are running."
        }
    } else {
        Write-Host "The file path $FilePath does not exist. Cannot monitor."
    }

    Start-Sleep -Seconds 30
}