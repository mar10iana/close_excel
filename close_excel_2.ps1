# Path to the shared file
$FilePath = "\\nas10\s_estomatologia_Escalas_urgencias$\0. BLOCOS 2025\AL.xlsx"

While ($true) {
    if (Test-Path $FilePath) {
        $FileLastModified = (Get-Item $FilePath).LastWriteTime
        $TimeSinceLastModified = (Get-Date) - $FileLastModified

        if ($TimeSinceLastModified.TotalMinutes -ge 1) { # Threshold

            try {
                # Connect to running Excel application
                $ExcelApp = [Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
                $WorkbookFound = $false

                foreach ($Workbook in $ExcelApp.Workbooks) {
                    if ($Workbook.FullName -eq $FilePath) {
                        Write-Host "Workbook found: $($Workbook.FullName). Saving and closing..."
                        $Workbook.Save()
                        $Workbook.Close()
                        $WorkbookFound = $true
                        break
                    }
                }

                if (-not $WorkbookFound) {
                    Write-Host "Workbook is not open. No action taken."
                }

                # Release COM object
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp) | Out-Null

            } catch {
                Write-Host "Excel is not open or workbook not found. Error: $_"
            }

        } else {
            Write-Host "File was modified recently ($([math]::Round($TimeSinceLastModified.TotalMinutes, 2)) minutes ago). No action."
        }

    } else {
        Write-Host "The file path $FilePath does not exist. Cannot monitor."
    }

    Start-Sleep -Seconds 30
}
