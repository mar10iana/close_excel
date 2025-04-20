# Path to the shared file
$FilePath = "C:\Users\61840\Desktop\teste.xlsx"
$IdleThresholdMinutes = 1

# Normalizar o caminho (evitar problemas com barras)
$NormalizedFilePath = [System.IO.Path]::GetFullPath($FilePath)

Write-Host "versao 3 Monitoring file: $NormalizedFilePath"

While ($true) {
    if (Test-Path $NormalizedFilePath) {
        $FileLastModified = (Get-Item $NormalizedFilePath).LastWriteTime
        $TimeSinceLastModified = (Get-Date) - $FileLastModified

        Write-Host "Time since last modified: $([math]::Round($TimeSinceLastModified.TotalMinutes,2)) minutes"

        if ($TimeSinceLastModified.TotalMinutes -ge $IdleThresholdMinutes) {
            Write-Host "Threshold exceeded. Searching for Excel processes..."

            # Obter todos os processos excel.exe
            $ExcelProcesses = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "EXCEL.EXE" }

            $TargetProcessFound = $false

            foreach ($Process in $ExcelProcesses) {
                if ($Process.CommandLine) {
                    Write-Host "Process command line: $($Process.CommandLine)"

                    # Verificar se o processo está a usar o ficheiro
                    if ($Process.CommandLine -like "*$($NormalizedFilePath)*") {
                        Write-Host "Target Excel process found. PID: $($Process.ProcessId). Closing it..."

                        # Fechar apenas esse processo
                        Stop-Process -Id $Process.ProcessId -Force -ErrorAction SilentlyContinue

                        $TargetProcessFound = $true
                    }
                }
            }

            if (-not $TargetProcessFound) {
                Write-Host "No Excel process is using the target file."
            }

            # Depois de agir (fechar ou não), seguir direto sem esperar muito
            Write-Host "Waiting 10 seconds before next check after action..."
            Start-Sleep -Seconds 10
        }
        else {
            Write-Host "File modified recently. No action needed."
            Start-Sleep -Seconds 10
        }
    }
    else {
        Write-Host "File path does not exist: $NormalizedFilePath"
        Start-Sleep -Seconds 10
    }
}
