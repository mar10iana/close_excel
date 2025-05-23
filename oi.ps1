# Caminho para o ficheiro partilhado lala
$FilePath = "\\nas10\s_estomatologia_Escalas_urgencias`$\0. BLOCOS 2025\AL.xlsx"
$IdleThresholdMinutes = 5
$GracePeriodMinutes = 5

# Dicionários de estado por PID
$WarnedPIDs = @{}
$FallbackStartTimes = @{}

Add-Type -AssemblyName PresentationFramework

Write-Host "Versão 9 - StartTime persistente para CreationDate inválido"

while ($true) {
    $ExcelProcesses = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "EXCEL.EXE" }

    foreach ($Process in $ExcelProcesses) {
        if ($Process.CommandLine -and $Process.CommandLine -like "*$($FilePath)*") {

            $ProcID = $Process.ProcessId
            $StartTime = $null

            try {
                if ($Process.CreationDate) {
                    $StartTime = [Management.ManagementDateTimeConverter]::ToDateTime($Process.CreationDate)
                }
            } catch {
                # Ignorado
            }

            # Fallback se não conseguimos obter StartTime
            if (-not $StartTime) {
                if ($FallbackStartTimes.ContainsKey($ProcID)) {
                    $StartTime = $FallbackStartTimes[$ProcID]
                } else {
                    $StartTime = Get-Date
                    $FallbackStartTimes[$ProcID] = $StartTime
                    Write-Host "Processo PID $ProcID tem CreationDate inválida. A usar hora atual como fallback inicial."
                }
            }

            $RunningTime = (Get-Date) - $StartTime
            $RunningMinutes = [math]::Round($RunningTime.TotalMinutes, 2)

            Write-Host "Process PID $ProcID using file for $RunningMinutes minutes"

            if ($RunningTime.TotalMinutes -ge $IdleThresholdMinutes) {
                if (-not $WarnedPIDs.ContainsKey($ProcID)) {
                    $response = [System.Windows.MessageBox]::Show(
                        "O ficheiro '$FilePath' está aberto há mais de $IdleThresholdMinutes minutos.`nPretende adiar o fecho por mais $GracePeriodMinutes minutos?",
                        "Aviso de fecho automático",
                        'YesNo',
                        'Warning'
                    )

                    if ($response -eq 'Yes') {
                        $WarnedPIDs[$ProcID] = Get-Date
                        Write-Host "Utilizador adiou fecho para PID $ProcID. Tempo extra concedido."
                    }
                    else {
                        $WarnedPIDs[$ProcID] = (Get-Date).AddMinutes(-$GracePeriodMinutes - 1)
                        Write-Host "Utilizador recusou adiar. O processo será encerrado após tempo de graça."
                    }
                }
                else {
                    $WarnTime = $WarnedPIDs[$ProcID]
                    $ElapsedSinceWarning = (Get-Date) - $WarnTime
                    if ($ElapsedSinceWarning.TotalMinutes -ge $GracePeriodMinutes) {
                        Write-Host "Tempo de graça terminado para PID $ProcID. Fechando..."
                        Stop-Process -Id $ProcID -Force -ErrorAction SilentlyContinue
                        $WarnedPIDs.Remove($ProcID)
                        $FallbackStartTimes.Remove($ProcID)
                    }
                    else {
                        $Remaining = [math]::Round($GracePeriodMinutes - $ElapsedSinceWarning.TotalMinutes, 1)
                        Write-Host "Ainda em tempo de graça para PID $ProcID. Faltam $Remaining minutos."
                    }
                }
            }
        }
    }

    # Limpar PIDs que já não existem
    $ExistingPIDs = $ExcelProcesses.ProcessId
    $PIDsToRemove = $WarnedPIDs.Keys | Where-Object { $_ -notin $ExistingPIDs }
    foreach ($OldPID in $PIDsToRemove) {
        $WarnedPIDs.Remove($OldPID)
        $FallbackStartTimes.Remove($OldPID)
    }

    Start-Sleep -Seconds 10
}
