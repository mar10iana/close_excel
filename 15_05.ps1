# Caminho para o ficheiro partilhado
$FilePath = "\\nas10\s_estomatologia_Escalas_urgencias`$\0. BLOCOS 2025\AL.xlsx"
$IdleThresholdMinutes = 1
$GracePeriodMinutes = 1

# Dicionário de PIDs e os seus tempos de aviso e adiamento
$WarnedPIDs = @{}

Add-Type -AssemblyName PresentationFramework

Write-Host "Versão 7 - Monitor com opção de adiar fecho por 5 minutos"

while ($true) {
    $ExcelProcesses = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq "EXCEL.EXE" }

    foreach ($Process in $ExcelProcesses) {
        if ($Process.CommandLine -and $Process.CommandLine -like "*$($FilePath)*") {
            $StartTime = [Management.ManagementDateTimeConverter]::ToDateTime($Process.CreationDate)
            $RunningTime = (Get-Date) - $StartTime
            $RunningMinutes = [math]::Round($RunningTime.TotalMinutes, 2)
            $PID = $Process.ProcessId

            Write-Host "Process PID $PID using file for $RunningMinutes minutes"

            if ($RunningTime.TotalMinutes -ge $IdleThresholdMinutes) {
                if (-not $WarnedPIDs.ContainsKey($PID)) {
                    # Primeira vez que o processo ultrapassa o limite
                    $response = [System.Windows.MessageBox]::Show(
                        "O ficheiro '$FilePath' está aberto há mais de $IdleThresholdMinutes minutos.`nPretende adiar o fecho por mais $GracePeriodMinutes minutos?",
                        "Aviso de fecho automático",
                        'YesNo',
                        'Warning'
                    )

                    if ($response -eq 'Yes') {
                        $WarnedPIDs[$PID] = Get-Date
                        Write-Host "Utilizador adiou fecho para PID $PID. Tempo extra concedido."
                    }
                    else {
                        $WarnedPIDs[$PID] = (Get-Date).AddMinutes(-$GracePeriodMinutes - 1)
                        Write-Host "Utilizador recusou adiar. O processo será encerrado após o tempo de graça."
                    }
                }
                else {
                    $WarnTime = $WarnedPIDs[$PID]
                    $ElapsedSinceWarning = (Get-Date) - $WarnTime
                    if ($ElapsedSinceWarning.TotalMinutes -ge $GracePeriodMinutes) {
                        Write-Host "Tempo de graça terminado para PID $PID. Fechando..."
                        Stop-Process -Id $PID -Force -ErrorAction SilentlyContinue
                        $WarnedPIDs.Remove($PID)
                    }
                    else {
                        $Remaining = [math]::Round($GracePeriodMinutes - $ElapsedSinceWarning.TotalMinutes, 1)
                        Write-Host "Ainda em tempo de graça para PID $PID. Faltam $Remaining minutos."
                    }
                }
            }
        }
    }

    # Remover PIDs que já não existem
    $ExistingPIDs = $ExcelProcesses.ProcessId
    $PIDsToRemove = $WarnedPIDs.Keys | Where-Object { $_ -notin $ExistingPIDs }
    foreach ($OldPID in $PIDsToRemove) {
        $WarnedPIDs.Remove($OldPID)
    }

    Start-Sleep -Seconds 10
}
