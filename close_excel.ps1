# Caminho para o ficheiro partilhado
$FilePath = "\\nas10\s_estomatologia_Escalas_urgencias`$\0. BLOCOS 2025\AL.xlsx"
$IdleThresholdMinutes = 1
$GracePeriodMinutes = 1

# Dicion�rios por PID
$WarnedPIDs = @{}
$FallbackStartTimes = @{}

Add-Type -AssemblyName System.Windows.Forms

function Show-TimedPrompt {
    param (
        [string]$message,
        [string]$title,
        [int]$timeoutSeconds = 30
    )

    $form = New-Object Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object Drawing.Size(400,150)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.ControlBox = $false


    $label = New-Object Windows.Forms.Label
    $label.Text = $message
    $label.Size = New-Object Drawing.Size(360,60)
    $label.Location = New-Object Drawing.Point(20,10)
    $form.Controls.Add($label)

    $yesButton = New-Object Windows.Forms.Button
    $yesButton.Text = "Adiar"
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $yesButton.Location = New-Object Drawing.Point(80,80)
    $form.Controls.Add($yesButton)

    $noButton = New-Object Windows.Forms.Button
    $noButton.Text = "Fechar"
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $noButton.Location = New-Object Drawing.Point(200,80)
    $form.Controls.Add($noButton)

    $form.AcceptButton = $yesButton
    $form.CancelButton = $noButton

    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = $timeoutSeconds * 1000
    $timer.Add_Tick({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::No
        $form.Close()
    })
    $timer.Start()

    $result = $form.ShowDialog()
    $timer.Stop()
    return $result
}

Write-Host "Monitor com aviso com timeout (30s) e adiamento de fecho"

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
            } catch {}

            if (-not $StartTime) {
                if ($FallbackStartTimes.ContainsKey($ProcID)) {
                    $StartTime = $FallbackStartTimes[$ProcID]
                } else {
                    $StartTime = Get-Date
                    $FallbackStartTimes[$ProcID] = $StartTime
                    Write-Host "Processo PID $ProcID sem CreationDate v�lida. Usar hora atual como in�cio."
                }
            }

            $RunningTime = (Get-Date) - $StartTime
            $RunningMinutes = [math]::Round($RunningTime.TotalMinutes, 2)
            Write-Host "Process PID $ProcID usando o ficheiro h� $RunningMinutes minutos."

            if ($RunningTime.TotalMinutes -ge $IdleThresholdMinutes) {
                if (-not $WarnedPIDs.ContainsKey($ProcID)) {
                    $response = Show-TimedPrompt -message "O ficheiro '$FilePath' est� aberto h� mais de $IdleThresholdMinutes minutos.`nPretende adiar o fecho por mais $GracePeriodMinutes minutos?" -title "Aviso de fecho autom�tico" -timeoutSeconds 30

                    if ($response -eq 'Yes') {
                        $WarnedPIDs[$ProcID] = Get-Date
                        Write-Host "Utilizador adiou fecho para PID $ProcID."
                    } else {
                        $WarnedPIDs[$ProcID] = (Get-Date).AddMinutes(-$GracePeriodMinutes - 1)
                        Write-Host "Sem resposta ou utilizador recusou adiar. Fechar� ap�s toler�ncia."
                    }
                }
                else {
                    $WarnTime = $WarnedPIDs[$ProcID]
                    $Elapsed = (Get-Date) - $WarnTime
                    if ($Elapsed.TotalMinutes -ge $GracePeriodMinutes) {
                        Write-Host "Fecho autom�tico do Excel PID $ProcID..."
                        Stop-Process -Id $ProcID -Force -ErrorAction SilentlyContinue
                        $WarnedPIDs.Remove($ProcID)
                        $FallbackStartTimes.Remove($ProcID)
                    }
                    else {
                        $Remaining = [math]::Round($GracePeriodMinutes - $Elapsed.TotalMinutes, 1)
                        Write-Host "Aguardando fim da toler�ncia: $Remaining minutos restantes."
                    }
                }
            }
        }
    }

    # Limpeza de PIDs terminados
    $ExistingPIDs = $ExcelProcesses.ProcessId
    $ToRemove = $WarnedPIDs.Keys | Where-Object { $_ -notin $ExistingPIDs }
    foreach ($OldPID in $ToRemove) {
        $WarnedPIDs.Remove($OldPID)
        $FallbackStartTimes.Remove($OldPID)
    }

    Start-Sleep -Seconds 10
}