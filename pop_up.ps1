Add-Type -AssemblyName System.Windows.Forms

# Caminho para o ficheiro
$FilePath = "C: \Users\61840\Desktop\teste.xlsx"

# Diretório onde o ficheiro está
$FolderPath = Split-Path $FilePath

# Threshold de inatividade
$IdleThresholdMinutes = 1

While ($true) {
    if (Test-Path $FilePath) {
        $FileLastModified = (Get-Item $FilePath).LastWriteTime
        $TimeSinceLastModified = (Get-Date) - $FileLastModified

        if ($TimeSinceLastModified.TotalMinutes -ge $IdleThresholdMinutes) {
            Write-Host "O ficheiro não foi alterado há $([math]::Round($TimeSinceLastModified.TotalMinutes,2)) minutos. Mostrar alerta."

            # Criar a janela pop-up
            $Form = New-Object System.Windows.Forms.Form
            $Form.Text = "Alerta: Ficheiro parado"
            $Form.Size = New-Object System.Drawing.Size(400,200)
            $Form.StartPosition = "CenterScreen"
            $Form.Topmost = $true

            # Mensagem
            $Label = New-Object System.Windows.Forms.Label
            $Label.Text = "O ficheiro 'AL.xlsx' não foi alterado há mais de $IdleThresholdMinutes minuto(s)."
            $Label.AutoSize = $true
            $Label.Location = New-Object System.Drawing.Point(30,30)
            $Form.Controls.Add($Label)

            # Botão Abrir Pasta
            $ButtonOpenFolder = New-Object System.Windows.Forms.Button
            $ButtonOpenFolder.Text = "Abrir Pasta"
            $ButtonOpenFolder.Size = New-Object System.Drawing.Size(100,30)
            $ButtonOpenFolder.Location = New-Object System.Drawing.Point(50,100)
            $ButtonOpenFolder.Add_Click({
                Start-Process "explorer.exe" $FolderPath
                $Form.Close()
            })
            $Form.Controls.Add($ButtonOpenFolder)

            # Botão Fechar
            $ButtonClose = New-Object System.Windows.Forms.Button
            $ButtonClose.Text = "Fechar"
            $ButtonClose.Size = New-Object System.Drawing.Size(100,30)
            $ButtonClose.Location = New-Object System.Drawing.Point(200,100)
            $ButtonClose.Add_Click({
                $Form.Close()
            })
            $Form.Controls.Add($ButtonClose)

            # Mostrar o Form
            $Form.ShowDialog()

            # Espera 5 minutos antes de voltar a verificar
            Start-Sleep -Seconds 300

        } else {
            Write-Host "O ficheiro foi modificado recentemente ($([math]::Round($TimeSinceLastModified.TotalMinutes,2)) minutos atrás)."
        }
    } else {
        Write-Host "O caminho $FilePath não existe. Não é possível monitorizar."
    }

    Start-Sleep -Seconds 30
}
