#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form                     = New-Object System.Windows.Forms.Form
$form.Text                = "AD Recon Tool (SharpHound + Fallback LDAP)"
$form.Size                = New-Object System.Drawing.Size(480,380)
$form.StartPosition        = "CenterScreen"
$form.FormBorderStyle      = "FixedDialog"
$form.MaximizeBox          = $false

$label                    = New-Object System.Windows.Forms.Label
$label.Location           = New-Object System.Drawing.Point(20,15)
$label.Size               = New-Object System.Drawing.Size(420,25)
$label.Text               = "Elige carpeta y pulsa Recolectar. Si SharpHound falla, se usará LDAP simple."
$form.Controls.Add($label)

$btn                      = New-Object System.Windows.Forms.Button
$btn.Location             = New-Object System.Drawing.Point(160,50)
$btn.Size                 = New-Object System.Drawing.Size(120,40)
$btn.Text                 = "Recolectar"
$form.Controls.Add($btn)

$log                      = New-Object System.Windows.Forms.TextBox
$log.Location             = New-Object System.Drawing.Point(20,100)
$log.Size                 = New-Object System.Drawing.Size(420,210)
$log.Multiline            = $true
$log.ScrollBars           = "Vertical"
$log.ReadOnly             = $true
$form.Controls.Add($log)

function Write-Log {
    param($msg)
    $log.AppendText("$msg`r`n")
    $log.Refresh()
}

$btn.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Carpeta donde guardar la recolecta"
    if ($dlg.ShowDialog() -ne "OK") { return }
    $dir = $dlg.SelectedPath
    $log.Clear()
    Write-Log "Carpeta seleccionada: $dir"

    # Deshabilitar botón mientras corre el job
    $btn.Enabled = $false

    # Iniciar job en background
    $job = Start-Job -ArgumentList $dir -ScriptBlock {
        param($outDir)
        
        # SharpHound download and execution
        try {
            Write-Output "Descargando SharpHound …"
            $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/BloodHoundAD/SharpHound/releases/latest" -TimeoutSec 15 -ErrorAction Stop
            $asset = $rel.assets | Where-Object { $_.name -like "*SharpHound.exe" } | Select-Object -First 1
            $url = $asset.browser_download_url
            
            if (-not $url) { 
                Write-Output "No se encontró SharpHound.exe en la release" 
            } else {
                $exe = Join-Path $outDir "SharpHound.exe"
                Invoke-WebRequest -Uri $url -OutFile $exe -UseBasicParsing -ErrorAction Stop
                Write-Output "✅ SharpHound descargado: $exe"
                Write-Output "Lanzando SharpHound …"
                
                $p = Start-Process -FilePath $exe -WorkingDirectory $outDir -ArgumentList "-c","All","--ZipFileName","SharpOut" -NoNewWindow -PassThru -Wait
                
                if ($p.ExitCode -eq 0 -and (Test-Path (Join-Path $outDir "SharpOut.zip"))) {
                    Write-Output "✅ SharpHound finalizó correctamente"
                } else {
                    Write-Output "⚠️ SharpHound finalizó con ExitCode $($p.ExitCode)"
                }
            }
        } catch {
            Write-Output "⚠️  Fallo SharpHound: $($_.Exception.Message)"
        }

        # Fallback LDAP simple
        try {
            Write-Output "Recuperando objetos básicos por LDAP …"
            $dom  = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $root = $dom.GetDirectoryEntry()
            $src  = New-Object System.DirectoryServices.DirectorySearcher($root)

            $src.Filter = "(objectCategory=user)"
            $users = $src.FindAll()
            $userList = @()
            foreach ($u in $users) {
                $val = $u.Properties["samaccountname"]
                if ($val -and $val.Count -gt 0) { $userList += $val[0] }
            }
            $userList | Set-Content (Join-Path $outDir "LDAP_Users.txt")
            $users.Dispose()

            $src.Filter = "(objectCategory=computer)"
            $computers = $src.FindAll()
            $compList = @()
            foreach ($c in $computers) {
                $val = $c.Properties["name"]
                if ($val -and $val.Count -gt 0) { $compList += $val[0] }
            }
            $compList | Set-Content (Join-Path $outDir "LDAP_Computers.txt")
            $computers.Dispose()

            $src.Filter = "(objectCategory=group)"
            $groups = $src.FindAll()
            $groupList = @()
            foreach ($g in $groups) {
                $val = $g.Properties["samaccountname"]
                if ($val -and $val.Count -gt 0) { $groupList += $val[0] }
            }
            $groupList | Set-Content (Join-Path $outDir "LDAP_Groups.txt")
            $groups.Dispose()

            $src.Dispose()
            $root.Dispose()
            Write-Output "✅ Datos simples exportados"
        } catch {
            Write-Output "❌ Error LDAP simple: $($_.Exception.Message)"
        }
        
        Write-Output "PROCESO_COMPLETADO"
    }

    # Variables para el timer
    $script:currentJob = $job
    $script:jobCompleted = $false
    
    # Timer para monitorear el job
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    
    $timer.Add_Tick({
        try {
            # Verificar si el job aún existe
            if ($null -eq $script:currentJob) {
                $this.Stop()
                $btn.Enabled = $true
                return
            }
            
            # Intentar obtener el job de forma segura
            $jobState = $null
            try {
                $jobObj = Get-Job -Id $script:currentJob.Id -ErrorAction Stop
                $jobState = $jobObj.State
            } catch {
                # El job ya no existe
                $script:currentJob = $null
                $this.Stop()
                $btn.Enabled = $true
                Write-Log "Job finalizado"
                [System.Windows.Forms.MessageBox]::Show("Recolecta finalizada.", "Hecho", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                return
            }
            
            # Leer salida disponible
            try {
                $output = Receive-Job -Id $script:currentJob.Id -ErrorAction SilentlyContinue
                if ($output) {
                    foreach ($line in $output) {
                        Write-Log $line
                        if ($line -eq "PROCESO_COMPLETADO") {
                            $script:jobCompleted = $true
                        }
                    }
                }
            } catch {
                Write-Log "Error leyendo salida: $($_.Exception.Message)"
            }
            
            # Verificar si el job terminó
            if ($jobState -ne 'Running') {
                try {
                    # Recoger salida final
                    $finalOutput = Receive-Job -Id $script:currentJob.Id -ErrorAction SilentlyContinue
                    if ($finalOutput) {
                        foreach ($line in $finalOutput) {
                            Write-Log $line
                        }
                    }
                    
                    # Limpiar job
                    Remove-Job -Id $script:currentJob.Id -Force -ErrorAction SilentlyContinue
                } catch {}
                
                $script:currentJob = $null
                $this.Stop()
                $this.Dispose()
                $btn.Enabled = $true
                
                if ($script:jobCompleted) {
                    [System.Windows.Forms.MessageBox]::Show("Recolecta finalizada. Revisa los archivos en la carpeta seleccionada.", "Hecho", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Proceso terminado. Revisa el log para más detalles.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                }
            }
            
        } catch {
            Write-Log "Error en timer: $($_.Exception.Message)"
            $this.Stop()
            $btn.Enabled = $true
            $script:currentJob = $null
        }
    })
    
    $timer.Start()
})

[void]$form.ShowDialog()