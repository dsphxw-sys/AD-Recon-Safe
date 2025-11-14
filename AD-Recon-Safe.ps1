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
        
        # Force TLS 1.2 for GitHub API
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $sharpHoundSuccess = $false
        
        # SharpHound download and execution
        try {
            Write-Output "Descargando SharpHound …"
            $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/BloodHoundAD/SharpHound/releases/latest" -TimeoutSec 15 -ErrorAction Stop
            $asset = $rel.assets | Where-Object { $_.name -like "*SharpHound.exe" } | Select-Object -First 1
            $url = $asset.browser_download_url
            
            if (-not $url) { 
                Write-Output "No se encontró SharpHound.exe en la release." 
            } else {
                $exe = Join-Path $outDir "SharpHound.exe"
                Invoke-WebRequest -Uri $url -OutFile $exe -UseBasicParsing -ErrorAction Stop
                Write-Output "✅ SharpHound descargado: $exe"
                Write-Output "Lanzando SharpHound …"
                
                $p = Start-Process -FilePath $exe -WorkingDirectory $outDir -ArgumentList "-c","All","--ZipFileName","SharpOut" -NoNewWindow -PassThru -Wait
                
                if ($p.ExitCode -eq 0 -and (Test-Path (Join-Path $outDir "SharpOut.zip"))) {
                    Write-Output "✅ SharpHound finalizó correctamente."
                    $sharpHoundSuccess = $true
                } else {
                    Write-Output "⚠️ SharpHound finalizó con ExitCode $($p.ExitCode)."
                }
            }
        } catch {
            Write-Output "⚠️ Fallo en la fase de SharpHound: $($_.Exception.Message)"
        }

        # Fallback LDAP simple (RSAT-less)
        if (-not $sharpHoundSuccess) {
            try {
                Write-Output "SharpHound falló. Intentando recolección LDAP simple..."
                
                # Find the default naming context without RSAT
                $rootDSE = New-Object System.DirectoryServices.DirectoryEntry("LDAP://rootDSE")
                $defaultNamingContext = $rootDSE.Properties["defaultNamingContext"].Value
                if (-not $defaultNamingContext) {
                    throw "No se pudo obtener el contexto de nombres predeterminado del dominio."
                }
                $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$defaultNamingContext")
                
                $src  = New-Object System.DirectoryServices.DirectorySearcher($root)
                $src.PageSize = 1000

                # Get Users
                Write-Output "Obteniendo usuarios..."
                $src.Filter = "(&(objectCategory=person)(objectClass=user))"
                $users = $src.FindAll()
                $userList = foreach ($u in $users) {
                    $u.Properties["samaccountname"][0]
                }
                $userList | Set-Content (Join-Path $outDir "LDAP_Users.txt")
                if ($users) { $users.Dispose() }

                # Get Computers
                Write-Output "Obteniendo equipos..."
                $src.Filter = "(objectCategory=computer)"
                $computers = $src.FindAll()
                $compList = foreach ($c in $computers) {
                    $c.Properties["name"][0]
                }
                $compList | Set-Content (Join-Path $outDir "LDAP_Computers.txt")
                if ($computers) { $computers.Dispose() }
                
                # Get Groups
                Write-Output "Obteniendo grupos..."
                $src.Filter = "(objectCategory=group)"
                $groups = $src.FindAll()
                $groupList = foreach ($g in $groups) {
                    $g.Properties["samaccountname"][0]
                }
                $groupList | Set-Content (Join-Path $outDir "LDAP_Groups.txt")
                if ($groups) { $groups.Dispose() }

                if ($src) { $src.Dispose() }
                if ($root) { $root.Dispose() }
                if ($rootDSE) { $rootDSE.Dispose() }
                
                Write-Output "✅ Datos LDAP simples exportados."
            } catch {
                Write-Output "❌ Error en recolección LDAP simple: $($_.Exception.Message)"
            }
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
