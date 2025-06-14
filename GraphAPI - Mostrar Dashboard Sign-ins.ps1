<#
.SYNOPSIS
Dashboard gráfico para visualizar inicios de sesión exitosos y fallidos en Microsoft 365.

.DESCRIPTION
Este script se conecta a Microsoft Graph con permisos para leer logs de inicio de sesión.
Obtiene los eventos de inicio de sesión dentro del rango de días indicado (máximo 30).
Muestra un formulario con:
 - Nombre del tenant
 - Contadores de inicios exitosos y fallidos
 - Gráfico de barras con el resumen
 - Listado detallado de los últimos intentos fallidos con usuario, fecha y código de error
 - Botones para exportar y copiar al portapapeles los intentos fallidos

.REQUIREMENTS
- PowerShell 7 o superior
- Módulo Microsoft.Graph instalado y actualizado
- Permisos delegados: AuditLog.Read.All

.NOTES
- El rango máximo de días es 30 para cumplir con limitaciones de API.
- Los intentos fallidos se muestran ordenados de más recientes a más antiguos.
- El listado muestra hasta 50 registros fallidos para mantener la legibilidad.

.EXAMPLE
Lanza el dashboard interactivo para revisar inicios de sesión.

#>

Import-Module Microsoft.Graph.Authentication

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

function Get-ValidDays([int]$inputDays) {
    if ($inputDays -le 0) { return 1 }
    if ($inputDays -gt 30) { return 30 }
    return $inputDays
}

Clear-Host
Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "AuditLog.Read.All" -NoWelcome | Out-Null

$tenantObj = Get-MgOrganization
$tenantName = if ($tenantObj) { $tenantObj.DisplayName } else { "Desconocido" }

$formInput = New-Object System.Windows.Forms.Form
$formInput.Text = "Seleccione rango de días para inicios de sesión"
$formInput.Size = New-Object System.Drawing.Size(350, 150)
$formInput.StartPosition = "CenterScreen"
$formInput.FormBorderStyle = 'FixedDialog'
$formInput.MaximizeBox = $false
$formInput.MinimizeBox = $false

$labelPrompt = New-Object System.Windows.Forms.Label
$labelPrompt.Text = "Indique número de días (1-30):"
$labelPrompt.Location = New-Object System.Drawing.Point(20, 20)
$labelPrompt.Size = New-Object System.Drawing.Size(300, 20)
$formInput.Controls.Add($labelPrompt)

$textBoxDays = New-Object System.Windows.Forms.TextBox
$textBoxDays.Location = New-Object System.Drawing.Point(20, 50)
$textBoxDays.Size = New-Object System.Drawing.Size(300, 25)
$textBoxDays.Text = "30"
$formInput.Controls.Add($textBoxDays)

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = "Aceptar"
$btnOk.Location = New-Object System.Drawing.Point(120, 85)
$btnOk.Size = New-Object System.Drawing.Size(100, 30)
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$formInput.Controls.Add($btnOk)

$formInput.AcceptButton = $btnOk

if ($formInput.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Operación cancelada." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit
}

$inputDays = [int]$textBoxDays.Text
$daysBack = Get-ValidDays $inputDays

$since = (Get-Date).AddDays(-$daysBack).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "Obteniendo logs de inicio de sesión últimos $daysBack días..." -ForegroundColor Cyan
$signIns = Get-MgAuditLogSignIn -Filter "createdDateTime ge $since" -All

$exitososCount = ($signIns | Where-Object { $_.Status.ErrorCode -eq 0 }).Count
$fallidos = $signIns | Where-Object { $_.Status.ErrorCode -ne 0 }
$fallidosCount = $fallidos.Count

$form = New-Object System.Windows.Forms.Form
$form.Text = "Dashboard Inicios de Sesión Microsoft 365"
$form.Size = New-Object System.Drawing.Size(1250, 880)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.AutoScroll = $true

$labelTenant = New-Object System.Windows.Forms.Label
$labelTenant.Text = "Tenant: $tenantName"
$labelTenant.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$labelTenant.Size = New-Object System.Drawing.Size(1210, 40)
$labelTenant.Location = New-Object System.Drawing.Point(20, 10)
$labelTenant.ForeColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
$labelTenant.TextAlign = 'MiddleLeft'
$form.Controls.Add($labelTenant)

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "Resumen de inicios de sesión últimos $daysBack días"
$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Italic)
$labelTitle.Size = New-Object System.Drawing.Size(1210, 30)
$labelTitle.Location = New-Object System.Drawing.Point(20, 60)
$form.Controls.Add($labelTitle)

$panelCounters = New-Object System.Windows.Forms.Panel
$panelCounters.Location = New-Object System.Drawing.Point(20, 90)
$panelCounters.Size = New-Object System.Drawing.Size(1210, 70)
$panelCounters.BackColor = [System.Drawing.Color]::FromArgb(224, 234, 251)
$form.Controls.Add($panelCounters)

$lblExitosos = New-Object System.Windows.Forms.Label
$lblExitosos.Text = "Inicios exitosos:`n$exitososCount"
$lblExitosos.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblExitosos.ForeColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
$lblExitosos.Size = New-Object System.Drawing.Size(600, 60)
$lblExitosos.TextAlign = 'MiddleCenter'
$panelCounters.Controls.Add($lblExitosos)

$lblFallidos = New-Object System.Windows.Forms.Label
$lblFallidos.Text = "Inicios fallidos:`n$fallidosCount"
$lblFallidos.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblFallidos.ForeColor = [System.Drawing.Color]::FromArgb(239, 68, 68)
$lblFallidos.Size = New-Object System.Drawing.Size(600, 60)
$lblFallidos.Location = New-Object System.Drawing.Point(610, 0)
$lblFallidos.TextAlign = 'MiddleCenter'
$panelCounters.Controls.Add($lblFallidos)

$chartSignIns = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chartSignIns.Location = New-Object System.Drawing.Point(20, 170)
$chartSignIns.Size = New-Object System.Drawing.Size(1210, 300)
$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartArea.AxisX.LabelStyle.Angle = 45
$chartArea.AxisX.Interval = 1
$chartSignIns.ChartAreas.Add($chartArea)

$series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$series.Name = "SignInStatus"
$series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$series.IsValueShownAsLabel = $true
$series.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$chartSignIns.Series.Add($series)

foreach ($point in $series.Points) {
    $point.LabelForeColor = [System.Drawing.Color]::Black
    $point.LabelBackColor = [System.Drawing.Color]::FromArgb(200, 255, 255, 255)
    $point.LabelBorderColor = [System.Drawing.Color]::LightGray
    $point.LabelBorderWidth = 1
}

$series.Points.AddXY("Exitosos", $exitososCount) | Out-Null
$series.Points[0].Color = [System.Drawing.Color]::FromArgb(21, 128, 61) # verde

$series.Points.AddXY("Fallidos", $fallidosCount) | Out-Null
$series.Points[1].Color = [System.Drawing.Color]::FromArgb(239, 68, 68) # rojo

$form.Controls.Add($chartSignIns)

$labelFallos = New-Object System.Windows.Forms.Label
$labelFallos.Text = "Últimos intentos fallidos:"
$labelFallos.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$labelFallos.Location = New-Object System.Drawing.Point(20, 480)
$labelFallos.Size = New-Object System.Drawing.Size(1210, 30)
$form.Controls.Add($labelFallos)

$textBoxFallos = New-Object System.Windows.Forms.TextBox
$textBoxFallos.Multiline = $true
$textBoxFallos.ScrollBars = "Vertical"
$textBoxFallos.ReadOnly = $true
$textBoxFallos.Font = New-Object System.Drawing.Font("Consolas", 10)
$textBoxFallos.Location = New-Object System.Drawing.Point(20, 520)
$textBoxFallos.Size = New-Object System.Drawing.Size(1210, 300)
$form.Controls.Add($textBoxFallos)

# Botón Exportar a archivo
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Exportar a archivo"
$btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$btnExport.Size = New-Object System.Drawing.Size(180, 35)
$btnExport.Location = New-Object System.Drawing.Point(20, 830)
$form.Controls.Add($btnExport)

$btnExport.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Archivos de texto (*.txt)|*.txt"
    $saveFileDialog.FileName = "IntentosFallidos_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $textBoxFallos.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Archivo exportado correctamente.`n$($saveFileDialog.FileName)", "Exportar", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al exportar el archivo.`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
})

# Botón Copiar al portapapeles
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Copiar texto"
$btnCopy.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$btnCopy.Size = New-Object System.Drawing.Size(180, 35)
$btnCopy.Location = New-Object System.Drawing.Point(210, 830)
$form.Controls.Add($btnCopy)

$btnCopy.Add_Click({
    try {
        [System.Windows.Forms.Clipboard]::SetText($textBoxFallos.Text)
        [System.Windows.Forms.MessageBox]::Show("Texto copiado al portapapeles.", "Copiar", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al copiar al portapapeles.`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$maxFallosMostrar = 50
$fallidos | Sort-Object CreatedDateTime -Descending | Select-Object -First $maxFallosMostrar | ForEach-Object {
    $dt = (Get-Date $_.CreatedDateTime).ToString("dd/MM/yyyy HH:mm:ss")
    $displayName = if ($_.UserDisplayName) { $_.UserDisplayName } else { "[Sin Nombre]" }
    $userPrincipalName = if ($_.UserPrincipalName) { $_.UserPrincipalName } else { "[Sin UPN]" }
    $errorCode = $_.Status.ErrorCode
    $errorMsg = $_.Status.FailureReason

    $line = "- [$displayName ($userPrincipalName)]: $dt | Código: $errorCode | $errorMsg"
    $textBoxFallos.AppendText("$line`r`n")
}

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

Disconnect-MgGraph | Out-Null
Write-Host "Desconectado de Microsoft Graph." -ForegroundColor Cyan
