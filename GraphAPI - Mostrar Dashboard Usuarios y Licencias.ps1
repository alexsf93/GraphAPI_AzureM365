<#
.SYNOPSIS
Dashboard gráfico para visualizar información de usuarios y licencias de Microsoft 365 mediante Microsoft Graph.

.DESCRIPTION
Este script se conecta a Microsoft Graph con permisos para leer usuarios, grupos y directorio.
Obtiene datos del tenant, usuarios activos e inactivos, y licencias asignadas.
Muestra un formulario con:
 - Nombre del tenant
 - Contadores de usuarios (total, activos, inactivos)
 - Gráficos de barras con estado de usuarios y distribución de licencias (consumo vs total)
Las barras de licencias usan superposición para mostrar claramente el total (gris más sólido) detrás y el consumo delante.

.REQUIREMENTS
- PowerShell 7 o superior
- Módulo Microsoft.Graph instalado y actualizado
- Permisos delegados: User.Read.All, Group.Read.All, Directory.Read.All

.NOTES
- El script gestiona automáticamente la paginación para usuarios y licencias.
- Las licencias sin asignar se muestran con la etiqueta "Sin Licencia".
- El máximo visualizado para licencias totales está limitado a 100 para mejor presentación.

.EXAMPLE
Ejecuta el dashboard, solicita autenticación interactiva si es necesario.

#>

Import-Module Microsoft.Graph.Authentication

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

Clear-Host
Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","Directory.Read.All" -NoWelcome | Out-Null

Write-Host "Obteniendo información del tenant..." -ForegroundColor Cyan
$tenantObj = Get-MgOrganization
$tenantName = if ($tenantObj) { $tenantObj.DisplayName } else { "Desconocido" }

Write-Host "Obteniendo datos de usuarios y grupos..." -ForegroundColor Cyan

$skuMap = @{
    "BUSINESS_BASIC"="Microsoft 365 Business Basic"; "BUSINESS_STANDARD"="Microsoft 365 Business Standard"; "BUSINESS_PREMIUM"="Microsoft 365 Business Premium"
    "M365_APPS_BUSINESS"="Microsoft 365 Apps for Business"; "M365_APPS"="Microsoft 365 Apps for enterprise"
    "M365_F3"="Microsoft 365 F3"; "M365_E3"="Microsoft 365 E3"; "M365_E5"="Microsoft 365 E5"
    "M365_E5_SECURITY"="Microsoft 365 E5 Security"; "M365_E5_COMPLIANCE"="Microsoft 365 E5 Compliance"
    "ENTERPRISEPACK"="Office 365 E1"; "ENTERPRISEPREMIUM"="Office 365 E3/E5"; "DESKLESS"="Office 365 F1"
    "EDUCATION_A1"="Office 365 A1 (Education)"; "EDUCATION_A3"="Office 365 A3 (Education)"; "EDUCATION_A5"="Office 365 A5 (Education)"
    "M365_EDU_A1"="Microsoft 365 Education A1"; "M365_EDU_A3"="Microsoft 365 Education A3"; "M365_EDU_A5"="Microsoft 365 Education A5"
    "EMS"="Microsoft Intune"; "EMS_E3"="Enterprise Mobility + Security E3"; "EMS_E5"="Enterprise Mobility + Security E5"
    "FREE"="Azure AD Free"; "BASIC"="Azure AD Basic"; "PREMIUM_P1"="Azure AD Premium P1"; "PREMIUM_P2"="Azure AD Premium P2"
    "EXTERNAL_IDENTS"="Azure AD External Identities"
    "DEFENDER_PLAN1"="Defender for Office 365 P1"; "DEFENDER_PLAN2"="Defender for Office 365 P2"
    "DEFENDER_FOR_ENDPOINT"="Defender for Endpoint"; "DEFENDER_FOR_IDENTITY"="Defender for Identity"
    "POWER_BI_PRO"="Power BI Pro"; "POWER_BI_PREMIUM"="Power BI Premium"; "POWERAPPS_PER_USER"="Power Apps Per User Plan"
    "FLOW_PER_USER"="Power Automate Per User Plan"; "TEAMS_EXPLORATORY"="Teams Exploratory"
    "PROJECT_PLAN_1"="Project Plan 1"; "PROJECT_PLAN_3"="Project Plan 3"; "PROJECT_PLAN_5"="Project Plan 5"
    "VISIO_PLAN_1"="Visio Plan 1"; "VISIO_PLAN_2"="Visio Plan 2"; "STREAM"="Microsoft Stream"
    "VIVA_INSIGHTS"="Viva Insights"; "VIVA_LEARNING"="Viva Learning"; "VIVA_TOPICS"="Viva Topics"
    "MCAS"="Defender for Cloud Apps"; "BUSINESS_VOICE"="Business Voice"; "POWER_VIRTUAL_AGENTS"="Power Virtual Agents"
    "EXCHANGE_S_STANDARD"="Exchange Online P1"; "EXCHANGE_S_ENTERPRISE"="Exchange Online P2"
    "YAMMER_ENTERPRISE"="Yammer Enterprise"; "FORMS_PRO"="Forms Pro"; "MYANALYTICS"="MyAnalytics"
    "KAIZALA_PRO"="Kaizala Pro"; "M365_E5_DEV"="M365 E5 Developer"; "DEVELOPERPACK_E5"="M365 E5 Developer"
    "ENTERPRISEDEV"="Office 365 E3 Developer"; "ENTERPRISEPREMIUMDEV"="Office 365 E5 Developer"
    "VISUAL_STUDIO_ENTERPRISE"="VS Enterprise"; "VISUAL_STUDIO_PROFESSIONAL"="VS Professional"
    "VISUAL_STUDIO_TESTPRO"="VS Test Professional"; "POWERAPPS_DEVELOPER_PLAN"="Power Apps Developer Plan"
    "DYNAMICS_365_DEVELOPER"="Dynamics 365 Developer"
    "FLOW_FREE" = "Power Automate Free"
    "Dynamics_365_Sales_Premium_Viral_Trial" = "Dynamics 365 Sales Premium Viral Trial"
}

$allUsers = Get-MgUser -Property "Id,AccountEnabled,UserPrincipalName,DisplayName,AssignedLicenses,SignInActivity" -All
$usuariosActivos = $allUsers | Where-Object { $_.PSObject.Properties.Match('AccountEnabled') -and $_.AccountEnabled -eq $true }
$usuariosInactivos = $allUsers | Where-Object { $_.PSObject.Properties.Match('AccountEnabled') -and $_.AccountEnabled -eq $false }
$totalUsuarios = $usuariosActivos.Count + $usuariosInactivos.Count

$licenseUserCount = @{}
foreach ($user in $usuariosActivos) {
    $licenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
    if ($licenses) {
        $addedSkus = @{}
        foreach ($lic in $licenses) {
            $skuRaw = $lic.SkuPartNumber
            $skuFriendly = if ($skuMap.ContainsKey($skuRaw)) { $skuMap[$skuRaw] } else { $skuRaw }
            if (-not $addedSkus.ContainsKey($skuFriendly)) {
                if ($licenseUserCount.ContainsKey($skuFriendly)) {
                    $licenseUserCount[$skuFriendly]++
                } else {
                    $licenseUserCount[$skuFriendly] = 1
                }
                $addedSkus[$skuFriendly] = $true
            }
        }
    } else {
        if ($licenseUserCount.ContainsKey("Sin Licencia")) {
            $licenseUserCount["Sin Licencia"]++
        } else {
            $licenseUserCount["Sin Licencia"] = 1
        }
    }
}

$skusDisponiblesRaw = Get-MgSubscribedSku -ErrorAction SilentlyContinue
$skuTotals = @{}
foreach ($sku in $skusDisponiblesRaw) {
    $skuPart = $sku.SkuPartNumber
    $friendlyName = if ($skuMap.ContainsKey($skuPart)) { $skuMap[$skuPart] } else { $skuPart }
    $skuTotals[$friendlyName] = $sku.PrepaidUnits.Enabled
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Dashboard Licencias y Uso Microsoft 365"
$form.Size = New-Object System.Drawing.Size(900, 1000)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.AutoScroll = $true

$labelTenant = New-Object System.Windows.Forms.Label
$labelTenant.Location = New-Object System.Drawing.Point(20, 10)
$labelTenant.Size = New-Object System.Drawing.Size(860, 40)
$labelTenant.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$labelTenant.ForeColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
$labelTenant.TextAlign = "MiddleLeft"
$labelTenant.Text = "Tenant: $tenantName"
$form.Controls.Add($labelTenant)

$panelCounters = New-Object System.Windows.Forms.Panel
$panelCounters.Location = New-Object System.Drawing.Point(20, 60)
$panelCounters.Size = New-Object System.Drawing.Size(860, 70)
$panelCounters.BackColor = [System.Drawing.Color]::FromArgb(224, 234, 251)
$form.Controls.Add($panelCounters)

$lblTotal = New-Object System.Windows.Forms.Label
$lblTotal.Location = New-Object System.Drawing.Point(0, 10)
$lblTotal.Size = New-Object System.Drawing.Size(280, 50)
$lblTotal.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTotal.ForeColor = [System.Drawing.Color]::FromArgb(21, 128, 61)
$lblTotal.TextAlign = "MiddleCenter"
$lblTotal.Text = "Total usuarios:`n$totalUsuarios"
$lblTotal.AutoSize = $false
$panelCounters.Controls.Add($lblTotal)

$lblActivos = New-Object System.Windows.Forms.Label
$lblActivos.Location = New-Object System.Drawing.Point(290, 10)
$lblActivos.Size = New-Object System.Drawing.Size(280, 50)
$lblActivos.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblActivos.ForeColor = [System.Drawing.Color]::FromArgb(21, 101, 192)
$lblActivos.TextAlign = "MiddleCenter"
$lblActivos.Text = "Usuarios activos:`n$($usuariosActivos.Count)"
$lblActivos.AutoSize = $false
$panelCounters.Controls.Add($lblActivos)

$lblInactivos = New-Object System.Windows.Forms.Label
$lblInactivos.Location = New-Object System.Drawing.Point(580, 10)
$lblInactivos.Size = New-Object System.Drawing.Size(280, 50)
$lblInactivos.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblInactivos.ForeColor = [System.Drawing.Color]::FromArgb(239, 68, 68)
$lblInactivos.TextAlign = "MiddleCenter"
$lblInactivos.Text = "Usuarios inactivos:`n$($usuariosInactivos.Count)"
$lblInactivos.AutoSize = $false
$panelCounters.Controls.Add($lblInactivos)

$line1 = New-Object System.Windows.Forms.Label
$line1.BorderStyle = 'Fixed3D'
$line1.Location = New-Object System.Drawing.Point(20, 140)
$line1.Size = New-Object System.Drawing.Size(860, 2)
$form.Controls.Add($line1)

$lblExpEstados = New-Object System.Windows.Forms.Label
$lblExpEstados.Location = New-Object System.Drawing.Point(20, 150)
$lblExpEstados.Size = New-Object System.Drawing.Size(860, 30)
$lblExpEstados.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Italic)
$lblExpEstados.Text = "Estado de Usuarios: Número de usuarios activos e inactivos."
$form.Controls.Add($lblExpEstados)

$chartEstados = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chartEstados.Location = New-Object System.Drawing.Point(20, 190)
$chartEstados.Size = New-Object System.Drawing.Size(860, 350)

$chartAreaEst = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartAreaEst.AxisX.LabelStyle.Angle = 45
$chartAreaEst.AxisX.Interval = 1
$chartEstados.ChartAreas.Add($chartAreaEst)

$seriesEst = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$seriesEst.Name = "EstadoUsuarios"
$seriesEst.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$seriesEst.IsValueShownAsLabel = $true
$seriesEst.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$chartEstados.Series.Add($seriesEst)

$chartEstados.Legends.Clear()

$pointIndex = $seriesEst.Points.AddXY("Activos", $usuariosActivos.Count)
$seriesEst.Points[$pointIndex].Color = [System.Drawing.Color]::FromArgb(0, 120, 215)
$seriesEst.Points[$pointIndex].LabelForeColor = [System.Drawing.Color]::Black
$seriesEst.Points[$pointIndex].LabelBackColor = [System.Drawing.Color]::FromArgb(220, 255, 255, 255)

$pointIndex = $seriesEst.Points.AddXY("Inactivos", $usuariosInactivos.Count)
$seriesEst.Points[$pointIndex].Color = [System.Drawing.Color]::FromArgb(255, 140, 0)
$seriesEst.Points[$pointIndex].LabelForeColor = [System.Drawing.Color]::Black
$seriesEst.Points[$pointIndex].LabelBackColor = [System.Drawing.Color]::FromArgb(220, 255, 255, 255)

$form.Controls.Add($chartEstados)

$line2 = New-Object System.Windows.Forms.Label
$line2.BorderStyle = 'Fixed3D'
$line2.Location = New-Object System.Drawing.Point(20, 550)
$line2.Size = New-Object System.Drawing.Size(860, 2)
$form.Controls.Add($line2)

$lblExpLicencias = New-Object System.Windows.Forms.Label
$lblExpLicencias.Location = New-Object System.Drawing.Point(20, 560)
$lblExpLicencias.Size = New-Object System.Drawing.Size(860, 30)
$lblExpLicencias.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Italic)
$lblExpLicencias.Text = "Distribución de Licencias: Consumo vs Total disponibles."
$form.Controls.Add($lblExpLicencias)

$chartLicencias = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chartLicencias.Location = New-Object System.Drawing.Point(20, 600)
$chartLicencias.Size = New-Object System.Drawing.Size(860, 350)

$chartAreaLic = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartAreaLic.AxisX.LabelStyle.Angle = 45
$chartAreaLic.AxisX.Interval = 1
$chartLicencias.ChartAreas.Add($chartAreaLic)

$seriesTotal = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$seriesTotal.Name = "TotalLicencias"
$seriesTotal.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$seriesTotal.Color = [System.Drawing.Color]::FromArgb(220, 200, 200, 200) # gris más sólido
$seriesTotal.IsValueShownAsLabel = $true
$seriesTotal.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$seriesTotal["PixelPointWidth"] = 40
$seriesTotal["DrawSideBySide"] = "false"

$seriesConsumo = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$seriesConsumo.Name = "LicenciasUsadas"
$seriesConsumo.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$seriesConsumo.IsValueShownAsLabel = $true
$seriesConsumo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$seriesConsumo["PixelPointWidth"] = 40
$seriesConsumo["DrawSideBySide"] = "false"

$colorPalette = @(
    [System.Drawing.Color]::FromArgb(0, 120, 215),
    [System.Drawing.Color]::FromArgb(255, 140, 0),
    [System.Drawing.Color]::FromArgb(196, 0, 0),
    [System.Drawing.Color]::FromArgb(0, 153, 188),
    [System.Drawing.Color]::FromArgb(104, 33, 122),
    [System.Drawing.Color]::FromArgb(0, 153, 0),
    [System.Drawing.Color]::FromArgb(255, 192, 0),
    [System.Drawing.Color]::FromArgb(0, 0, 128),
    [System.Drawing.Color]::FromArgb(128, 0, 0)
)

$index = 0
foreach ($skuName in $licenseUserCount.Keys) {
    $realTotal = if ($skuTotals.ContainsKey($skuName)) {
        $skuTotals[$skuName]
    } elseif ($skuName -eq "Sin Licencia") {
        $totalUsuarios
    } else {
        0
    }
    $totalQty = [Math]::Min($realTotal, 100)
    $usedQty = $licenseUserCount[$skuName]

    $ptTotal = $seriesTotal.Points.AddXY($skuName, $totalQty)
    $seriesTotal.Points[$ptTotal].LabelForeColor = [System.Drawing.Color]::DarkGray
    $seriesTotal.Points[$ptTotal].LabelBackColor = [System.Drawing.Color]::FromArgb(180, 240, 240, 240)

    $ptConsumo = $seriesConsumo.Points.AddXY($skuName, $usedQty)
    $seriesConsumo.Points[$ptConsumo].Color = $colorPalette[$index % $colorPalette.Count]
    $seriesConsumo.Points[$ptConsumo].LabelForeColor = [System.Drawing.Color]::Black
    $seriesConsumo.Points[$ptConsumo].LabelBackColor = [System.Drawing.Color]::FromArgb(220, 255, 255, 255)

    $index++
}

$chartLicencias.Series.Clear()
$chartLicencias.Series.Add($seriesTotal)
$chartLicencias.Series.Add($seriesConsumo)

$form.Controls.Add($chartLicencias)

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

Disconnect-MgGraph | Out-Null
Write-Host "Desconectado de Microsoft Graph." -ForegroundColor Cyan
