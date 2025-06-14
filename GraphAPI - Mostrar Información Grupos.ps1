<#
.SYNOPSIS
    Informe detallado de grupos Microsoft 365 con interfaz gráfica en PowerShell.

.DESCRIPTION
    Este script se conecta a Microsoft Graph y obtiene todos los grupos (de seguridad y Microsoft 365) del tenant.
    Proporciona una interfaz gráfica moderna y adaptativa para buscar y seleccionar grupos.
    Al seleccionar un grupo, muestra detalles como:
      - Nombre, descripción, tipo (con detalle dinámico y asignable)
      - Propietarios con DisplayName y UPN
      - Miembros con tipo (Miembro, Invitado, Grupo Anidado)
      - Roles asignados al grupo
      - Visibilidad, seguridad, correo habilitado y fecha de creación

    El diseño visual y disposición son similares al script original para usuarios, con filtro, selector y panel de información.

.REQUIREMENTS
    - PowerShell 7.x o superior recomendado.
    - Módulo Microsoft.Graph.Authentication instalado.
    - Permisos necesarios: Group.Read.All, Directory.Read.All, Directory.ReadWrite.All.

.NOTES
    - El script no muestra licencias asignadas a grupos porque la API es inconsistente.

#>

Import-Module Microsoft.Graph.Authentication

Clear-Host
Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes "Group.Read.All","Directory.Read.All","Directory.ReadWrite.All"

Write-Host "Cargando grupos del tenant..." -ForegroundColor Cyan

# Obtener el nombre del tenant para mostrar arriba
$tenantName = (Get-MgOrganization).DisplayName

# Obtener todos los grupos de seguridad y Microsoft 365 (Unified)
$allGroups = Get-MgGroup -All | Where-Object { $_.SecurityEnabled -or $_.GroupTypes -contains "Unified" }

function Get-GroupTypeText {
    param($group)
    if ($group.SecurityEnabled -and -not $group.MailEnabled) { return "Grupo de Seguridad" }
    elseif (-not $group.SecurityEnabled -and $group.MailEnabled -and $group.GroupTypes -contains "Unified") { return "Grupo Microsoft 365 (Unificado)" }
    elseif ($group.MailEnabled -and -not $group.SecurityEnabled) { return "Grupo de Distribución" }
    elseif ($group.SecurityEnabled -and $group.MailEnabled) { return "Grupo de Seguridad con correo" }
    else { return "Tipo desconocido" }
}

function Get-GroupTypeExtendedText {
    param($group)

    $baseType = Get-GroupTypeText $group

    if ($group.MembershipRule) {
        $rule = $group.MembershipRule.ToLower()
        $dynamicType = if ($rule -match "device") { " - Dinámico Dispositivo" } else { " - Dinámico Usuario" }
    } else {
        $dynamicType = ""
    }

    $assignableText = if ($group.IsAssignableToRole) { " (Asignable)" } else { "" }

    return "$baseType$dynamicType$assignableText"
}

# Construcción lista para selector
$groupList = $allGroups | ForEach-Object {
    [PSCustomObject]@{
        Id = $_.Id
        DisplayName = $_.DisplayName
        Label = "$($_.DisplayName)"
    }
} | Sort-Object DisplayName

# Guardar y leer tamaño/posición ventana
$configPath = "$PSScriptRoot\groupform.config"
$formWidth = 1050
$formHeight = 820
$formLeft = -1
$formTop = -1
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        $formWidth = $config.Width
        $formHeight = $config.Height
        $formLeft = $config.Left
        $formTop = $config.Top
    } catch { }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Informes Grupos Microsoft 365"
$form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
$form.MinimumSize = New-Object System.Drawing.Size(700,600)
if ($formLeft -ge 0 -and $formTop -ge 0) {
    $form.StartPosition = 'Manual'
    $form.Left = $formLeft
    $form.Top = $formTop
} else {
    $form.StartPosition = "CenterScreen"
}
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f4f6fa")

# Etiqueta principal (Tenant y total grupos)
$labelTenant = New-Object System.Windows.Forms.Label
$labelTenant.Text = "Tenant: $tenantName | Total grupos: $($groupList.Count)"
$labelTenant.AutoSize = $false
$labelTenant.Size = New-Object System.Drawing.Size(970, 38)
$labelTenant.Location = New-Object System.Drawing.Point(40, 15)
$labelTenant.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$labelTenant.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#2563eb")
$labelTenant.Anchor = 'Top, Left, Right'
$labelTenant.TextAlign = "MiddleLeft"
$labelTenant.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#e0e7ef")

# Panel superior con filtro y selector
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Location = New-Object System.Drawing.Point(30, 65)
$panelTop.Size = New-Object System.Drawing.Size(970, 120)
$panelTop.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f8fafc")
$panelTop.Anchor = 'Top, Left, Right'

$labelFilter = New-Object System.Windows.Forms.Label
$labelFilter.Text = "Buscar grupo:"
$labelFilter.AutoSize = $true
$labelFilter.Location = New-Object System.Drawing.Point(15,15)
$labelFilter.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$labelFilter.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#334155")
$labelFilter.Anchor = 'Top, Left'

$textBoxFilter = New-Object System.Windows.Forms.TextBox
$textBoxFilter.Width = 800
$textBoxFilter.Location = New-Object System.Drawing.Point(15,40)
$textBoxFilter.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$textBoxFilter.BackColor = [System.Drawing.Color]::White
$textBoxFilter.Anchor = 'Top, Left, Right'

$labelSelector = New-Object System.Windows.Forms.Label
$labelSelector.Text = "Selecciona el grupo:"
$labelSelector.AutoSize = $true
$labelSelector.Location = New-Object System.Drawing.Point(15,75)
$labelSelector.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$labelSelector.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#334155")
$labelSelector.Anchor = 'Top, Left'

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Width = 800
$comboBox.Location = New-Object System.Drawing.Point(200,73)
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$comboBox.BackColor = [System.Drawing.Color]::White
$comboBox.Anchor = 'Top, Left, Right'

$panelTop.Controls.AddRange(@($labelFilter, $textBoxFilter, $labelSelector, $comboBox))

# Label para sección de información
$labelInfo = New-Object System.Windows.Forms.Label
$labelInfo.Text = "INFORMACIÓN DEL GRUPO"
$labelInfo.AutoSize = $false
$labelInfo.Size = New-Object System.Drawing.Size(970, 30)
$labelInfo.Location = New-Object System.Drawing.Point(30,195)
$labelInfo.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$labelInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#2563eb")
$labelInfo.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#e0e7ef")
$labelInfo.TextAlign = "MiddleLeft"
$labelInfo.Anchor = 'Top, Left, Right'

# Textbox para mostrar info detallada
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = 'Both'
$textBox.ReadOnly = $true
$textBox.Font = New-Object System.Drawing.Font("Consolas", 13)
$textBox.Location = New-Object System.Drawing.Point(30, 235)
$textBox.Size = New-Object System.Drawing.Size(970, 520)
$textBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f1f5f9")
$textBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1e293b")
$textBox.BorderStyle = 'FixedSingle'
$textBox.Anchor = 'Top, Bottom, Left, Right'

$form.Controls.AddRange(@(
    $labelTenant, $panelTop, $labelInfo, $textBox
))

function Format-CenteredText($text, $width = 70) {
    $len = $text.Length
    if ($len -ge $width) { return $text }
    $pad = [int](($width - $len) / 2)
    return (" " * $pad) + $text
}

function Update-ComboBoxItems {
    param([string]$filterText)
    $filtered = if (![string]::IsNullOrWhiteSpace($filterText)) {
        $groupList | Where-Object { $_.DisplayName -like "*$filterText*" }
    } else { $groupList }
    $comboBox.Items.Clear(); $comboBox.Items.AddRange($filtered.Label)
    $comboBox.Tag = $filtered
    if ($comboBox.Items.Count -gt 0) { $comboBox.SelectedIndex = 0 }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Enabled = $false
$textBoxFilter.Add_TextChanged({
    $timer.Stop()
    $timer.Start()
})
$timer.Add_Tick({
    $timer.Stop()
    Update-ComboBoxItems $textBoxFilter.Text
})
Update-ComboBoxItems ""

function Get-MemberDisplayInfo($member) {
    if ($member.DisplayName) {
        $displayName = $member.DisplayName
    } elseif ($member.AdditionalProperties.ContainsKey('displayName')) {
        $displayName = $member.AdditionalProperties['displayName']
    } elseif ($member.Mail) {
        $displayName = $member.Mail
    } elseif ($member.UserPrincipalName) {
        $displayName = $member.UserPrincipalName
    } else {
        $displayName = "[Sin nombre]"
    }

    if ($member.UserPrincipalName) {
        $upn = " ($($member.UserPrincipalName))"
    } elseif ($member.AdditionalProperties.ContainsKey('userPrincipalName')) {
        $upn = " ($($member.AdditionalProperties['userPrincipalName']))"
    } else {
        $upn = ""
    }

    return @{DisplayName = $displayName; UPN = $upn}
}

$comboBox.Add_SelectedIndexChanged({
    $selectedIndex = $comboBox.SelectedIndex
    if ($selectedIndex -lt 0) { $textBox.Text = ""; return }
    $group = $comboBox.Tag[$selectedIndex]

    $grpDetails = Get-MgGroup -GroupId $group.Id -Property "Id,DisplayName,Description,CreatedDateTime,MailEnabled,SecurityEnabled,GroupTypes,Visibility,MembershipRule,IsAssignableToRole"
    $grpOwnersRaw = Get-MgGroupOwner -GroupId $group.Id -All
    $grpMembersRaw = Get-MgGroupMember -GroupId $group.Id -All

    if ($grpOwnersRaw) {
        $ownersList = $grpOwnersRaw | ForEach-Object {
            if ($_.DisplayName) {
                $name = $_.DisplayName
            } elseif ($_.AdditionalProperties.ContainsKey('displayName')) {
                $name = $_.AdditionalProperties['displayName']
            } else {
                $name = "[Sin nombre]"
            }

            if ($_.UserPrincipalName) {
                $upn = " ($($_.UserPrincipalName))"
            } elseif ($_.AdditionalProperties.ContainsKey('userPrincipalName')) {
                $upn = " ($($_.AdditionalProperties['userPrincipalName']))"
            } else {
                $upn = ""
            }

            "- $name$upn"
        }
    } else {
        $ownersList = @("- Sin propietarios asignados")
    }

    $grpMembers = $grpMembersRaw | ForEach-Object {
        $odataType = if ($_.AdditionalProperties.ContainsKey('@odata.type')) {
            $_.AdditionalProperties['@odata.type']
        } elseif ($_.PSObject.Properties.Name -contains '@odata.type') {
            $_.'@odata.type'
        } else {
            ""
        }

        $tipo = ""
        if ($odataType -eq "#microsoft.graph.user") {
            $tipoRaw = if ($_.UserType) { $_.UserType } elseif ($_.AdditionalProperties.ContainsKey('userType')) { $_.AdditionalProperties['userType'] } else { "Desconocido" }
            $tipo = switch ($tipoRaw.ToLower()) {
                "member" { "Miembro" }
                "guest" { "Invitado" }
                default { "" }
            }
        } elseif ($odataType -eq "#microsoft.graph.group") {
            $tipo = "Grupo Anidado"
        }

        $info = Get-MemberDisplayInfo $_
        if ($tipo) {
            "- ${tipo}: $($info.DisplayName)$($info.UPN)"
        } else {
            "- $($info.DisplayName)$($info.UPN)"
        }
    }
    if (-not $grpMembers) { $grpMembers = @("- Sin miembros") }

    $directoryRoles = Get-MgDirectoryRole -All
    $groupRoles = @()
    foreach ($role in $directoryRoles) {
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
        if ($members.Id -contains $grpDetails.Id) {
            $groupRoles += "- $($role.DisplayName)"
        }
    }
    if (-not $groupRoles) { $groupRoles = @("- Sin roles asignados") }

    $tipoGrupo = Get-GroupTypeExtendedText $grpDetails
    $visibilidadTexto = switch ($grpDetails.Visibility) {
        "Public" { "Público" }
        "Private" { "Privado" }
        default { "Desconocido" }
    }
    $seguridadTexto = if ($grpDetails.SecurityEnabled) { "Habilitado" } else { "No" }
    $fechaCreacion = $grpDetails.CreatedDateTime ? (Get-Date $grpDetails.CreatedDateTime).ToString("dd/MM/yyyy HH:mm:ss") : "-"

    $sepFull = "┏" + ("━" * 70) + "┓"
    $sepFullEnd = "┗" + ("━" * 70) + "┛"
    $sepThin = "┃" + (" " * 70) + "┃"
    $div = "┃" + ("─" * 70) + "┃"
    $ancho = 70

    $info = @"
$sepFull
$(Format-CenteredText "Información del grupo" $ancho)
$sepThin
$div
$(Format-CenteredText "Identidad y Configuración" $ancho)
$div
┃ Nombre           : $($grpDetails.DisplayName)
┃ Descripción      : $($grpDetails.Description -replace "`r`n"," " -replace "`n"," ")
┃ Tipo             : $tipoGrupo
┃ Visibilidad      : $visibilidadTexto
┃ Seguridad        : $seguridadTexto
┃ Correo habilitado: $(if($grpDetails.MailEnabled){"Sí"}else{"No"})
┃ Creación         : $fechaCreacion
$div
$(Format-CenteredText "Propietarios" $ancho)
$div
$($ownersList -join "`r`n")
$div
$(Format-CenteredText "Miembros" $ancho)
$div
$($grpMembers -join "`r`n")
$div
$(Format-CenteredText "Roles asignados al grupo" $ancho)
$div
$($groupRoles -join "`r`n")
$sepFullEnd
"@

    $textBox.Text = $info
})

$form.Add_FormClosed({
    Disconnect-MgGraph
    Write-Host "Sesión Graph desconectada." -ForegroundColor Yellow
    $conf = @{
        Width = $form.Width
        Height = $form.Height
        Left = $form.Left
        Top = $form.Top
    } | ConvertTo-Json
    Set-Content -Path $configPath -Value $conf
})

$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
