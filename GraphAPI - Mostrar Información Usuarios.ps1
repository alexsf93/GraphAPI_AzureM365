<#
.SYNOPSIS
    Script de PowerShell para conectar y extraer información detallada de usuarios y grupos desde Microsoft 365 a través de Microsoft Graph.

.DESCRIPTION
    Este script utiliza el módulo Microsoft.Graph.Authentication para autenticarse con permisos adecuados y recopilar datos relevantes sobre usuarios activos, inactivos y eliminados dentro del tenant de Microsoft 365.
    Incluye un formulario GUI moderno que permite filtrar usuarios, seleccionar uno y visualizar su información detallada, incluyendo licencias asignadas, roles, grupos estáticos y dinámicos, datos de contacto y últimos intentos de inicio de sesión con códigos de error explicados.
    El formulario es adaptable en tamaño y posición, almacenando la configuración para la próxima ejecución.

.PARAMETER
    No requiere parámetros. El script solicitará autenticación con los permisos definidos.

.REQUIREMENTS
    - PowerShell 7.x o superior recomendado.
    - Módulo Microsoft.Graph.Authentication instalado.
    - Permisos adecuados en Azure AD para los scopes: User.ReadWrite.All, Directory.Read.All, GroupMember.Read.All, AuditLog.Read.All.

.NOTES
    - Guarda la configuración del formulario en un archivo JSON local para mantener tamaño y posición.
    - Incluye interpretación de códigos de error comunes en intentos de inicio de sesión.
    - Permite manejar usuarios activos y eliminados.

.EXAMPLE
    Ejecuta el script para abrir el formulario interactivo y navegar entre usuarios con información detallada.

#>

# Requiere módulo Microsoft.Graph
Import-Module Microsoft.Graph.Authentication

Clear-Host
Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.Read.All","GroupMember.Read.All","AuditLog.Read.All" -NoWelcome

Write-Host "Consultando información del tenant..." -ForegroundColor Cyan

# Tenant Info
$tenantName = (Get-MgOrganization).DisplayName

# Usuarios activos
$activeUsers = Get-MgUser -Property "Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,JobTitle,Department,MobilePhone,City,Country,SignInActivity,CreatedDateTime" -Top 999

# Papelera (usuarios eliminados, REST para paginación)
function Get-AllDeletedUsers {
    $uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user"
    $all = @()
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
        $all += $resp.value
        $uri = $resp.'@odata.nextLink'
    } while ($uri)
    return $all
}
$deletedUsersRaw = Get-AllDeletedUsers
$deletedUsers = $deletedUsersRaw | Select-Object -ExpandProperty id -Unique

# Contadores
$usuariosActivosCount   = ($activeUsers | Where-Object AccountEnabled).Count
$usuariosInactivosCount = ($activeUsers | Where-Object { -not $_.AccountEnabled }).Count
$usuariosEliminadosCount = $deletedUsers.Count

# Licencias amigables
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

# Diccionario ampliado de mensajes para códigos de error de inicio de sesión
$signInErrorMessages = @{
    0      = "Inicio de sesión exitoso."
    50034  = "Nombre de usuario no encontrado."
    50053  = "Cuenta bloqueada por demasiados intentos fallidos."
    50056  = "Contraseña incorrecta."
    50074  = "Requiere autenticación multifactor."
    50126  = "Nombre de usuario o contraseña incorrectos."
    50140  = "Acceso bloqueado por política de acceso condicional o restricción de seguridad."
    50158  = "Token de acceso inválido o expirado."
    700001 = "No autorizado: token inválido o caducado."
    70002  = "No se encontró el recurso solicitado."
    90002  = "Timeout de conexión con el servidor de autenticación."
    90014  = "Error interno del servidor de autenticación."
}

# Construcción de lista de usuarios (activos + eliminados)
$usuariosFullList = @(
    $activeUsers | ForEach-Object {
        [PSCustomObject]@{
            Id = $_.Id
            DisplayName = $_.DisplayName
            UserPrincipalName = $_.UserPrincipalName
            Label = "$($_.DisplayName) ($($_.UserPrincipalName))"
            IsDeleted = $false
        }
    }
    $deletedUsersRaw | ForEach-Object {
        [PSCustomObject]@{
            Id = $_.Id
            DisplayName = $_.DisplayName ? $_.DisplayName : "[Sin nombre]"
            UserPrincipalName = $_.UserPrincipalName ? $_.UserPrincipalName : "[Sin UPN]"
            Label = "$($_.DisplayName) ($($_.UserPrincipalName)) - ELIMINADO"
            IsDeleted = $true
        }
    }
) | Sort-Object DisplayName

# Obtener todos los grupos y sus miembros
Write-Host "Cargando grupos y miembros..." -ForegroundColor Cyan
$allGroups = Get-MgGroup -All | Where-Object { $_.SecurityEnabled -or $_.GroupTypes -contains "Unified" }
$DynamicGroupIds = $allGroups | Where-Object MembershipRule | ForEach-Object { $_.Id }
$UserGroupsDict_Static = @{}; $UserGroupsDict_Dynamic = @{}

foreach ($g in $allGroups) {
    try {
        $members = Get-MgGroupMember -GroupId $g.Id -All
        $isDynamic = $DynamicGroupIds -contains $g.Id
        foreach ($m in $members) {
            if ($m.AdditionalProperties['@odata.type'] -eq "#microsoft.graph.user") {
                $uid = $m.Id
                if ($isDynamic) {
                    $UserGroupsDict_Dynamic[$uid] = $UserGroupsDict_Dynamic[$uid] + @($g.DisplayName)
                } else {
                    $UserGroupsDict_Static[$uid] = $UserGroupsDict_Static[$uid] + @($g.DisplayName)
                }
            }
        }
    } catch { }
}

# ---------- FORMULARIO Y CONTROLES MODERNOS Y ADAPTATIVOS, RECORDANDO TAMAÑO ----------

$configPath = "$PSScriptRoot\userform.config"
$formWidth = 1880
$formHeight = 950  # Más alto para dar más espacio vertical
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
$form.Text = "Informes Usuarios Microsoft 365"
$form.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 700)
if ($formLeft -ge 0 -and $formTop -ge 0) {
    $form.StartPosition = 'Manual'
    $form.Left = $formLeft
    $form.Top = $formTop
} else {
    $form.StartPosition = "CenterScreen"
}
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f4f6fa")

# Etiqueta principal
$labelTenant = New-Object System.Windows.Forms.Label
$labelTenant.Text = "Tenant: $tenantName | Activos: $usuariosActivosCount | Inactivos: $usuariosInactivosCount | Eliminados: $usuariosEliminadosCount"
$labelTenant.AutoSize = $false
$labelTenant.Size = New-Object System.Drawing.Size(1700, 38)
$labelTenant.Location = New-Object System.Drawing.Point(40, 15)
$labelTenant.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$labelTenant.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#2563eb")
$labelTenant.Anchor = 'Top, Left, Right'
$labelTenant.TextAlign = "MiddleLeft"
$labelTenant.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#e0e7ef")

# Panel superior
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Location = New-Object System.Drawing.Point(30, 65)
$panelTop.Size = New-Object System.Drawing.Size(1720, 120)
$panelTop.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f8fafc")
$panelTop.Anchor = 'Top, Left, Right'

# Label y Textbox filtro
$labelFilter = New-Object System.Windows.Forms.Label
$labelFilter.Text = "Buscar usuario:"
$labelFilter.AutoSize = $true
$labelFilter.Location = New-Object System.Drawing.Point(15, 15)
$labelFilter.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$labelFilter.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#334155")
$labelFilter.Anchor = 'Top, Left'

$textBoxFilter = New-Object System.Windows.Forms.TextBox
$textBoxFilter.Width = 1550
$textBoxFilter.Location = New-Object System.Drawing.Point(15, 40)
$textBoxFilter.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$textBoxFilter.BackColor = [System.Drawing.Color]::White
$textBoxFilter.Anchor = 'Top, Left, Right'

$labelSelector = New-Object System.Windows.Forms.Label
$labelSelector.Text = "Selecciona el usuario:"
$labelSelector.AutoSize = $true
$labelSelector.Location = New-Object System.Drawing.Point(15, 75)
$labelSelector.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$labelSelector.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#334155")
$labelSelector.Anchor = 'Top, Left'

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Width = 1550
$comboBox.Location = New-Object System.Drawing.Point(200, 73)
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$comboBox.BackColor = [System.Drawing.Color]::White
$comboBox.Anchor = 'Top, Left, Right'

$panelTop.Controls.AddRange(@($labelFilter, $textBoxFilter, $labelSelector, $comboBox))

# Label info
$labelInfo = New-Object System.Windows.Forms.Label
$labelInfo.Text = "INFORMACIÓN DE USUARIO"
$labelInfo.AutoSize = $false
$labelInfo.Size = New-Object System.Drawing.Size(1720, 30)
$labelInfo.Location = New-Object System.Drawing.Point(30, 195)
$labelInfo.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$labelInfo.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#2563eb")
$labelInfo.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#e0e7ef")
$labelInfo.TextAlign = "MiddleLeft"
$labelInfo.Anchor = 'Top, Left, Right'

# Avisos
$labelAvisoSignIn = New-Object System.Windows.Forms.Label
$labelAvisoSignIn.Text = "¡Este usuario nunca ha iniciado sesión!"
$labelAvisoSignIn.AutoSize = $true
$labelAvisoSignIn.ForeColor = [System.Drawing.Color]::Red
$labelAvisoSignIn.Location = New-Object System.Drawing.Point(45, 235)
$labelAvisoSignIn.Visible = $false
$labelAvisoSignIn.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$labelAvisoSignIn.Anchor = 'Top, Left'

$labelAvisoDisabled = New-Object System.Windows.Forms.Label
$labelAvisoDisabled.Text = "¡Esta cuenta está DESHABILITADA!"
$labelAvisoDisabled.AutoSize = $true
$labelAvisoDisabled.ForeColor = [System.Drawing.Color]::Red
$labelAvisoDisabled.Location = New-Object System.Drawing.Point(45, 265)
$labelAvisoDisabled.Visible = $false
$labelAvisoDisabled.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$labelAvisoDisabled.Anchor = 'Top, Left'

$labelDeleted = New-Object System.Windows.Forms.Label
$labelDeleted.Text = "¡Esta cuenta está ELIMINADA o en proceso de eliminación!"
$labelDeleted.AutoSize = $true
$labelDeleted.ForeColor = [System.Drawing.Color]::DarkRed
$labelDeleted.Location = New-Object System.Drawing.Point(45, 295)
$labelDeleted.Visible = $false
$labelDeleted.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$labelDeleted.Anchor = 'Top, Left'

# Textbox área info principal (más alto para mejor espacio)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = 'Both'
$textBox.ReadOnly = $true
$textBox.Font = New-Object System.Drawing.Font("Consolas", 13)
$textBox.Location = New-Object System.Drawing.Point(30, 335)
$textBox.Size = New-Object System.Drawing.Size(1800, 580)  # altura aumentada
$textBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f1f5f9")
$textBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1e293b")
$textBox.BorderStyle = 'FixedSingle'
$textBox.Anchor = 'Top, Bottom, Left, Right'

$form.Controls.AddRange(@(
    $labelTenant, $panelTop, $labelInfo,
    $labelAvisoSignIn, $labelAvisoDisabled, $labelDeleted, $textBox
))

# Función para centrar texto (verbo aprobado)
function Format-CenteredText($text, $width = 140) {
    $len = $text.Length
    if ($len -ge $width) { return $text }
    $pad = [int](($width - $len) / 2)
    return (" " * $pad) + $text
}

# Función para actualizar ComboBox filtrando
function Update-ComboBoxItems {
    param([string]$filterText)
    $filtered = if (![string]::IsNullOrWhiteSpace($filterText)) {
        $usuariosFullList | Where-Object { $_.DisplayName -like "*$filterText*" -or $_.UserPrincipalName -like "*$filterText*" }
    } else { $usuariosFullList }
    $comboBox.Items.Clear(); $comboBox.Items.AddRange($filtered.Label)
    $comboBox.Tag = $filtered
    if ($comboBox.Items.Count -gt 0) { $comboBox.SelectedIndex = 0 }
}

# Temporizador filtro
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500; $timer.Enabled = $false
$textBoxFilter.Add_TextChanged({ $timer.Stop(); $timer.Start() })
$timer.Add_Tick({ $timer.Stop(); Update-ComboBoxItems $textBoxFilter.Text })
Update-ComboBoxItems ""

# Evento selección usuario
$comboBox.Add_SelectedIndexChanged({
    $selectedIndex = $comboBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        $usuario = $comboBox.Tag[$selectedIndex]
        $labelDeleted.Visible = $usuario.IsDeleted
        $labelAvisoDisabled.Visible = $false
        $labelAvisoSignIn.Visible = $false
        $sepFull = "┏" + ("━" * 140) + "┓"
        $sepFullEnd = "┗" + ("━" * 140) + "┛"
        $sepThin = "┃" + (" " * 140) + "┃"
        $div = "┃" + ("─" * 140) + "┃"
        $ancho = 140

        if ($usuario.IsDeleted) {
            $info = @"
$sepFull
$(Format-CenteredText "Información de usuario" $ancho)
$sepThin
Nombre:   $($usuario.DisplayName)
Usuario:  $($usuario.UserPrincipalName)
ID:       $($usuario.Id)
$sepThin
(No hay más información para usuarios eliminados.)
$sepFullEnd
"@
            $textBox.Text = $info
        } else {
            $u = $activeUsers | Where-Object { $_.Id -eq $usuario.Id }
            if (!$u) { $textBox.Text = ""; return }
            $estado = if ($u.AccountEnabled) {"Habilitado"} else {"Deshabilitado"}
            $tipoUsuario = if ($u.UserType -eq "Guest") {"Invitado"} elseif ($u.UserType -eq "Member") {"Miembro"} else {$u.UserType}

            # Licencias amigables
            $licenses = (Get-MgUserLicenseDetail -UserId $u.Id | Select-Object -ExpandProperty SkuPartNumber) | ForEach-Object { if ($skuMap.ContainsKey($_)) {"- $($skuMap[$_])"} else {"- $_"} }
            if (!$licenses) { $licenses = "- Sin licencias" }
            # Roles
            $userRoles = @(Get-MgDirectoryRole -All | ForEach-Object {
                if ((Get-MgDirectoryRoleMember -DirectoryRoleId $_.Id | Where-Object { $_.Id -eq $u.Id })) { "- $($_.DisplayName)" }
            })
            if (!$userRoles) { $userRoles = "- Sin roles asignados" }
            # Grupos (con cantidad)
            $miGruposEstaticos = $UserGroupsDict_Static[$u.Id]
            $miGruposDinamicos = $UserGroupsDict_Dynamic[$u.Id]
            $gruposEstaticosCount = if ($miGruposEstaticos) { $miGruposEstaticos.Count } else { 0 }
            $gruposDinamicosCount = if ($miGruposDinamicos) { $miGruposDinamicos.Count } else { 0 }
            $gruposEstaticos = if ($miGruposEstaticos) {
                $miGruposEstaticos | Sort-Object | ForEach-Object { "- $_" }
            } else {
                "- Sin grupos asignados"
            }
            $gruposDinamicos = if ($miGruposDinamicos) {
                $miGruposDinamicos | Sort-Object | ForEach-Object { "- $_" }
            } else {
                "- Sin grupos dinámicos"
            }
            # Fechas y avisos
            $lastSignIn = $u.SignInActivity?.LastSignInDateTime; $avisoSignIn = $false
            if ($lastSignIn) {
                $lastSignIn = (Get-Date $lastSignIn).ToString("dd/MM/yyyy HH:mm:ss")
            } else {
                $lastSignIn = "- Sin información"
                $avisoSignIn = $true
            }
            $labelAvisoSignIn.Visible = $avisoSignIn
            $labelAvisoDisabled.Visible = !$u.AccountEnabled
            $fechaCreacion = $u.CreatedDateTime ? (Get-Date $u.CreatedDateTime).ToString("dd/MM/yyyy HH:mm:ss") : "-"

            # Últimos intentos de inicio de sesión con país y ciudad nativos de Graph
            $signInLogs = Get-MgAuditLogSignIn -Filter "userId eq '$($u.Id)'" -Top 5
            $ultimosIniciosSesion = if ($signInLogs) {
                $signInLogs | ForEach-Object {
                    $dt = (Get-Date $_.CreatedDateTime).ToString("dd/MM/yyyy HH:mm:ss")
                    $ip = $_.IpAddress
                    $app = $_.AppDisplayName
                    $code = $_.Status.ErrorCode
                    $message = if ($signInErrorMessages.ContainsKey($code)) {
                        $signInErrorMessages[$code]
                    } else {
                        "Error desconocido"
                    }
                    $result = if ($code -eq 0) {"OK"} else {"ERROR ($code) - $message"}
                    $pais = $_.Location.CountryOrRegion
                    $ciudad = $_.Location.City
                    " - $dt | $ip | $pais | $ciudad | $app | $result"
                }
            } else { "- Sin registros" }

            $info = @"
$sepFull
$(Format-CenteredText "Información de usuario" $ancho)
$sepThin
$div
$(Format-CenteredText "Identidad" $ancho)
$div
┃ Nombre completo   : $($u.DisplayName)
┃ Usuario           : $($u.UserPrincipalName)
┃ Tipo              : $tipoUsuario
┃ Estado cuenta     : $estado
┃ Creación          : $fechaCreacion
┃ Último inicio     : $lastSignIn
┃ ID usuario        : $($u.Id)
$div
$(Format-CenteredText "Licencias" $ancho)
$div
$($licenses -join "`r`n")
$div
$(Format-CenteredText "Roles asignados" $ancho)
$div
$($userRoles -join "`r`n")
$div
$(Format-CenteredText "Grupos Estáticos ($gruposEstaticosCount)" $ancho)
$div
$($gruposEstaticos -join "`r`n")
$div
$(Format-CenteredText "Grupos Dinámicos ($gruposDinamicosCount)" $ancho)
$div
$($gruposDinamicos -join "`r`n")
$div
$(Format-CenteredText "Cargo y departamento" $ancho)
$div
┃ Cargo             : $($u.JobTitle)
┃ Departamento      : $($u.Department)
$div
$(Format-CenteredText "Contacto" $ancho)
$div
┃ Móvil             : $($u.MobilePhone)
┃ Ciudad            : $($u.City)
┃ País              : $($u.Country)
$div
$(Format-CenteredText "Últimos intentos de inicio de sesión" $ancho)
$div
$(($ultimosIniciosSesion -join "`r`n"))
$sepFullEnd
"@
            $textBox.Text = $info
        }
    }
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
