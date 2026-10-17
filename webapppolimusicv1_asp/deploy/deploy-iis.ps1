#Requires -RunAsAdministrator
<#
  PoliMusic - Despliegue en IIS (ASP clásico)

  Ejecutar en PowerShell COMO ADMINISTRADOR:
    powershell -ExecutionPolicy Bypass -File deploy\deploy-iis.ps1

  Qué hace:
    1. Verifica/instala la característica ASP de IIS.
    2. Copia el sitio a C:\inetpub\wwwroot\PoliMusic (sin borrar los medios ya subidos).
    3. Crea el grupo de aplicaciones "PoliMusicPool" (sin código administrado, 64 bits).
    4. Crea la aplicación /PoliMusic en "Default Web Site".
    5. Configura ASP: UTF-8, sesión de 30 min, subidas de hasta 50 MB, errores ocultos.
    6. Da permisos de escritura en las carpetas de medios (subida de audio/imágenes).
    7. Aplica database\01_seed_roles_usuarios.sql (userAdmin = ADMIN, paul.davila = USER).
    8. Comprueba que http://localhost/PoliMusic/login.asp responde.
#>
param(
  [string]$SiteName   = 'Default Web Site',
  [string]$AppName    = 'PoliMusic',
  [string]$TargetRoot = 'C:\inetpub\wwwroot',
  [string]$PoolName   = 'PoliMusicPool',
  [switch]$SkipDatabase
)

$ErrorActionPreference = 'Stop'
$repo    = Split-Path $PSScriptRoot -Parent
$source  = Join-Path $repo 'PoliMusic'
$target  = Join-Path $TargetRoot $AppName
$appcmd  = Join-Path $env:windir 'System32\inetsrv\appcmd.exe'
$media   = @('songFiles', 'albumImages', 'ArtistImages', 'userImages', 'playlistImages')

function Step($msg) { Write-Host "`n==> $msg" -ForegroundColor Magenta }
function AppCmd { & $appcmd @args; if ($LASTEXITCODE -ne 0) { throw "appcmd $($args -join ' ') fallo ($LASTEXITCODE)" } }

Step 'Verificando IIS y ASP clásico'
if (-not (Test-Path $appcmd)) { throw 'IIS no está instalado (no se encontró appcmd.exe).' }
if (-not (Test-Path (Join-Path $env:windir 'System32\inetsrv\asp.dll'))) {
  Write-Host 'Instalando la característica IIS-ASP...'
  Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP -All -NoRestart | Out-Null
}
$prov = Test-Path 'Registry::HKEY_CLASSES_ROOT\MSOLEDBSQL19'
if (-not $prov) { Write-Warning 'No se encontró el proveedor MSOLEDBSQL19 (Microsoft OLE DB Driver 19 for SQL Server). Instálelo o ajuste DB_CONN en includes\config.asp.' }

Step "Copiando archivos a $target"
New-Item -ItemType Directory -Force $target | Out-Null
# Código: espejo exacto excepto carpetas de medios (para no borrar archivos subidos en producción)
robocopy $source $target /MIR /XD ($media | ForEach-Object { Join-Path $source $_ }) /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy (código) fallo: $LASTEXITCODE" }
# Medios: solo agrega archivos nuevos
foreach ($m in $media) {
  New-Item -ItemType Directory -Force (Join-Path $target $m) | Out-Null
  if (Test-Path (Join-Path $source $m)) {
    robocopy (Join-Path $source $m) (Join-Path $target $m) /E /XO /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy ($m) fallo: $LASTEXITCODE" }
  }
}
$global:LASTEXITCODE = 0

Step "Grupo de aplicaciones $PoolName"
# appcmd escribe sus errores en stdout: se compara contra el listado completo
$poolExists = @(& $appcmd list apppool) -match ('^APPPOOL "' + [regex]::Escape($PoolName) + '"')
if (-not $poolExists) { AppCmd add apppool /name:$PoolName }
AppCmd set apppool $PoolName /managedRuntimeVersion:"" /enable32BitAppOnWin64:false /processModel.idleTimeout:00:00:00

Step "Aplicación /$AppName en '$SiteName'"
$appId = "$SiteName/$AppName"
$appExists = @(& $appcmd list app) -match ('^APP "' + [regex]::Escape($appId) + '"')
if (-not $appExists) {
  AppCmd add app /site.name:"$SiteName" /path:"/$AppName" /physicalPath:"$target" /applicationPool:$PoolName
} else {
  AppCmd set app "$appId" /applicationPool:$PoolName
  AppCmd set vdir "$appId/" /physicalPath:"$target"
}

Step 'Configuración de ASP clásico'
AppCmd set config "$appId" /section:system.webServer/asp /codePage:65001 /enableParentPaths:false /scriptErrorSentToBrowser:false /session.timeout:00:30:00 /limits.maxRequestEntityAllowed:52494336 /limits.bufferingLimit:67108864 /commit:apphost
AppCmd set config "$appId" /section:system.webServer/security/authentication/anonymousAuthentication /enabled:true /commit:apphost

Step 'Permisos de escritura en carpetas de medios'
foreach ($m in $media) {
  $dir = Join-Path $target $m
  icacls $dir /grant 'IUSR:(OI)(CI)M' "IIS AppPool\${PoolName}:(OI)(CI)M" /T /Q | Out-Null
}

if (-not $SkipDatabase) {
  Step 'Base de datos: roles y usuarios'
  & (Join-Path $repo 'tools\run-sql.ps1') -File (Join-Path $repo 'database\01_seed_roles_usuarios.sql')
}

Step 'Reiniciando el grupo de aplicaciones'
& $appcmd recycle apppool /apppool.name:$PoolName | Out-Null

Step 'Comprobación'
$url = "http://localhost/$AppName/login.asp"
try {
  $r = Invoke-WebRequest $url -UseBasicParsing -TimeoutSec 30
  if ($r.StatusCode -eq 200 -and $r.Content -match 'Inicia sesi') { Write-Host "OK  $url responde correctamente." -ForegroundColor Green }
  else { Write-Warning "Respuesta inesperada de $url (HTTP $($r.StatusCode))." }
} catch { Write-Warning "No se pudo abrir ${url}: $($_.Exception.Message)" }

Write-Host "`nDespliegue terminado. Abra http://localhost/$AppName/" -ForegroundColor Green
Write-Host '  Administrador:  userAdmin   / 12345678'
Write-Host '  Usuario normal: paul.davila / 12345678'
