<#  Fix-OfficeStart-0x3.ps1  (PS 5.1 / Server 2019)
    Arregla arranque de Office con error 0x3-0x0 (ruta/ClickToRun).
    Hace: detectar Office -> asegurar ClickToRunSvc (si C2R) -> cerrar procesos ->
          autorreparar C2R -> /regserver de Word/Excel/PPT -> recrear accesos en ProgramData.
#>

[CmdletBinding()]
param(
  [switch]$SkipShortcuts,
  [switch]$AttemptC2RRepair,            # intenta autorreparación C2R si la detección falla
  [string]$OfficeFolderName = 'Microsoft Office'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) { throw "Ejecuta este script como **Administrador**." }
}

function Start-TranscriptSafe {
  try {
    if (-not (Test-Path 'C:\Temp')) { New-Item -ItemType Directory -Path 'C:\Temp' | Out-Null }
    $Global:TranscriptPath = "C:\Temp\Fix-OfficeStart-0x3-{0}.log" -f (Get-Date -f 'yyyyMMdd-HHmmss')
    Start-Transcript -Path $Global:TranscriptPath -Force | Out-Null
  } catch {}
}

function Get-OfficeInstallInfo {
  # Devuelve objeto con Type='C2R'/'MSI', RootPath y Version
  $info = [pscustomobject]@{ Type=$null; RootPath=$null; Version=$null }

  try {
    # 1) App Paths -> WINWORD.EXE
    $ap = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE' -ErrorAction SilentlyContinue
    if ($ap) {
      $exe = $ap.'(Default)'
      if (-not $exe -and $ap.Path) { $exe = Join-Path $ap.Path 'WINWORD.EXE' }
      if ($exe -and (Test-Path $exe)) {
        $root = Split-Path -Parent $exe
        $ver  = (Get-Item $exe).VersionInfo.ProductVersion
        $type = 'MSI'
        if ($root -match '\\root\\Office16') { $type = 'C2R' }
        return [pscustomobject]@{ Type=$type; RootPath=$root; Version=$ver }
      }
    }

    # 2) Click-to-Run (algunas builds no tienen InstallPath)
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c2r) {
      $pf86 = ${env:ProgramFiles(x86)}
      $roots = @()
      if ($c2r.InstallPath)      { $roots += (Join-Path $c2r.InstallPath 'Office16') }
      if ($c2r.InstallationPath) { $roots += (Join-Path $c2r.InstallationPath 'Office16') }
      $roots += (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16')
      if ($pf86) { $roots += (Join-Path $pf86 'Microsoft Office\root\Office16') }
      $roots = $roots | Select-Object -Unique

      foreach ($r in $roots) {
        $exe = Join-Path $r 'WINWORD.EXE'
        if (Test-Path $exe) {
          $ver = (Get-Item $exe).VersionInfo.ProductVersion
          $verReport = $null
          try { $verReport = $c2r.ClientVersionToReport } catch {}
          if (-not $verReport) { $verReport = $ver }
          return [pscustomobject]@{ Type='C2R'; RootPath=$r; Version=$verReport }
        }
      }
    }

    # 3) MSI InstallRoot
    $msi = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot' -ErrorAction SilentlyContinue
    if ($msi -and $msi.Path) {
      $exe = Join-Path $msi.Path 'WINWORD.EXE'
      if (Test-Path $exe) {
        $root = $msi.Path.TrimEnd('\')
        $ver  = (Get-Item $exe).VersionInfo.ProductVersion
        return [pscustomobject]@{ Type='MSI'; RootPath=$root; Version=$ver }
      }
    }

    # 4) Rutas conocidas MSI
    $pf86 = ${env:ProgramFiles(x86)}
    $fallback = @((Join-Path $env:ProgramFiles 'Microsoft Office\Office16'))
    if ($pf86) { $fallback += (Join-Path $pf86 'Microsoft Office\Office16') }
    foreach ($r in $fallback) {
      $exe = Join-Path $r 'WINWORD.EXE'
      if (Test-Path $exe) {
        $ver = (Get-Item $exe).VersionInfo.ProductVersion
        return [pscustomobject]@{ Type='MSI'; RootPath=$r; Version=$ver }
      }
    }
  } catch {}
  return $info
}

function Ensure-ClickToRunSvc {
  $svc = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
  if (-not $svc) { return $false }
  try { Set-Service ClickToRunSvc -StartupType Automatic } catch {}
  if ($svc.Status -ne 'Running') {
    try { Start-Service ClickToRunSvc } catch {}
  }
  $svc = Get-Service ClickToRunSvc -ErrorAction SilentlyContinue
  if ($svc -and $svc.Status -eq 'Running') { return $true } else { return $false }
}

function New-Shortcut {
  param([string]$LinkPath,[string]$TargetPath,[string]$IconPath=$null,[string]$WorkingDir=$null)
  $ws = New-Object -ComObject WScript.Shell
  $lnk = $ws.CreateShortcut($LinkPath)
  $lnk.TargetPath = $TargetPath
  if ($IconPath)   { $lnk.IconLocation = $IconPath }
  if ($WorkingDir) { $lnk.WorkingDirectory = $WorkingDir }
  $lnk.Save()
}

# ================= MAIN =================
try {
  Assert-Admin
  Start-TranscriptSafe

  Write-Host "==> 1) Detectando Office..." -ForegroundColor Cyan
  $off = Get-OfficeInstallInfo
  if (-not $off.Type -and $AttemptC2RRepair) {
    Write-Host "    No detectado. Intentando autorreparación C2R..." -ForegroundColor Yellow
    $c2rClient = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    if (Test-Path $c2rClient) {
      Start-Process $c2rClient -ArgumentList "/update user displaylevel=false forceappshutdown=true" -Wait -WindowStyle Hidden
      $off = Get-OfficeInstallInfo
    }
  }
  if (-not $off.Type) { throw "No se localizan binarios de Office. Si debe estar instalado, repara o reinstala." }
  Write-Host ("    Detectado: {0} {1} | {2}" -f $off.Type,$off.Version,$off.RootPath)

  if ($off.Type -eq 'C2R') {
    Write-Host "==> 2) Asegurando servicio ClickToRunSvc..." -ForegroundColor Cyan
    if (-not (Ensure-ClickToRunSvc)) { Write-Warning "ClickToRunSvc no está disponible o no pudo iniciarse." }
  }

  Write-Host "==> 3) Cerrando procesos Office y cliente C2R..." -ForegroundColor Cyan
  'winword','excel','powerpnt','outlook','onenote','officec2rclient','integratedoffice' |
    ForEach-Object { Get-Process $_ -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }

  if ($off.Type -eq 'C2R') {
    Write-Host "==> 4) Forzando autorreparación/actualización de Click-to-Run..." -ForegroundColor Cyan
    $c2rClient = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    if (Test-Path $c2rClient) {
      Start-Process $c2rClient -ArgumentList "/update user displaylevel=false forceappshutdown=true" -Wait -WindowStyle Hidden
    } else {
      Write-Warning "OfficeC2RClient.exe no encontrado; si es MSI, omite este aviso."
    }
  }

  Write-Host "==> 5) Re-registrando Word/Excel/PowerPoint (/regserver)..." -ForegroundColor Cyan
  foreach ($exe in 'WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE') {
    $p = Join-Path $off.RootPath $exe
    if (Test-Path $p) {
      try { Start-Process $p -ArgumentList '/regserver' -WindowStyle Hidden -Wait }
      catch { Write-Warning ("No se pudo re-registrar {0}: {1}" -f $exe, $_.Exception.Message) }
    }
  }

  if (-not $SkipShortcuts) {
    Write-Host "==> 6) Recreando accesos en Menú Inicio (ProgramData)..." -ForegroundColor Cyan
    $progData = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'
    $folder   = Join-Path $progData $OfficeFolderName
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    foreach ($exe in 'WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE','OUTLOOK.EXE','ONENOTE.EXE') {
      $p = Join-Path $off.RootPath $exe
      if (Test-Path $p) {
        switch ($exe) {
          'WINWORD.EXE'  { $appName = 'Microsoft Word' }
          'EXCEL.EXE'    { $appName = 'Microsoft Excel' }
          'POWERPNT.EXE' { $appName = 'Microsoft PowerPoint' }
          'OUTLOOK.EXE'  { $appName = 'Microsoft Outlook' }
          'ONENOTE.EXE'  { $appName = 'Microsoft OneNote' }
          default        { $appName = "Microsoft " + ($exe -replace '\.EXE$','') }
        }
        $lnk = Join-Path $folder ($appName + '.lnk')
        New-Shortcut -LinkPath $lnk -TargetPath $p -IconPath $p -WorkingDir $off.RootPath
      }
    }
  }

  Write-Host "==> 7) Validación" -ForegroundColor Cyan
  if ($off.Type -eq 'C2R') {
    $svc = Get-Service ClickToRunSvc -ErrorAction SilentlyContinue
    $svcStatus = if ($svc) { $svc.Status } else { 'No encontrado' }
    Write-Host ("    ClickToRunSvc: {0}" -f $svcStatus)
  }
  Write-Host "    Abre ahora Word/Excel. Si aún aparece 0x3-0x0, lanza una **Reparación rápida** desde Programas y características."

  Write-Host "`n=== COMPLETADO ==="
  if ($Global:TranscriptPath) { Write-Host ("Log: {0}" -f $Global:TranscriptPath) }
}
catch {
  Write-Error $_.Exception.Message
}
finally {
  try { Stop-Transcript | Out-Null } catch {}
}
