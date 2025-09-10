<# 
  Script Name : Fix-OfficeInstall.ps1
  Purpose     : Restaurar accesos directos de Office y re-registrar apps; opcionalmente reparar C2R
  Compat      : PowerShell 5.1 (Windows Server 2019)
#>

[CmdletBinding()]
param(
  [switch]$CreateShortcuts     = $true,
  [switch]$RepairAssociations  = $true,
  [switch]$AttemptC2RRepair    = $false,
  [string]$OfficeFolderName    = 'Microsoft Office'   # subcarpeta del Menú Inicio
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) { throw "Este script debe ejecutarse como **Administrador**." }
}

function Start-MyTranscript {
  try {
    if (-not (Test-Path 'C:\Temp')) { New-Item -ItemType Directory -Path 'C:\Temp' | Out-Null }
    $Global:TranscriptPath = "C:\Temp\Fix-OfficeInstall-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss")
    Start-Transcript -Path $Global:TranscriptPath -Force | Out-Null
  } catch { Write-Warning ("No se pudo iniciar transcript: {0}" -f $_.Exception.Message) }
}

function Get-OfficeInstallInfo {
  # Devuelve objeto con Type='C2R'/'MSI', RootPath (carpeta con los EXE) y Version
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

    # 2) Click-to-Run (varía entre builds; InstallPath a veces falta)
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c2r) {
      $pf86 = ${env:ProgramFiles(x86)}
      $candidateRoots = @()
      if ($c2r.InstallPath)      { $candidateRoots += (Join-Path $c2r.InstallPath 'Office16') }
      if ($c2r.InstallationPath) { $candidateRoots += (Join-Path $c2r.InstallationPath 'Office16') }
      $candidateRoots += (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16')
      if ($pf86) { $candidateRoots += (Join-Path $pf86 'Microsoft Office\root\Office16') }
      $candidateRoots = $candidateRoots | Select-Object -Unique

      foreach ($root in $candidateRoots) {
        $exe = Join-Path $root 'WINWORD.EXE'
        if (Test-Path $exe) {
          $ver = (Get-Item $exe).VersionInfo.ProductVersion
          $verReport = $null
          try { $verReport = $c2r.ClientVersionToReport } catch {}
          if (-not $verReport) { $verReport = $ver }
          return [pscustomobject]@{ Type='C2R'; RootPath=$root; Version=$verReport }
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

    # 4) Rutas conocidas MSI (por si el registro está "mínimo")
    $pf86 = ${env:ProgramFiles(x86)}
    $fallbackRoots = @((Join-Path $env:ProgramFiles 'Microsoft Office\Office16'))
    if ($pf86) { $fallbackRoots += (Join-Path $pf86 'Microsoft Office\Office16') }

    foreach ($root in $fallbackRoots) {
      $exe = Join-Path $root 'WINWORD.EXE'
      if (Test-Path $exe) {
        $ver = (Get-Item $exe).VersionInfo.ProductVersion
        return [pscustomobject]@{ Type='MSI'; RootPath=$root; Version=$ver }
      }
    }
  } catch {
    Write-Warning "Detección de Office lanzó excepción: $($_.Exception.Message)"
  }

  return $info
}

function New-Shortcut {
  param(
    [Parameter(Mandatory=$true)][string]$LinkPath,
    [Parameter(Mandatory=$true)][string]$TargetPath,
    [string]$Arguments = '',
    [string]$IconPath = $null,
    [string]$WorkingDir = $null
  )
  $ws = New-Object -ComObject WScript.Shell
  $lnk = $ws.CreateShortcut($LinkPath)
  $lnk.TargetPath = $TargetPath
  if ($Arguments)  { $lnk.Arguments = $Arguments }
  if ($IconPath)   { $lnk.IconLocation = $IconPath }
  if ($WorkingDir) { $lnk.WorkingDirectory = $WorkingDir }
  $lnk.Save()
}

function Restore-OfficeShortcutsAndRegister {
  param([Parameter(Mandatory=$true)][pscustomobject]$OfficeInfo)
  if (-not $OfficeInfo.Type -or -not $OfficeInfo.RootPath) {
    Write-Warning "Office no detectado (Type/RootPath nulos)."
    return $false
  }

  Write-Host ("[+] Office detectado: {0} {1} en {2}" -f $OfficeInfo.Type, $OfficeInfo.Version, $OfficeInfo.RootPath) -ForegroundColor Green

  # 1) Accesos directos en ProgramData
  if ($CreateShortcuts) {
    $progData = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'
    $officeFolder = Join-Path $progData $OfficeFolderName
    if (-not (Test-Path $officeFolder)) { New-Item -ItemType Directory -Path $officeFolder -Force | Out-Null }

    $apps = @(
      @{ Name='Microsoft Word';       Exe='WINWORD.EXE'  },
      @{ Name='Microsoft Excel';      Exe='EXCEL.EXE'    },
      @{ Name='Microsoft PowerPoint'; Exe='POWERPNT.EXE' },
      @{ Name='Microsoft Outlook';    Exe='OUTLOOK.EXE'  },
      @{ Name='Microsoft OneNote';    Exe='ONENOTE.EXE'  }  # si existe
    )

    foreach ($a in $apps) {
      $exePath = Join-Path $OfficeInfo.RootPath $a.Exe
      if (Test-Path $exePath) {
        $lnk = Join-Path $officeFolder ($a.Name + '.lnk')
        New-Shortcut -LinkPath $lnk -TargetPath $exePath -IconPath $exePath -WorkingDir $OfficeInfo.RootPath
        Write-Host ("    Acceso directo creado: {0}" -f $lnk)
      }
    }
  }

  # 2) Re-registrar asociaciones/COM
  if ($RepairAssociations) {
    $regApps = @('WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE')
    foreach ($exe in $regApps) {
      $p = Join-Path $OfficeInfo.RootPath $exe
      if (Test-Path $p) {
        try { Start-Process $p -ArgumentList '/regserver' -WindowStyle Hidden -Wait; Write-Host ("    Re-registrado: {0}" -f $exe) }
        catch { Write-Warning ("    No se pudo re-registrar {0}: {1}" -f $exe, $_.Exception.Message) }
      }
    }
  }

  return $true
}

function Repair-OfficeClickToRun {
  $client = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
  if (-not (Test-Path $client)) { return $false }
  Write-Host "Intentando reparación Click-to-Run (silenciosa)..." -ForegroundColor Yellow
  try {
    # Fuerza comprobación/autorrepair; cierra apps si es necesario
    Start-Process $client -ArgumentList "/update user displaylevel=false forceappshutdown=true" -Wait -WindowStyle Hidden
    return $true
  } catch {
    Write-Warning "Reparación C2R falló: $($_.Exception.Message)"
    return $false
  }
}

# ========================= MAIN =========================

try {
  Assert-Admin
  Start-MyTranscript

  Write-Host "==> Detección de Office..." -ForegroundColor Cyan
  $info = Get-OfficeInstallInfo
  if (-not $info.Type) {
    Write-Warning "No se detectó Office por rutas estándar/registro."
    if ($AttemptC2RRepair) {
      if (Repair-OfficeClickToRun) {
        Write-Host "Reparación C2R ejecutada. Reintentando detección..." -ForegroundColor Yellow
        $info = Get-OfficeInstallInfo
      }
    }
  }

  if (-not $info.Type) {
    throw "No se ha podido localizar una instalación válida de Office (binarios ausentes)."
  }

  Write-Host "==> Restaurando accesos directos y registro..." -ForegroundColor Cyan
  $ok = Restore-OfficeShortcutsAndRegister -OfficeInfo $info
  if (-not $ok) {
    throw "No fue posible restaurar Office (detección sin RootPath)."
  }

  # Validación rápida
  $menuPath = Join-Path $env:ProgramData ("Microsoft\Windows\Start Menu\Programs\{0}" -f $OfficeFolderName)
  $links = @(Get-ChildItem $menuPath -Filter *.lnk -ErrorAction SilentlyContinue)
  Write-Host ("==> Validación: {0} accesos directos en {1}" -f $links.Count, $menuPath) -ForegroundColor Cyan

  Write-Host "`n=== COMPLETADO ===" -ForegroundColor Green
  if ($Global:TranscriptPath) { Write-Host ("Log: {0}" -f $Global:TranscriptPath) }
}
catch {
  Write-Error $_.Exception.Message
}
finally {
  try { Stop-Transcript | Out-Null } catch {}
}
