# README — Scripts de reparación de Microsoft Office (Windows Server 2019)

Este repositorio incluye **dos utilidades PowerShell** para corregir problemas típicos de Microsoft Office en **Windows Server 2019** (RDS/VDI/servidor estándar) con **PowerShell 5.1**:

* `Fix-OfficeInstall.ps1` – Restaura accesos del Menú Inicio y asociaciones cuando Office “desaparece” visualmente pero **los binarios siguen instalados**.
* `Fix-OfficeStart-0x3.ps1` – Corrige el error de inicio **0x3-0x0** (Click-to-Run) y problemas de arranque relacionados con el servicio **ClickToRunSvc**.

Ambos scripts generan log (**transcript**) en `C:\Temp\` y son **idempotentes** (puedes ejecutarlos varias veces sin efectos adversos).

---

## Compatibilidad y requisitos

* **SO:** Windows Server 2019 (build 17763.x).
* **PowerShell:** 5.1 (por defecto en WS2019).
* **Permisos:** ejecutar **como Administrador**.
* **Espacio para logs:** `C:\Temp\`.
* **Conectividad (solo Click-to-Run):** para autorreparación/actualización puede requerirse acceso al **CDN de Office** (o a vuestro repositorio interno ODT).

---

## Uso rápido

Abrir **PowerShell (Administrador)** y ejecutar:

```powershell
# 1) “Reaparecer” Office (accesos + asociaciones/COM)
PowerShell -ExecutionPolicy Bypass -File "C:\Ruta\Fix-OfficeInstall.ps1"

# 2) Resolver error 0x3-0x0 (Click-to-Run) y arranque
PowerShell -ExecutionPolicy Bypass -File "C:\Ruta\Fix-OfficeStart-0x3.ps1"
```

> Logs:
> `C:\Temp\Fix-OfficeInstall-YYYYMMDD-HHMMSS.log`
> `C:\Temp\Fix-OfficeStart-0x3-YYYYMMDD-HHMMSS.log`

---

## Detalles de cada script

### `Fix-OfficeInstall.ps1`

**Propósito**

Restaurar la “visibilidad” de Office cuando, tras reparaciones del sistema o cambios de imagen, **faltan accesos del Menú Inicio** o se rompen **asociaciones/COM**, aunque los ejecutables siguen presentes.

**Qué hace (resumen)**

1. **Detección robusta** de Office (C2R/MSI):

   * Registro `App Paths\WINWORD.EXE`.
   * Claves de **ClickToRun** (`HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration`).
   * **InstallRoot** MSI (`HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot`).
   * Rutas conocidas: `...\root\Office16` (C2R) y `...\Office16` (MSI).

2. **Recrea accesos directos** en *Todos los usuarios*:

   * `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\`

3. **Re-registra** Word/Excel/PowerPoint:

   * `WINWORD.EXE /regserver`, `EXCEL.EXE /regserver`, `POWERPNT.EXE /regserver`.

4. **(Opcional)** intenta **reparación Click-to-Run** silenciosa:

   * `OfficeC2RClient.exe /update user displaylevel=false forceappshutdown=true`.

**Parámetros**

* `-CreateShortcuts` (predeterminado: **true**)
* `-RepairAssociations` (predeterminado: **true**)
* `-AttemptC2RRepair` (predeterminado: **false**)
* `-OfficeFolderName 'Microsoft Office'`

**Ejemplos**

```powershell
# Estándar: accesos + /regserver
PowerShell -ExecutionPolicy Bypass -File ".\Fix-OfficeInstall.ps1"

# Con intento de reparación C2R
PowerShell -ExecutionPolicy Bypass -File ".\Fix-OfficeInstall.ps1" -AttemptC2RRepair

# Solo accesos, sin re-registro
PowerShell -ExecutionPolicy Bypass -File ".\Fix-OfficeInstall.ps1" -RepairAssociations:$false
```

---

### `Fix-OfficeStart-0x3.ps1`

**Propósito**

Corregir el error de inicio **0x3-0x0** y fallos de arranque de Office asociados a **Click-to-Run** (servicio parado, instalación incoherente, rutas desalineadas).

**Qué hace (resumen)**

1. **Detecta** Office (C2R/MSI) igual que el script anterior.
2. **Asegura** el servicio **ClickToRunSvc** (Automático + Running) cuando es C2R.
3. **Cierra** procesos de Office/C2R para evitar bloqueos.
4. **Fuerza autorreparación/actualización** del cliente C2R (`OfficeC2RClient.exe`).
5. **Re-registra** Word/Excel/PPT (`/regserver`).
6. **Recrea accesos** en ProgramData (opcional).

**Parámetros**

* `-SkipShortcuts` (no crear accesos)
* `-AttemptC2RRepair` (si la detección inicial falla, intenta reparar C2R y reintenta)
* `-OfficeFolderName 'Microsoft Office'`

**Ejemplos**

```powershell
# Corrección completa del 0x3-0x0
PowerShell -ExecutionPolicy Bypass -File ".\Fix-OfficeStart-0x3.ps1"

# Si no detecta Office al inicio, probar autorreparación C2R primero
PowerShell -ExecutionPolicy Bypass -File ".\Fix-OfficeStart-0x3.ps1" -AttemptC2RRepair

# Solo reparar arranque sin tocar accesos
PowerShell -ExecutionPolicy Bypass -File ".\Fix-OfficeStart-0x3.ps1" -SkipShortcuts
```

---

## Verificación posterior

* **Accesos presentes**
  `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\*.lnk`

* **Lanzamiento**
  Abrir **Word**, **Excel** y **PowerPoint** desde el Menú Inicio (sin errores).

* **Servicio Click-to-Run (si aplica)**

  ```powershell
  Get-Service ClickToRunSvc
  ```

  Resultado esperado: **Running** y **Automatic**.

* **Asociaciones**
  Abrir `.docx`, `.xlsx` y `.pptx` desde el Explorador.

---

## Resolución de problemas

* **Sigue el error 0x3-0x0 al abrir Office**

  1. Ejecuta `Fix-OfficeStart-0x3.ps1` con `-AttemptC2RRepair`.
  2. Si persiste, ejecuta una **Reparación rápida** desde *Programas y características* (C2R) o reinstala con vuestro paquete **ODT/ConfigMgr/Intune**.

* **Los scripts no detectan Office**
  Los ejecutables no están presentes → **reinstalar/reparar** según vuestro estándar.
  Verifica conectividad al **CDN de Office** si usas C2R.

* **No aparecen accesos**
  Revisa permisos en `C:\ProgramData\Microsoft\Windows\Start Menu\Programs` y que no haya directivas/perfil redirigido que sobrescriba el Menú Inicio.

* **Dónde ver los logs**

  * `C:\Temp\Fix-OfficeInstall-*.log`
  * `C:\Temp\Fix-OfficeStart-0x3-*.log`

---

## Notas operativas

* Los scripts **no desinstalan** Office ni tocan datos de usuario.
* `Fix-OfficeStart-0x3.ps1` **cierra** procesos de Office: guarda trabajos abiertos antes.
* Son **seguros de re-ejecutar**; recrean accesos y re-registran sin duplicar entradas.

---

## Mantenimiento

* Probar primero en **máquinas de prueba** o sesiones no críticas.
* Actualizar el repositorio si adoptáis **nuevas versiones de Office** (cambios de ruta/versión).
* Si trabajáis con **WSUS/ODT offline**, documentad la fuente de reparación de C2R para evitar bloqueos de red.
