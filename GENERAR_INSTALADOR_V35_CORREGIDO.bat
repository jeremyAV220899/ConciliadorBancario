@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"
set "ROOT=%CD%"
set "PROJECT="
if exist "%ROOT%\ConciliadorBancario.csproj" set "PROJECT=%ROOT%\ConciliadorBancario.csproj"
if not defined PROJECT if exist "%ROOT%\ConciliadorBancario\ConciliadorBancario.csproj" set "PROJECT=%ROOT%\ConciliadorBancario\ConciliadorBancario.csproj"
set "OUT=%ROOT%\Publicacion_Instalador"
set "SETUPDIR=%ROOT%\Instalador"
set "ISS=%SETUPDIR%\ConciliadorBancario_V35_auto.iss"
set "ISCC="

echo ==========================================================
echo   CONCILIADOR BANCARIO V35 - GENERADOR DE INSTALADOR
echo ==========================================================
echo.
if not defined PROJECT (
  echo ERROR: No se encuentra ConciliadorBancario.csproj en esta carpeta ni en ConciliadorBancario\.
  goto :error
)
where dotnet.exe >nul 2>&1
if errorlevel 1 (
  echo ERROR: Falta el SDK de .NET o no se encuentra en PATH.
  echo Instala el SDK compatible con tu .csproj, y vuelve a ejecutar.
  goto :error
)

call :find_inno
if not defined ISCC (
  echo No se encontro Inno Setup 6 o 7.
  where winget.exe >nul 2>&1
  if not errorlevel 1 (
    echo.
    choice /C SN /N /M "Deseas instalar Inno Setup con winget ahora? [S/N]: "
    if errorlevel 2 goto :manual_inno
    echo.
    echo Instalando Inno Setup desde el repositorio de winget...
    winget install --id JRSoftware.InnoSetup --exact --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
      echo No se pudo instalar con winget.
      goto :manual_inno
    )
    call :find_inno
  ) else (
    goto :manual_inno
  )
)
if not defined ISCC goto :manual_inno

:ready
echo.
echo Se usara: "%ISCC%"
echo.
echo [1/3] Publicando el Conciliador para Windows x64...
dotnet publish "%PROJECT%" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -p:PublishTrimmed=false -p:PublishReadyToRun=false -o "%OUT%"
if errorlevel 1 (
  echo ERROR: Fallo la publicacion. Comprueba los errores de .NET mostrados arriba.
  goto :error
)
if not exist "%OUT%\ConciliadorBancario.exe" (
  echo ERROR: No se encuentra ConciliadorBancario.exe en "%OUT%".
  echo Comprueba AssemblyName y OutputType en el .csproj.
  goto :error
)

if not exist "%SETUPDIR%" mkdir "%SETUPDIR%"
if not exist "%SETUPDIR%" (
  echo ERROR: No se pudo crear la carpeta Instalador.
  goto :error
)

echo [2/3] Preparando el archivo de instalacion...
> "%ISS%" echo #define MyAppName "Conciliador Bancario"
>> "%ISS%" echo #define MyAppVersion "35.0.0"
>> "%ISS%" echo #define MyAppExeName "ConciliadorBancario.exe"
>> "%ISS%" echo.
>> "%ISS%" echo [Setup]
>> "%ISS%" echo AppId={{7D72979E-A1CB-4D2A-ABA3-A4490544A4FD}
>> "%ISS%" echo AppName={#MyAppName}
>> "%ISS%" echo AppVersion={#MyAppVersion}
>> "%ISS%" echo DefaultDirName={autopf}\Conciliador Bancario
>> "%ISS%" echo DefaultGroupName={#MyAppName}
>> "%ISS%" echo OutputDir=Salida
>> "%ISS%" echo OutputBaseFilename=Setup_Conciliador_Bancario_V35
>> "%ISS%" echo Compression=lzma2
>> "%ISS%" echo SolidCompression=yes
>> "%ISS%" echo WizardStyle=modern
>> "%ISS%" echo ArchitecturesAllowed=x64compatible
>> "%ISS%" echo ArchitecturesInstallIn64BitMode=x64compatible
>> "%ISS%" echo PrivilegesRequired=admin
>> "%ISS%" echo UninstallDisplayName={#MyAppName} V35
>> "%ISS%" echo.
>> "%ISS%" echo [Languages]
>> "%ISS%" echo Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
>> "%ISS%" echo.
>> "%ISS%" echo [Tasks]
>> "%ISS%" echo Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce
>> "%ISS%" echo.
>> "%ISS%" echo [Files]
>> "%ISS%" echo Source: "..\Publicacion_Instalador\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
>> "%ISS%" echo.
>> "%ISS%" echo [Icons]
>> "%ISS%" echo Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
>> "%ISS%" echo Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
>> "%ISS%" echo.
>> "%ISS%" echo [Run]
>> "%ISS%" echo Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

 echo [3/3] Generando Setup.exe...
"%ISCC%" "%ISS%"
if errorlevel 1 (
  echo ERROR: Inno Setup no pudo compilar el instalador. Revisa el mensaje superior.
  goto :error
)
if not exist "%SETUPDIR%\Salida\Setup_Conciliador_Bancario_V35.exe" (
  echo ERROR: No se encuentra Setup_Conciliador_Bancario_V35.exe.
  goto :error
)
echo.
echo ==========================================================
echo   INSTALADOR GENERADO CORRECTAMENTE
echo   "%SETUPDIR%\Salida\Setup_Conciliador_Bancario_V35.exe"
echo ==========================================================
explorer "%SETUPDIR%\Salida"
pause
exit /b 0

:find_inno
for %%V in (6 7) do (
  if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup %%V\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup %%V\ISCC.exe"
  if not defined ISCC if exist "%ProgramFiles%\Inno Setup %%V\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup %%V\ISCC.exe"
  if defined LOCALAPPDATA if not defined ISCC if exist "%LOCALAPPDATA%\Programs\Inno Setup %%V\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup %%V\ISCC.exe"
)
if not defined ISCC for %%I in (ISCC.exe) do if not "%%~$PATH:I"=="" set "ISCC=%%~$PATH:I"
exit /b

:manual_inno
echo.
echo Inno Setup no esta disponible en las rutas habituales.
echo Descargalo desde: https://jrsoftware.org/isdl.php
echo O indica aqui la ruta completa a ISCC.exe si ya lo instalaste.
set "CUSTOM="
set /p "CUSTOM=Ruta a ISCC.exe [ENTER para salir]: "
if not defined CUSTOM goto :error

if not exist "%CUSTOM%" (
  echo ERROR: El archivo indicado no existe.
  goto :error
)
for %%F in ("%CUSTOM%") do if /I not "%%~nxF"=="ISCC.exe" (
  echo ERROR: Debes seleccionar ISCC.exe.
  goto :error
)
set "ISCC=%CUSTOM%"
goto :ready

:error
echo.
echo No se genero el instalador. Corrige el error y vuelve a ejecutar.
pause
exit /b 1
