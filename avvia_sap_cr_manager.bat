@echo off
setlocal

pushd "%~dp0SAPCRManager"

set SAP_CR_MANAGER_PORT=5055
set SAP_CR_MANAGER_DEBUG=0
set APP_URL=http://127.0.0.1:%SAP_CR_MANAGER_PORT%
set HEALTH_URL=%APP_URL%/health

where py >nul 2>nul
if errorlevel 1 (
	echo Il comando py non e disponibile. Installa Python dal sito ufficiale e riprova.
	pause
	popd
	endlocal
	exit /b 1
)

py -c "import flask" >nul 2>nul
if errorlevel 1 (
	echo Flask non trovato. Installo le dipendenze richieste...
	py -m pip install -r requirements.txt
	if errorlevel 1 (
		echo Installazione dipendenze non riuscita.
		pause
		popd
		endlocal
		exit /b 1
	)
)

start "SAP CR Manager" cmd /k set SAP_CR_MANAGER_PORT=%SAP_CR_MANAGER_PORT%^& set SAP_CR_MANAGER_DEBUG=%SAP_CR_MANAGER_DEBUG%^& py app.py

for /l %%I in (1,1,30) do (
	py -c "import sys, urllib.request; urllib.request.urlopen(r'%HEALTH_URL%', timeout=1); sys.exit(0)" >nul 2>nul
	if not errorlevel 1 goto open_browser
	timeout /t 1 /nobreak >nul
)

echo L'app non ha risposto in tempo. Verifica la finestra del server.
pause
goto cleanup

:open_browser
start "" "%APP_URL%"

:cleanup
popd
endlocal