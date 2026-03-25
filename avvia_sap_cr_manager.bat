@echo off
setlocal

pushd "%~dp0SAPCRManager"

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

start "SAP CR Manager" cmd /k py app.py
timeout /t 2 /nobreak >nul
start "" http://127.0.0.1:5000

popd
endlocal